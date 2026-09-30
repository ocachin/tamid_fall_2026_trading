"""Fair value and arbitrage signals. Same logic for every ETF; only the underlying differs.

  fair value  = anchor x (1 + underlying move since yesterday's close)
  anchor      = CLOSE (default): the ETF's own T-1 close, so the screen shows today's move in the
                premium. NAV: the T-1 NAV, i.e. the true creation/redemption value (needs a NAV
                you trust; Yahoo's lags sometimes, so it falls back to CLOSE when they disagree >2%).
  underlying  = BASKET: sum(w_i r_i) over live holdings, uncovered weight x proxy return
                PROXY:  beta x hedge return (beta = OLS of daily ETF returns on hedge returns)
  sell edge   = bid vs fair value - cost    -> sell ETF / buy underlying  (AP creates)
  buy edge    = fair value vs ask - cost    -> buy ETF / sell underlying  (AP redeems)

Equities default to BASKET: holdings trade live, so the basket is the fair value.
Fixed income defaults to PROXY: bonds don't trade live, so a liquid rates/credit ETF moves the
anchor forward. Bond ETFs sit at a persistent premium/discount to NAV (NAV is marked at the bid),
which is why CLOSE is the default anchor: it nets out yesterday's premium.
"""
from __future__ import annotations

import datetime as dt
import math
from dataclasses import dataclass, field
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd


@dataclass
class Settings:
    refresh_seconds: float = 20
    equity_cost_bps: float = 3.0
    fi_cost_bps: float = 8.0
    equity_min_edge_bps: float = 2.0
    fi_min_edge_bps: float = 5.0
    anchor: str = "CLOSE"
    max_basket_names: int = 600
    history_period: str = "1y"
    beta_lookback: int = 120
    z_lookback: int = 60
    stale_minutes: float = 20
    exdiv_flag_bps: float = 30
    chart_seconds: float = 120
    vol_lookback: int = 60
    holdings_shown: int = 3
    live: bool = True
    log_live: bool = True


@dataclass
class Model:
    etf: str
    asset_class: str                    # EQUITY / FI
    kind: str                           # BASKET / PROXY
    hedge: str | None                   # PROXY: hedge; BASKET: proxy for the uncovered weight
    peers: list[str]
    leverage: float | None = None
    weights: pd.Series = field(default_factory=pd.Series)
    beta: float = 1.0
    coverage: float = np.nan            # basket weight with price data
    resid_mean: float = 0.0             # daily ETF - beta x underlying, in return units
    resid_std: float = np.nan
    und_hist: pd.Series | None = None   # daily underlying returns (for charts)

    @property
    def is_fi(self) -> bool:
        return self.asset_class == "FI"


def _f(x) -> float:
    try:
        return float(x) if x is not None else np.nan
    except (TypeError, ValueError):
        return np.nan


def basket_returns(close: pd.DataFrame, w: pd.Series, proxy: str | None) -> pd.Series:
    """Daily basket returns holding today's shares constant (weights drift back with prices),
    so past days are not weighted by today's winners."""
    w = w[w.index.isin(close.columns)]
    px = close[w.index].ffill()
    drift = w * px.shift(1).div(px.iloc[-1])            # weight at the start of each day
    drift = drift.div(drift.sum(axis=1), axis=0) * w.sum()
    r = (px.pct_change() * drift).sum(axis=1, min_count=1)
    if proxy and proxy in close:
        return r + (1.0 - w.sum()) * close[proxy].pct_change()
    return r / w.sum() if w.sum() > 0 else r


def build_model(row: dict, holdings: pd.DataFrame, close: pd.DataFrame, s: Settings) -> Model:
    etf, cls = row["etf"], row["class"]
    kind = row["model"]
    w = holdings.loc[holdings["etf"] == etf].set_index("symbol")["weight"].astype(float)
    w = w.groupby(level=0).sum().sort_values(ascending=False).head(s.max_basket_names)
    if kind == "AUTO":
        kind = "BASKET" if cls == "EQUITY" and len(w) else "PROXY"
    lev = row.get("leverage")
    if lev:
        kind = "PROXY"                      # leveraged / inverse: fixed multiple of the underlying
    m = Model(etf, cls, kind, row["hedge"] or None, row["peers"], leverage=lev)
    if etf not in close:
        return m
    close = close[close.index.date < dt.date.today()]   # completed days only
    rets = close.pct_change()
    y = rets[etf]
    if kind == "BASKET":
        m.weights = w[w.index.isin(rets.columns)]
        m.coverage = m.weights.sum()
        u = basket_returns(close, m.weights, m.hedge) if len(m.weights) else None
    else:
        u = rets[m.hedge] if m.hedge in rets else None
        if u is not None and lev:
            m.beta = float(lev)
        elif u is not None:
            xy = pd.concat([y, u], axis=1).dropna().tail(s.beta_lookback)
            if len(xy) > 20 and xy.iloc[:, 1].var() > 0:
                m.beta = float(np.cov(xy.iloc[:, 0], xy.iloc[:, 1])[0, 1] / xy.iloc[:, 1].var())
    if u is not None:
        m.und_hist = u
        res = (y - m.beta * u).dropna().tail(s.z_lookback)
        if len(res) > 10:
            m.resid_mean, m.resid_std = float(res.mean()), float(res.std())
    return m


def live_underlying(m: Model, r: dict[str, float]) -> tuple[float, float]:
    """Underlying move since T-1 close and the share of it that is priced live."""
    if m.kind == "BASKET":
        live = [(wt, r[s]) for s, wt in m.weights.items() if np.isfinite(r.get(s, np.nan))]
        if not live:
            return np.nan, 0.0
        wsum = sum(wt for wt, _ in live)
        move = sum(wt * x for wt, x in live)
        hr = r.get(m.hedge, np.nan) if m.hedge else np.nan
        move = move + (1 - wsum) * hr if np.isfinite(hr) else move / wsum
        return move, wsum
    hr = r.get(m.hedge, np.nan)
    return (m.beta * hr, 1.0) if np.isfinite(hr) else (np.nan, 0.0)


def evaluate(m: Model, q: dict, r: dict[str, float], s: Settings, now: dt.datetime) -> dict:
    bid, ask, last = (_f(q.get(k)) for k in ("bid", "ask", "last"))
    two_sided = np.isfinite(bid) and np.isfinite(ask) and ask >= bid > 0
    if two_sided and np.isfinite(last):
        # A quote far from the last trade is stale (Yahoo's bid/ask often is): don't trade off it.
        mid = (bid + ask) / 2
        two_sided = abs(mid / last - 1) <= max(1.5 * (ask - bid) / mid, 0.0010)
    px_src = "BID/ASK" if two_sided else "LAST"
    mid = (bid + ask) / 2 if two_sided else last
    if not two_sided:
        bid = ask = last
    nav, prev = _f(q.get("nav")), _f(q.get("prev_close"))
    prem_t1 = (prev / nav - 1) * 1e4 if np.isfinite(nav) and np.isfinite(prev) else np.nan
    use_nav = s.anchor.upper() == "NAV" and np.isfinite(prem_t1) and abs(prem_t1) < 200
    anchor, anchor_src = (nav, "NAV") if use_nav else (prev, "CLOSE")

    und, cov = live_underlying(m, r)
    fv = anchor * (1 + und) if np.isfinite(und) else np.nan
    cost = s.fi_cost_bps if m.is_fi else s.equity_cost_bps
    sell = (bid / fv - 1) * 1e4 - cost
    buy = (1 - ask / fv) * 1e4 - cost

    r_etf = r.get(m.etf, np.nan)
    z = (r_etf - und - m.resid_mean) / m.resid_std if np.isfinite(m.resid_std) and m.resid_std > 0 else np.nan
    peer_r = [r[p] for p in m.peers if np.isfinite(r.get(p, np.nan))]
    peer_gap = (r_etf - np.mean(peer_r)) * 1e4 if peer_r else np.nan

    t = q.get("time")
    stale = t is None or (now - t).total_seconds() > s.stale_minutes * 60
    edge_min = s.fi_min_edge_bps if m.is_fi else s.equity_min_edge_bps
    leg = "BASKET" if m.kind == "BASKET" else (m.hedge or "HEDGE")
    if not np.isfinite(fv):
        signal = "NO FAIR VALUE"
    elif stale:
        signal = "STALE / CLOSED"
    elif sell > edge_min:
        signal = f"SELL ETF / BUY {leg}"
    elif buy > edge_min:
        signal = f"BUY ETF / SELL {leg}"
    else:
        signal = ""
    vol, avg = q.get("volume", np.nan), q.get("avg_volume", np.nan)
    return {
        "ETF": m.etf, "Model": m.kind, "Hedge / Proxy": m.hedge or "",
        "Last": last, "Bid": bid if two_sided else None, "Ask": ask if two_sided else None,
        "Bid Sz": q.get("bid_size") if two_sided else None, "Ask Sz": q.get("ask_size") if two_sided else None,
        "Spread (bps)": (ask - bid) / mid * 1e4 if two_sided else np.nan,
        "Chg %": (last / prev - 1) if np.isfinite(prev) else np.nan,
        "Volume": vol, "Avg Vol": avg, "Rel Vol": vol / avg if avg else np.nan,
        "NAV T-1": nav, "T-1 Prem (bps)": prem_t1,
        "Und Move (bps)": und * 1e4, "Und Move": und, "Fair Value": fv,
        "Mid vs FV (bps)": (mid / fv - 1) * 1e4 if np.isfinite(fv) else np.nan,
        "Sell Edge (bps)": sell, "Buy Edge (bps)": buy,
        "Resid Z": z, "Resid σ (bps)": m.resid_std * 1e4,
        "Beta": m.beta, "Coverage": cov if m.kind == "BASKET" else np.nan,
        "Peer Gap (bps)": peer_gap, "Yield": q.get("yield"),
        "Signal": signal, "Px Src": px_src, "Anchor": anchor_src, "Quote Time": t,
    }


def history_frame(m: Model, close: pd.DataFrame, vol: pd.DataFrame, days: int) -> pd.DataFrame:
    """Daily ETF vs underlying index, cumulative residual spread and volume, for charts."""
    if m.etf not in close or m.und_hist is None:
        return pd.DataFrame()
    r = close[m.etf].pct_change()
    u = m.beta * m.und_hist
    df = pd.DataFrame({"etf_ret": r, "und_ret": u}).dropna().tail(days)
    df["ETF"] = (1 + df["etf_ret"]).cumprod() * 100
    df["Underlying"] = (1 + df["und_ret"]).cumprod() * 100
    df["Spread (bps)"] = ((df["etf_ret"] - df["und_ret"]) * 1e4).cumsum()
    for p in m.peers:
        if p in close:
            df[p] = (1 + close[p].pct_change().reindex(df.index).fillna(0)).cumprod() * 100
    df["Volume"] = vol[m.etf].reindex(df.index) if m.etf in vol else np.nan
    return df.drop(columns=["etf_ret", "und_ret"])


# ------------------------------------------------------------------ sigma bands
ET = ZoneInfo("America/New_York")


def session_left(now: dt.datetime | None = None) -> float:
    """Share of today's 9:30-16:00 ET session still to trade (1 before the open, 0 after the close)."""
    t = (now or dt.datetime.now(ET)).astimezone(ET)
    if t.weekday() >= 5:
        return 0.0
    o = t.replace(hour=9, minute=30, second=0, microsecond=0)
    c = t.replace(hour=16, minute=0, second=0, microsecond=0)
    return 1.0 if t < o else 0.0 if t >= c else (c - t).total_seconds() / (c - o).total_seconds()


def _phi(x: float) -> float:
    return 0.5 * (1 + math.erf(x / math.sqrt(2)))


def band_probs(move: float, sigma: float, left: float) -> tuple[float, float, float, float]:
    """Move so far in daily sigmas, and P(close beyond +/-1, 2, 3 sigma of yesterday's close).

    The rest of the day is a normal with variance sigma^2 x (session left), so the probability
    tightens toward 0 or 1 as the close approaches."""
    if not (np.isfinite(move) and np.isfinite(sigma) and sigma > 0):
        return (np.nan,) * 4
    z = move / sigma
    if left <= 0:
        return (z, *(float(abs(z) > k) for k in (1, 2, 3)))
    s = math.sqrt(left)
    return (z, *(_phi((-k - z) / s) + 1 - _phi((k - z) / s) for k in (1, 2, 3)))

