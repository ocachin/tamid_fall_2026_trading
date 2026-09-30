"""Market data: live quotes, daily history and ETF holdings.

Providers
- Yahoo (yfinance): free, no key. Batch quotes (last, volume, prev close) for ETFs and every
  underlying, NAV and yield, daily history, top-10 holdings. Its bid/ask fields are often stale.
- Nasdaq (api.nasdaq.com): free, no key, real-time Nasdaq bid/ask/sizes for the ETFs.
  Unofficial public endpoint (the one nasdaq.com uses), one request per ETF.
- Alpaca (optional): free key, real-time IEX top-of-book bid/ask for ETFs and underlyings.
  Enabled automatically when ALPACA_API_KEY / ALPACA_SECRET_KEY are set (env or etf_arb/.env).
Holdings: SPDR and Invesco daily files (full basket), Yahoo top 10, any xlsx/csv URL, or rows
typed in Excel.
"""
from __future__ import annotations

import datetime as dt
import io
import logging
import os
import re
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import requests
import yfinance as yf

logging.getLogger("yfinance").setLevel(logging.CRITICAL)

HERE = Path(__file__).resolve().parent
UA = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
EQUITY_SYMBOL = re.compile(r"^[A-Z]{1,5}([.\-][A-Z]{1,2})?$")
LOCAL_TZ = dt.datetime.now().astimezone().tzinfo
NOT_STOCKS = {"CASH", "USD", "XTSLA", "MSFUT", "FGXXX", "DGCXX"}


def load_env() -> None:
    env = HERE / ".env"
    if env.exists():
        for line in env.read_text().splitlines():
            if "=" in line and not line.strip().startswith("#"):
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def yahoo_symbol(s: str) -> str:
    s = str(s).strip().upper()
    return s.replace(".", "-") if EQUITY_SYMBOL.match(s) else s  # BRK.B -> BRK-B


def _num(x) -> float:
    try:
        x = float(x)
        return x if np.isfinite(x) else np.nan
    except (TypeError, ValueError):
        return np.nan


# --------------------------------------------------------------------------- Yahoo
class Yahoo:
    name = "yfinance"
    QUOTE = "https://query1.finance.yahoo.com/v7/finance/quote"

    def __init__(self):
        from yfinance.data import YfData  # yfinance's session handles Yahoo's cookie/crumb
        self.yd = YfData()

    def quotes(self, symbols: list[str]) -> dict[str, dict]:
        """Batch quotes: 250 symbols per request. Bid/ask here are often stale."""
        out: dict[str, dict] = {}
        for i in range(0, len(symbols), 250):
            r = self.yd.get(self.QUOTE, params={"symbols": ",".join(symbols[i:i + 250])}, timeout=15)
            for q in r.json().get("quoteResponse", {}).get("result", []):
                t = q.get("regularMarketTime")
                out[q["symbol"]] = {
                    "last": _num(q.get("regularMarketPrice")),
                    "bid": _num(q.get("bid")) or np.nan, "ask": _num(q.get("ask")) or np.nan,
                    "bid_size": _num(q.get("bidSize")), "ask_size": _num(q.get("askSize")),
                    "prev_close": _num(q.get("regularMarketPreviousClose")),
                    "volume": _num(q.get("regularMarketVolume")),
                    "avg_volume": _num(q.get("averageDailyVolume3Month")),
                    "time": dt.datetime.fromtimestamp(t) if isinstance(t, (int, float)) else None,
                    "type": q.get("quoteType"),
                }
        return out

    @staticmethod
    def _meta(sym: str) -> dict:
        try:
            i = yf.Ticker(sym).info
        except Exception:
            return {}
        return {"nav": _num(i.get("navPrice")), "yield": _num(i.get("yield")), "name": i.get("shortName") or ""}

    def etf_meta(self, symbols: list[str]) -> dict[str, dict]:
        """NAV (T-1), distribution yield, name. Changes once a day, so callers cache it."""
        with ThreadPoolExecutor(max_workers=8) as ex:
            return dict(zip(symbols, ex.map(self._meta, symbols)))

    def daily_history(self, symbols: list[str], period: str) -> tuple[pd.DataFrame, pd.DataFrame]:
        """Dividend-adjusted closes and volumes, one column per symbol."""
        d = yf.download(sorted(set(symbols)), period=period, interval="1d", progress=False,
                        threads=True, auto_adjust=True)
        if d.empty:
            return pd.DataFrame(), pd.DataFrame()
        if not isinstance(d.columns, pd.MultiIndex):
            d.columns = pd.MultiIndex.from_product([d.columns, symbols])
        close, vol = d["Close"].dropna(how="all", axis=1), d["Volume"]
        close.index = pd.to_datetime(close.index).tz_localize(None)
        vol.index = close.index
        return close, vol


# --------------------------------------------------------------------------- Nasdaq
class Nasdaq:
    name = "nasdaq"
    URL = "https://api.nasdaq.com/api/quote/{}/info?assetclass={}"
    HEAD = {**UA, "Accept": "application/json"}
    ET = ZoneInfo("America/New_York")

    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update(self.HEAD)

    @staticmethod
    def _px(x) -> float:
        return _num(str(x).replace("$", "").replace(",", "")) if x not in (None, "", "N/A") else np.nan

    def _one(self, sym: str, cls: str = "etf") -> dict:
        try:
            d = self.session.get(self.URL.format(sym.replace("-", "."), cls), timeout=8).json()["data"]["primaryData"]
        except Exception:
            return {}
        try:
            ts = dt.datetime.strptime(d["lastTradeTimestamp"].replace(" ET", ""), "%b %d, %Y %I:%M %p")
            ts = ts.replace(tzinfo=self.ET).astimezone(LOCAL_TZ).replace(tzinfo=None)
        except Exception:
            ts = None
        return {"bid": self._px(d.get("bidPrice")) or np.nan, "ask": self._px(d.get("askPrice")) or np.nan,
                "bid_size": self._px(d.get("bidSize")), "ask_size": self._px(d.get("askSize")),
                "last": self._px(d.get("lastSalePrice")), "time": ts}

    def quotes(self, symbols: list[str], stocks: set[str] = frozenset()) -> dict[str, dict]:
        """Real-time bid/ask per symbol; `stocks` are queried as stocks, the rest as ETFs."""
        syms = [s for s in symbols if EQUITY_SYMBOL.match(s)]
        cls = ["stocks" if s in stocks else "etf" for s in syms]
        with ThreadPoolExecutor(max_workers=16) as ex:
            return {s: q for s, q in zip(syms, ex.map(self._one, syms, cls)) if q}

    def _watch(self, chunk: list[tuple[str, str]]) -> list[dict]:
        try:
            r = self.session.get("https://api.nasdaq.com/api/quote/watchlist", timeout=20,
                                 params=[("symbol", f"{s.replace('-', '.').lower()}|{c}") for s, c in chunk])
            return r.json().get("data") or []
        except Exception:
            return []

    def batch(self, symbols: list[str], etfs: set[str]) -> dict[str, dict]:
        """Last / prev close / volume, 20 symbols per request. Fallback when Yahoo throttles."""
        pairs = [(s, "etf" if s in etfs else "stocks") for s in symbols if EQUITY_SYMBOL.match(s)]
        chunks = [pairs[i:i + 20] for i in range(0, len(pairs), 20)]
        out: dict[str, dict] = {}
        with ThreadPoolExecutor(max_workers=8) as ex:
            for rows in ex.map(self._watch, chunks):
                for d in rows:
                    try:
                        ts = pd.Timestamp(d["lastTradeTimestampDateTime"]).tz_convert(LOCAL_TZ)
                        ts = ts.tz_localize(None).to_pydatetime()
                    except Exception:
                        ts = None
                    out[yahoo_symbol(d["symbol"])] = {
                        "last": self._px(d.get("lastSalePrice")), "prev_close": self._px(d.get("previousClosePrice")),
                        "volume": self._px(d.get("volume")), "time": ts}
        return out


# --------------------------------------------------------------------------- Alpaca
class Alpaca:
    """Real-time IEX quotes. Note IEX is one venue: its spread is wider than the NBBO."""
    name = "alpaca-iex"
    URL = "https://data.alpaca.markets/v2/stocks/snapshots"

    def __init__(self, key: str, secret: str, feed: str = "iex"):
        self.h = {"APCA-API-KEY-ID": key, "APCA-API-SECRET-KEY": secret}
        self.feed = feed

    @classmethod
    def from_env(cls) -> "Alpaca | None":
        load_env()
        k, s = os.environ.get("ALPACA_API_KEY"), os.environ.get("ALPACA_SECRET_KEY")
        return cls(k, s, os.environ.get("ALPACA_FEED", "iex")) if k and s else None

    def quotes(self, symbols: list[str]) -> dict[str, dict]:
        out: dict[str, dict] = {}
        syms = [s for s in symbols if EQUITY_SYMBOL.match(s)]
        for i in range(0, len(syms), 200):
            chunk = {s.replace("-", "."): s for s in syms[i:i + 200]}
            r = requests.get(self.URL, headers=self.h, timeout=10,
                             params={"symbols": ",".join(chunk), "feed": self.feed})
            r.raise_for_status()
            data = r.json()
            data = data.get("snapshots", data)
            for a_sym, snap in data.items():
                if not snap:
                    continue
                q, t = snap.get("latestQuote") or {}, snap.get("latestTrade") or {}
                ts = q.get("t") or t.get("t")
                out[chunk.get(a_sym, a_sym)] = {
                    "bid": _num(q.get("bp")) or np.nan, "ask": _num(q.get("ap")) or np.nan,
                    "bid_size": _num(q.get("bs")), "ask_size": _num(q.get("as")),
                    "last": _num(t.get("p")),
                    "time": pd.Timestamp(ts).tz_convert(LOCAL_TZ).tz_localize(None).to_pydatetime() if ts else None,
                }
        return out


# --------------------------------------------------------------------------- Holdings
def _parse_holdings_table(raw: pd.DataFrame) -> pd.DataFrame:
    """Find the header row that has a Ticker/Symbol and a Weight column (SPDR, iShares-style files)."""
    for i in range(min(len(raw), 40)):
        cells = [str(c).strip().lower() for c in raw.iloc[i]]
        sym = next((j for j, c in enumerate(cells) if c in ("ticker", "symbol")), None)
        wgt = next((j for j, c in enumerate(cells) if c.startswith("weight") or c == "% of net assets"), None)
        if sym is None or wgt is None:
            continue
        name = next((j for j, c in enumerate(cells) if c in ("name", "security name", "holding")), None)
        body = raw.iloc[i + 1:]
        df = pd.DataFrame({
            "symbol": body.iloc[:, sym].astype(str).str.strip().str.upper(),
            "name": body.iloc[:, name].astype(str) if name is not None else "",
            "weight": pd.to_numeric(body.iloc[:, wgt].astype(str).str.replace("%", "").str.replace(",", ""),
                                    errors="coerce"),
        }).dropna(subset=["weight"])
        if df["weight"].sum() > 1.5:
            df["weight"] /= 100.0
        return df
    raise ValueError("no Ticker/Weight header found")


def fetch_holdings(etf: str, source: str) -> pd.DataFrame:
    """Columns symbol, name, weight (fraction of NAV), sorted by weight. Stocks only."""
    source = (source or "").strip()
    src = source.upper()
    if src == "SPDR":
        source = ("https://www.ssga.com/us/en/intermediary/library-content/products/fund-data/"
                  f"etfs/us/holdings-daily-us-en-{etf.lower()}.xlsx")
    if source.lower().startswith("http"):
        r = requests.get(source, headers=UA, timeout=30)
        r.raise_for_status()
        if r.content[:2] == b"PK":
            raw = pd.read_excel(io.BytesIO(r.content), header=None)
        else:
            raw = pd.read_csv(io.StringIO(r.text), header=None, on_bad_lines="skip", engine="python")
        df = _parse_holdings_table(raw)
    elif src == "INVESCO":
        r = requests.get("https://dng-api.invesco.com/cache/v1/accounts/en_US/shareclasses/"
                         f"{etf.upper()}/holdings/fund?idType=ticker&interval=monthly&productType=ETF", headers=UA, timeout=30)
        r.raise_for_status()
        h = pd.DataFrame(r.json()["holdings"])
        df = pd.DataFrame({"symbol": h["ticker"].astype(str).str.upper(), "name": h["issuerName"],
                           "weight": pd.to_numeric(h["percentageOfTotalNetAssets"], errors="coerce") / 100})
    elif src == "YAHOO":
        th = yf.Ticker(etf).funds_data.top_holdings
        df = pd.DataFrame({"symbol": th.index.astype(str), "name": th["Name"], "weight": th["Holding Percent"]})
    else:
        return pd.DataFrame(columns=["symbol", "name", "weight"])
    df = df[df["symbol"].str.match(EQUITY_SYMBOL) & ~df["symbol"].isin(NOT_STOCKS)]
    df["symbol"] = df["symbol"].map(yahoo_symbol)
    return df.groupby("symbol", as_index=False).agg(name=("name", "first"), weight=("weight", "sum")) \
             .sort_values("weight", ascending=False).reset_index(drop=True)
