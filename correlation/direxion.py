"""Direxion fund data: daily holdings (exact NAV and swap exposure), distributions and 30-day NAV history.

    https://www.direxion.com/holdings/{TICKER}.csv   today's holdings; positions and prices are the prior
                                                     close. Ticker must be upper case (else HTTP 403).
                                                     Published around 10:00 ET, so early in the morning
                                                     the file can still describe the close before that.
    https://www.direxion.com/all-etfs                fund directory (Next.js JSON): distributions, product URL
    https://www.direxion.com{product url}            thirtyDayPricing: last 30 NAVs, rounded to the cent

Only today's holdings file exists, so every run saves a copy under cache/direxion/; the exact NAVs
build up from those copies.
"""
from __future__ import annotations

import io
import json
import re
import time
from functools import lru_cache
from pathlib import Path

import numpy as np
import pandas as pd
import requests

BASE = "https://www.direxion.com"
ARCHIVE = Path(__file__).resolve().parent / "cache" / "direxion"
HEADERS = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/128.0 Safari/537.36",
           "Accept": "text/html,application/xhtml+xml,text/csv,*/*;q=0.8"}
COUNTERPARTY = {"BCS": "Barclays", "GSS": "Goldman Sachs", "CTS": "Citibank", "MLS": "BofA Merrill Lynch",
                "BPS": "BNP Paribas", "UBS": "UBS"}


def _get(url: str) -> requests.Response:
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    time.sleep(0.5)                                         # be polite: one request at a time
    return r


def _next_data(html: str) -> dict:
    m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', html, re.S)
    if not m:
        raise ValueError("Direxion page layout changed: no __NEXT_DATA__ block")
    try:
        return json.loads(m.group(1))["props"]["pageProps"]["ssrData"]
    except KeyError as err:
        raise ValueError(f"Direxion page layout changed: no {err} in __NEXT_DATA__") from None


@lru_cache(maxsize=1)
def _funds() -> dict:
    return {f["Ticker"]: f for f in _next_data(_get(f"{BASE}/all-etfs").text)["funds"]}


def distributions(ticker: str) -> pd.Series:
    """Cash paid per share (income + capital gains + return of capital), indexed by ex-date."""
    fund = _funds()[ticker.upper()]
    if "Distributions" not in fund:                         # a renamed key must not read as "never paid"
        raise ValueError("Direxion fund record has no 'Distributions' field: page layout changed")
    rows = fund["Distributions"] or []
    amt = ["IncomeDividend", "ShortTermCapitalGain", "LongTermCapitalGain", "ReturnOfCapital"]
    s = pd.Series({pd.to_datetime(d["ExDate"], format="%m%d%Y"): sum(float(d.get(a) or 0) for a in amt) for d in rows},
                  dtype=float)
    return s[s > 0].groupby(level=0).sum().sort_index()


def nav_history(ticker: str) -> pd.Series:
    """Direxion's last 30 NAVs (rounded to the cent), indexed by NAV date."""
    data = _next_data(_get(BASE + _funds()[ticker.upper()]["Url"]).text)["thirtyDayPricing"].get(ticker.upper())
    if not data:
        raise ValueError(f"Direxion product page has no NAV history for {ticker.upper()}: page layout changed")
    return pd.Series({pd.to_datetime(d["TradeDate"], format="%m%d%Y"): float(d["Nav"]) for d in data}, dtype=float).sort_index()


def parse_holdings(text: str) -> pd.DataFrame:
    """One row per position; attrs: shares_outstanding, trade_date (publication day)."""
    h = pd.read_csv(io.StringIO(text[text.index('"TradeDate"'):]))
    df = pd.DataFrame({"fund": h.AccountTicker, "description": h.SecurityDescription, "cusip": h.Cusip,
                       "shares": h.Shares.astype(float), "price": h.Price.astype(float),
                       "market_value": h.MarketValue.astype(float), "pct_net_assets": h.HoldingsPercent.astype(float)})
    swap = df.description.str.contains("SWAP", case=False)
    df["kind"] = np.where(swap, "swap", np.where(df.price.eq(1.0), "money market", "other"))
    df["counterparty"] = np.where(swap, df.cusip.str[-3:].map(COUNTERPARTY).fillna(df.cusip.str[-3:]), "")
    df.attrs["shares_outstanding"] = int(re.search(r"Shares Outstanding:\s*([\d,]+)", text).group(1).replace(",", ""))
    df.attrs["trade_date"] = pd.to_datetime(h.TradeDate.iloc[0].split()[0], format="%m/%d/%Y")
    return df


def holdings(ticker: str) -> pd.DataFrame:
    """Today's holdings file (saved to the archive as {TICKER}_{publication date}.csv)."""
    text = _get(f"{BASE}/holdings/{ticker.upper()}.csv").text
    df = parse_holdings(text)
    ARCHIVE.mkdir(parents=True, exist_ok=True)
    (ARCHIVE / f"{ticker.upper()}_{df.attrs['trade_date']:%Y%m%d}.csv").write_text(text)
    return df


def exact_nav(df: pd.DataFrame) -> float:
    """Net assets / shares. Net assets = market value / weight, read off the money-market rows."""
    mm = df[df.kind == "money market"]
    return float(mm.market_value.sum() / (mm.pct_net_assets.sum() / 100) / df.attrs["shares_outstanding"])


def nav_date(df: pd.DataFrame, und_close: pd.Series) -> pd.Timestamp | None:
    """The close a holdings file describes: the session whose underlying close equals the swap price.

    Needed because an early-morning file can still describe the close two sessions back."""
    swaps = df[df.kind == "swap"]
    if swaps.empty:
        return None
    px = float(swaps.price.iloc[0])
    prior = und_close[und_close.index < df.attrs["trade_date"]].dropna().tail(5)
    hit = prior[(prior - px).abs() <= 0.006]
    return hit.index[-1] if len(hit) else None


def archived_navs(ticker: str, und_close: pd.Series) -> pd.Series:
    """Exact NAVs from every saved holdings file, indexed by the close they describe."""
    out = {}
    for p in sorted(ARCHIVE.glob(f"{ticker.upper()}_*.csv")):
        df = parse_holdings(p.read_text())
        d = nav_date(df, und_close)
        if d is not None:
            out[d] = exact_nav(df)
    return pd.Series(out, dtype=float).sort_index()
