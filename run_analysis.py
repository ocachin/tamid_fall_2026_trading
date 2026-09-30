from __future__ import annotations

import argparse
import datetime as dt
import os
import sys
import time
from pathlib import Path

os.environ.setdefault("MPLBACKEND", "Agg")

import xlwings as xw

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / "xlwings_lite"))
import main as lite

SOURCE = HERE / "StockDatabase.xlsx"
OUTPUT = HERE / "output"
TIMEFRAMES = ("DAILY", "WEEKLY", "MONTHLY")


def open_book(path: Path) -> xw.Book:
    for app in xw.apps:
        for book in app.books:
            if book.name == path.name:
                return book
    return xw.Book(str(path))


def _fingerprint(sh):
    try:
        head = sh["A1:F2"].value
        if head[0][0] != "Date" or not isinstance(head[1][0], dt.datetime) or not isinstance(head[1][4], (int, float)):
            return None
        last_row = sh["A1"].expand("down").last_cell.row
        last_close = sh.range((last_row, 5)).value
    except Exception:
        return None
    return (head[1][0], head[1][4], last_row, last_close) if isinstance(last_close, (int, float)) else None


def select_and_wait(book: xw.Book, ticker: str, timeframe: str, timeout: float = 90.0) -> None:
    sh = book.sheets["Data"]
    if str(sh["O1"].value).upper() == ticker and str(sh["O4"].value).upper() == timeframe:
        return
    before = _fingerprint(sh)
    sh["O4"].value = timeframe
    sh["O1"].value = ticker
    book.app.calculate()
    deadline, prev = time.time() + timeout, None
    while True:
        fp = _fingerprint(sh)
        if fp is not None and fp != before and fp == prev:
            return
        if time.time() > deadline:
            raise TimeoutError(f"STOCKHISTORY returned no data for {ticker} ({timeframe}) in {timeout:.0f}s. "
                               "Check the ticker and your internet connection.")
        prev = fp
        time.sleep(1.5)


def build_one(source: xw.Book, ticker: str, timeframe: str) -> Path:
    select_and_wait(source, ticker, timeframe)
    ticker, timeframe, raw = lite.read_data_sheet(source)

    out = OUTPUT / f"Analysis_{ticker}_{timeframe}_{dt.date.today():%Y-%m-%d}.xlsx"
    for app in xw.apps:
        for bk in app.books:
            if bk.name == out.name:
                bk.close()

    book = xw.Book()
    blank = book.sheets[0].name
    lite.write_analysis_sheet(book, ticker, timeframe, raw)
    book.sheets[blank].delete()
    book.save(str(out))
    return out


def main() -> None:
    p = argparse.ArgumentParser(description="Lognormal return analysis, one workbook per ticker.")
    p.add_argument("tickers", nargs="*", help="e.g. META ORCL (default: the ticker in Data!O1)")
    p.add_argument("-t", "--timeframe", type=str.upper, choices=TIMEFRAMES, help="default: the timeframe in Data!O4")
    args = p.parse_args()

    OUTPUT.mkdir(exist_ok=True)
    source = open_book(SOURCE)
    data = source.sheets["Data"]
    original = (data["O1"].value, data["O4"].value)
    tickers = [t.upper() for t in args.tickers] or [str(original[0]).strip().upper()]
    timeframe = args.timeframe or str(original[1]).strip().upper()
    try:
        for t in tickers:
            print(f"{t} {timeframe}: loading...", flush=True)
            out = build_one(source, t, timeframe)
            print(f"saved {out.relative_to(HERE)}")
    finally:
        data["O4"].value = original[1]
        data["O1"].value = original[0]


if __name__ == "__main__":
    main()
