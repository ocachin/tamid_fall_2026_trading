"""Intraday correlation of Direxion Bear 1X ETFs against their underlyings.

Pairs: AVS/AVGO, AAPD/AAPL, METD/META, MSFD/MSFT, SPDN/^GSPC, AMZD/AMZN, PLTD/PLTR.

Python only downloads prices and writes returns. Every correlation, beta and R-squared in the
workbook is an Excel formula (CORREL / SLOPE / RSQ) over those return columns, so each number can
be clicked and audited.

    .venv/bin/python correlation/intraday_corr.py          # build the workbook
    .venv/bin/python correlation/intraday_corr.py --open    # ... and open it in Excel
"""
from __future__ import annotations

import argparse
import datetime as dt
import logging
import os
from pathlib import Path

os.environ.setdefault("MPLBACKEND", "Agg")

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import yfinance as yf
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

logging.getLogger("yfinance").setLevel(logging.CRITICAL)

HERE = Path(__file__).resolve().parent
BOOK = HERE / "Correlation_Intraday.xlsx"
FONT, THIN = "Arial", Side(style="thin", color="000000")

PAIRS = [("AVS", "AVGO"), ("AAPD", "AAPL"), ("METD", "META"), ("MSFD", "MSFT"),
         ("SPDN", "^GSPC"), ("AMZD", "AMZN"), ("PLTD", "PLTR")]
ETFS = [e for e, _ in PAIRS]
UNDER = [u for _, u in PAIRS]
SYMS = ETFS + UNDER

EPPS_MINUTES = [1, 2, 3, 5, 10, 15, 30, 60]
# sheet, yfinance interval, period, resample rule (None = use as downloaded)
GRIDS = ([(f"R{m}", "1m", "5d", None if m == 1 else f"{m}min") for m in EPPS_MINUTES]
         + [("R5_60d", "5m", "60d", None), ("R15_60d", "15m", "60d", None), ("RD", "1d", "1y", None)])


def fetch(interval: str, period: str) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Close prices, traded flags (False = the symbol printed no bar) and volumes."""
    d = yf.download(SYMS, period=period, interval=interval, progress=False,
                    auto_adjust=False, prepost=False, threads=True)
    close, vol = d["Close"][SYMS], d["Volume"][SYMS]
    traded = close.notna()
    if interval != "1d":                      # regular hours only
        close, traded, vol = (x.between_time("09:30", "16:00") for x in (close, traded, vol))
    return close, traded, vol


def returns(close: pd.DataFrame, traded: pd.DataFrame, rule: str | None, intraday: bool):
    """Simple returns per bar, with a per-pair flag for bars where both legs actually traded.

    Prices are forward filled inside each session (a missing bar means no print, not a gap in
    value); the flag marks the bars where that fill was needed, so Excel can exclude them."""
    if intraday:
        day = close.index.date
        close = close.groupby(day).ffill()
        if rule:
            agg = close.resample(rule, label="right", closed="right").last()
            traded = traded.resample(rule, label="right", closed="right").max()
            close = agg.dropna(how="all")
            traded = traded.reindex(close.index).fillna(0).astype(bool)
        r = close.pct_change()
        r[pd.Series(close.index.date, index=close.index).ne(
            pd.Series(close.index.date, index=close.index).shift()).values] = np.nan  # drop overnight
    else:
        r = close.pct_change()
    traded = traded.fillna(0).astype(bool)
    both = pd.DataFrame({e: traded[e] & traded[u] for e, u in PAIRS}, index=traded.index)
    return r.dropna(how="all"), both.reindex(r.dropna(how="all").index).fillna(False)


def epps_curve(close: pd.DataFrame, minutes: list[int]) -> pd.DataFrame:
    """Correlation of each pair as the sampling interval lengthens."""
    day = close.index.date
    filled = close.groupby(day).ffill()
    out = {}
    for m in minutes:
        px = filled if m == 1 else filled.resample(f"{m}min", label="right", closed="right").last()
        r = px.pct_change()
        r[pd.Series(px.index.date, index=px.index).ne(
            pd.Series(px.index.date, index=px.index).shift()).values] = np.nan
        out[m] = {f"{e}/{u}": r[e].corr(r[u]) for e, u in PAIRS}
    return pd.DataFrame(out)


# ------------------------------------------------------------------ workbook
def write_returns(ws, r: pd.DataFrame, both: pd.DataFrame, bold, base) -> int:
    ws.cell(1, 1, "Time").font = bold
    ws.column_dimensions["A"].width = 17
    for k, (e, u) in enumerate(PAIRS):
        for j, lab in enumerate((e, u, f"{e} both traded", f"{e} traded-only", f"{u} traded-only")):
            c = ws.cell(1, 2 + NCOL * k + j, lab)
            c.font, c.alignment = bold, Alignment(horizontal="right", wrap_text=True)
            ws.column_dimensions[get_column_letter(2 + NCOL * k + j)].width = 11
    rows = []
    assert NCOL == 5
    for t in r.index:
        row = [t.tz_localize(None) if getattr(t, "tzinfo", None) else t]
        for e, u in PAIRS:
            a, b = r.at[t, e], r.at[t, u]
            row += [None if pd.isna(a) else float(a), None if pd.isna(b) else float(b),
                    int(bool(both.at[t, e])), None, None]   # 5 slots per pair: the last two
                                                            # are filled with the traded-only formulas
        rows.append(row)
    assert all(len(row) == 1 + NCOL * len(PAIRS) for row in rows), "row width must match headers"
    for i, row in enumerate(rows, start=2):
        for j, v in enumerate(row, start=1):
            c = ws.cell(i, j, v)
            c.font = base
            c.number_format = "yyyy-mm-dd hh:mm" if j == 1 else ("0" if (j - 1) % NCOL == 0 else "0.0000%")
        for k in range(len(PAIRS)):           # traded-only copies: blank unless both legs printed
            a, b, f = (get_column_letter(2 + NCOL * k + x) for x in range(3))
            for x, src in enumerate((a, b)):
                c = ws.cell(i, 2 + NCOL * k + 3 + x, f'=IF({f}{i}=1,{src}{i},"")')
                c.font, c.number_format = base, "0.0000%"
    ws.freeze_panes = "B2"
    return len(rows) + 1


NCOL = 5                                       # columns per pair on a returns sheet


def col_of(_unused: int, k: int, which: int) -> str:
    return get_column_letter(2 + NCOL * k + which)


def build(data: dict, epps: pd.DataFrame, quality: pd.DataFrame, png: Path) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    base, bold = Font(name=FONT, size=10), Font(name=FONT, size=10, bold=True)
    sheets = {}
    for name, *_ in GRIDS:
        sheets[name] = wb.create_sheet(name)
    summary = wb.create_sheet("Summary", 0)
    chart_ws = wb.create_sheet("Chart", 1)
    qual = wb.create_sheet("Data Quality", 2)
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False

    last = {}
    for name, *_ in GRIDS:
        r, both = data[name]
        last[name] = write_returns(sheets[name], r, both, bold, base)

    # ---- Summary: every cell a formula over the return sheets
    heads = ["Pair", "ETF", "Underlying", "ρ 1m", "ρ 1m traded-only", "ρ 5m", "ρ 15m",
             "ρ 5m (60d)", "ρ 15m (60d)", "ρ daily (1y)", "β 15m", "R² 15m", "β daily", "R² daily",
             "n 1m", "n 15m", "ETF minutes traded"]
    for j, h in enumerate(heads, start=1):
        c = summary.cell(1, j, h)
        c.font, c.border = bold, Border(bottom=THIN)
        c.alignment = Alignment(horizontal="left" if j <= 3 else "right", wrap_text=True)
        summary.column_dimensions[get_column_letter(j)].width = 13 if j > 3 else 11
    for k, (e, u) in enumerate(PAIRS):
        row = 2 + k
        a, b, f, ta, tb = (col_of(0, k, i) for i in range(5))
        def rng(sheet: str, col: str) -> str:
            return f"{sheet}!${col}$2:${col}${last[sheet]}"
        vals = {
            1: f"{e}/{u}", 2: e, 3: u,
            4: f"=CORREL({rng('R1', a)},{rng('R1', b)})",
            5: f"=CORREL({rng('R1', ta)},{rng('R1', tb)})",
            6: f"=CORREL({rng('R5', a)},{rng('R5', b)})",
            7: f"=CORREL({rng('R15', a)},{rng('R15', b)})",
            8: f"=CORREL({rng('R5_60d', a)},{rng('R5_60d', b)})",
            9: f"=CORREL({rng('R15_60d', a)},{rng('R15_60d', b)})",
            10: f"=CORREL({rng('RD', a)},{rng('RD', b)})",
            11: f"=SLOPE({rng('R15', a)},{rng('R15', b)})",
            12: f"=RSQ({rng('R15', a)},{rng('R15', b)})",
            13: f"=SLOPE({rng('RD', a)},{rng('RD', b)})",
            14: f"=RSQ({rng('RD', a)},{rng('RD', b)})",
            15: f"=COUNT({rng('R1', a)})",
            16: f"=COUNT({rng('R15', a)})",
            17: f"=AVERAGE({rng('R1', f)})",
        }
        for j, v in vals.items():
            c = summary.cell(row, j, v)
            c.font = base
            c.number_format = ("0.000" if 4 <= j <= 14 else "0%" if j == 17 else "#,##0" if j >= 15 else "General")
    summary.freeze_panes = "D2"

    # ---- Data quality
    qual.cell(1, 1, "Symbol").font = bold
    for j, h in enumerate(["Symbol", "1m bars present", "of possible", "% of minutes traded",
                           "median volume / min", "price", "tick (bps)", "σ 1m (bps)",
                           "tick / σ"], start=1):
        c = qual.cell(1, j, h)
        c.font, c.border = bold, Border(bottom=THIN)
        qual.column_dimensions[get_column_letter(j)].width = 18
    for i, (sym, row) in enumerate(quality.iterrows(), start=2):
        qual.cell(i, 1, sym).font = base
        for j, v in enumerate(row, start=2):
            c = qual.cell(i, j, float(v) if pd.notna(v) else None)
            c.font = base
            c.number_format = {4: "0%", 6: "#,##0.00", 7: "0.0", 8: "0.0", 9: "0.00"}.get(j, "#,##0")

    # ---- Epps table + chart picture
    chart_ws.cell(1, 1, "Correlation by sampling interval").font = bold
    chart_ws.cell(2, 1, "Interval (min)").font = bold
    chart_ws.column_dimensions["A"].width = 16
    for j, m in enumerate(epps.columns, start=2):
        c = chart_ws.cell(2, j, m)
        c.font = bold
        chart_ws.column_dimensions[get_column_letter(j)].width = 8
    for i, (pair, row) in enumerate(epps.iterrows(), start=3):
        chart_ws.cell(i, 1, pair).font = base
        for j, v in enumerate(row, start=2):
            c = chart_ws.cell(i, j, None if pd.isna(v) else float(v))
            c.font, c.number_format = base, "0.000"
    # ---- correlation per session (5-minute returns), one row per pair
    byday = wb.create_sheet("By Day", 3)
    idx5 = data["R5"][0].index
    days = sorted({t.date() for t in idx5})
    spans = {d0: (int(np.argmax([t.date() == d0 for t in idx5])) + 2,
                  int(len(idx5) - np.argmax([t.date() == d0 for t in idx5][::-1])) + 1) for d0 in days}
    byday.cell(1, 1, "Pair").font = bold
    byday.column_dimensions["A"].width = 13
    for j, d0 in enumerate(days, start=2):
        c = byday.cell(1, j, d0)
        c.font, c.border, c.number_format = bold, Border(bottom=THIN), "ddd dd mmm"
        byday.column_dimensions[get_column_letter(j)].width = 11
    for k, (e, u) in enumerate(PAIRS):
        a, b = col_of(0, k, 0), col_of(0, k, 1)
        byday.cell(2 + k, 1, f"{e}/{u}").font = base
        for j, d0 in enumerate(days, start=2):
            lo, hi = spans[d0]
            c = byday.cell(2 + k, j, f"=CORREL(R5!${a}${lo}:${a}${hi},R5!${b}${lo}:${b}${hi})")
            c.font, c.number_format = base, "0.000"

    img = XLImage(str(png))
    img.anchor = f"A{len(epps) + 5}"
    chart_ws.add_image(img)

    write_daily(wb.create_sheet("Daily", 4), days, last, bold, base)
    wb.save(BOOK)


def write_daily(ws, days: list, last: dict, bold, base) -> None:
    """One session at a time: ρ per pair at every sampling interval, using only that day's bars.

    B3 picks the session, B4 the last bar time included (move it before 16:00 to drop the close).
    Rows 15-17 find that window on each return sheet with COUNTIF over the sorted time column;
    the correlations are CORREL over INDEX(first):INDEX(last)."""
    ws.column_dimensions["A"].width = 16
    ws.cell(1, 1, "Correlation for one session, by sampling interval").font = bold
    ws.cell(3, 1, "Session").font = bold
    ws.cell(4, 1, "Bars up to").font = bold
    sel, cut = ws.cell(3, 2, days[-1]), ws.cell(4, 2, dt.time(16, 0))
    sel.number_format, cut.number_format = "ddd dd mmm yyyy", "hh:mm"
    for c in (sel, cut):
        c.font, c.border = bold, Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
    ws.column_dimensions["B"].width = 16
    ws.cell(3, 4, "← pick a session").font = base
    ws.cell(4, 4, "← last bar included (e.g. 15:45 drops the close)").font = base

    # list of sessions for the dropdown
    lc = get_column_letter(4 + len(EPPS_MINUTES) + 3)
    ws.cell(6, 4 + len(EPPS_MINUTES) + 3, "Sessions").font = bold
    ws.column_dimensions[lc].width = 14
    for i, d0 in enumerate(days, start=7):
        c = ws.cell(i, 4 + len(EPPS_MINUTES) + 3, d0)
        c.font, c.number_format = base, "ddd dd mmm yyyy"
    dv = DataValidation(type="list", formula1=f"${lc}$7:${lc}${6 + len(days)}", allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B3")
    tv = DataValidation(type="time", operator="between", formula1="0.3958333", formula2="0.6666667")
    ws.add_data_validation(tv)
    tv.add("B4")

    head = 6
    ws.cell(head, 1, "Pair").font = bold
    for j, m in enumerate(EPPS_MINUTES, start=2):
        c = ws.cell(head, j, f"{m} min")
        c.font, c.border, c.alignment = bold, Border(bottom=THIN), Alignment(horizontal="right")
        ws.column_dimensions[get_column_letter(j)].width = max(ws.column_dimensions[get_column_letter(j)].width or 0, 9)
    jt = 2 + len(EPPS_MINUTES)
    c = ws.cell(head, jt, "1 min traded-only")
    c.font, c.border = bold, Border(bottom=THIN)
    c.alignment = Alignment(horizontal="right", wrap_text=True)
    ws.column_dimensions[get_column_letter(jt)].width = 11
    ws.cell(head, 1).border = Border(bottom=THIN)

    # window on each return sheet: first / last data row of the chosen session
    lo_r, hi_r, n_r = head + len(PAIRS) + 3, head + len(PAIRS) + 4, head + len(PAIRS) + 5
    ws.cell(lo_r - 1, 1, "Rows used on each return sheet").font = bold
    for r, lab in ((lo_r, "first row"), (hi_r, "last row"), (n_r, "returns (n)")):
        ws.cell(r, 1, lab).font = base
    for j, m in enumerate(EPPS_MINUTES, start=2):
        s, n = f"R{m}", last[f"R{m}"]
        t = f"{s}!$A$2:$A${n}"
        col = get_column_letter(j)
        ws.cell(lo_r, j, f'=COUNTIF({t},"<"&$B$3)+2').font = base
        ws.cell(hi_r, j, f'=COUNTIF({t},"<"&($B$3+$B$4+1/2880))+1').font = base
        ws.cell(n_r, j, f"=COUNT(INDEX({s}!$B:$B,{col}${lo_r}):INDEX({s}!$B:$B,{col}${hi_r}))").font = base

    def span(sheet: str, c: str, jcol: str) -> str:
        return f"INDEX({sheet}!${c}:${c},{jcol}${lo_r}):INDEX({sheet}!${c}:${c},{jcol}${hi_r})"

    for k, (e, u) in enumerate(PAIRS):
        row = head + 1 + k
        ws.cell(row, 1, f"{e}/{u}").font = base
        a, b, _f, ta, tb = (col_of(0, k, i) for i in range(5))
        for j, m in enumerate(EPPS_MINUTES, start=2):
            jc = get_column_letter(j)
            c = ws.cell(row, j, f'=IFERROR(CORREL({span(f"R{m}", a, jc)},{span(f"R{m}", b, jc)}),"")')
            c.font, c.number_format = base, "0.000"
        c = ws.cell(row, jt, f'=IFERROR(CORREL({span("R1", ta, "B")},{span("R1", tb, "B")}),"")')
        c.font, c.number_format = base, "0.000"
    ws.freeze_panes = f"B{head + 1}"

    ch = LineChart()
    ch.title = "Correlation vs sampling interval, selected session"
    ch.y_axis.title, ch.x_axis.title = "ρ", "sampling interval"
    ch.y_axis.scaling.min, ch.y_axis.scaling.max = -1, 1
    ch.y_axis.delete = ch.x_axis.delete = False
    ch.height, ch.width = 9, 20
    last_col = 1 + len(EPPS_MINUTES)
    ch.add_data(Reference(ws, min_col=1, max_col=last_col, min_row=head + 1, max_row=head + len(PAIRS)),
                from_rows=True, titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, max_col=last_col, min_row=head))
    for s in ch.series:
        s.marker.symbol, s.smooth = "circle", False
    ws.add_chart(ch, f"A{n_r + 3}")


def draw(epps: pd.DataFrame, r1: pd.DataFrame, r15: pd.DataFrame, png: Path) -> None:
    plt.rcParams.update({"font.family": FONT, "font.size": 9, "axes.spines.top": False,
                         "axes.spines.right": False, "axes.titlesize": 10,
                         "axes.titleweight": "bold", "legend.frameon": False})
    fig, ax = plt.subplots(1, 3, figsize=(16, 4.6))

    a = ax[0]
    for pair, row in epps.iterrows():
        a.plot(row.index, row.values.astype(float), marker="o", ms=3, lw=1.2, label=pair)
    a.axhline(-1, color="black", lw=0.6, ls=":")
    a.set_xscale("log")
    a.set_xticks(list(epps.columns))
    a.set_xticklabels([str(m) for m in epps.columns])
    a.set_xlabel("sampling interval (minutes)")
    a.set_title("Correlation vs sampling interval")
    a.legend(fontsize=7, ncol=2)

    a = ax[1]
    y = [abs(epps.loc[f"{e}/{u}", 1]) for e, u in PAIRS]
    y15 = [abs(epps.loc[f"{e}/{u}", 15]) for e, u in PAIRS]
    x = np.arange(len(PAIRS))
    a.bar(x - 0.2, y, 0.4, color="#8faadc", label="1 min")
    a.bar(x + 0.2, y15, 0.4, color="#1f3864", label="15 min")
    a.set_xticks(x)
    a.set_xticklabels([e for e, _ in PAIRS], rotation=45, ha="right")
    a.set_ylim(0, 1)
    a.set_title("|correlation|: 1 min vs 15 min")
    a.legend(fontsize=8)

    a = ax[2]
    e, u = PAIRS[1]
    a.scatter(r15[u] * 100, r15[e] * 100, s=8, alpha=0.5, color="#1f3864")
    lim = np.nanpercentile(np.abs(r15[u] * 100), 99) * 1.3
    a.plot([-lim, lim], [lim, -lim], color="#c55a11", lw=1)
    a.set_xlim(-lim, lim)
    a.set_ylim(-lim, lim)
    a.set_xlabel(f"{u} 15-min return (%)")
    a.set_ylabel(f"{e} 15-min return (%)")
    a.set_title(f"{e} vs {u}, 15-min returns")
    fig.tight_layout()
    fig.savefig(png, dpi=110)
    plt.close(fig)


def main() -> None:
    p = argparse.ArgumentParser(description="Intraday correlation, Bear 1X ETFs vs underlyings")
    p.add_argument("--open", action="store_true", help="open the workbook in Excel when done")
    args = p.parse_args()

    raw: dict[tuple[str, str], tuple[pd.DataFrame, pd.DataFrame]] = {}
    data: dict[str, tuple[pd.DataFrame, pd.DataFrame]] = {}
    for name, interval, period, rule in GRIDS:
        key = (interval, period)
        if key not in raw:
            print(f"downloading {interval} / {period} ...")
            raw[key] = fetch(interval, period)
        close, traded, _vol = raw[key]
        data[name] = returns(close, traded, rule, intraday=interval != "1d")
        print(f"  {name}: {len(data[name][0])} bars")

    close1, traded1, vol1 = raw[("1m", "5d")]
    filled1 = close1.groupby(close1.index.date).ffill()
    sigma1 = filled1.pct_change().std() * 1e4
    quality = pd.DataFrame({
        "bars": traded1.sum(),
        "possible": len(traded1),
        "share": traded1.mean(),
        "med_vol": vol1.median(),
        "price": filled1.iloc[-1],
        "tick_bps": 0.01 / filled1.iloc[-1] * 1e4,
        "sigma_bps": sigma1,
        "tick_over_sigma": (0.01 / filled1.iloc[-1] * 1e4) / sigma1,
    })
    quality.loc[[u for _, u in PAIRS], ["tick_bps", "tick_over_sigma"]] = np.nan
    epps = epps_curve(close1, EPPS_MINUTES)
    png = HERE / "correlation.png"
    draw(epps, data["R1"][0], data["R15"][0], png)
    build(data, epps, quality, png)
    print(f"\nwrote {BOOK}")
    print(epps.round(3).to_string())

    if args.open:
        import xlwings as xw
        xw.Book(str(BOOK)).sheets["Summary"].activate()


if __name__ == "__main__":
    main()
