# TAMID: Lognormal Return Distributions

```
StockDatabase.xlsx      data source (Data!O1 = ticker, Data!O4 = DAILY / WEEKLY / MONTHLY)
xlwings_lite/main.py    the analysis + tab layout (pasted into xlwings Lite, also used by run_analysis.py)
run_analysis.py         VS Code: one new workbook per ticker in output/
output/                 Analysis_<TICKER>_<TIMEFRAME>_<date>.xlsx
docs/                   guide (PDF) and technical report (Word)
requirements.txt, .venv Python for run_analysis.py
```

Both ways read the prices from StockDatabase's Data sheet and build the exact same tab.

## A) VS Code: separate workbook per ticker
Keep StockDatabase.xlsx open, then in the VS Code terminal:

```bash
.venv/bin/python run_analysis.py META ORCL
```

Creates `output/Analysis_META_DAILY_<date>.xlsx` and `output/Analysis_ORCL_DAILY_<date>.xlsx`, one tab each.
Add `-t WEEKLY` or `-t MONTHLY` for other timeframes. The script sets O1/O4 itself, waits for the
refresh, reads the sheet, and puts O1/O4 back when done.

## B) Inside Excel: xlwings Lite
1. Data sheet: set O1 (ticker) and O4 (timeframe); wait for the table to refresh.
2. xlwings Lite pane → run **lognormal_current_ticker**.
3. A tab `<TICKER>_<TIMEFRAME>_<date>` appears inside StockDatabase (xlwings Lite can't create separate files).

If you edit `xlwings_lite/main.py`, paste it into the xlwings Lite pane again so both stay identical.
