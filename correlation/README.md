# Intraday correlation: Bear 1X ETFs vs underlyings

```bash
.venv/bin/python correlation/intraday_corr.py
```

Rebuilds `Correlation_Intraday.xlsx` (add `--open` to open it in Excel). Pairs: AVS/AVGO,
AAPD/AAPL, METD/META, MSFD/MSFT, SPDN/^GSPC, AMZD/AMZN, PLTD/PLTR.

Python downloads prices and writes returns; every correlation, beta and R² in the workbook is an
Excel formula (CORREL / SLOPE / RSQ) over those columns.

| Sheet | Contents |
|---|---|
| Summary | One row per pair: ρ at 1m / 5m / 15m (last 5 sessions), ρ at 5m / 15m over 60 days, ρ daily over 1 year, β and R², sample sizes, and the share of minutes the ETF actually traded |
| Chart | Correlation vs sampling interval (1–60 min), 1m vs 15m bars, and a 15-minute scatter |
| Data Quality | Bars present, % of minutes traded, median volume, price, tick size in bps, 1-minute σ, and tick/σ |
| By Day | Correlation per session (5-minute returns), to show stability |
| Daily | Pick one session (dropdown in B3): ρ per pair at 1, 2, 3, 5, 10, 15, 30 and 60 min using only that day's bars, plus a chart. B4 sets the last bar included (e.g. 15:45 drops the close) |
| R1, R2, R3, R5, R10, R15, R30, R60 | Return columns from 1-minute data (last 5 sessions), resampled. Per pair: ETF, underlying, both-traded flag, and traded-only copies |
| R5_60d, R15_60d, RD | 5-minute and 15-minute returns over 60 days, and daily returns over 1 year |

Windows are what Yahoo allows: 1-minute data goes back 7 days, 5/15-minute 60 days, daily 1 year.

Two microstructure effects drive the intraday numbers, both measured on Data Quality:
non-trading (the ETFs print in 27–68% of minutes, the underlyings in ~100%) and price
discreteness (a 1-cent tick is 8–21 bps on an $5–13 ETF).
