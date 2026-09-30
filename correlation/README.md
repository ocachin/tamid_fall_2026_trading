# Bear 1X ETFs vs underlyings: correlation, beta, theo price

Pairs: AVS/AVGO, AAPD/AAPL, METD/META, MSFD/MSFT, SPDN/SPX, AMZD/AMZN, PLTD/PLTR (Direxion Daily Bear 1X).

```bash
.venv/bin/python correlation/intraday_corr_resolved.py
```
```bash
.venv/bin/python correlation/theo_price.py --replay correlation/cache/bars_2026-09-29.pkl
```

The first builds `Correlation_Resolved.xlsx` from the last 7 sessions (`--sessions 20` for more; Yahoo keeps
about 30 days of 1-minute bars) and saves the completed sessions to `cache/bars_<last session>.pkl`, merged
with any earlier file of that name. Keep those files: they are the only copy once Yahoo drops the bars. The second
builds `Theo_Price.xlsx` from those bars plus Direxion's holdings, NAVs and distributions (saved to
`cache/funds_<date>.pkl`). Add `--open` to open a workbook in Excel. Python writes prices; every return,
correlation, beta, theo price and error statistic in both workbooks is an Excel formula.

## Why the two correlation scripts disagreed

`etf intra script.py` (repo root) and `intraday_corr.py` download identical Yahoo bars. On the same bars,
each reproduces its own published numbers exactly. They differ only in method:

```bash
.venv/bin/python correlation/reconcile.py --replay correlation/cache/bars_2026-09-23.pkl
```

rebuilds both published tables from the 09-15..09-23 bars both scripts used, and splits the gap between
them (mean over pairs; Shapley shares):

| Factor | 1m | 5m | 15m |
|---|---|---|---|
| `etf intra script.py`: per-session ρ of price levels | -0.986 | -0.981 | -0.977 |
| price levels → returns | +0.328 | +0.143 | +0.060 |
| drop untraded minutes → forward fill | +0.068 | +0.011 | +0.001 |
| window 7 → 5 sessions | +0.005 | +0.006 | +0.007 |
| Yahoo bars → 1m resampled label/closed='right' | 0.000 | -0.005 | -0.003 |
| per session → pooled | +0.001 | -0.004 | -0.004 |
| `intraday_corr.py`: pooled ρ of forward-filled returns | -0.584 | -0.828 | -0.916 |

A correlation of price levels is close to -1 for any pair of prices that trend within the day,
unrelated ones included, so it says little about how the ETF tracks. Forward filling a thin ETF's price
(it trades in 30-70% of minutes) scores zero returns against a moving underlying, which drags ρ and β
toward 0.

## What the resolved version does

| Choice | Resolved | Haas script | Original script |
|---|---|---|---|
| Statistic | returns within a session | price levels | returns within a session |
| Untraded minutes | a return runs between minutes in which both legs traded (refresh time) | rows dropped, levels correlated | forward filled (zero returns) |
| 2-60 minute bars | from the 09:30 minute, bins (09:30, 09:35], ... labelled by their last minute, each holding the last minute in which both legs traded | Yahoo bars (09:30 to the first bin's end not measured) | 1m resampled label/closed='right' |
| 60-minute bins | anchored at 09:30; the last bin runs 15:31-15:59 | n/a | anchored at midnight (09:30-10:00 lost) |
| Sessions | each session and all pooled | each session | pooled |
| Daily returns | dividend-adjusted closes | n/a | unadjusted closes (distributions read as moves) |

## Correlation_Resolved.xlsx

| Sheet | Contents |
|---|---|
| Summary | Per pair, all sessions pooled: ρ and β at 1m / 5m / 15m and daily (1 year), R², n, share of minutes the ETF traded |
| By Day | One row per session and pair: ρ, β and n at 1m / 5m / 15m (filterable) |
| 1m / 5m / 15m Matrix | ρ per pair and session (the layout of `New ETF Data.xlsx`) |
| Epps | Pooled ρ and β at 1, 2, 3, 5, 10, 15, 30, 60 minutes, with a chart |
| Data Quality | Minutes traded, median volume per minute, tick in bps, σ per minute, tick/σ, share of zero returns |
| R1 ... R60 | Per pair: both legs' prices at the last minute in each bin in which both traded (the row time is the bin's last minute), a both-traded flag, and the two returns |
| RD | Dividend-adjusted daily closes and returns |

## Theo_Price.xlsx

A -1x daily fund resets its exposure at every close, so during session d the ETF should be worth:

```
R_U        = (underlying now + its dividend if ex today) / underlying close(d-1) - 1
Theo (β)   = ETF close(d-1) × (1 + β × R_U + carry) - distribution if the fund is ex today
Theo (NAV) = NAV(d-1)       × (1 -     R_U + carry) - distribution
```

| Sheet | Contents |
|---|---|
| Summary | Per pair: tick in bps, and the mean, SD and RMSE (bps) of the ETF's price against each theo |
| Beta | Contract (-1), daily β over the year before the window with its SE, and β of the ETF's move since the prior close on the underlying's |
| By Day | The same statistics per session, with the distribution paid that day |
| Today | The next session's anchors; type an underlying price (blue) to get the theo |
| Holdings | Direxion's holdings files; NAV, swap exposure and money-market share worked out from them |
| NAV | Direxion's NAVs next to the official closes |
| Sessions | Each session's prior closes, distribution, underlying dividend and NAV |
| Theo 1m | Every minute: prices, returns since the prior close, the three theos and the premiums |
| Config | Carry (bps per calendar day) and exposure (-1) |

Each fund holds total-return swaps on its underlying equal to -100% of net assets (reset at every
close) plus government money-market funds. `direxion.py` reads the daily holdings file
(`direxion.com/holdings/<TICKER>.csv`, published around 10:00 ET for the prior close) and saves a copy to
`cache/direxion/`. Only today's file is online, so the exact NAVs build up from those copies. Before that,
the theo uses Direxion's 30 published NAVs, which are rounded to the cent (10-20 bps on these prices).

## Earlier version

`intraday_corr.py` builds `Correlation_Intraday.xlsx` with the original method (forward-filled returns,
pooled over 5 sessions, plus 60-day 5m/15m sheets and a session picker). It is kept for reference.
