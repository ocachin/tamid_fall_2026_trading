# ETF Arbitrage Monitor (Excel + xlwings)

```bash
.venv/bin/python etf_arb/tracker.py
```

Refreshes every 20 s until Ctrl+C. Other commands:

```bash
.venv/bin/python etf_arb/tracker.py --once
```
```bash
.venv/bin/python etf_arb/tracker.py holdings
```
```bash
.venv/bin/python etf_arb/tracker.py setup --force
```

Unattended mode refreshes holdings at 8:30 ET, streams 9:25–16:05 ET on weekdays and idles
otherwise. Running it under `caffeinate` keeps the Mac awake while it runs:

```bash
nohup caffeinate -i .venv/bin/python -u etf_arb/tracker.py auto > etf_arb/logs/auto.log 2>&1 &
```

To stop it:

```bash
pkill -f "tracker.py auto"
```

To chart any ETF, pick it in Charts!B1 (it redraws within seconds, in or out of market hours), or run:

```bash
.venv/bin/python etf_arb/tracker.py chart TLT
```

`--once` does a single refresh. `holdings` re-downloads the baskets (do it daily before the open).
`setup --force` rebuilds a clean workbook.

## How it's split

Python is only the data feed. It writes raw numbers to the **Feed** sheet: last, bid, ask, sizes,
prev close, volume, NAV, yield, daily σ, beta, residual mean/SD. **Every number on the Dashboard is
an Excel formula** on Feed, Holdings and Config. You can click any cell to see how it's computed.
Change a cost or edge on Config and the Dashboard updates immediately, without Python.

| Sheet | What |
|---|---|
| Dashboard | Each ETF (bold) with its underlying directly below. Basket ETFs show the basket plus the top holdings; proxy and leveraged ETFs show their underlying (e.g. AAPU → AAPL). |
| Universe | Your tickers (blue = input, dropdowns). |
| Holdings | Basket weights; `Chg %`, contribution and priced weight are formulas. |
| Charts | Pick an ETF in B1 (dropdown): vs underlying, tracking spread, daily-move distribution in σ with realised >1/2/3σ frequencies, volume, 20d σ, intraday premium. |
| Chart Data | The series behind the charts. |
| Config | Costs, min edges, anchor, lookbacks, refresh, pause. Blue cells are inputs. |
| Feed | Raw data written by Python. Don't edit it. |

## Adding a ticker

Type the ticker in column A of **Universe** and leave the rest blank. On the next refresh the
tracker fills in class, model and holdings source (SPDR / Invesco / Yahoo). For leveraged ETFs it
reads the multiple and underlying from the fund name: NVDL → 2x NVDA, AAPU → 2x AAPL. Anything it
fills in you can overwrite with the dropdowns.

| Column | Values |
|---|---|
| Class | EQUITY or FI (sets which cost/edge applies) |
| Model | BASKET (price holdings), PROXY (β × underlying), AUTO |
| Underlying | Hedge for PROXY, e.g. VGLT for TLT, AAPL for AAPU |
| Leverage | Fixed multiple (2, 3, -1 …). Blank = β estimated by regression |
| Holdings Source | SPDR, INVESCO, YAHOO (top 10 only), a URL, or MANUAL (type rows in Holdings) |

## Formulas (Dashboard)

```
Und Move    = β × Chg % of the row below (basket: SUMIFS over Holdings / priced weight)
Fair Value  = prev close (or NAV if Config anchor = NAV) × (1 + Und Move)
Sell Edge   = (bid / FV − 1) × 10000 − cost      → SELL ETF / BUY underlying
Buy Edge    = (1 − ask / FV) × 10000 − cost      → BUY ETF / SELL underlying
Move σ      = Chg % / σ 1D     (σ = daily stdev over the vol lookback)
P>kσ        = probability today's close ends beyond ±kσ from yesterday's close, given the move so
              far and the share of the session left (normal):
              NORM.S.DIST((−k−z)/√left) + 1 − NORM.S.DIST((k−z)/√left)
```

Signal = STALE if the quote is old; EX-DIV? if an equity ETF's T-1 close is more than the
Config "Ex-div flag" (30 bps) away from NAV, which almost always means it went ex-dividend and
Yahoo hasn't caught up; otherwise SELL/BUY when an edge beats the minimum.

Colours: P>kσ white → amber → red, Move σ red ↔ green, Signal red (sell) / green (buy), positive
edges green.

## Data sources

| What | Source | Key |
|---|---|---|
| Real-time bid/ask/sizes (ETFs + displayed underlyings) | Nasdaq public quote API | none |
| All other prices, volume, prev close | Yahoo batch quote (yfinance session), 250 symbols/request | none |
| NAV, yield, daily history | Yahoo / yfinance | none |
| Holdings | SPDR daily xlsx, Invesco API, Yahoo top 10 | none |
| Optional real-time bid/ask for everything | Alpaca (`etf_arb/.env`: `ALPACA_API_KEY`, `ALPACA_SECRET_KEY`) | free |

Nasdaq and Yahoo are unofficial. If Yahoo throttles you, the feed switches to Nasdaq (slower), and
the note in Dashboard!A1 says so. Quotes far from the last trade are treated as stale and blanked.

## Caveats

This is a screen, not an execution system. It ignores creation-unit size, AP fees, borrow, and
bond-ETF cash creations. On an ETF's ex-dividend day the T-1 premium to NAV is distorted. After
the close, the P>kσ columns are 0% or 100% because the day's outcome is known.
