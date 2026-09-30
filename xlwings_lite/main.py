import datetime as dt

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import xlwings as xw
from matplotlib.ticker import PercentFormatter
from scipy import stats as st
from xlwings import script

PERIODS_PER_YEAR = {"DAILY": 252, "WEEKLY": 52, "MONTHLY": 12}
NAVY, BLUE = "#1F3A5F", "#2f6db5"


def read_data_sheet(book):
    sh = book.sheets["Data"]
    ticker = str(sh["O1"].value).strip().upper()
    timeframe = str(sh["O4"].value).strip().upper()
    raw = sh["A1"].expand("down").resize(column_size=6).options(pd.DataFrame, index=False, header=True).value
    return ticker, timeframe, raw


def prepare(raw):
    df = raw.copy()
    df.columns = [str(c).strip().title() for c in df.columns]
    if pd.api.types.is_numeric_dtype(df["Date"]):
        df["Date"] = pd.to_datetime(df["Date"], unit="D", origin="1899-12-30")
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Close"] = pd.to_numeric(df["Close"], errors="coerce")
    df = df.dropna(subset=["Date", "Close"])
    df = df[df["Close"] > 0].sort_values("Date").drop_duplicates("Date").reset_index(drop=True)
    df["LogReturn"] = np.log(df["Close"] / df["Close"].shift(1))
    df["Z"] = (df["LogReturn"] - df["LogReturn"].mean()) / df["LogReturn"].std()
    return df


def compute(df, ticker, timeframe):
    r = df["LogReturn"].dropna()
    ppy = PERIODS_PER_YEAR.get(timeframe, 252)
    mu, sd = float(r.mean()), float(r.std(ddof=1))
    skew, kurt = float(st.skew(r, bias=False)), float(st.kurtosis(r, bias=False))
    n, p0 = int(len(r)), float(df["Close"].iloc[-1])
    summary = [
        ("Ticker", ticker, None),
        ("Timeframe", timeframe, None),
        ("First date", df["Date"].iloc[0].to_pydatetime(), "m/d/yyyy"),
        ("Last date", df["Date"].iloc[-1].to_pydatetime(), "m/d/yyyy"),
        ("Return observations", n, "#,##0"),
        ("Last close", p0, "$#,##0.00"),
        ("Mean log return (μ)", mu, "0.000%"),
        ("Std. deviation (σ)", sd, "0.00%"),
        ("Annualized drift", mu * ppy, "0.00%"),
        ("Annualized volatility", sd * np.sqrt(ppy), "0.00%"),
        ("Skewness", skew, "0.000"),
        ("Excess kurtosis", kurt, "0.000"),
        ("Jarque-Bera p-value", float(st.chi2.sf(n / 6 * (skew**2 + kurt**2 / 4), 2)), "0.0000"),
    ]
    bounds = []
    for k in (1, 2, 3):
        lo, hi = mu - k * sd, mu + k * sd
        bounds.append([f"±{k}σ", float(2 * st.norm.cdf(k) - 1), float(np.exp(lo) - 1), float(np.exp(hi) - 1),
                       float(p0 * np.exp(lo)), float(p0 * np.exp(hi)), float(((r >= lo) & (r <= hi)).mean())])
    return {"mu": mu, "sd": sd, "summary": summary, "bounds": bounds}


def hist_fig(df, mu, sd, ticker, timeframe):
    r = df["LogReturn"].dropna()
    fig, ax = plt.subplots(figsize=(7.5, 4))
    lo, hi = mu - 4.5 * sd, mu + 4.5 * sd
    ax.hist(r.clip(lo, hi), bins=np.linspace(lo, hi, 60), density=True, color=BLUE, alpha=0.55, edgecolor="white")
    x = np.linspace(lo, hi, 300)
    ax.plot(x, st.norm.pdf(x, mu, sd), color=NAVY, lw=1.6, label="Normal fit")
    for k, c in zip((1, 2, 3), ("#2e9e5b", "#e0a800", "#d64545")):
        ax.axvline(mu - k * sd, color=c, ls="--", lw=1, label=f"±{k}σ")
        ax.axvline(mu + k * sd, color=c, ls="--", lw=1)
    ax.xaxis.set_major_formatter(PercentFormatter(1.0, decimals=0))
    ax.set_title(f"{ticker} {timeframe.lower()} log returns (σ = {sd:.2%})", loc="left")
    ax.legend(frameon=False, fontsize=8)
    fig.tight_layout()
    return fig


def price_fig(df, ticker):
    fig, ax = plt.subplots(figsize=(7.5, 3))
    ax.plot(df["Date"], df["Close"], color=BLUE, lw=1.1)
    ax.set_title(f"{ticker} close price", loc="left")
    fig.tight_layout()
    return fig


def write_analysis_sheet(book, ticker, timeframe, raw):
    df = prepare(raw)
    res = compute(df, ticker, timeframe)

    run_date = f"{dt.date.today():%Y-%m-%d}"
    name = f"{ticker}_{timeframe}_{run_date}"[:31]
    if name in [s.name for s in book.sheets]:
        book.sheets[name].delete()
    sh = book.sheets.add(name)

    sh["A1"].value = f"{ticker}: lognormal return distribution ({timeframe.lower()})"
    sh["A1"].font.bold = True
    sh["A1"].font.size = 14
    sh["A2"].value = f"Generated {dt.datetime.now():%Y-%m-%d %H:%M} from the Data sheet"
    sh["A3"].value = f"File name: Analysis_{ticker}_{timeframe}_{run_date}.xlsx"

    sh["A4"].value = [["Statistic", "Value"]] + [[label, value] for label, value, _ in res["summary"]]
    sh["A4:B4"].font.bold = True
    for i, (_, _, fmt) in enumerate(res["summary"]):
        if fmt:
            sh[f"B{5 + i}"].number_format = fmt

    sh["A19"].value = [["Band", "Normal confidence", "Lower % move", "Upper % move", "Lower price",
                        "Upper price", "Empirical coverage"]] + res["bounds"]
    sh["A19:G19"].font.bold = True
    sh["B20:D22"].number_format = "0.00%"
    sh["E20:F22"].number_format = "$#,##0.00"
    sh["G20:G22"].number_format = "0.00%"

    rows = [["Date", "Close", "Log Return", "Z-Score"]] + [
        [d.to_pydatetime(), float(c), None if pd.isna(lr) else float(lr), None if pd.isna(z) else float(z)]
        for d, c, lr, z in df[["Date", "Close", "LogReturn", "Z"]].itertuples(index=False)]
    sh["I1"].value = rows
    last = len(rows)
    sh["I1:L1"].font.bold = True
    sh[f"I2:I{last}"].number_format = "m/d/yy"
    sh[f"J2:J{last}"].number_format = "$#,##0.00"
    sh[f"K2:K{last}"].number_format = "0.00%"
    sh[f"L2:L{last}"].number_format = "0.00"

    for fig, pic, cell in ((hist_fig(df, res["mu"], res["sd"], ticker, timeframe), "Distribution", "A25"),
                           (price_fig(df, ticker), "Price", "A50")):
        sh.pictures.add(fig, name=pic, update=True, anchor=sh[cell])
        plt.close(fig)
    return sh


@script
def lognormal_current_ticker(book: xw.Book):
    ticker, timeframe, raw = read_data_sheet(book)
    sh = write_analysis_sheet(book, ticker, timeframe, raw)
    sh.activate()
