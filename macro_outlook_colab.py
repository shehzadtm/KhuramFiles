"""Macro dashboard builder for Google Colab.

Run this file in Colab (or paste it into a notebook cell). It downloads daily
September prices from Yahoo Finance and creates Macro_Outlook_September_2026.xlsx.
"""

import sys
import time
import subprocess
from calendar import monthrange
from datetime import date


def ensure_packages():
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "yfinance", "xlsxwriter"])


ensure_packages()

import numpy as np
import pandas as pd
import yfinance as yf

# ---- User settings ----------------------------------------------------------
YEAR = 2026
MONTH = 9
OUTPUT_FILE = f"Macro_Outlook_{date(YEAR, MONTH, 1):%B_%Y}.xlsx"
PAUSE_BETWEEN_REQUESTS = 1.25  # gentler on Yahoo Finance than bulk downloads
MAX_RETRIES = 4

TICKERS = {
    "SPY": "S&P 500", "RSP": "S&P 500 Equal Weight", "IWM": "US Small Caps",
    "^VIX": "CBOE VIX", "^VVIX": "VVIX", "GLD": "Gold", "HYG": "High Yield",
    "LQD": "Investment Grade Credit", "TLT": "20+ Year Treasury", "IEF": "7-10 Year Treasury",
    "UUP": "US Dollar", "EFA": "Developed ex-US", "EEM": "Emerging Markets",
    "KRE": "Regional Banks", "XLY": "Consumer Discretionary", "XLP": "Consumer Staples",
    "SOXX": "Semiconductors", "USO": "Oil", "^TNX": "US 10Y Yield",
}


def download_one(ticker, start, end):
    """Download one series with backoff; avoids large parallel Yahoo requests."""
    for attempt in range(MAX_RETRIES):
        try:
            frame = yf.Ticker(ticker).history(start=start, end=end, auto_adjust=True, raise_errors=True)
            if not frame.empty and "Close" in frame:
                s = frame["Close"].copy()
                s.index = pd.to_datetime(s.index).tz_localize(None)
                return s.rename(ticker)
        except Exception as exc:
            print(f"{ticker}: attempt {attempt + 1}/{MAX_RETRIES} failed ({exc})")
            time.sleep(4 * (attempt + 1))
        time.sleep(PAUSE_BETWEEN_REQUESTS)
    print(f"Skipping {ticker}: no usable data returned.")
    return None


start = date(YEAR, MONTH, 1)
end = date(YEAR, MONTH, monthrange(YEAR, MONTH)[1]) + pd.Timedelta(days=1)
print(f"Downloading Yahoo Finance data from {start} to {end - pd.Timedelta(days=1)} …")

series = []
for ticker in TICKERS:
    item = download_one(ticker, start.isoformat(), end.isoformat())
    if item is not None:
        series.append(item)

if not series:
    raise RuntimeError("Yahoo Finance returned no data. Try again later; Yahoo may be rate-limiting requests.")

prices = pd.concat(series, axis=1).sort_index()


def metric(ticker):
    s = prices[ticker].dropna()
    if len(s) < 2:
        return None
    ret = s.iloc[-1] / s.iloc[0] - 1
    dd = (s / s.cummax() - 1).min()
    return {"start": s.iloc[0], "last": s.iloc[-1], "return": ret, "drawdown": dd, "obs": len(s)}


stats = {t: metric(t) for t in prices.columns}


def ratio_change(numerator, denominator):
    x = prices[[numerator, denominator]].dropna()
    if len(x) < 2:
        return np.nan
    ratio = x[numerator] / x[denominator]
    return ratio.iloc[-1] / ratio.iloc[0] - 1


def last_value(ticker):
    return stats[ticker]["last"] if ticker in stats and stats[ticker] else np.nan


def monthly_change(ticker):
    return stats[ticker]["return"] if ticker in stats and stats[ticker] else np.nan


signals = [
    ("Equity breadth", "SPY/RSP", ratio_change("SPY", "RSP"), "Positive favors mega-cap leadership; negative favors broader participation."),
    ("Small-cap participation", "IWM/SPY", ratio_change("IWM", "SPY"), "Positive suggests more cyclical risk appetite."),
    ("Credit risk appetite", "HYG/LQD", ratio_change("HYG", "LQD"), "Falling indicates high-yield credit is lagging investment grade."),
    ("Consumer cyclicality", "XLY/XLP", ratio_change("XLY", "XLP"), "Falling indicates defensive consumer leadership."),
    ("Regional-bank stress", "KRE/SPY", ratio_change("KRE", "SPY"), "Falling can flag domestic-credit or funding concerns."),
    ("Semiconductor leadership", "SOXX/SPY", ratio_change("SOXX", "SPY"), "Falling indicates technology leadership is fading."),
    ("Dollar regime", "UUP", monthly_change("UUP"), "A stronger dollar can tighten global financial conditions."),
    ("Inflation / geopolitics", "GLD", monthly_change("GLD"), "Rising gold can reflect falling real yields, inflation concern, or risk aversion."),
    ("Oil impulse", "USO", monthly_change("USO"), "A sharp rise can add to inflation pressure."),
    ("Long-duration rates", "TLT", monthly_change("TLT"), "Falling TLT generally means rising long yields."),
    ("10Y yield", "^TNX", monthly_change("^TNX"), "^TNX is quoted in percent; rising yields pressure long-duration assets."),
    ("Volatility", "^VIX", monthly_change("^VIX"), "Rising VIX implies higher expected near-term equity volatility."),
]


def regime(value, positive_is_risk_on=True, neutral_band=0.01):
    if pd.isna(value):
        return "No data"
    if value > neutral_band:
        return "Risk-on" if positive_is_risk_on else "Caution"
    if value < -neutral_band:
        return "Caution" if positive_is_risk_on else "Risk-on"
    return "Neutral"


# ^VIX and ^TNX rising are intentionally treated as caution.
signal_rows = []
for category, measure, change, reading in signals:
    inverse = measure in {"UUP", "USO", "^TNX", "^VIX"}
    signal_rows.append([category, measure, change, regime(change, not inverse), reading])

summary_rows = []
for ticker in prices.columns:
    m = stats[ticker]
    if m:
        summary_rows.append([ticker, TICKERS[ticker], m["start"], m["last"], m["return"], m["drawdown"], m["obs"]])
summary = pd.DataFrame(summary_rows, columns=["Ticker", "Asset / Signal", "Sep Start", "Last Close", "Sep Return", "Max Drawdown", "Observations"])
summary = summary.sort_values("Sep Return", ascending=False)

# ---- Excel workbook ---------------------------------------------------------
with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
    workbook = writer.book
    dashboard = workbook.add_worksheet("Dashboard")
    analysis = workbook.add_worksheet("Analysis")
    prices_ws = workbook.add_worksheet("Prices")
    method = workbook.add_worksheet("Methodology")
    writer.sheets.update({"Dashboard": dashboard, "Analysis": analysis, "Prices": prices_ws, "Methodology": method})

    navy = "#172033"; blue = "#2F75B5"; light_blue = "#D9EAF7"; green = "#198754"; red = "#C62828"; amber = "#C98700"; grey = "#6B7280"
    title = workbook.add_format({"bold": True, "font_size": 20, "font_color": "#FFFFFF", "bg_color": navy, "align": "left", "valign": "vcenter"})
    subtitle = workbook.add_format({"font_color": "#DCE6F1", "bg_color": navy, "italic": True})
    section = workbook.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": blue, "align": "left"})
    header = workbook.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#305496", "border": 0})
    label = workbook.add_format({"bold": True, "font_color": "#334155"})
    note = workbook.add_format({"font_color": grey, "italic": True, "text_wrap": True, "valign": "top"})
    currency = workbook.add_format({"num_format": "$#,##0.00;[Red]($#,##0.00)", "font_color": "#1F2937"})
    number = workbook.add_format({"num_format": "0.00", "font_color": "#1F2937"})
    pct = workbook.add_format({"num_format": "0.0%;[Red]-0.0%"})
    pct_green = workbook.add_format({"num_format": "0.0%", "font_color": green, "bold": True})
    pct_red = workbook.add_format({"num_format": "0.0%", "font_color": red, "bold": True})
    card_label = workbook.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": blue, "align": "center"})
    card_value = workbook.add_format({"bold": True, "font_size": 16, "font_color": navy, "bg_color": light_blue, "align": "center", "num_format": "0.0%;[Red]-0.0%"})
    caution_fmt = workbook.add_format({"font_color": red, "bold": True})
    riskon_fmt = workbook.add_format({"font_color": green, "bold": True})
    neutral_fmt = workbook.add_format({"font_color": amber, "bold": True})

    # Dashboard
    dashboard.set_tab_color(blue)
    dashboard.set_column("A:A", 23); dashboard.set_column("B:B", 16); dashboard.set_column("C:C", 16); dashboard.set_column("D:D", 16); dashboard.set_column("E:E", 23); dashboard.set_column("F:F", 46)
    dashboard.set_row(0, 30)
    dashboard.merge_range("A1:F1", "Macro Outlook Dashboard", title)
    dashboard.merge_range("A2:F2", f"Yahoo Finance data | {prices.index.min():%d %b %Y} – {prices.index.max():%d %b %Y} | Generated {date.today():%d %b %Y}", subtitle)
    dashboard.merge_range("A4:F4", "September market snapshot", section)

    card_metrics = [("S&P 500", monthly_change("SPY")), ("High Yield / IG", ratio_change("HYG", "LQD")), ("Equity Breadth", ratio_change("SPY", "RSP")), ("VIX Change", monthly_change("^VIX"))]
    for idx, (name, val) in enumerate(card_metrics):
        col = idx * 1 + 0
        dashboard.write(4, col, name, card_label)
        dashboard.write(5, col, val, card_value)
    dashboard.merge_range("A8:F8", "Macro signal board", section)
    dashboard.write_row("A9", ["Category", "Measure", "September change", "Regime", "How to read it", "Portfolio implication"], header)
    implications = {
        "Risk-on": "Keep risk controls; no signal-driven action alone.",
        "Neutral": "Monitor for confirmation across other indicators.",
        "Caution": "Review concentration, liquidity, and hedge sizing.",
        "No data": "Refresh when Yahoo data becomes available.",
    }
    for row, (category, measure, change, tag, reading) in enumerate(signal_rows, start=9):
        dashboard.write(row, 0, category)
        dashboard.write(row, 1, measure)
        dashboard.write(row, 2, change, pct)
        fmt = riskon_fmt if tag == "Risk-on" else caution_fmt if tag == "Caution" else neutral_fmt
        dashboard.write(row, 3, tag, fmt)
        dashboard.write(row, 4, reading, note)
        dashboard.write(row, 5, implications[tag], note)
    dashboard.conditional_format(9, 2, 9 + len(signal_rows) - 1, 2, {"type": "3_color_scale", "min_color": "#F8696B", "mid_color": "#FFEB84", "max_color": "#63BE7B"})
    dashboard.merge_range(23, 0, 24, 5, "This workbook is an informational market-monitoring tool, not investment advice. Use the signals as a cross-check with your allocation, time horizon, taxes, liquidity needs, and risk tolerance.", note)
    dashboard.freeze_panes(9, 0)

    # Analysis
    analysis.set_tab_color(green)
    analysis.set_column("A:A", 12); analysis.set_column("B:B", 28); analysis.set_column("C:G", 15); analysis.set_column("H:H", 14)
    analysis.merge_range("A1:H1", "September Performance & Risk Analysis", title)
    analysis.merge_range("A2:H2", "Returns use adjusted daily close where Yahoo Finance supplies it. Max drawdown is measured within the selected month.", subtitle)
    analysis.write_row("A4", list(summary.columns), header)
    for r, values in enumerate(summary.itertuples(index=False), start=4):
        analysis.write(r, 0, values[0]); analysis.write(r, 1, values[1]); analysis.write(r, 2, values[2], currency if values[0] != "^TNX" else number); analysis.write(r, 3, values[3], currency if values[0] != "^TNX" else number)
        analysis.write(r, 4, values[4], pct); analysis.write(r, 5, values[5], pct); analysis.write(r, 6, values[6], number)
    analysis.conditional_format(4, 4, 3 + len(summary), 4, {"type": "3_color_scale", "min_color": "#F8696B", "mid_color": "#FFEB84", "max_color": "#63BE7B"})
    analysis.conditional_format(4, 5, 3 + len(summary), 5, {"type": "3_color_scale", "min_color": "#F8696B", "mid_color": "#FFEB84", "max_color": "#63BE7B"})
    analysis.autofilter(3, 0, 3 + len(summary), len(summary.columns) - 1)
    analysis.freeze_panes(4, 2)

    # Prices
    prices_ws.set_tab_color("#808080")
    prices_ws.write(0, 0, "Date", header)
    for c, t in enumerate(prices.columns, start=1):
        prices_ws.write(0, c, t, header)
    for r, (dt, row) in enumerate(prices.iterrows(), start=1):
        prices_ws.write_datetime(r, 0, dt.to_pydatetime(), workbook.add_format({"num_format": "yyyy-mm-dd"}))
        for c, val in enumerate(row, start=1):
            if pd.notna(val): prices_ws.write_number(r, c, float(val), number)
    prices_ws.set_column(0, 0, 13); prices_ws.set_column(1, len(prices.columns), 13); prices_ws.freeze_panes(1, 1)
    prices_ws.autofilter(0, 0, len(prices), len(prices.columns))

    # Methodology
    method.set_tab_color("#808080")
    method.set_column("A:A", 26); method.set_column("B:B", 105)
    method.merge_range("A1:B1", "Methodology & Refresh Notes", title)
    method.write("A3", "Data source", header); method.write("B3", "Yahoo Finance, downloaded through the yfinance Python library. Yahoo data can be delayed, corrected, incomplete, or temporarily rate-limited.", note)
    method.write("A4", "Period", header); method.write("B4", f"{prices.index.min():%d %B %Y} through {prices.index.max():%d %B %Y}.", note)
    method.write("A5", "Return", header); method.write("B5", "Last adjusted close divided by first adjusted close, minus one.", note)
    method.write("A6", "Max drawdown", header); method.write("B6", "Largest peak-to-trough decline using daily adjusted closing prices inside the selected month.", note)
    method.write("A7", "Ratios", header); method.write("B7", "Ratio changes compare the month-end ratio with the first available ratio of the month. They are directional signals, not forecasts.", note)
    method.write("A8", "Refresh", header); method.write("B8", "Change YEAR and MONTH near the top of the script, rerun in Colab, and download the newly generated workbook.", note)
    method.write("A10", "Ticker", header); method.write("B10", "Description", header)
    for r, (ticker, desc) in enumerate(TICKERS.items(), start=10):
        method.write(r, 0, ticker); method.write(r, 1, desc)

    # Chart of normalised major-market series on Dashboard.
    major = [t for t in ["SPY", "RSP", "IWM", "HYG", "TLT", "GLD", "UUP"] if t in prices.columns]
    if major:
        norm_start = len(prices.columns) + 3
        prices_ws.write(0, norm_start, "Date", header)
        for c, t in enumerate(major, start=norm_start + 1): prices_ws.write(0, c, t, header)
        normalised = prices[major].apply(lambda col: col / col.dropna().iloc[0] * 100 if not col.dropna().empty else col)
        for r, (dt, row) in enumerate(normalised.iterrows(), start=1):
            prices_ws.write_datetime(r, norm_start, dt.to_pydatetime(), workbook.add_format({"num_format": "yyyy-mm-dd"}))
            for c, val in enumerate(row, start=norm_start + 1):
                if pd.notna(val): prices_ws.write_number(r, c, float(val), number)
        chart = workbook.add_chart({"type": "line"})
        for c, t in enumerate(major, start=norm_start + 1):
            chart.add_series({"name": ["Prices", 0, c], "categories": ["Prices", 1, norm_start, len(prices), norm_start], "values": ["Prices", 1, c, len(prices), c]})
        chart.set_title({"name": "Selected assets (rebased to 100)"}); chart.set_y_axis({"name": "Index (first day = 100)"}); chart.set_legend({"position": "bottom"}); chart.set_size({"width": 760, "height": 330})
        dashboard.insert_chart("A27", chart)

print(f"Created {OUTPUT_FILE}")
try:
    from google.colab import files
    files.download(OUTPUT_FILE)
except ImportError:
    pass
