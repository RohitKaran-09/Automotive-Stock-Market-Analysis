"""Task 2: Fundamental and Technical Analysis — Mahindra & Mahindra Ltd.

Downloads actual daily market data, calculates common technical indicators, and
generates a referenced Word report plus reproducible CSV outputs. No synthetic
market prices are created. Fundamental FY26/FY25 data are transcribed from M&M's
official FY2025-26 annual report and audited results release.
"""
from datetime import datetime
from pathlib import Path
import json
import warnings

import numpy as np
import pandas as pd
import yfinance as yf
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "outputs"
CHARTS = ROOT / "visualizations"
REPORT = ROOT / "Fundamental_and_Technical_Analysis_Mahindra_Mahindra.docx"
SUBMISSION_DOCX = ROOT / "Submission_Description.docx"
SUBMISSION_TXT = ROOT / "submission_description.txt"

TICKERS = {
    "M&M.NS": "Mahindra & Mahindra",
    "MARUTI.NS": "Maruti Suzuki",
    "^NSEI": "Nifty 50",
}
HISTORY_PERIOD = "5y"
TRADING_DAYS = 252

# Consolidated financial figures (₹ crore unless noted) from M&M FY26 Integrated Annual Report.
FINANCIALS = {
    "FY26": {
        "income_from_operations": 198638.55,
        "revenue_from_operations": 197792.78,
        "other_income": 3445.79,
        "total_income": 202084.34,
        "ebitda_like": 40425.66,
        "depreciation_amortisation": 7322.02,
        "profit_before_finance_share_exceptional_tax": 33103.64,
        "finance_costs": 9590.85,
        "profit_before_tax": 25184.76,
        "profit_for_year": 18621.71,
        "pat_attributable_to_owners": 17098.85,
        "roe_pct_company_reported": 20.1,
        "eps_rupees_company_reported": 152.2,
        "cfo": 11657.36,
        "capex_ppe_intangibles": 9605.45,
        "cash_end": 4292.24,
        "total_assets": 317634.64,
        "total_current_assets": 147139.62,
        "total_current_liabilities": 116098.56,
        "total_equity": 109491.00,
        "owner_equity": 93096.35,
        "noncurrent_borrowings": 73469.93,
        "current_borrowings": 56144.57,
        "operating_margin_pct_company_reported": None,
    },
    "FY25": {
        "income_from_operations": 159210.82,
        "revenue_from_operations": 158749.75,
        "other_income": 2181.05,
        "total_income": 161391.87,
        "ebitda_like": 32699.24,
        "depreciation_amortisation": 6073.65,
        "profit_before_finance_share_exceptional_tax": 26625.59,
        "finance_costs": 9083.39,
        "profit_before_tax": 19079.62,
        "profit_for_year": 14073.17,
        "pat_attributable_to_owners": 12929.10,
        "roe_pct_company_reported": 16.4,
        "eps_rupees_company_reported": None,
        "cfo": 3175.81,
        "capex_ppe_intangibles": 10391.95,
        "cash_end": 4923.89,
        "total_assets": 277586.27,
        "total_current_assets": 126573.27,
        "total_current_liabilities": 90322.01,
        "total_equity": 89098.00,
        "owner_equity": 77038.86,
        "noncurrent_borrowings": 79707.19,
        "current_borrowings": 41465.76,
        "operating_margin_pct_company_reported": None,
    }
}

SOURCES = [
    {
        "id": "F1",
        "title": "Mahindra & Mahindra Ltd., Integrated Annual Report 2025–26, Board's Report financial highlights (page 98)",
        "url": "https://www.mahindra.com/annual-report-FY2026/98/",
        "use": "Audited FY26/FY25 consolidated results, income, profit, tax and attributable earnings."
    },
    {
        "id": "F2",
        "title": "Mahindra & Mahindra Ltd., Integrated Annual Report 2025–26, Consolidated Statement of Cash Flows (page 397)",
        "url": "https://www.mahindra.com/annual-report-FY2026/397/",
        "use": "Operating cash flow, capex on property, plant, equipment and intangible assets, and investing cash flows."
    },
    {
        "id": "F3",
        "title": "Mahindra & Mahindra, M&M Results Q4 F26 and FY26, 5 May 2026",
        "url": "https://www.mahindra.com/news-room/press-release/en/m-and-m-results-q4-f26-and-fy26",
        "use": "Reported FY26 consolidated revenue ₹198,639 crore, PAT ₹17,099 crore, ROE 20.1%, EPS ₹152.2, SUV and LCV market-share highlights."
    },
    {
        "id": "F4",
        "title": "NSE India, Mahindra & Mahindra Limited quote and historical stock information",
        "url": "https://www.nseindia.com/get-quote/equity/M%26M/Mahindra-%26-Mahindra-Limited",
        "use": "Exchange reference for market price, 52-week range and valuation snapshot; values change over time."
    },
    {
        "id": "F5",
        "title": "Mahindra & Mahindra FY2025–26 Integrated Annual Report",
        "url": "https://www.mahindra.com/annual-report-FY2026/index.html",
        "use": "Management narrative, sector overview, market share, segment and strategic context."
    },
    {
        "id": "F6",
        "title": "Mahindra & Mahindra audited FY26 consolidated balance sheet, in Q4 FY26 results PDF",
        "url": "https://www.mahindra.com/sites/default/files/2026-05/Outcome%20of%20Board%20Meeting%20and%20Financial%20Results%20-%205th%20May%202026.pdf",
        "use": "Total assets, equity, current assets/liabilities and current/non-current borrowings at 31 March 2026 and 31 March 2025."
    },
    {
        "id": "T1",
        "title": "Yahoo Finance public historical market data interface (accessed at run time)",
        "url": "https://finance.yahoo.com/",
        "use": "Daily OHLCV data downloaded through yfinance; the script saves exact coverage metadata. Technical indicator results depend on the data source and observation date."
    },
    {
        "id": "P1",
        "title": "Maruti Suzuki India Ltd., annual report and financial disclosures",
        "url": "https://www.marutisuzuki.com/corporate/investors",
        "use": "Peer company disclosures; peer price history is used only for market-performance context here."
    },
]

def normalize_yf_frame(frame, ticker):
    if frame is None or frame.empty:
        return pd.DataFrame()
    df = frame.copy()
    if isinstance(df.columns, pd.MultiIndex):
        if ticker in df.columns.get_level_values(0):
            df = df[ticker].copy()
        elif ticker in df.columns.get_level_values(1):
            df = df.xs(ticker, axis=1, level=1).copy()
        elif "Close" in df.columns.get_level_values(0):
            df.columns = df.columns.get_level_values(0)
        elif "Close" in df.columns.get_level_values(1):
            df.columns = df.columns.get_level_values(1)
    return df

def download_data():
    raw = yf.download(
        list(TICKERS), period=HISTORY_PERIOD, interval="1d",
        auto_adjust=False, group_by="ticker", progress=False, threads=True, timeout=45
    )
    frames = {}
    diagnostics = {}
    for ticker, name in TICKERS.items():
        df = normalize_yf_frame(raw, ticker)
        if df.empty or "Close" not in df or df["Close"].dropna().empty:
            try:
                df = yf.Ticker(ticker).history(period=HISTORY_PERIOD, interval="1d", auto_adjust=False, actions=False)
                df = normalize_yf_frame(df, ticker)
            except Exception as exc:
                diagnostics[ticker] = {"error": str(exc)}
                continue
        if df.empty or "Close" not in df:
            diagnostics[ticker] = {"shape": list(df.shape), "columns": [str(x) for x in df.columns]}
            continue
        df.index = pd.to_datetime(df.index)
        if getattr(df.index, "tz", None) is not None:
            df.index = df.index.tz_localize(None)
        df = df[~df.index.duplicated(keep="last")].sort_index()
        df = df.loc[:, ~df.columns.duplicated()]
        df["Close"] = pd.to_numeric(df["Close"], errors="coerce")
        if "Adj Close" in df:
            df["Adj Close"] = pd.to_numeric(df["Adj Close"], errors="coerce")
        df = df.dropna(subset=["Close"])
        if len(df) >= 100:
            df["Ticker"] = ticker
            df["Name"] = name
            frames[ticker] = df
            diagnostics[ticker] = {
                "name": name,
                "rows": int(len(df)),
                "start": df.index.min().strftime("%Y-%m-%d"),
                "end": df.index.max().strftime("%Y-%m-%d"),
                "latest_close": float(df["Close"].iloc[-1]),
            }
    if "M&M.NS" not in frames:
        raise RuntimeError("Failed to retrieve adequate actual historical price/volume data for M&M.NS; no synthetic prices used. Diagnostics: " + json.dumps(diagnostics))
    if "^NSEI" not in frames:
        warnings.warn("Nifty 50 was unavailable; technical analysis will use M&M only.")
    return frames, diagnostics

def technical_indicators(df):
    d = df.copy()
    close = d["Close"].astype(float)
    d["SMA20"] = close.rolling(20).mean()
    d["SMA50"] = close.rolling(50).mean()
    d["SMA200"] = close.rolling(200).mean()
    delta = close.diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)
    avg_gain = gain.ewm(alpha=1/14, min_periods=14, adjust=False).mean()
    avg_loss = loss.ewm(alpha=1/14, min_periods=14, adjust=False).mean()
    rs = avg_gain / avg_loss.replace(0, np.nan)
    d["RSI14"] = 100 - (100 / (1 + rs))
    ema12 = close.ewm(span=12, adjust=False).mean()
    ema26 = close.ewm(span=26, adjust=False).mean()
    d["MACD"] = ema12 - ema26
    d["MACD_signal"] = d["MACD"].ewm(span=9, adjust=False).mean()
    d["MACD_hist"] = d["MACD"] - d["MACD_signal"]
    d["BB_mid"] = close.rolling(20).mean()
    roll_std = close.rolling(20).std(ddof=1)
    d["BB_upper"] = d["BB_mid"] + 2 * roll_std
    d["BB_lower"] = d["BB_mid"] - 2 * roll_std
    d["Daily_return"] = close.pct_change()
    d["Volume_SMA20"] = pd.to_numeric(d["Volume"], errors="coerce").rolling(20).mean() if "Volume" in d else np.nan
    d["Drawdown"] = close / close.cummax() - 1
    return d

def fundamentals_frame():
    rows = []
    for fy, d in FINANCIALS.items():
        revenue = d["income_from_operations"]
        rows.append({
            "FY": fy,
            "Income from operations (₹ crore)": revenue,
            "Income growth (%)": np.nan,
            "EBITDA-like profit before D&A/finance/associates/exceptionals/tax (₹ crore)": d["ebitda_like"],
            "EBITDA-like margin (%)": d["ebitda_like"] / revenue * 100,
            "PAT attributable to owners (₹ crore)": d["pat_attributable_to_owners"],
            "PAT margin (%)": d["pat_attributable_to_owners"] / revenue * 100,
            "Profit before tax (₹ crore)": d["profit_before_tax"],
            "CFO (₹ crore)": d["cfo"],
            "Capex on PPE/intangibles (₹ crore)": d["capex_ppe_intangibles"],
            "CFO less PPE/intangible capex proxy (₹ crore)": d["cfo"] - d["capex_ppe_intangibles"],
            "ROE (%) company reported": d["roe_pct_company_reported"],
            "EPS (₹) company reported": d["eps_rupees_company_reported"],
            "Finance costs (₹ crore)": d["finance_costs"],
            "Total assets (₹ crore)": d["total_assets"],
            "Current assets (₹ crore)": d["total_current_assets"],
            "Current liabilities (₹ crore)": d["total_current_liabilities"],
            "Total equity including NCI (₹ crore)": d["total_equity"],
            "Borrowings (current + non-current, ₹ crore)": d["current_borrowings"] + d["noncurrent_borrowings"],
            "Current ratio (x)": d["total_current_assets"] / d["total_current_liabilities"],
            "Gross borrowings / total equity (x)": (d["current_borrowings"] + d["noncurrent_borrowings"]) / d["total_equity"],
        })
    f = pd.DataFrame(rows).set_index("FY")
    f.loc["FY26", "Income growth (%)"] = (FINANCIALS["FY26"]["income_from_operations"] / FINANCIALS["FY25"]["income_from_operations"] - 1) * 100
    f.loc["FY25", "Income growth (%)"] = np.nan
    return f

def calculate_technical_summary(frames, tech_frames):
    summary = {}
    for ticker, d in tech_frames.items():
        name = TICKERS[ticker]
        last = d.dropna(subset=["Close"]).iloc[-1]
        hist = d.dropna(subset=["Close"])
        start_1y = hist.loc[hist.index >= hist.index.max() - pd.DateOffset(years=1)]
        start_3y = hist.loc[hist.index >= hist.index.max() - pd.DateOffset(years=3)]
        start_5y = hist.loc[hist.index >= hist.index.max() - pd.DateOffset(years=5)]
        returns = lambda x: float(x["Close"].iloc[-1] / x["Close"].iloc[0] - 1) if len(x) >= 2 else None
        rsi = float(last["RSI14"]) if pd.notna(last["RSI14"]) else None
        macd_cross = "MACD above signal" if last["MACD"] > last["MACD_signal"] else "MACD below signal"
        trend = (
            "Price above SMA20, SMA50 and SMA200"
            if all(pd.notna(last[k]) and last["Close"] > last[k] for k in ["SMA20", "SMA50", "SMA200"])
            else "Price below SMA20, SMA50 and SMA200"
            if all(pd.notna(last[k]) and last["Close"] < last[k] for k in ["SMA20", "SMA50", "SMA200"])
            else "Mixed moving-average picture"
        )
        if rsi is None: rsi_state = "RSI unavailable"
        elif rsi >= 70: rsi_state = "RSI in conventional overbought zone (>=70)"
        elif rsi <= 30: rsi_state = "RSI in conventional oversold zone (<=30)"
        else: rsi_state = "RSI in middle range (30–70)"
        volume = float(last["Volume"]) if "Volume" in d and pd.notna(last.get("Volume")) else None
        volavg = float(last["Volume_SMA20"]) if pd.notna(last.get("Volume_SMA20")) else None
        summary[ticker] = {
            "name": name,
            "last_date": hist.index.max().strftime("%Y-%m-%d"),
            "close": float(last["Close"]),
            "high_52w": float(hist.loc[hist.index >= hist.index.max()-pd.DateOffset(years=1), "High"].max()) if "High" in hist and len(start_1y) else None,
            "low_52w": float(hist.loc[hist.index >= hist.index.max()-pd.DateOffset(years=1), "Low"].min()) if "Low" in hist and len(start_1y) else None,
            "return_1y": returns(start_1y),
            "return_3y": returns(start_3y),
            "return_5y": returns(start_5y),
            "SMA20": float(last["SMA20"]) if pd.notna(last["SMA20"]) else None,
            "SMA50": float(last["SMA50"]) if pd.notna(last["SMA50"]) else None,
            "SMA200": float(last["SMA200"]) if pd.notna(last["SMA200"]) else None,
            "RSI14": rsi,
            "MACD": float(last["MACD"]) if pd.notna(last["MACD"]) else None,
            "MACD_signal": float(last["MACD_signal"]) if pd.notna(last["MACD_signal"]) else None,
            "MACD_state": macd_cross,
            "trend_state": trend,
            "RSI_state": rsi_state,
            "volume": volume,
            "volume_sma20": volavg,
            "max_drawdown_5y": float(hist["Drawdown"].min()),
            "support_60d": float(hist.loc[hist.index >= hist.index.max()-pd.Timedelta(days=90), "Low"].min()) if "Low" in hist else None,
            "resistance_60d": float(hist.loc[hist.index >= hist.index.max()-pd.Timedelta(days=90), "High"].max()) if "High" in hist else None,
        }
    return summary

def create_charts(tech, fundamentals, comparison):
    CHARTS.mkdir(parents=True, exist_ok=True)
    sns.set_theme(style="whitegrid", context="notebook")
    paths = []
    mm = tech["M&M.NS"]
    close = mm["Close"].dropna()

    # Fundamental financial trend
    fig, ax = plt.subplots(figsize=(8.5, 5.2))
    x = ["FY25", "FY26"]
    revenue = [FINANCIALS[y]["income_from_operations"] / 1000 for y in x]
    pat = [FINANCIALS[y]["pat_attributable_to_owners"] / 1000 for y in x]
    ax.bar(np.array([0,1])-0.18, revenue, width=.36, label="Income from operations")
    ax.bar(np.array([0,1])+0.18, pat, width=.36, label="PAT attributable to owners")
    ax.set_xticks([0,1], x)
    ax.set_ylabel("₹ thousand crore")
    ax.set_title("1. M&M Consolidated Revenue and Owner PAT")
    ax.legend()
    fig.tight_layout()
    p=CHARTS/"01_fundamental_revenue_profit.png";fig.savefig(p,dpi=170);plt.close(fig);paths.append(p)

    # Profit margins and cash flow
    fig, ax = plt.subplots(figsize=(8.5,5.2))
    keys=["FY25","FY26"]
    ebitda_m=[FINANCIALS[k]["ebitda_like"]/FINANCIALS[k]["income_from_operations"]*100 for k in keys]
    pat_m=[FINANCIALS[k]["pat_attributable_to_owners"]/FINANCIALS[k]["income_from_operations"]*100 for k in keys]
    loc=np.array([0,1])
    ax.bar(loc-.18,ebitda_m,width=.36,label="EBITDA-like margin")
    ax.bar(loc+.18,pat_m,width=.36,label="Owner PAT margin")
    ax.set_xticks(loc,keys);ax.set_ylabel("Percent");ax.set_title("2. Margin Comparison");ax.legend()
    fig.tight_layout()
    p=CHARTS/"02_fundamental_margins.png";fig.savefig(p,dpi=170);plt.close(fig);paths.append(p)

    # Cash flow and capital expenditure
    fig, ax = plt.subplots(figsize=(8.5,5.2))
    cfo=[FINANCIALS[k]["cfo"] for k in keys]
    capex=[FINANCIALS[k]["capex_ppe_intangibles"] for k in keys]
    fcf=[cfo[i]-capex[i] for i in range(2)]
    ax.bar(loc-.25,cfo,width=.25,label="Operating cash flow")
    ax.bar(loc,capex,width=.25,label="PPE/intangible capex")
    ax.bar(loc+.25,fcf,width=.25,label="CFO less capex proxy")
    ax.set_xticks(loc,keys);ax.set_ylabel("₹ crore");ax.set_title("3. Cash Generation and Capex");ax.legend(fontsize=8)
    fig.tight_layout()
    p=CHARTS/"03_cashflow_capex.png";fig.savefig(p,dpi=170);plt.close(fig);paths.append(p)

    # Price and moving averages
    d=mm.dropna(subset=["Close"]).copy()
    fig, ax = plt.subplots(figsize=(10,5.8))
    recent=d.loc[d.index >= d.index.max()-pd.DateOffset(years=2)]
    ax.plot(recent.index,recent["Close"],label="M&M close",linewidth=1.5)
    for col in ["SMA20","SMA50","SMA200"]:
        ax.plot(recent.index,recent[col],label=col,linewidth=1.1)
    ax.set_title("4. M&M Price Trend with 20/50/200-Day Moving Averages")
    ax.set_ylabel("Price (₹)");ax.legend(ncol=2)
    fig.tight_layout()
    p=CHARTS/"04_price_moving_averages.png";fig.savefig(p,dpi=170);plt.close(fig);paths.append(p)

    # RSI
    fig, ax = plt.subplots(figsize=(10,4.3))
    ax.plot(recent.index,recent["RSI14"],label="RSI (14)",color="#4C78A8")
    ax.axhline(70,ls="--",color="#E45756",label="70 reference")
    ax.axhline(30,ls="--",color="#54A24B",label="30 reference")
    ax.set_ylim(0,100);ax.set_title("5. M&M Relative Strength Index (14-day)");ax.set_ylabel("RSI");ax.legend()
    fig.tight_layout()
    p=CHARTS/"05_rsi14.png";fig.savefig(p,dpi=170);plt.close(fig);paths.append(p)

    # MACD
    fig, ax = plt.subplots(figsize=(10,4.3))
    ax.plot(recent.index,recent["MACD"],label="MACD")
    ax.plot(recent.index,recent["MACD_signal"],label="Signal")
    colors=np.where(recent["MACD_hist"]>=0,"#54A24B","#E45756")
    ax.bar(recent.index,recent["MACD_hist"],label="Histogram",color=colors,width=1.2,alpha=.35)
    ax.axhline(0,color="gray",linewidth=.8);ax.set_title("6. M&M MACD Momentum");ax.legend(ncol=3)
    fig.tight_layout()
    p=CHARTS/"06_macd.png";fig.savefig(p,dpi=170);plt.close(fig);paths.append(p)

    # Bollinger
    fig, ax = plt.subplots(figsize=(10,5.3))
    ax.plot(recent.index,recent["Close"],label="Close",linewidth=1.4)
    ax.plot(recent.index,recent["BB_mid"],label="20-day midline")
    ax.fill_between(recent.index,recent["BB_lower"].to_numpy(dtype=float),recent["BB_upper"].to_numpy(dtype=float),alpha=.18,label="±2 standard deviations")
    ax.set_title("7. M&M Bollinger Bands (20-day, ±2σ)");ax.set_ylabel("Price (₹)");ax.legend()
    fig.tight_layout()
    p=CHARTS/"07_bollinger_bands.png";fig.savefig(p,dpi=170);plt.close(fig);paths.append(p)

    # Volume
    fig, ax = plt.subplots(figsize=(10,4.4))
    if "Volume" in recent:
        ax.bar(recent.index,recent["Volume"],width=1.2,alpha=.45,label="Daily volume")
        ax.plot(recent.index,recent["Volume_SMA20"],label="20-day average volume",color="#E45756")
    ax.set_title("8. M&M Trading Volume and 20-Day Average");ax.set_ylabel("Shares");ax.legend()
    fig.tight_layout()
    p=CHARTS/"08_volume.png";fig.savefig(p,dpi=170);plt.close(fig);paths.append(p)

    # Peer relative returns indexed to 100
    cols=[k for k in comparison if k in ["M&M.NS","MARUTI.NS","^NSEI"]]
    fig,ax=plt.subplots(figsize=(10,5.5))
    for ticker in cols:
        d=comparison[ticker].dropna()
        if len(d):
            d=d.loc[d.index >= d.index.max()-pd.DateOffset(years=3)]
            ax.plot(d.index,d/d.iloc[0]*100,label=TICKERS[ticker])
    ax.set_title("9. Three-Year Relative Price Performance");ax.set_ylabel("Indexed price (base = 100)");ax.legend()
    fig.tight_layout()
    p=CHARTS/"09_relative_performance.png";fig.savefig(p,dpi=170);plt.close(fig);paths.append(p)
    return paths

def build_report(fin_df, tech, summary, paths, metadata):
    doc=Document()
    sec=doc.sections[0];sec.top_margin=Inches(.6);sec.bottom_margin=Inches(.6);sec.left_margin=Inches(.7);sec.right_margin=Inches(.7)
    doc.styles["Normal"].font.name="Aptos";doc.styles["Normal"].font.size=Pt(9.4)
    doc.styles["Title"].font.name="Aptos Display";doc.styles["Title"].font.size=Pt(23)
    doc.styles["Heading 1"].font.name="Aptos Display";doc.styles["Heading 1"].font.size=Pt(15)
    doc.styles["Heading 2"].font.name="Aptos Display";doc.styles["Heading 2"].font.size=Pt(11.5)
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run("AUTOMOTIVE STOCK MARKET ANALYSIS INTERN");r.bold=True;r.font.size=Pt(11);r.font.color.rgb=RGBColor(230,100,0)
    p=doc.add_heading("Fundamental and Technical Analysis of Automotive Stocks",0);p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p=doc.add_paragraph("Primary company: Mahindra & Mahindra Ltd. (NSE: M&M) | Peer and benchmark context");p.alignment=WD_ALIGN_PARAGRAPH.CENTER

    fy26=FINANCIALS["FY26"];fy25=FINANCIALS["FY25"]
    rev_growth=(fy26["income_from_operations"]/fy25["income_from_operations"]-1)*100
    pat_growth=(fy26["pat_attributable_to_owners"]/fy25["pat_attributable_to_owners"]-1)*100
    margin26=fy26["pat_attributable_to_owners"]/fy26["income_from_operations"]*100
    margin25=fy25["pat_attributable_to_owners"]/fy25["income_from_operations"]*100
    ebitda26=fy26["ebitda_like"]/fy26["income_from_operations"]*100
    ebitda25=fy25["ebitda_like"]/fy25["income_from_operations"]*100
    fc26=fy26["cfo"]-fy26["capex_ppe_intangibles"]
    fc25=fy25["cfo"]-fy25["capex_ppe_intangibles"]
    s=summary["M&M.NS"]
    d_summary=summary.get("^NSEI",{})
    maruti=summary.get("MARUTI.NS",{})
    fmt_pct=lambda x: "n/a" if x is None or pd.isna(x) else f"{x*100:.1f}%"
    fmt_num=lambda x,n=2: "n/a" if x is None or pd.isna(x) else f"{x:,.{n}f}"

    doc.add_heading("Executive Summary",1)
    doc.add_paragraph(
      f"M&M's FY2025–26 consolidated income from operations was ₹{fy26['income_from_operations']:,.2f} crore, up {rev_growth:.1f}% from FY25, while profit attributable to owners reached ₹{fy26['pat_attributable_to_owners']:,.2f} crore, up {pat_growth:.1f}%. The owner PAT margin improved from {margin25:.2f}% to {margin26:.2f}%, and company-reported FY26 return on equity was 20.1% with EPS of ₹152.2 [F1–F3]. "
      f"The statement of cash flows reports ₹{fy26['cfo']:,.2f} crore operating cash flow and ₹{fy26['capex_ppe_intangibles']:,.2f} crore PPE/intangible acquisition outflow, a simple CFO-minus-capex proxy of ₹{fc26:,.2f} crore [F2]. "
      f"Technical analysis uses daily OHLCV data through {s['last_date']}. The latest recorded close was ₹{s['close']:,.2f}; price was {s['trend_state'].lower()}, RSI(14) was {fmt_num(s['RSI14'])}, and MACD was {s['MACD_state'].lower()}. These signals are descriptive, lagging and not a standalone buy/sell call."
    )
    doc.add_heading("Key messages",2)
    for item in [
      f"Fundamental growth: consolidated income from operations rose {rev_growth:.1f}% and owner-attributable PAT rose {pat_growth:.1f}% in FY26.",
      f"Profitability: owner PAT margin was {margin26:.2f}% versus {margin25:.2f}% in FY25. The report separately labels its EBITDA-like measure to avoid confusion with a standardized reported EBITDA figure.",
      f"Cash quality: CFO improved to ₹{fy26['cfo']:,.0f} crore. CFO less PPE/intangible capex was ₹{fc26:,.0f} crore, but this is a simplified proxy and consolidated cash flows include financial-services effects.",
      f"Technical state: the latest price was ₹{s['close']:,.2f}; RSI {fmt_num(s['RSI14'])}; {s['MACD_state']}. Review the actual charts and price date before interpreting a short-term signal.",
      "The company is diversified beyond vehicle manufacturing. Consolidated finance costs, assets and working-capital ratios should not be read as if M&M were a pure-play automaker."
    ]: doc.add_paragraph(item,style="List Bullet")

    doc.add_heading("1. Company profile and analytical scope",1)
    doc.add_paragraph("Mahindra & Mahindra Limited (M&M) is a major Indian automotive and farm-equipment business within a diversified group. The FY26 group reporting also reflects financial services, technology and other interests. This makes consolidated results useful to understand group earnings but requires care when inferring the economics of only the automotive business. The report uses M&M as the main case, Maruti Suzuki as a listed automotive peer for share-price context, and Nifty 50 as the broad-market comparison. Comparisons are not perfectly like-for-like: companies have different segment mixes and a stock-price chart does not replace a peer's financial statements.")
    doc.add_paragraph("The assignment is split into two distinct methods. Fundamental analysis asks what the business earned, how margins and cash generation changed, what risks appear in the statements, and what valuation context implies. Technical analysis asks how the market price and trading volume behaved, and what moving-average, momentum and volatility-band indicators currently show. The fundamental period is FY25 vs FY26; technical calculations use rolling daily data up to the latest downloaded date. This avoids mixing annual accounting periods with the rolling trading window.")
    doc.add_heading("2. Data sources and methodology",1)
    doc.add_paragraph("Fundamental figures come from M&M's FY2025–26 Integrated Annual Report and audited results announcement published on 5 May 2026 [F1–F3]. The report distinguishes income from operations from revenue from operations as labeled in the annual report; consolidated figures include the group subsidiaries and financial-services business. Market price/volume series are downloaded at run time from Yahoo Finance through yfinance [T1]. NSE India provides a separate official quote and market-statistics reference [F4]. Maruti's investor disclosures are linked as peer context [P1].")
    doc.add_paragraph(f"Price-data window available: {metadata['observed_start']} to {metadata['observed_end']}. Technical indicators are calculated on the adjusted daily date index using close prices (moving averages, RSI, MACD, Bollinger Bands) and daily trading volume. For performance comparisons, close-price returns are price returns and do not include all dividends unless reflected by adjusted prices. Exact ticker coverage and row counts are saved in outputs/run_metadata.json.")
    doc.add_heading("3. Fundamental analysis",1)
    doc.add_heading("3.1 Income statement trend",2)
    doc.add_paragraph(
      f"Income from operations increased from ₹{fy25['income_from_operations']:,.2f} crore in FY25 to ₹{fy26['income_from_operations']:,.2f} crore in FY26 ({rev_growth:.1f}% growth). Profit before depreciation, finance costs, share of profit from associates/joint ventures, exceptional items and tax—the report's explicitly labeled EBITDA-like measure—rose from ₹{fy25['ebitda_like']:,.2f} crore to ₹{fy26['ebitda_like']:,.2f} crore. Its ratio to income from operations was {ebitda25:.2f}% in FY25 and {ebitda26:.2f}% in FY26. Because this subtotal is defined by the annual report's line item and may not match analyst-adjusted EBITDA, it is labelled rather than presented as a standardized metric [F1]."
    )
    doc.add_paragraph(
      f"Profit attributable to owners grew from ₹{fy25['pat_attributable_to_owners']:,.2f} crore to ₹{fy26['pat_attributable_to_owners']:,.2f} crore ({pat_growth:.1f}%). Owner PAT margin grew from {margin25:.2f}% to {margin26:.2f}%. Profit for the year before attribution to non-controlling interests is different from PAT attributable to owners; the report uses owner PAT consistently for the owner margin calculation. M&M's release reports consolidated PAT of ₹17,099 crore, RoE of 20.1% and EPS of ₹152.2 for FY26 [F3]."
    )
    doc.add_heading("3.2 Fundamental ratios",2)
    ratio_rows=[
      ["Income from operations growth",f"{rev_growth:.2f}%","(FY26 income from operations / FY25 income from operations − 1) × 100"],
      ["Owner PAT growth",f"{pat_growth:.2f}%","(FY26 owner PAT / FY25 owner PAT − 1) × 100"],
      ["EBITDA-like margin FY25",f"{ebitda25:.2f}%","EBITDA-like subtotal / income from operations × 100"],
      ["EBITDA-like margin FY26",f"{ebitda26:.2f}%","EBITDA-like subtotal / income from operations × 100"],
      ["Owner PAT margin FY25",f"{margin25:.2f}%","Owner-attributable PAT / income from operations × 100"],
      ["Owner PAT margin FY26",f"{margin26:.2f}%","Owner-attributable PAT / income from operations × 100"],
      ["CFO change FY25 → FY26",f"{(fy26['cfo']/fy25['cfo']-1)*100:.1f}%","(FY26 CFO / FY25 CFO − 1) × 100; comparison is from a low base"],
      ["CFO less PPE/intangible capex FY25",f"₹{fc25:,.2f} cr","CFO − purchase of PPE and intangible assets; simplified proxy"],
      ["CFO less PPE/intangible capex FY26",f"₹{fc26:,.2f} cr","CFO − purchase of PPE and intangible assets; simplified proxy"],
      ["Company-reported RoE FY26","20.1%","M&M FY26 results release; denominator methodology per company report"],
      ["Company-reported EPS FY26","₹152.20","M&M FY26 results release"],
      ["Total assets FY25","₹277,586.27 cr","Audited consolidated balance sheet at 31 March 2025"],
      ["Total assets FY26","₹317,634.64 cr","Audited consolidated balance sheet at 31 March 2026"],
      ["Total current assets / current liabilities FY25","1.40x","Current assets divided by current liabilities; group-wide indicator only"],
      ["Total current assets / current liabilities FY26","1.27x","Current assets divided by current liabilities; group-wide indicator only"],
      ["Gross borrowings / total equity FY25","1.36x","Current plus non-current borrowings / total equity including NCI; not an industrial net-debt ratio"],
      ["Gross borrowings / total equity FY26","1.18x","Current plus non-current borrowings / total equity including NCI; not an industrial net-debt ratio"]
    ]
    table=doc.add_table(rows=1,cols=3);table.style="Light Shading Accent 1"
    for c,t in zip(table.rows[0].cells,["Ratio / measure","Result","Formula / basis"]):c.text=t
    for row in ratio_rows:
      cells=table.add_row().cells
      for cell,value in zip(cells,row):cell.text=value
    doc.add_paragraph("The audited consolidated balance sheet reports total assets of ₹277,586.27 crore at March 2025 and ₹317,634.64 crore at March 2026; total equity including non-controlling interests increased from ₹89,098.00 crore to ₹109,491.00 crore. Total current assets / total current liabilities decreased from 1.40x to 1.27x, while current plus non-current borrowings divided by total equity decreased from approximately 1.36x to 1.18x [F6]. These are broad group ratios only: the group includes a large lending business, whose loans, funding and liquidity profile are fundamentally different from the inventory-and-trade-credit model of a vehicle manufacturer. They should not be interpreted as a standalone automotive working-capital or industrial gearing test.")
    doc.add_heading("3.3 Cash flow and capital allocation",2)
    doc.add_paragraph(
      f"Net cash flow from operating activities increased from ₹{fy25['cfo']:,.2f} crore in FY25 to ₹{fy26['cfo']:,.2f} crore in FY26 [F2]. Purchase of property, plant, equipment and intangible assets was ₹{fy26['capex_ppe_intangibles']:,.2f} crore in FY26, lower than ₹{fy25['capex_ppe_intangibles']:,.2f} crore in FY25. On the simple CFO-minus-capex proxy, the result moved from ₹{fc25:,.2f} crore to ₹{fc26:,.2f} crore. The increase in operating cash flow is encouraging, yet the proxy should not be over-read because changes in financial-services receivables are material in consolidated operating cash flow [F2]."
    )
    doc.add_paragraph("The analyst should compare cash generation with reported earnings over multiple years, not one period. Additional checks include capex commitments, investment cash flows, acquisition-related outflows, dividend commitments, debt maturities, and whether profits are converted into cash. M&M's group structure means that investing cash flows can be large and volatile because purchases and sales of financial investments are substantial; those lines do not equal vehicle-manufacturing capex.")
    doc.add_heading("3.4 Business positioning and peer context",2)
    doc.add_paragraph("The FY26 company release reported SUV revenue market share of 25.3%, LCV below 3.5-tonne market share of 52.3%, tractor market share of 43.6% and electric three-wheeler market share of 40.0% under the stated company/industry definitions [F3]. It also reported growth across auto and farm, alongside other group businesses. These are relevant operating indicators but must be read with definitions and periods in the release. Revenue market share is not the same metric as unit share; neither proves that the shares will persist.")
    doc.add_paragraph("Maruti Suzuki is included as a peer for market-price context because it has meaningful passenger-vehicle overlap, but its business mix and reporting details are not identical to M&M. A complete valuation comparison would require aligned FY statements, diluted EPS, debt/cash, segment economics and forecast estimates for both companies. This project avoids drawing a firm valuation conclusion from a single price-to-earnings snapshot.")
    doc.add_heading("3.5 Valuation and fundamental risks",2)
    doc.add_paragraph("The NSE quote page on 12 October 2026 showed M&M at ₹2,813, with a reported symbol P/E of 17.18 and 52-week range of ₹2,744.30–₹3,839.90 [F4]. These are timestamped market observations and will change. P/E is most useful when paired with expected earnings growth, capital intensity, cyclicality, competitive positioning and the quality of earnings. A lower P/E can reflect a lower expected growth rate or higher risk; a higher P/E can reflect expectations already embedded in the share price.")
    doc.add_paragraph("Fundamental risks include cyclical vehicle demand, financing affordability, commodity/input costs, discounting, product recalls/quality, model-cycle timing, competition in SUVs and electric vehicles, export-market fluctuations, regulation and group-level complexity. Financial services can add funding, credit-cost and asset-quality exposure. The analyst should monitor monthly sales and quarterly segment disclosures and test whether market-share gains translate into profitability and cash flow.")
    doc.add_page_break()
    doc.add_heading("3.6 Fundamental visualizations",1)
    fundamental_figures = [
      (paths[0], "Figure 1. Consolidated Income and Owner Profit", "Income from operations and profit attributable to owners both increased from FY25 to FY26. Read this together with margin changes; absolute growth alone does not explain returns or cash quality."),
      (paths[1], "Figure 2. EBITDA-Like and Owner PAT Margins", "The report's explicitly defined EBITDA-like margin eased slightly while owner PAT margin improved. This emphasizes why multiple profit measures are needed."),
      (paths[2], "Figure 3. Operating Cash Flow and Capex", "Operating cash flow rose strongly and PPE/intangible capex eased. CFO less capex is only a simplified proxy because consolidated cash flows include financial-services movements.")
    ]
    for idx, (img, title, note) in enumerate(fundamental_figures):
      if idx == 2:
        doc.add_page_break()
      hp = doc.add_heading(title,2); hp.paragraph_format.keep_with_next = True
      doc.add_picture(str(img),width=Inches(5.9))
      doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
      doc.paragraphs[-1].paragraph_format.keep_together = True
      np = doc.add_paragraph(note); np.paragraph_format.keep_together = True

    doc.add_heading("4. Technical analysis",1)
    doc.add_heading("4.1 Trend and moving averages",2)
    doc.add_paragraph(
      f"The latest downloaded M&M closing price was ₹{s['close']:,.2f} on {s['last_date']}. The 20-day simple moving average was ₹{fmt_num(s['SMA20'])}, the 50-day average ₹{fmt_num(s['SMA50'])}, and the 200-day average ₹{fmt_num(s['SMA200'])}. Current configuration: {s['trend_state']}. Shorter moving averages react faster but are noisier; the 200-day average is slower and may lag major turning points. A moving-average cross is a rule-based description, not proof of an upcoming move."
    )
    doc.add_heading("4.2 RSI momentum",2)
    doc.add_paragraph(f"RSI(14) was {fmt_num(s['RSI14'])}, classified as {s['RSI_state'].lower()}. RSI compares the average magnitude of recent up days with down days and is bounded between 0 and 100. Traders often use 70 and 30 as overbought/oversold reference levels, but a strong trend can keep RSI elevated or depressed, and an oversold reading alone is not a buy signal.")
    doc.add_heading("4.3 MACD",2)
    doc.add_paragraph(f"The latest MACD was {fmt_num(s['MACD'])} and its signal line {fmt_num(s['MACD_signal'])}; the relative state was {s['MACD_state']}. MACD is the difference between 12- and 26-period exponential moving averages; the signal line is a 9-period EMA of MACD. A MACD line above the signal may indicate positive short-term momentum, but the indicator is lagging and false crosses occur during sideways markets.")
    doc.add_heading("4.4 Bollinger Bands and volatility",2)
    doc.add_paragraph("The chart uses a 20-session moving average with bands at two rolling standard deviations above and below the midline. A widening band suggests higher realized price dispersion, while narrowing bands show a quieter period that may precede renewed movement—but the direction is unknown from band width alone. Price touching an upper band does not automatically mean it is overvalued; it can accompany a strong trend.")
    doc.add_heading("4.5 Volume, support/resistance and drawdown",2)
    doc.add_paragraph(f"The latest observation's volume was {fmt_num(s['volume'],0)} shares versus a 20-day average of {fmt_num(s['volume_sma20'],0)} where available. Volume helps assess whether a price move attracted more or less participation, but data vendor conventions and block trades can affect the series. The one-year high/low context was ₹{fmt_num(s['low_52w'])}–₹{fmt_num(s['high_52w'])}; a rough recent 60-trading-day support/resistance proxy based on the last 90 calendar days is ₹{fmt_num(s['support_60d'])}/₹{fmt_num(s['resistance_60d'])}. The five-year maximum drawdown was {fmt_pct(s['max_drawdown_5y'])}. These are historical zones, not precise guarantees or trade instructions.")
    doc.add_heading("4.6 Relative performance",2)
    if d_summary:
      doc.add_paragraph(f"Over the observed one-year window, M&M price return was {fmt_pct(s['return_1y'])} and Nifty 50 was {fmt_pct(d_summary['return_1y'])}; over three years, M&M was {fmt_pct(s['return_3y'])} and Nifty 50 {fmt_pct(d_summary['return_3y'])}. For Maruti Suzuki, the one-year price return was {fmt_pct(maruti.get('return_1y'))}. These are window-specific price returns, not forecasts or total-return comparisons.")
    doc.add_page_break()
    doc.add_heading("4.7 Technical visualizations",1)
    technical_figures = [
      (paths[3], "Figure 4. M&M Price Trend and Moving Averages", "The closing price is plotted with 20-, 50- and 200-day moving averages. The three averages provide short-, medium- and long-horizon trend context; they lag price and should not be used in isolation."),
      (paths[4], "Figure 5. RSI (14)", "RSI near/below 30 is conventionally called oversold, but this is not a buy signal. Momentum indicators can remain depressed in strong downtrends."),
      (paths[5], "Figure 6. MACD Momentum", "The MACD line below its signal line indicates weaker recent momentum under this rule. False crosses occur, particularly when price is moving sideways."),
      (paths[6], "Figure 7. Bollinger Bands", "The band width describes recent volatility. A touch of a band is not by itself a reversal signal; price can track a band during a persistent trend."),
      (paths[7], "Figure 8. Trading Volume", "Volume is compared with its 20-session average to help assess participation behind moves. Single-day volume may be distorted by block trades or data-provider conventions."),
      (paths[8], "Figure 9. Three-Year Relative Price Performance", "M&M, Maruti Suzuki and Nifty 50 are rebased to 100 for easier comparison. Different business mixes and the use of price returns limit direct conclusions.")
    ]
    for idx, (img, title, note) in enumerate(technical_figures):
      if idx in (2,4):
        doc.add_page_break()
      hp = doc.add_heading(title,2); hp.paragraph_format.keep_with_next = True
      doc.add_picture(str(img),width=Inches(5.9))
      doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
      doc.paragraphs[-1].paragraph_format.keep_together = True
      np = doc.add_paragraph(note); np.paragraph_format.keep_together = True

    doc.add_heading("5. Fundamental and technical analysis together",1)
    doc.add_paragraph("Fundamental evidence works over earnings cycles: FY26 revenue and profit grew, reported RoE was strong, and operating cash flow improved. Technical indicators describe how the market has re-priced the stock during the rolling window. The two can diverge because prices respond to future expectations, valuation changes, market-wide risk appetite, and short-term news. A positive fundamental report does not guarantee positive short-term price action; a weak chart does not invalidate the audited accounts.")
    doc.add_paragraph(f"The current technical snapshot—{s['trend_state'].lower()}, RSI {fmt_num(s['RSI14'])}, and {s['MACD_state'].lower()}—should be treated as a dashboard, not a categorical recommendation. In practice, the analyst should specify a time horizon, scenario, price/volume invalidation level and valuation assumptions before forming a view. It is also important to check that the latest daily data are not stale and that official exchange quotes agree with third-party sources.")
    doc.add_heading("6. Monitoring checklist and recommendations",1)
    recs=[
      "Monthly: compare M&M SUV, LCV, tractor and total auto dispatches with segment context, industry growth, exports and inventory/retail indicators.",
      "Quarterly: track revenue, EBITDA-like margins, auto/farm segment PBIT, warranty/quality costs, discounts, receivables, capex, debt and cash conversion.",
      "Valuation: compare forward earnings expectations, P/E history and peer multiples with expected growth, ROE, cash flow and business cyclicality.",
      "Technical: update SMA20/50/200, RSI, MACD, Bollinger width, volume, relative performance and drawdown using fresh market observations.",
      "Risk control: distinguish investment horizon and risk tolerance from chart signals; predefine position-sizing and loss tolerance instead of chasing a price breakout.",
      "Source control: use company filings and NSE/BSE disclosures for final facts, record retrieval dates, and investigate discrepancies before changing a conclusion."
    ]
    for item in recs: doc.add_paragraph(item,style="List Bullet")
    doc.add_heading("7. Limitations",1)
    for item in [
      "The financial analysis focuses on FY25 vs FY26 and is not a complete ten-year historical model or forward forecast.",
      "Consolidated M&M results combine businesses with different economics; consolidated finance costs, assets, and cash flow cannot be attributed solely to the automotive segment.",
      "The EBITDA-like subtotal follows an annual-report line-item definition and may differ from analyst-adjusted EBITDA.",
      "CFO less PPE/intangible capex is a simplified cash-flow proxy, not a fully standardized free-cash-flow measure.",
      "Technical indicators are lagging, parameter-sensitive and prone to false signals in volatile or range-bound markets.",
      "Yahoo Finance data may contain gaps, adjustments or volume differences; the script records data coverage, and current values should be cross-checked against official exchange records.",
      "Past returns do not imply future returns. This report is educational analysis, not personalized investment advice."
    ]:doc.add_paragraph(item,style="List Bullet")
    doc.add_heading("8. Conclusion",1)
    doc.add_paragraph("FY26 provides a constructive fundamental base for examining M&M: consolidated income from operations rose by about one quarter, owner-attributable PAT grew faster, the owner PAT margin increased, and reported operating cash flow improved. Company-reported market-share indicators provide operating context, but future performance still depends on demand, product mix, pricing, costs, credit conditions and execution. Technical charts supply a separate view of trend, momentum and volatility; none can forecast a price with certainty. The most defensible analyst output is a documented, regularly updated thesis that joins audited financial evidence to market behavior, and states what would change the view.")
    doc.add_heading("9. References",1)
    for source in SOURCES:
      doc.add_paragraph(f"[{source['id']}] {source['title']}. {source['use']} URL: {source['url']}")
    doc.add_paragraph(f"Market-data run timestamp: {metadata['generated_at_local']}; requested history {HISTORY_PERIOD}; observed price range {metadata['observed_start']} to {metadata['observed_end']}. Exact tickers and observations are in outputs/run_metadata.json.")
    doc.save(REPORT)

def build_submission_doc():
    content=SUBMISSION_TXT.read_text(encoding="utf-8").strip()
    doc=Document()
    sec=doc.sections[0];sec.top_margin=Inches(.7);sec.bottom_margin=Inches(.7);sec.left_margin=Inches(.85);sec.right_margin=Inches(.85)
    doc.styles["Normal"].font.name="Aptos";doc.styles["Normal"].font.size=Pt(11)
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run("AUTOMOTIVE STOCK MARKET ANALYSIS INTERN");r.bold=True;r.font.size=Pt(11);r.font.color.rgb=RGBColor(230,100,0)
    p=doc.add_heading("Task 2 — Submission Description",0);p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    for block in content.split("\n\n"):
        p=doc.add_paragraph(block);p.paragraph_format.space_after=Pt(10);p.paragraph_format.line_spacing=1.12
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    r=p.add_run(f"Word count: {len(content.split())}");r.italic=True;r.font.size=Pt(9)
    doc.save(SUBMISSION_DOCX)

def main():
    OUT.mkdir(parents=True,exist_ok=True);CHARTS.mkdir(parents=True,exist_ok=True)
    frames,diagnostics=download_data()
    all_close={}
    tech_frames={}
    all_returns={}
    for ticker,df in frames.items():
        tech_frames[ticker]=technical_indicators(df)
        close=df["Close"].copy()
        if "Adj Close" in df and ticker not in ("^NSEI",):
            close=df["Adj Close"].fillna(df["Close"])
        all_close[ticker]=close.rename(TICKERS[ticker])
        all_returns[ticker]=all_close[ticker].pct_change()
    price_table=pd.concat(all_close.values(),axis=1)
    price_table.to_csv(OUT/"historical_close_prices.csv",index_label="Date")
    for ticker,df in tech_frames.items():
        df.to_csv(OUT/(ticker.replace("^","").replace(".NS","").replace("&","")+"_technical_indicators.csv"),index_label="Date")
    fin_df=fundamentals_frame()
    fin_df.to_csv(OUT/"fundamental_ratios.csv")
    summary=calculate_technical_summary(frames,tech_frames)
    (OUT/"technical_snapshot.json").write_text(json.dumps(summary,indent=2),encoding="utf-8")
    end=min(df.index.max() for df in frames.values())
    start=max(df.index.min() for df in frames.values())
    meta={
      "generated_at_local":datetime.now().astimezone().isoformat(timespec="seconds"),
      "source":"Yahoo Finance via yfinance; OHLCV series downloaded at run time",
      "period_requested":HISTORY_PERIOD,
      "observed_start":start.strftime("%Y-%m-%d"),
      "observed_end":end.strftime("%Y-%m-%d"),
      "tickers":diagnostics,
      "fundamental_periods":["FY2025","FY2026"],
      "fundamental_source":"M&M Integrated Annual Report FY2025-26 and audited results release",
      "synthetic_price_data_used":False,
      "charts":[
        "01_fundamental_revenue_profit.png","02_fundamental_margins.png","03_cashflow_capex.png",
        "04_price_moving_averages.png","05_rsi14.png","06_macd.png","07_bollinger_bands.png",
        "08_volume.png","09_relative_performance.png"
      ]
    }
    (OUT/"run_metadata.json").write_text(json.dumps(meta,indent=2),encoding="utf-8")
    (OUT/"sources.json").write_text(json.dumps(SOURCES,indent=2),encoding="utf-8")
    paths=create_charts(tech_frames,fin_df,all_close)
    build_report(fin_df,tech_frames,summary,paths,meta)
    build_submission_doc()
    print("M&M FY26/FY25 data loaded from official published statement values.")
    print(f"Market data downloaded through {meta['observed_end']}; tickers: {list(frames)}")
    print("Synthetic market prices used: False")
    print(fin_df.to_string())
    print(json.dumps(summary["M&M.NS"],indent=2))
    print(f"Created Word report and {len(paths)} charts.")

if __name__=="__main__":
    main()
