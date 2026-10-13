# Automotive Stock Market Analysis

**Task 1 — Automotive Market Overview & Trend Analysis**

An independent Python project analyzing India's listed automotive equities using five years of publicly available daily market data, public industry research, return/risk metrics, visualizations, and strategic insights.

## Coverage
- Nifty 50 benchmark and an equal-weighted sample basket of selected Indian automotive companies.
- Companies include Maruti Suzuki, Mahindra & Mahindra, Bajaj Auto, Eicher Motors, Hero MotoCorp, TVS Motor, Ashok Leyland, Bharat Forge, Bosch, and Exide Industries.
- Historical returns, annualized volatility, maximum drawdown, daily return correlations, and monthly performance.
- Industry context from SIAM, Government of India notices, PM E-DRIVE, company results, investment announcements, exports and input-cost risks.

## Data-source note
The project downloads historical prices through Yahoo Finance via `yfinance`. In the documented run, Yahoo supplied only one usable historical observation for the `^CNXAUTO` symbol; therefore the historical comparison uses a clearly labelled equal-weighted basket of the downloaded automotive shares compared with Nifty 50. This proxy is **not** the official Nifty Auto index. No synthetic price history is used, and the limitation is documented in the report.

## Run locally
```bash
python -m pip install -r requirements.txt
python automotive_market_analysis.py
```

The script generates `Automotive_Market_Overview_and_Trend_Analysis_Report.docx`, `Submission_Description.docx`, seven charts in `visualizations/`, and data/metrics/source files in `outputs/`.

## Interpretation
Historical returns are descriptive, not forecasts. This is educational market research, not personalized investment advice.
