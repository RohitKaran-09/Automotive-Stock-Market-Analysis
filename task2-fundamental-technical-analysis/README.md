# Task 2 — Fundamental and Technical Analysis of Automotive Stocks

This project continues the independent Automotive Stock Market Analysis internship repository. It is separate from the earlier Student Performance coursework and adds a dual-analysis report focused on Mahindra & Mahindra Ltd. (NSE: M&M), with market comparisons against Nifty 50 and selected auto peers.

## Deliverables
- A 2,000+ word Word report with separate fundamental and technical analysis sections.
- FY2025–26 income statement, balance-sheet and cash-flow analysis using M&M's audited annual report and company results release.
- Financial ratios: revenue growth, EBITDA and PAT margins, EPS growth, return on equity, current ratio, debt/equity, CFO, capex and an explained free-cash-flow proxy.
- Technical indicators calculated from actual daily price history: SMA 20/50/200, RSI(14), MACD, Bollinger Bands, volume, relative performance, drawdown, and 52-week high/low context.
- Fundamental trend charts, price/technical-indicator charts, volume, return comparison and indicator snapshots.
- Python source, data tables, chart outputs, metadata, and a separate 200+ word submission description.

## Data sources and limits
Fundamental figures are anchored to M&M's FY2025–26 published results/Integrated Annual Report (year ended 31 March 2026). Share prices and daily technical indicators are downloaded at run time from Yahoo Finance through yfinance; output metadata records the observation window. The report cites official NSE, M&M, and Maruti sources plus dated news for industry context.

M&M is a diversified group with Automotive, Farm Equipment, Financial Services and other investments. Consolidated working-capital and debt ratios are affected by the lending businesses and must not be interpreted as pure-auto ratios. The report explicitly explains this caveat. Technical signals are descriptive and not a stand-alone buy/sell recommendation.

## Run
From this folder:
```bash
python -m pip install -r requirements.txt
python fundamental_technical_analysis.py
```

If market-price download fails, the script exits rather than fabricating prices.
