# Task 3 — Financial Health Evaluation of Selected Automotive Companies

This independent project compares the FY2025–26 standalone financial health of **Mahindra & Mahindra Limited (M&M)** and **Maruti Suzuki India Limited (MSIL)**. It extends the automotive market-analysis repository and is not connected to the earlier Student Performance coursework.

## Purpose and comparison basis
The project evaluates profitability, liquidity, solvency, operating efficiency and cash generation across FY25 and FY26 using the companies' published financial disclosures. Standalone financial statements are used to focus on the listed operating companies and avoid mixing M&M's large consolidated financial-services and diversified group businesses with Maruti's vehicle-manufacturing profile. The report explains that M&M standalone still includes farm equipment and investment income, while Maruti standalone reflects its own operations after the Suzuki Motor Gujarat amalgamation/restatement.

## Deliverables
- A 2,000+ word DOCX report with executive summary, methodology, two-company comparative analysis, key ratios, findings, risks and conclusion.
- Comparison tables covering revenue, EBIT/operating margins, PAT, ROE, current ratio, working capital, leverage, interest cover and cash conversion.
- Charts for revenue/PAT, profitability, liquidity, leverage/coverage and cash flow/capex.
- Reproducible Python source, source-data JSON, calculated ratios CSV, chart PNGs and separate 200+ word submission description in TXT and DOCX formats.

## Key data sources
- M&M audited standalone results for the year ended 31 March 2026, including standalone balance sheet, cash flow and SEBI ratio disclosures: https://www.mahindra.com/sites/default/files/2026-05/Quarterly%20Result%20for%20Q4F26.pdf
- M&M Integrated Annual Report FY2025–26: https://www.mahindra.com/annual-report-FY2026/
- Maruti Suzuki Annual Integrated Report FY2025–26: https://www.marutisuzuki.com/corporate/media/press-releases/2026/august/maruti-suzuki-annual-integrated-report-2025-26
- Maruti Suzuki FY2025–26 audited results: https://www.marutisuzuki.com/corporate/media/press-releases/2026/april/maruti-suzuki-announces-financial-results-for-fy2025-26
- Maruti Q2/H1 FY2026–27 results and audited 31 March 2026 comparative balance sheet: https://www.marutisuzuki.com/corporate/media/press-releases/2026/october/maruti-suzuki-india-limited-financial-results-q2-and-h1-fy-2026-27

## Reproduce
```bash
python -m pip install -r requirements.txt
python financial_health_analysis.py
```

The script uses disclosed/verified financial values and performs all calculations and chart generation locally. It does not generate or invent financial statements. Amounts are normalized to ₹ crore unless the table specifies otherwise.

## Important interpretation limits
Company-defined operating margins and financial ratios may use different denominators or definitions. M&M's report explicitly excludes financial-services business from selected debt/coverage ratios; Maruti's near-zero borrowings mainly comprise lease liabilities. These distinctions are documented in the report. This is an educational financial-health review, not investment advice or a target-price recommendation.

## Published summary
See [analysis-summary.md](analysis-summary.md) for key FY25/FY26 comparison tables, derived ratios, conclusions and public references. The detailed submission files are included in the complete project ZIP supplied with this task.
