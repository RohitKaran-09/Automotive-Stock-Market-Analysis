# Week 4 — Risk and Opportunity Assessment in the Automotive Stock Market

India-focused sector research assessing internal/external risks and medium- to long-term opportunities across listed OEMs, components, EVs and mobility infrastructure.

**Research cut-off:** 2 November 2026.

## Deliverables
- `Automotive_Risk_and_Opportunity_Assessment_Week4_Report.docx` — detailed report (3,400+ words) with 12-item risk register, 9-item opportunity register, risk heatmap, prioritisation chart, opportunity map, combined risk-reward matrix, SWOT, scenario analysis and monitoring dashboard.
- `Submission_Description.docx` and `submission_description.txt` — 400-word portal description.
- `visualizations/` — seven PNG charts.
- `outputs/` — CSV risk/opportunity/monitoring registers and source metadata.
- `risk_opportunity_assessment.py` — reproducible report and chart generator.

## Main evidence used
SIAM Q2 FY2026–27 industry results; PM E-DRIVE scheme notifications and October 2026 factsheet; September 2026 CAFE norms; IEA Global EV Outlook 2026; dated reporting on energy/supply-chain disruption and India manufacturing investment. See `outputs/source_register.json` and citations in the report.

## Scoring methodology
Risk scores = qualitative likelihood (1–5) × impact (1–5). Opportunity scores = analyst-assessed readiness/probability (1–5) × potential impact (1–5). Scores are prioritisation judgements, not actuarial probabilities, forecasts, target prices or expected returns. The scenario scores are directional stress illustrations only.

## Reproduce
```bash
python -m pip install -r requirements.txt
python risk_opportunity_assessment.py
```

The project is educational market research, not personalised investment advice or a buy/sell recommendation.
