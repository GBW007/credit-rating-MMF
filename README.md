# M M Forgings Ltd — Independent Credit Rating Analysis (real public data)

Rebuilds the credit rating of a listed Indian forging company (NSE: MMFL) from public financials and benchmarks it against CARE Ratings' published rating (CARE A; Stable, 30-Jun-2026).

**Result:** conservative scorecard = **BBB**; after a transparent notch bridge (net liquidity, CARE's debt definition, forward-looking basis, CARE's accrual expectation) = **A-**, one notch below CARE. Conclusion: CARE A is defensible but has limited headroom and depends on accruals rising from ₹202 Cr to ~₹290 Cr.

| File | Contents |
|---|---|
| `MMForgings_Credit_Rating_Analysis.xlsx` | Dashboard · Source_Data (every number cited) · FY27E_Base · Analysis (ratios FY22–FY27E) · Debt_Service (5 scenarios) · Rating (scorecard, gross/net toggle) · Rating_Bridge · CARE_Benchmark (reconciliation + trigger distances) |
| `MMForgings_Credit_Rating_Note.pdf/.docx` | 5-page rating note |
| `python/rating_bridge.py` | Replicates the scorecard in Python (matches Excel exactly) and computes the notch bridge |
| `python/build_mmf.py`, `python/charts.py` | Rebuild the workbook and charts |

**Notable findings:** debt +₹479 Cr vs PBILDT +₹7 Cr over FY23–26 · interest coverage 7.0x → 3.6x · FY26 debt ₹160 Cr higher than CARE's implied figure (definition gap) · CARE text/table interest-coverage inconsistency (3.98x vs 3.58x) · Q1 FY27 PAT inflated by ₹63 Cr one-off other income.

Sources: Screener.in consolidated financials (C-MOTS data); CARE Ratings press release, 30-Jun-2026. Educational analysis — not an official rating or investment advice.
