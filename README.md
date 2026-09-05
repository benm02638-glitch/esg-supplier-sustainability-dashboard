# ESG Supplier Sustainability Dashboard

A Streamlit decision-support dashboard for evidence-based semiconductor supplier sustainability assessment using two-level AHP weights and benefit-type TOPSIS ranking.

## Features

- Executive overview with assessment completeness and headline results
- Supplier comparison with heatmap, radar chart, and E/S/G pillar scores
- AHP pillar and global indicator weights
- AHP consistency-ratio checks
- TOPSIS ranking recalculated in Python
- Optional Evidence Register / Evidence Explorer interface
- Session-only score editing without changing the source workbook

## Project files

- `app.py` — Streamlit application and calculation logic
- `requirements.txt` — Python dependencies
- `.gitignore` — excludes local environments, caches, spreadsheets, and data files

The assessment workbook is intentionally not included in this repository. Upload it through the dashboard or configure `ESG_WORKBOOK_PATH`.

## Requirements

- Python 3.10 or newer
- A compatible AHP–TOPSIS Excel workbook

## Run locally

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m streamlit run app.py
```

Open [http://localhost:8501](http://localhost:8501) after the server starts.

On Windows, use:

```powershell
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python -m streamlit run app.py
```

## Workbook structure

The workbook must contain these exact sheet names:

- `Level1_AHP`
- `E_AHP`
- `S_AHP`
- `G_AHP`
- `Rubric`
- `TOPSIS`

The `Level1_AHP` sheet provides the final indicator framework and global weights. The `Rubric` sheet provides the evidence-based 0–5 scoring rules. The `TOPSIS` sheet provides supplier names and raw scores for the 12 indicators.

## Indicator codes

| Pillar | Code | Indicator |
|---|---|---|
| Environment | E1 | GHG & Carbon Footprint |
| Environment | E2 | Energy Transition |
| Environment | E3 | Water Management |
| Environment | E4 | Waste & Hazardous Waste |
| Environment | E5 | Chemical Management |
| Social | S1 | Occupational Health & Safety |
| Social | S2 | Labor Rights |
| Social | S3 | Human Rights & Training |
| Governance | G1 | ESG Transparency |
| Governance | G2 | Ethics & Anti-Corruption |
| Governance | G3 | Compliance & Risk Management |
| Governance | G4 | Traceability |

## Using the dashboard

1. Start the app and upload a compatible `.xlsx` workbook from the sidebar if no default workbook is found.
2. Enter integer scores from 0 to 5 in the working score table, or populate the TOPSIS input area in Excel before uploading.
3. Complete all 12 indicators for at least two suppliers before interpreting the TOPSIS ranking.
4. Use **Executive Overview** for headline results and completeness.
5. Use **Supplier Comparison** for indicator and pillar-level comparisons.
6. Use **AHP Weights** to inspect priorities and consistency checks.
7. Use **TOPSIS Ranking** for closeness coefficients and final rank.
8. Use **Evidence Explorer** to review an optional Evidence Register.

Session edits update the dashboard immediately but are not written back to Excel. Save permanent scores in the workbook.

## Evidence Register interface

The dashboard accepts either an `Evidence_Register` worksheet or a CSV containing fields such as:

- Supplier
- Code
- Indicator
- Evidence Type
- Source / URL
- Reporting Year
- Evidence Summary
- Score
- Verification Status
- Reviewer Notes
- Date Accessed

The Evidence Explorer also provides a downloadable blank CSV template.

## TOPSIS inclusion rule

All 12 indicators are treated as benefit criteria. Only suppliers with complete scores are included in the ranking. At least two fully scored suppliers are required for a meaningful comparison.

## Data privacy

Spreadsheet and CSV files are ignored by Git. Review any evidence and supplier data before sharing or deploying the application.
