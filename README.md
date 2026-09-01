# COLUMN FOOTING Streamlit App

Professional Excel-derived isolated column-footing design application based on the visible `FOOTING` worksheet of `COLUMN FOOTING.xlsx`.

## Run locally

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Architecture

- `app.py` — Streamlit UI, manual Run workflow, stale-result invalidation, comparison and downloads.
- `models.py` — canonical structured inputs/results.
- `engine.py` — single-source engineering calculation engine.
- `code_basis.py` — code/material provisions used by the engine.
- `optimizer.py` — authorised-variable safe-design search; all candidates rerun `engine.calculate()`.
- `reports.py` — PDF + Excel report generation using the same result object.
- `source_map.py` — Excel cell-to-engine mapping.
- `excel_baseline.py` — baseline workbook values for Excel-vs-Web comparison only.
- `correction_notes.md` — approved QD decisions.
- `tests/` — regression tests.

## Scope

All visible content of the `FOOTING` sheet is in scope. Hidden rows/cells are excluded and are not used by the engine.

## Important status wording

`SAFE` means safe for all mandatory checks implemented in this application and within the visible workbook scope. Development length is displayed as information-only per QD-006.

## Deployment

Push this folder to GitHub and select `app.py` in Streamlit Community Cloud. Future changes normally require replacing only the files listed in the change-impact note, not redeploying from scratch.

## UI reference

The Streamlit UI has been aligned to the established SBR WALL W1 design-app visual system: dark engineering dashboard, charcoal sidebar, coral primary actions, compact bordered cards/tables, styled SAFE/UNSAFE/stale status panels, and consistent tab/input styling. This is a UI-only adaptation; the Column Footing engineering engine remains independent and unchanged.
