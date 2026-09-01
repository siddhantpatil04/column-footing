# Workbook Audit Summary

Source workbook: `COLUMN FOOTING.xlsx`

- Visible worksheet: `FOOTING`.
- Visible working scope: all visible cells in the sheet.
- Hidden rows/cells: excluded by user instruction.
- No visible formula in the reviewed calculation chain requires the excluded hidden rows.
- Visible calculation sequence reviewed through footing sizing, soil pressures, two-way bending/depth, punching shear, reinforcement, development length, one-way shear and column-footing bearing.
- Legacy external/broken defined names are not reproduced in the application engine because they do not feed the visible calculation chain.

The central engine is mapped back to workbook cells in `source_map.py`.
