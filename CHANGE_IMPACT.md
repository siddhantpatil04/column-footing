# Change Impact

This is the first reconciled Streamlit build for the uploaded `COLUMN FOOTING.xlsx` workbook.

## Files created

- `app.py` - new Streamlit UI and state workflow.
- `models.py` - canonical input/result models.
- `engine.py` - single-source engineering calculation engine.
- `code_basis.py` - material/code provisions.
- `optimizer.py` - authorised-variable SAFE-design optimiser.
- `reports.py` - PDF and Excel reporting.
- `source_map.py` - Excel source mapping.
- `excel_baseline.py` - workbook baseline comparison values only.
- documentation and tests - new.

## Engineering impact

Approved QD corrections have been implemented. Items marked IGNORE remain on workbook logic. QD-013 is user-selectable between the existing Excel method and the IS 456 footing/solid-slab method.

No hidden-cell calculation has been introduced.

## UI update — SBR WALL W1 visual system

### Files changed
- `app.py` — changed: visual theme/layout styling only.
- `CHANGE_IMPACT.md` — changed: records this UI-only modification.
- `engine.py` — unchanged.
- `models.py` — unchanged.
- `code_basis.py` — unchanged.
- `optimizer.py` — unchanged.
- `reports.py` — unchanged.
- `source_map.py` — unchanged.
- `excel_baseline.py` — unchanged.

### Engineering impact
UI-only modification. No engineering formula, QD reconciliation, SAFE/UNSAFE rule, optimiser calculation, PDF calculation content, Excel calculation content, source mapping, or workbook baseline value was changed.

The Column Footing app now follows the SBR WALL W1 interface language: dark charcoal canvas, #24252e sidebar, #262730 inputs, #ff4b4b primary actions, bordered metric cards/tables/expanders, coral active tabs, SAFE/UNSAFE/stale status panels, and dark graphical input cards.
