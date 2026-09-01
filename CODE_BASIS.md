# Code Basis

Primary code basis: **IS 456:2000** as represented by the uploaded workbook and approved reconciliation decisions.

Implemented code-linked provisions include:

- limiting neutral-axis depth ratio and derived `Ru,max` for Fe250 / Fe415 / Fe500;
- design bond stress by concrete grade, with 60% increase for deformed bars in tension;
- footing-column bearing enhancement `0.45 fck sqrt(A1/A2)` with enhancement capped at 2.0;
- optional solid-slab minimum reinforcement ratio for footing reinforcement;
- maximum main reinforcement spacing `min(3d, 300 mm)`;
- workbook one-way shear-strength expression retained as the source calculation method.

## QD-013 selectable minimum reinforcement

1. **Existing Excel method**: `Ast,min = 0.85 b d / fy`.
2. **IS 456 footing / solid slab method**: 0.12% of gross section for HYSD reinforcement, or 0.15% for mild steel.

The selected method is passed through the same central engine and therefore changes reinforcement results, mandatory checks, optimiser verification, PDF report and Excel report consistently.
