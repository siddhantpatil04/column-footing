from __future__ import annotations

import math
from dataclasses import replace
from typing import Dict, List

from code_basis import (
    design_bond_stress_tension,
    max_main_bar_spacing_mm,
    min_slab_reinf_ratio,
    one_way_shear_strength,
    ru_max,
)
from models import CheckResult, DesignInputs, DesignResult, FormulaTrace


def _check(check_id: str, name: str, ok: bool, demand, capacity, unit: str, note: str = "") -> CheckResult:
    return CheckResult(check_id, name, "SAFE" if ok else "UNSAFE", demand, capacity, unit, note)


def _flexural_ast(mu_knm: float, fck: float, fy: float, b_mm: float, d_mm: float) -> float:
    if b_mm <= 0 or d_mm <= 0:
        return float("inf")
    inside = 1.0 - (4.6 * mu_knm * 1e6) / (fck * b_mm * d_mm**2)
    if inside <= 0:
        return float("inf")
    return (1.0 - math.sqrt(inside)) * ((0.5 * fck) / fy) * (b_mm * d_mm)


def _spacing(clear_dim_mm: float, cover_mm: float, dia_mm: float, n: int) -> float:
    if n <= 1:
        return float("inf")
    available_centres = clear_dim_mm - 2.0 * (cover_mm + dia_mm / 2.0)
    if available_centres <= 0:
        return float("inf")
    return available_centres / (n - 1)


def calculate(inp: DesignInputs) -> DesignResult:
    if inp.sbc_kn_m2 <= 0 or inp.design_load_kn <= 0:
        raise ValueError("SBC and design load must be positive.")
    if inp.column_width_mm <= 0 or inp.column_depth_mm <= 0:
        raise ValueError("Column dimensions must be positive.")
    if inp.footing_length_mm <= 0 or inp.footing_depth_mm <= 0:
        raise ValueError("Footing dimensions must be positive.")
    if inp.cover_mm < 0 or inp.bar_dia_mm <= 0:
        raise ValueError("Cover/bar diameter is invalid.")

    values: Dict[str, float | str] = {}
    trace: List[FormulaTrace] = []
    checks: List[CheckResult] = []

    # QD-012: live Ru.max.
    ru = ru_max(inp.fck_mpa, inp.fy_mpa)
    values["ru_max_mpa"] = ru

    # QD-010 retained: self-weight = 10% of design load.
    sw = 0.10 * inp.design_load_kn
    total_load = inp.design_load_kn + sw
    values.update(self_weight_kn=sw, total_load_kn=total_load)

    area_req_m2 = total_load / (inp.sbc_kn_m2 * 1.5)
    delta = inp.column_depth_mm - inp.column_width_mm
    req_l = delta / 2.0 + math.sqrt((delta**2) / 4.0 + area_req_m2 * 1e6)
    req_b = req_l - inp.column_depth_mm + inp.column_width_mm

    # QD-016: equal projection linkage for a rectangular column.
    bf = inp.footing_length_mm - inp.column_depth_mm + inp.column_width_mm
    lf = inp.footing_length_mm
    if bf <= 0:
        raise ValueError("Derived footing width is not positive. Increase footing length.")
    if lf <= inp.column_depth_mm or bf <= inp.column_width_mm:
        raise ValueError("Provided footing plan dimensions must exceed the corresponding column dimensions.")
    area_prov_m2 = lf * bf / 1e6
    values.update(required_area_m2=area_req_m2, required_length_mm=req_l,
                  required_width_mm=req_b, footing_width_mm=bf, provided_area_m2=area_prov_m2)
    checks.append(_check("CHK-AREA", "Provided footing area", area_prov_m2 >= area_req_m2,
                         area_req_m2, area_prov_m2, "m²", "Demand = required area; capacity = provided area."))

    pd = total_load / area_prov_m2
    z_x_m3 = (bf / 1000.0) * (lf / 1000.0) ** 2 / 6.0
    z_y_m3 = (lf / 1000.0) * (bf / 1000.0) ** 2 / 6.0
    pbx = inp.mux_knm / z_x_m3
    pby = inp.muy_knm / z_y_m3
    pmax_x, pmin_x = pd + pbx, pd - pbx
    pmax_y, pmin_y = pd + pby, pd - pby

    # QD-002: both working minimum pressures are now live.
    pmax_x_w, pmin_x_w = pmax_x / 1.5, pmin_x / 1.5
    pmax_y_w, pmin_y_w = pmax_y / 1.5, pmin_y / 1.5
    values.update(pd_kn_m2=pd, pbx_kn_m2=pbx, pby_kn_m2=pby,
                  pmax_x_kn_m2=pmax_x, pmin_x_kn_m2=pmin_x,
                  pmax_y_kn_m2=pmax_y, pmin_y_kn_m2=pmin_y,
                  pmax_x_working_kn_m2=pmax_x_w, pmin_x_working_kn_m2=pmin_x_w,
                  pmax_y_working_kn_m2=pmax_y_w, pmin_y_working_kn_m2=pmin_y_w)

    checks += [
        _check("CHK-BEAR-X-U", "Factored bearing pressure — X", pmax_x <= 1.5 * inp.sbc_kn_m2, pmax_x, 1.5 * inp.sbc_kn_m2, "kN/m²"),
        _check("CHK-BEAR-X-W", "Working bearing pressure — X", pmax_x_w <= inp.sbc_kn_m2, pmax_x_w, inp.sbc_kn_m2, "kN/m²"),
        _check("CHK-BEAR-Y-U", "Factored bearing pressure — Y", pmax_y <= 1.5 * inp.sbc_kn_m2, pmax_y, 1.5 * inp.sbc_kn_m2, "kN/m²"),
        _check("CHK-BEAR-Y-W", "Working bearing pressure — Y", pmax_y_w <= inp.sbc_kn_m2, pmax_y_w, inp.sbc_kn_m2, "kN/m²"),
    ]

    # QD-011: separate no-tension/eccentricity checks, consistent with QD-001 decision.
    ex_mm = abs(inp.mux_knm / total_load) * 1000.0
    ey_mm = abs(inp.muy_knm / total_load) * 1000.0
    values.update(eccentricity_x_mm=ex_mm, eccentricity_y_mm=ey_mm,
                  eccentricity_limit_x_mm=lf / 6.0, eccentricity_limit_y_mm=bf / 6.0)
    checks += [
        _check("CHK-CONTACT-X", "Full contact / no tension — X", pmin_x >= 0 and ex_mm <= lf / 6.0,
               ex_mm, lf / 6.0, "mm", f"pmin,x = {pmin_x:.3f} kN/m²"),
        _check("CHK-CONTACT-Y", "Full contact / no tension — Y", pmin_y >= 0 and ey_mm <= bf / 6.0,
               ey_mm, bf / 6.0, "mm", f"pmin,y = {pmin_y:.3f} kN/m²"),
    ]

    proj_x = (lf - inp.column_depth_mm) / 2.0
    proj_y = (bf - inp.column_width_mm) / 2.0
    values.update(projection_x_mm=proj_x, projection_y_mm=proj_y)

    # QD-019 retained: pressure at face equals edge maximum as in workbook.
    top_b = inp.column_width_mm + 100.0
    top_l = inp.column_depth_mm + 100.0
    mux1 = pmax_x * bf * (proj_x**2 / 2.0) / 1e9
    muy1 = pmax_y * lf * (proj_y**2 / 2.0) / 1e9
    req_dx = math.sqrt(mux1 * 1e6 / (ru * top_b))
    req_dy = math.sqrt(muy1 * 1e6 / (ru * top_l))
    values.update(moment_x_knm=mux1, moment_y_knm=muy1,
                  required_eff_depth_x_mm=req_dx, required_eff_depth_y_mm=req_dy,
                  required_effective_depth_mm=max(req_dx, req_dy))

    # QD-003: nomenclature only. Workbook comparison is preserved.
    checks.append(_check("CHK-DEPTH", "Footing depth vs required effective depth",
                         inp.footing_depth_mm >= max(req_dx, req_dy), max(req_dx, req_dy), inp.footing_depth_mm, "mm",
                         "QD-003: workbook formula retained; label corrected to required effective depth."))

    dx = inp.footing_depth_mm - (inp.cover_mm + inp.bar_dia_mm / 2.0)
    dy = inp.footing_depth_mm - (inp.cover_mm + 1.5 * inp.bar_dia_mm)
    if dx <= 0 or dy <= 0:
        raise ValueError("Footing depth is too small for the selected cover/bar arrangement.")
    values.update(effective_depth_x_mm=dx, effective_depth_y_mm=dy)

    # Two-way shear. QD-018: column/total load minus soil reaction within critical perimeter.
    crit_perimeter = 2.0 * ((inp.column_width_mm + dy) + (inp.column_depth_mm + dy))
    shear_area = crit_perimeter * dy
    tau_uc_prime = 0.25 * math.sqrt(inp.fck_mpa)
    ks = min(0.5 + inp.column_width_mm / inp.column_depth_mm, 1.0)
    tau_punch = ks * tau_uc_prime
    vuc2 = tau_punch * shear_area / 1000.0
    inside_area_m2 = (inp.column_width_mm + dy) * (inp.column_depth_mm + dy) / 1e6
    reaction_inside = pd * inside_area_m2
    vud2 = max(0.0, total_load - reaction_inside)
    values.update(punching_perimeter_mm=crit_perimeter, punching_tau_mpa=tau_punch,
                  punching_capacity_kn=vuc2, punching_inside_reaction_kn=reaction_inside,
                  punching_demand_kn=vud2)
    checks.append(_check("CHK-PUNCH", "Two-way (punching) shear", vuc2 >= vud2, vud2, vuc2, "kN"))

    ast_x_req = _flexural_ast(mux1, inp.fck_mpa, inp.fy_mpa, top_b, dx)
    ast_y_req = _flexural_ast(muy1, inp.fck_mpa, inp.fy_mpa, top_l, dy)

    if inp.min_reinf_method == "IS 456 footing / solid slab method":
        ratio = min_slab_reinf_ratio(inp.fy_mpa)
        ast_x_min = ratio * bf * inp.footing_depth_mm
        ast_y_min = ratio * lf * inp.footing_depth_mm
        min_formula = f"{ratio:.4f} × gross footing section"
        min_code = "IS 456:2000 Clauses 34.5.1 + 26.5.2.1"
    else:
        # QD-013 legacy option: exact Excel beam-style expression.
        ast_x_min = 0.85 * top_b * dx / inp.fy_mpa
        ast_y_min = 0.85 * top_l * dy / inp.fy_mpa
        min_formula = "0.85 b d / fy"
        min_code = "Existing Excel method (legacy selectable option)"

    # QD-004: correct directional references.
    ast_x_to_provide = max(ast_x_req, ast_x_min)
    ast_y_to_provide = max(ast_y_req, ast_y_min)
    area_bar = math.pi / 4.0 * inp.bar_dia_mm**2
    ast_x_prov = inp.bars_x * area_bar
    ast_y_prov = inp.bars_y * area_bar

    # QD-015: proper centre-to-centre spacing. X bars are distributed across Bf; Y across Lf.
    spacing_x = _spacing(bf, inp.cover_mm, inp.bar_dia_mm, inp.bars_x)
    spacing_y = _spacing(lf, inp.cover_mm, inp.bar_dia_mm, inp.bars_y)
    max_spacing_x = max_main_bar_spacing_mm(dx)
    max_spacing_y = max_main_bar_spacing_mm(dy)
    values.update(ast_x_req_mm2=ast_x_req, ast_y_req_mm2=ast_y_req,
                  ast_x_min_mm2=ast_x_min, ast_y_min_mm2=ast_y_min,
                  ast_x_to_provide_mm2=ast_x_to_provide, ast_y_to_provide_mm2=ast_y_to_provide,
                  ast_x_provided_mm2=ast_x_prov, ast_y_provided_mm2=ast_y_prov,
                  spacing_x_mm=spacing_x, spacing_y_mm=spacing_y,
                  max_spacing_x_mm=max_spacing_x, max_spacing_y_mm=max_spacing_y,
                  min_reinf_method=inp.min_reinf_method)
    checks += [
        _check("CHK-AST-X", "Bottom reinforcement — X", ast_x_prov >= ast_x_to_provide, ast_x_to_provide, ast_x_prov, "mm²"),
        _check("CHK-AST-Y", "Bottom reinforcement — Y", ast_y_prov >= ast_y_to_provide, ast_y_to_provide, ast_y_prov, "mm²"),
        _check("CHK-SP-X", "Bottom bar spacing — X", spacing_x <= max_spacing_x, spacing_x, max_spacing_x, "mm c/c"),
        _check("CHK-SP-Y", "Bottom bar spacing — Y", spacing_y <= max_spacing_y, spacing_y, max_spacing_y, "mm c/c"),
    ]

    # QD-007: live bond stress. QD-006: Ld remains INFO only, not an overall mandatory check.
    tau_bd = design_bond_stress_tension(inp.fck_mpa, inp.bar_type)
    ld_req = 0.87 * inp.fy_mpa * inp.bar_dia_mm / (4.0 * tau_bd)
    ld_avail = proj_x - 75.0
    values.update(bond_stress_mpa=tau_bd, ld_required_mm=ld_req, ld_available_mm=ld_avail,
                  development_length_status="INFO")

    # One-way shear, using workbook's shear-strength expression.
    pt_y = 100.0 * ast_y_prov / (lf * dy)
    tau_y = one_way_shear_strength(inp.fck_mpa, pt_y)
    vuc_y = tau_y * lf * dy / 1000.0
    # QD-005: Y check uses Y pressure.
    vud_y = pmax_y * lf * max(0.0, proj_y - dy) * 1e-6

    pt_x = 100.0 * ast_x_prov / (bf * dx)
    tau_x = one_way_shear_strength(inp.fck_mpa, pt_x)
    vuc_x = tau_x * bf * dx / 1000.0
    # QD-005: X check uses X pressure.
    vud_x = pmax_x * bf * max(0.0, proj_x - dx) * 1e-6
    values.update(pt_y_percent=pt_y, tau_c_y_mpa=tau_y, one_way_y_capacity_kn=vuc_y, one_way_y_demand_kn=vud_y,
                  pt_x_percent=pt_x, tau_c_x_mpa=tau_x, one_way_x_capacity_kn=vuc_x, one_way_x_demand_kn=vud_x)
    checks += [
        _check("CHK-1WAY-Y", "One-way shear — Y", vuc_y >= vud_y, vud_y, vuc_y, "kN"),
        _check("CHK-1WAY-X", "One-way shear — X", vuc_x >= vud_x, vud_x, vuc_x, "kN"),
    ]

    # QD-009: no second 1.5 factor at column base.
    bearing_demand = total_load / (inp.column_width_mm * inp.column_depth_mm) * 1000.0
    a1_footing = lf * bf
    a1_dispersion = (inp.column_width_mm + 4.0 * inp.footing_depth_mm) * (inp.column_depth_mm + 4.0 * inp.footing_depth_mm)
    a1 = min(a1_footing, a1_dispersion)
    a2 = inp.column_width_mm * inp.column_depth_mm
    sqrt_ratio = math.sqrt(a1 / a2)
    # QD-008 verified correction: cap at 2.0.
    enhancement = min(sqrt_ratio, 2.0)
    bearing_capacity = 0.45 * inp.fck_mpa * enhancement
    values.update(column_bearing_demand_mpa=bearing_demand, bearing_a1_mm2=a1,
                  bearing_a2_mm2=a2, bearing_sqrt_ratio=sqrt_ratio,
                  bearing_enhancement=enhancement, column_bearing_capacity_mpa=bearing_capacity)
    checks.append(_check("CHK-COL-BEAR", "Column-footing bearing", bearing_demand <= bearing_capacity,
                         bearing_demand, bearing_capacity, "N/mm²"))

    # Formula trace — key calculations only, all driven by the same values.
    trace.extend([
        FormulaTrace("Footing self weight", "Wf = 0.10 P", f"0.10 × {inp.design_load_kn:.3f}", sw, "kN", "FOOTING!E18", "Workbook assumption retained (QD-010)"),
        FormulaTrace("Required footing area", "Areq = Ptotal / (1.5 SBC)", f"{total_load:.3f} / (1.5 × {inp.sbc_kn_m2:.3f})", area_req_m2, "m²", "FOOTING!E22"),
        FormulaTrace("Linked footing width", "Bf = Lf - D + b", f"{lf:.1f} - {inp.column_depth_mm:.1f} + {inp.column_width_mm:.1f}", bf, "mm", "FOOTING!E32", "QD-016 approved correction"),
        FormulaTrace("Factored base pressure", "pd = Ptotal / Af", f"{total_load:.3f} / {area_prov_m2:.6f}", pd, "kN/m²", "FOOTING!E38"),
        FormulaTrace("Ru,max", "0.36 fck xu(1-0.42xu)", f"fck={inp.fck_mpa:.0f}, fy={inp.fy_mpa:.0f}", ru, "N/mm²", "FOOTING!E16", "QD-012 approved correction"),
        FormulaTrace("Required effective depth X", "dreq = √(M/(Ru b))", f"√({mux1:.4f}×10^6 / ({ru:.4f}×{top_b:.1f}))", req_dx, "mm", "FOOTING!E98/E117", "QD-003 relabel only"),
        FormulaTrace("Required effective depth Y", "dreq = √(M/(Ru b))", f"√({muy1:.4f}×10^6 / ({ru:.4f}×{top_l:.1f}))", req_dy, "mm", "FOOTING!E101/E119", "QD-003 relabel only"),
        FormulaTrace("Punching shear demand", "Vu = Ptotal - pd × Ainside", f"{total_load:.3f} - {pd:.3f}×{inside_area_m2:.6f}", vud2, "kN", "FOOTING!E163", "QD-018 approved correction"),
        FormulaTrace("Minimum steel X", min_formula, f"method={inp.min_reinf_method}", ast_x_min, "mm²", "FOOTING!E173", min_code),
        FormulaTrace("Minimum steel Y", min_formula, f"method={inp.min_reinf_method}", ast_y_min, "mm²", "FOOTING!E175", min_code),
        FormulaTrace("Development length", "Ld = 0.87 fy φ / (4 τbd)", f"0.87×{inp.fy_mpa:.0f}×{inp.bar_dia_mm:.0f}/(4×{tau_bd:.3f})", ld_req, "mm", "FOOTING!E195", "QD-007; QD-006 remains INFO"),
        FormulaTrace("Column bearing demand", "σ = Ptotal / (bD)", f"{total_load:.3f}×1000/({inp.column_width_mm:.0f}×{inp.column_depth_mm:.0f})", bearing_demand, "N/mm²", "FOOTING!E236", "QD-009 approved correction"),
        FormulaTrace("Column bearing capacity", "σbr = 0.45 fck min(√(A1/A2),2)", f"0.45×{inp.fck_mpa:.0f}×{enhancement:.3f}", bearing_capacity, "N/mm²", "FOOTING!E242:E245", "QD-008 verified correction"),
    ])

    mandatory = [c for c in checks if c.status in ("SAFE", "UNSAFE")]
    failures = [c.name for c in mandatory if c.status == "UNSAFE"]
    overall = "SAFE" if not failures else "UNSAFE"
    assumptions = [
        "Visible FOOTING sheet is the application scope; hidden rows/cells are excluded.",
        "Input design load is already 1.5(DL+LL), matching workbook E9.",
        "Footing self-weight is retained as 10% of design load (QD-010).",
        "X and Y bearing pressures are checked separately; no combined biaxial corner-pressure check is introduced (QD-001).",
        "Pressure at column face is retained as the corresponding maximum edge pressure (QD-019).",
        "Development length is reported as INFO and does not govern overall status (QD-006).",
        "Footing width is linked to length by equal projections for rectangular columns (QD-016).",
    ]

    return DesignResult(values=values, checks=checks, trace=trace,
                        overall_status=overall, governing_failures=failures,
                        assumptions=assumptions)
