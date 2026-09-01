from __future__ import annotations

import math


def xu_max_ratio(fy_mpa: float) -> float:
    """IS 456 limiting neutral-axis depth ratio for supported reinforcement grades."""
    if fy_mpa <= 250:
        return 0.53
    if fy_mpa <= 415:
        return 0.48
    if fy_mpa <= 500:
        return 0.46
    # The source workbook is Fe500. Do not silently extrapolate beyond its code basis.
    raise ValueError("Steel grade above Fe500 is not implemented in this app.")


def ru_max(fck_mpa: float, fy_mpa: float) -> float:
    x = xu_max_ratio(fy_mpa)
    return 0.36 * fck_mpa * x * (1.0 - 0.42 * x)


def plain_bar_bond_stress_tension(fck_mpa: float) -> float:
    """IS 456 Table 26 design bond stress for plain bars in tension."""
    if fck_mpa < 20:
        raise ValueError("Concrete grade below M20 is not supported.")
    if fck_mpa < 25:
        return 1.2
    if fck_mpa < 30:
        return 1.4
    if fck_mpa < 35:
        return 1.5
    if fck_mpa < 40:
        return 1.7
    return 1.9


def design_bond_stress_tension(fck_mpa: float, bar_type: str = "Deformed") -> float:
    tau = plain_bar_bond_stress_tension(fck_mpa)
    if bar_type.lower().startswith("deform"):
        tau *= 1.60
    return tau


def min_slab_reinf_ratio(fy_mpa: float) -> float:
    # IS 456 solid-slab minimum: HYSD 0.12%, mild steel 0.15%.
    return 0.0015 if fy_mpa <= 250 else 0.0012


def one_way_shear_strength(fck_mpa: float, pt_percent: float) -> float:
    """Workbook's IS 456 shear expression retained as the source method."""
    if pt_percent <= 0:
        return 0.0
    beta = (0.8 * fck_mpa) / (6.89 * pt_percent)
    return (0.85 * math.sqrt(0.8 * fck_mpa) * (math.sqrt(1.0 + 5.0 * beta) - 1.0)) / (6.0 * beta)


def max_main_bar_spacing_mm(effective_depth_mm: float) -> float:
    return min(3.0 * effective_depth_mm, 300.0)
