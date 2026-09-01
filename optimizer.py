from __future__ import annotations

from dataclasses import replace
from typing import Dict, Tuple

from engine import calculate
from models import DesignInputs, DesignResult


def recommend_safe_design(inp: DesignInputs, allow: Dict[str, bool], max_iterations: int = 120) -> Tuple[DesignInputs | None, DesignResult | None, list[str]]:
    """Greedy safe-design search. Every candidate is verified by the full central engine."""
    candidate = inp
    history: list[str] = []

    for _ in range(max_iterations):
        result = calculate(candidate)
        if result.overall_status == "SAFE":
            return candidate, result, history

        failed = {c.check_id for c in result.checks if c.status == "UNSAFE"}
        changed = False

        bearing_fail = bool(failed & {"CHK-AREA", "CHK-BEAR-X-U", "CHK-BEAR-X-W", "CHK-BEAR-Y-U", "CHK-BEAR-Y-W", "CHK-CONTACT-X", "CHK-CONTACT-Y"})
        depth_fail = bool(failed & {"CHK-DEPTH", "CHK-PUNCH", "CHK-1WAY-X", "CHK-1WAY-Y", "CHK-COL-BEAR"})
        x_rebar_fail = bool(failed & {"CHK-AST-X", "CHK-SP-X"})
        y_rebar_fail = bool(failed & {"CHK-AST-Y", "CHK-SP-Y"})

        if bearing_fail and allow.get("footing_length", False):
            candidate = replace(candidate, footing_length_mm=candidate.footing_length_mm + 50.0)
            history.append("Increased footing length by 50 mm; width followed equal-projection linkage.")
            changed = True

        if depth_fail and allow.get("footing_depth", False):
            candidate = replace(candidate, footing_depth_mm=candidate.footing_depth_mm + 25.0)
            history.append("Increased footing depth by 25 mm.")
            changed = True

        if x_rebar_fail and allow.get("bars_x", False):
            candidate = replace(candidate, bars_x=candidate.bars_x + 1)
            history.append("Increased X-direction bottom bars by 1.")
            changed = True

        if y_rebar_fail and allow.get("bars_y", False):
            candidate = replace(candidate, bars_y=candidate.bars_y + 1)
            history.append("Increased Y-direction bottom bars by 1.")
            changed = True

        # Coupled effect: depth increases minimum steel under slab method.
        if changed and allow.get("bars_x", False) and failed & {"CHK-AST-X", "CHK-SP-X"}:
            pass
        if not changed:
            return None, None, history

    return None, None, history
