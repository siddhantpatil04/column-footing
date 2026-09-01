from __future__ import annotations

from dataclasses import dataclass, asdict
from typing import Any, Dict, List, Optional


@dataclass(frozen=True)
class ProjectInfo:
    project: str = ""
    client: str = ""
    structure: str = "COLUMN FOOTING"
    document_no: str = ""
    revision: str = "R0"
    prepared_by: str = ""
    checked_by: str = ""
    approved_by: str = ""


@dataclass(frozen=True)
class DesignInputs:
    sbc_kn_m2: float = 90.0
    design_load_kn: float = 251.96  # Workbook E9: already 1.5(DL+LL)
    mux_knm: float = 5.04
    muy_knm: float = 21.77
    column_width_mm: float = 380.0  # b
    column_depth_mm: float = 380.0  # D
    fck_mpa: float = 30.0
    fy_mpa: float = 500.0
    footing_length_mm: float = 1650.0
    footing_depth_mm: float = 350.0
    cover_mm: float = 50.0
    bar_dia_mm: float = 10.0
    bars_x: int = 10
    bars_y: int = 10
    min_reinf_method: str = "Existing Excel method"
    bar_type: str = "Deformed"

    def as_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class CheckResult:
    check_id: str
    name: str
    status: str
    demand: Optional[float]
    capacity: Optional[float]
    unit: str
    note: str = ""


@dataclass(frozen=True)
class FormulaTrace:
    name: str
    formula: str
    substitution: str
    result: float | str
    unit: str
    excel_source: str
    code_source: str = ""


@dataclass
class DesignResult:
    values: Dict[str, float | str]
    checks: List[CheckResult]
    trace: List[FormulaTrace]
    overall_status: str
    governing_failures: List[str]
    assumptions: List[str]

    def check_map(self) -> Dict[str, CheckResult]:
        return {c.check_id: c for c in self.checks}
