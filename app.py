from __future__ import annotations

import hashlib
import json
from dataclasses import asdict

import pandas as pd
import streamlit as st

from engine import calculate
from models import DesignInputs, ProjectInfo
from optimizer import recommend_safe_design
from reports import build_excel, build_pdf

st.set_page_config(
    page_title="COLUMN FOOTING Design",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
:root { color-scheme: dark; }
html, body, [class*="css"] { font-family: "Segoe UI", Arial, sans-serif; }
.stApp { background: #0e1117; color: #f3f4f6; }
.block-container {
    padding-top: 2.0rem;
    padding-bottom: 3rem;
    max-width: 1480px;
}
[data-testid="stSidebar"] {
    background: #24252e;
    border-right: 1px solid #343640;
}
[data-testid="stSidebar"] > div:first-child { padding-top: 1.0rem; }
[data-testid="stSidebar"] .block-container { padding-top: 0.8rem; }

h1, h2, h3 { color: #ffffff; letter-spacing: -0.02em; }
h1 { font-size: 2.15rem !important; margin-bottom: .2rem !important; }
h2 { font-size: 1.72rem !important; margin-top: 1.2rem !important; }
h3 { font-size: 1.15rem !important; }
p, label, .stCaption { color: #d1d5db; }

[data-baseweb="input"] > div,
[data-baseweb="select"] > div,
[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input {
    background: #262730 !important;
    color: #ffffff !important;
    border-color: #30323c !important;
    border-radius: 7px !important;
}
[data-testid="stTextInput"] input:focus,
[data-testid="stNumberInput"] input:focus {
    border-color: #ff4b4b !important;
    box-shadow: 0 0 0 1px #ff4b4b !important;
}
[data-testid="stWidgetLabel"] p {
    color: #ffffff !important;
    font-size: .82rem !important;
    font-weight: 600 !important;
}

.stButton > button[kind="primary"],
.stDownloadButton > button[kind="primary"] {
    background: #ff4b4b !important;
    border: 1px solid #ff4b4b !important;
    color: #ffffff !important;
    font-weight: 700 !important;
    border-radius: 7px !important;
}
.stButton > button[kind="primary"]:hover,
.stDownloadButton > button[kind="primary"]:hover {
    background: #ff6262 !important;
    border-color: #ff6262 !important;
}
.stButton > button:not([kind="primary"]),
.stDownloadButton > button:not([kind="primary"]) {
    background: #171a21;
    border: 1px solid #343840;
    color: #f3f4f6;
    border-radius: 7px;
}

.st-key-apply_safe_design button {
    background: #15803d !important;
    border: 1px solid #22c55e !important;
    color: #ffffff !important;
    font-weight: 750 !important;
}
.st-key-apply_safe_design button:hover {
    background: #16a34a !important;
    border-color: #4ade80 !important;
}

[data-testid="stExpander"] {
    border: 1px solid #343840 !important;
    background: #11151c !important;
    border-radius: 8px !important;
}
[data-testid="stDataFrame"], [data-testid="stDataEditor"] {
    border: 1px solid #2f333d;
    border-radius: 8px;
    overflow: hidden;
}

[data-testid="stMetric"] {
    background: #11151c;
    border: 1px solid #343840;
    border-radius: 8px;
    padding: .85rem .95rem;
    min-height: 96px;
}
[data-testid="stMetricLabel"] { color: #e5e7eb; }
[data-testid="stMetricValue"] { color: #ffffff; }

button[data-baseweb="tab"] { color: #d1d5db !important; font-weight: 600; }
button[data-baseweb="tab"][aria-selected="true"] { color: #ff5a5f !important; }
[data-baseweb="tab-highlight"] { background-color: #ff4b4b !important; }

.section-divider {
    height: 1px;
    background: #343840;
    margin: 1.65rem 0 1.55rem 0;
}
.section-kicker {
    color: #ffffff;
    font-size: 1.72rem;
    line-height: 1.2;
    font-weight: 750;
    margin: .25rem 0 1.05rem 0;
}
.app-subtitle {
    color: #8f969f;
    font-size: .80rem;
    margin-top: -.25rem;
    margin-bottom: 1.25rem;
}
.status-safe {
    background: #123d2c;
    border: 1px solid #1f6849;
    color: #55e69d;
    padding: .82rem 1rem;
    border-radius: 7px;
    font-weight: 650;
    margin-bottom: .85rem;
}
.status-unsafe {
    background: #431c22;
    border: 1px solid #74303a;
    color: #ff8289;
    padding: .82rem 1rem;
    border-radius: 7px;
    font-weight: 650;
    margin-bottom: .85rem;
}
.status-stale {
    background: #453411;
    border: 1px solid #795d1d;
    color: #f7cb62;
    padding: .82rem 1rem;
    border-radius: 7px;
    font-weight: 600;
    margin-bottom: .85rem;
}
.sidebar-title {
    color: #ffffff;
    font-weight: 750;
    font-size: 1.0rem;
    margin: .25rem 0 .65rem 0;
}
.sidebar-note {
    color: #a8aeb8;
    font-size: .74rem;
    line-height: 1.45;
}
.code-ok {
    background: #123d2c;
    border: 1px solid #1f6849;
    color: #69e6a8;
    padding: .55rem .65rem;
    border-radius: 7px;
    font-size: .76rem;
}
.small-note { color: #9299a4; font-size: .78rem; }

.input-mode-bar {
    background: #11151c; border: 1px solid #343840; border-radius: 10px;
    padding: .75rem .9rem; margin: .2rem 0 1rem 0;
}
.diagram-card {
    background: linear-gradient(180deg, #121720 0%, #0f131a 100%);
    border: 1px solid #343840;
    border-radius: 12px;
    padding: .9rem 1rem;
    min-height: 100%;
}
.diagram-title { color:#ffffff; font-size:1rem; font-weight:750; margin-bottom:.15rem; }
.diagram-subtitle { color:#8f969f; font-size:.76rem; margin-bottom:.65rem; line-height:1.35; }
.draftsman-tip {
    background:#102a3b; border:1px solid #1e5875; color:#b8e7ff;
    padding:.7rem .85rem; border-radius:8px; font-size:.80rem; line-height:1.45;
    margin:.25rem 0 .9rem 0;
}
.graphical-input-panel {
    background:#11151c; border:1px solid #2e333d; border-radius:10px; padding:.75rem .85rem;
}
.summary-card {
    background:#11151c;
    border:1px solid #343840;
    border-radius:12px;
    padding:.9rem 1rem;
    min-height:100%;
}
.lock-pill {
    display:inline-block; color:#c9d0d8; background:#1a2028; border:1px solid #343840;
    border-radius:999px; padding:.2rem .55rem; margin:.15rem .15rem .15rem 0; font-size:.72rem;
}
</style>
""",
    unsafe_allow_html=True,
)


def section_title(number: int, text: str) -> None:
    st.markdown(f'<div class="section-kicker">{number}. {text}</div>', unsafe_allow_html=True)


def divider() -> None:
    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)


def sig(inp: DesignInputs) -> str:
    return hashlib.sha256(json.dumps(inp.as_dict(), sort_keys=True).encode()).hexdigest()


def _widget_values(inp: DesignInputs) -> dict[str, object]:
    """Map a DesignInputs object to the exact Streamlit widget keys used by this UI."""
    return {
        "sbc": float(inp.sbc_kn_m2),
        "design_load": float(inp.design_load_kn),
        "mux": float(inp.mux_knm),
        "muy": float(inp.muy_knm),
        "column_b": float(inp.column_width_mm),
        "column_d": float(inp.column_depth_mm),
        "fck": int(inp.fck_mpa),
        "fy": int(inp.fy_mpa),
        "footing_length": float(inp.footing_length_mm),
        "footing_depth": float(inp.footing_depth_mm),
        "cover": float(inp.cover_mm),
        "bar_dia": int(inp.bar_dia_mm),
        "bars_x": int(inp.bars_x),
        "bars_y": int(inp.bars_y),
        "min_reinf_method": inp.min_reinf_method,
    }


def _queue_design_update(
    inp: DesignInputs,
    *,
    result=None,
    calc_signature: str | None = None,
    optimizer_original=None,
) -> None:
    """Queue widget mutations for the next rerun.

    Streamlit does not allow a widget's session-state key to be changed after the
    widget has been instantiated in the same run. The optimiser and restore
    actions therefore write to non-widget pending keys, rerun, and the pending
    values are applied before any design widgets are created.
    """
    st.session_state["_pending_design_values"] = _widget_values(inp)
    st.session_state["_pending_result"] = result
    st.session_state["_pending_calc_sig"] = calc_signature
    st.session_state["_pending_optimizer_original"] = optimizer_original
    st.session_state["_pending_apply"] = True


def _apply_pending_design_update() -> None:
    if not st.session_state.get("_pending_apply", False):
        return
    values = st.session_state.pop("_pending_design_values", {})
    for key, value in values.items():
        st.session_state[key] = value
    st.session_state.result = st.session_state.pop("_pending_result", None)
    st.session_state.calc_sig = st.session_state.pop("_pending_calc_sig", None)
    st.session_state.optimizer_original = st.session_state.pop("_pending_optimizer_original", None)
    st.session_state.recommended = None
    st.session_state.pop("_pending_apply", None)


def footing_svg(inp: DesignInputs, bf: float) -> str:
    l, b, cw, cd = inp.footing_length_mm, bf, inp.column_width_mm, inp.column_depth_mm
    scale = 330 / max(l, b)
    fw, fh, colw, colh = l*scale, b*scale, cd*scale, cw*scale
    x0, y0 = (420-fw)/2, (350-fh)/2
    cx, cy = 210-colw/2, 175-colh/2
    lf_y = y0 + fh + 35
    return f'''<svg viewBox="0 0 420 405" width="100%" xmlns="http://www.w3.org/2000/svg">
    <rect x="{x0:.1f}" y="{y0:.1f}" width="{fw:.1f}" height="{fh:.1f}" fill="#20262D" stroke="#CBD5E1" stroke-width="2"/>
    <rect x="{cx:.1f}" y="{cy:.1f}" width="{colw:.1f}" height="{colh:.1f}" fill="#E35D4F" stroke="#F8FAFC" stroke-width="2"/>
    <line x1="{x0:.1f}" y1="{y0+fh+16:.1f}" x2="{x0+fw:.1f}" y2="{y0+fh+16:.1f}" stroke="#94A3B8"/>
    <circle cx="145" cy="{lf_y-4:.1f}" r="9" fill="#ff4b4b"/><text x="145" y="{lf_y-1:.1f}" text-anchor="middle" fill="white" font-size="10" font-weight="800">1</text>
    <text x="255" y="{lf_y:.1f}" fill="#E5E7EB" text-anchor="middle" font-size="13">Lf = {l:.0f} mm</text>
    <text x="25" y="175" fill="#E5E7EB" text-anchor="middle" font-size="13" transform="rotate(-90 25 175)">Bf = {b:.0f} mm (derived)</text>
    <circle cx="145" cy="163" r="9" fill="#ff4b4b"/><text x="145" y="166" text-anchor="middle" fill="white" font-size="10" font-weight="800">2</text>
    <text x="245" y="182" fill="white" text-anchor="middle" font-size="12">COLUMN {cw:.0f} × {cd:.0f}</text>
    <text x="210" y="397" fill="#94A3B8" text-anchor="middle" font-size="11">NOT TO SCALE</text></svg>'''


def make_inputs(mode: str) -> DesignInputs:
    if mode == "🖼 Graphical Input":
        st.markdown(
            '<div class="draftsman-tip"><b>How to use this view:</b> read the numbered callouts on the diagram, '
            'then enter the matching values beside it. The diagram is intentionally simple and is <b>not to scale</b>.</div>',
            unsafe_allow_html=True,
        )
        st.markdown("#### Loads & Materials")
        c1,c2,c3,c4,c5,c6 = st.columns(6)
        sbc = c1.number_input("SBC (kN/m²)", min_value=1.0, value=90.0, step=5.0, key="sbc")
        p = c2.number_input("Design Load (kN)", min_value=1.0, value=251.96, step=1.0, key="design_load", help="1.5(DL+LL)")
        mux = c3.number_input("Mux (kNm)", value=5.04, step=0.5, key="mux")
        muy = c4.number_input("Muy (kNm)", value=21.77, step=0.5, key="muy")
        fck = float(c5.selectbox("Concrete", [20,25,30,35,40], index=2, format_func=lambda x:f"M{x}", key="fck"))
        fy = float(c6.selectbox("Steel", [250,415,500], index=2, format_func=lambda x:f"Fe{x}", key="fy"))
        st.caption("Ru,max is derived automatically from fck and fy (QD-012).")

        divider()
        st.markdown("#### Column & Footing Geometry")
        diag, controls = st.columns([1.1, 1.0], gap="large")
        diagram_slot = diag.empty()
        with controls:
            st.markdown("##### Enter the numbered values")
            lf = st.number_input("① Provided Footing Length Lf (mm)", min_value=500.0, value=1650.0, step=50.0, key="footing_length")
            b = st.number_input("② Column Width b (mm)", min_value=100.0, value=380.0, step=10.0, key="column_b")
            dcol = st.number_input("② Column Depth D (mm)", min_value=100.0, value=380.0, step=10.0, key="column_d")
            with st.expander("More geometry", expanded=False):
                df = st.number_input("Overall Footing Depth Df (mm)", min_value=100.0, value=350.0, step=25.0, key="footing_depth")
                cover = st.number_input("Clear Cover (mm)", min_value=20.0, value=50.0, step=5.0, key="cover")
        bf_preview = lf - dcol + b
        if bf_preview > 0:
            diagram_slot.markdown(
                '<div class="diagram-card">'
                '<div class="diagram-title">Graphical Geometry</div>'
                '<div class="diagram-subtitle">Plan view driven by the same numerical inputs used by the calculation engine. NOT TO SCALE.</div>'
                + footing_svg(DesignInputs(footing_length_mm=lf, column_width_mm=b, column_depth_mm=dcol), bf_preview) +
                '</div>',
                unsafe_allow_html=True,
            )
        else:
            diagram_slot.error("Derived footing width is non-positive.")
        st.caption("Footing width is linked automatically by equal projections: Bf = Lf − D + b.")

        divider()
        st.markdown("#### Bottom Reinforcement")
        r1,r2,r3,r4 = st.columns(4)
        dia = float(r1.selectbox("Bar Diameter (mm)", [8,10,12,16,20,25,32], index=1, key="bar_dia"))
        nx = int(r2.number_input("No. of bars — X", min_value=2, value=10, step=1, key="bars_x"))
        ny = int(r3.number_input("No. of bars — Y", min_value=2, value=10, step=1, key="bars_y"))
        minmethod = r4.selectbox("Minimum Reinforcement Method", ["Existing Excel method","IS 456 footing / solid slab method"], index=0, key="min_reinf_method")
    else:
        st.markdown(
            '<div class="draftsman-tip"><b>Detailed Input View:</b> use this when an engineer/checker wants the '
            'traditional full form. It controls the same data as Graphical Input View.</div>',
            unsafe_allow_html=True,
        )
        st.markdown("#### Loads")
        c1,c2,c3,c4 = st.columns(4)
        sbc = c1.number_input("Safe Bearing Capacity (kN/m²)", min_value=1.0, value=90.0, step=5.0, key="sbc")
        p = c2.number_input("Design Load — 1.5(DL+LL) (kN)", min_value=1.0, value=251.96, step=1.0, key="design_load")
        mux = c3.number_input("Mux (kNm)", value=5.04, step=0.5, key="mux")
        muy = c4.number_input("Muy (kNm)", value=21.77, step=0.5, key="muy")

        st.markdown("#### Column")
        c1,c2 = st.columns(2)
        b = c1.number_input("Column Width b (mm)", min_value=100.0, value=380.0, step=10.0, key="column_b")
        dcol = c2.number_input("Column Depth D (mm)", min_value=100.0, value=380.0, step=10.0, key="column_d")

        st.markdown("#### Materials")
        c1,c2 = st.columns(2)
        fck = float(c1.selectbox("Concrete Grade", [20,25,30,35,40], index=2, format_func=lambda x:f"M{x}", key="fck"))
        fy = float(c2.selectbox("Steel Grade", [250,415,500], index=2, format_func=lambda x:f"Fe{x}", key="fy"))
        st.caption("Ru,max is derived automatically from fck and fy (QD-012).")

        st.markdown("#### Footing Geometry")
        c1,c2,c3 = st.columns(3)
        lf = c1.number_input("Provided Footing Length Lf (mm)", min_value=500.0, value=1650.0, step=50.0, key="footing_length")
        df = c2.number_input("Overall Footing Depth Df (mm)", min_value=100.0, value=350.0, step=25.0, key="footing_depth")
        cover = c3.number_input("Clear Cover (mm)", min_value=20.0, value=50.0, step=5.0, key="cover")
        st.caption("Footing width is linked automatically by equal projections: Bf = Lf − D + b.")

        st.markdown("#### Bottom Reinforcement")
        c1,c2,c3,c4 = st.columns(4)
        dia = float(c1.selectbox("Bar Diameter (mm)", [8,10,12,16,20,25,32], index=1, key="bar_dia"))
        nx = int(c2.number_input("No. of bars — X", min_value=2, value=10, step=1, key="bars_x"))
        ny = int(c3.number_input("No. of bars — Y", min_value=2, value=10, step=1, key="bars_y"))
        minmethod = c4.selectbox("Minimum Reinforcement Method", ["Existing Excel method","IS 456 footing / solid slab method"], index=0, key="min_reinf_method")

    return DesignInputs(sbc,p,mux,muy,b,dcol,fck,fy,lf,df,cover,dia,nx,ny,minmethod,"Deformed")


def project_inputs() -> ProjectInfo:
    st.sidebar.markdown('<div class="sidebar-title">📋 Project Information</div>', unsafe_allow_html=True)
    project = st.sidebar.text_input("Project")
    client = st.sidebar.text_input("Client")
    structure = st.sidebar.text_input("Structure", value="COLUMN FOOTING")
    doc = st.sidebar.text_input("Document No.")
    with st.sidebar.expander("Preparation / approval details", expanded=False):
        prep = st.text_input("Prepared By")
        chk = st.text_input("Checked By")
        appr = st.text_input("Approved By")
        rev = st.text_input("Revision", value="R0")
    st.sidebar.divider()
    st.sidebar.markdown('<div class="sidebar-title">📐 Design Basis</div>', unsafe_allow_html=True)
    st.sidebar.markdown('<div class="code-ok">✓ IS 456:2000 — visible FOOTING-sheet scope</div>', unsafe_allow_html=True)
    with st.sidebar.expander("Scope & basis", expanded=False):
        st.markdown(
            '<div class="sidebar-note">'
            '<b>Application scope:</b> visible FOOTING-sheet calculations only.<br><br>'
            '<b>Hidden workbook content:</b> excluded; not used by the engine.<br><br>'
            '<b>Minimum reinforcement (QD-013):</b> user-selectable between the existing Excel method '
            'and the IS 456 footing/solid-slab method.'
            '</div>',
            unsafe_allow_html=True,
        )
    return ProjectInfo(project,client,structure,doc,rev,prep,chk,appr)


# Apply queued optimiser/default/original values before any design widget is instantiated.
_apply_pending_design_update()
project = project_inputs()

st.title("🏗️ RCC Design — COLUMN FOOTING")
st.markdown(
    '<div class="app-subtitle">Excel-derived footing design system · IS 456:2000 · visible FOOTING-sheet scope only</div>',
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="input-mode-bar"><b>Input Mode</b> — Graphical mode is intended for draftsmen and occasional users; '
    'Detailed mode exposes the full engineering form.</div>',
    unsafe_allow_html=True,
)
input_view = st.radio(
    "Choose how you want to enter the design data",
    ["🖼 Graphical Input", "📋 Detailed Input"],
    index=0,
    horizontal=True,
    key="input_view",
    help="Both views use the same underlying input data. Switching view does not change the design.",
)

section_title(1, "Design Inputs")
inp = make_inputs(input_view)
current_sig = sig(inp)

if "result" not in st.session_state: st.session_state.result = None
if "calc_sig" not in st.session_state: st.session_state.calc_sig = None
if "recommended" not in st.session_state: st.session_state.recommended = None
if "optimizer_original" not in st.session_state: st.session_state.optimizer_original = None

stale = st.session_state.result is not None and st.session_state.calc_sig != current_sig
if stale:
    st.markdown(
        '<div class="status-stale">Inputs changed after the previous calculation. Run Design Calculation again.</div>',
        unsafe_allow_html=True,
    )
    st.session_state.recommended = None

c_run,c_reset,c_default = st.columns([1.35,1.35,3.3])
if c_run.button("Run Design Calculation", type="primary", use_container_width=True):
    try:
        st.session_state.result = calculate(inp)
        st.session_state.calc_sig = current_sig
        st.session_state.recommended = None
        st.rerun()
    except Exception as e:
        st.error(str(e))
if c_reset.button("Clear Calculated Result", use_container_width=True):
    st.session_state.result=None
    st.session_state.calc_sig=None
    st.session_state.recommended=None
    st.session_state.optimizer_original=None
    st.rerun()
if c_default.button("Restore Default Values", use_container_width=True):
    defaults = DesignInputs()
    _queue_design_update(defaults, result=None, calc_signature=None, optimizer_original=None)
    st.rerun()

divider()
section_title(2, "Design Overview")
bf_preview = inp.footing_length_mm - inp.column_depth_mm + inp.column_width_mm
if input_view == "📋 Detailed Input":
    left,right = st.columns([1.1,1])
    with left:
        if bf_preview > 0:
            st.markdown(
                '<div class="diagram-card">'
                '<div class="diagram-title">Graphical Geometry</div>'
                '<div class="diagram-subtitle">Plan view driven by the same numerical inputs used by the calculation engine. NOT TO SCALE.</div>'
                + footing_svg(inp,bf_preview) +
                '</div>',
                unsafe_allow_html=True,
            )
        else:
            st.error("Derived footing width is non-positive.")
    with right:
        st.markdown('<div class="summary-card">', unsafe_allow_html=True)
        st.markdown('<div class="diagram-title">Current Input Summary</div>', unsafe_allow_html=True)
        st.markdown('<div class="diagram-subtitle">Live input summary before running the engineering calculation.</div>', unsafe_allow_html=True)
        st.dataframe(pd.DataFrame([
            ["Design load",inp.design_load_kn,"kN"],["Column",f"{inp.column_width_mm:.0f} × {inp.column_depth_mm:.0f}","mm"],
            ["Footing",f"{inp.footing_length_mm:.0f} × {bf_preview:.0f} × {inp.footing_depth_mm:.0f}","mm"],
            ["Concrete / steel",f"M{inp.fck_mpa:.0f} / Fe{inp.fy_mpa:.0f}","—"],["Bottom reinforcement",f"X: {inp.bars_x}–{inp.bar_dia_mm:.0f}φ; Y: {inp.bars_y}–{inp.bar_dia_mm:.0f}φ","—"],
            ["Minimum steel method",inp.min_reinf_method,"—"],
        ],columns=["Parameter","Value","Unit"]),hide_index=True,use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="summary-card">', unsafe_allow_html=True)
    st.markdown('<div class="diagram-title">Current Input Summary</div>', unsafe_allow_html=True)
    st.markdown('<div class="diagram-subtitle">The geometry diagram is shown above; this summarizes every input before running the engineering calculation.</div>', unsafe_allow_html=True)
    st.dataframe(pd.DataFrame([
        ["Design load",inp.design_load_kn,"kN"],["Column",f"{inp.column_width_mm:.0f} × {inp.column_depth_mm:.0f}","mm"],
        ["Footing",f"{inp.footing_length_mm:.0f} × {bf_preview:.0f} × {inp.footing_depth_mm:.0f}","mm"],
        ["Concrete / steel",f"M{inp.fck_mpa:.0f} / Fe{inp.fy_mpa:.0f}","—"],["Bottom reinforcement",f"X: {inp.bars_x}–{inp.bar_dia_mm:.0f}φ; Y: {inp.bars_y}–{inp.bar_dia_mm:.0f}φ","—"],
        ["Minimum steel method",inp.min_reinf_method,"—"],
    ],columns=["Parameter","Value","Unit"]),hide_index=True,use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

result = None if stale else st.session_state.result
if result:
    divider()
    section_title(3, "Design Results")
    status_class = "status-safe" if result.overall_status=="SAFE" else "status-unsafe"
    status_note = "SAFE for all mandatory checks implemented in this application." if result.overall_status=="SAFE" else "UNSAFE — one or more mandatory implemented checks have failed."
    st.markdown(
        f'<div class="{status_class}">Overall Status: {result.overall_status}<br><span style="font-size:.78rem;font-weight:500">{status_note}</span></div>',
        unsafe_allow_html=True,
    )
    if result.governing_failures: st.error("Governing failures: " + "; ".join(result.governing_failures))
    st.caption("Status applies only to mandatory checks implemented in this application and within the visible workbook scope.")

    m1,m2,m3,m4 = st.columns(4)
    m1.metric("Required Area",f"{result.values['required_area_m2']:.3f} m²")
    m2.metric("Provided Area",f"{result.values['provided_area_m2']:.3f} m²")
    m3.metric("Punching Vu / Vuc",f"{result.values['punching_demand_kn']:.1f} / {result.values['punching_capacity_kn']:.1f} kN")
    m4.metric("Ru,max",f"{result.values['ru_max_mpa']:.3f} N/mm²")

    # Detailed Calculations, Excel vs Web and Source Map are deliberately excluded from the web UI.
    tabs=st.tabs(["Design Results","Reinforcement","Checks","Safe Design","Reports"])
    with tabs[0]:
        st.dataframe(pd.DataFrame([
            ["Required footing L × B",f"{result.values['required_length_mm']:.1f} × {result.values['required_width_mm']:.1f}","mm"],
            ["Provided footing L × B",f"{inp.footing_length_mm:.1f} × {result.values['footing_width_mm']:.1f}","mm"],
            ["Factored pressure X (max/min)",f"{result.values['pmax_x_kn_m2']:.3f} / {result.values['pmin_x_kn_m2']:.3f}","kN/m²"],
            ["Factored pressure Y (max/min)",f"{result.values['pmax_y_kn_m2']:.3f} / {result.values['pmin_y_kn_m2']:.3f}","kN/m²"],
            ["Required effective depth X",result.values['required_eff_depth_x_mm'],"mm"],["Required effective depth Y",result.values['required_eff_depth_y_mm'],"mm"],
            ["Provided effective depth X",result.values['effective_depth_x_mm'],"mm"],["Provided effective depth Y",result.values['effective_depth_y_mm'],"mm"],
            ["One-way shear Y demand/capacity",f"{result.values['one_way_y_demand_kn']:.3f} / {result.values['one_way_y_capacity_kn']:.3f}","kN"],
            ["One-way shear X demand/capacity",f"{result.values['one_way_x_demand_kn']:.3f} / {result.values['one_way_x_capacity_kn']:.3f}","kN"],
            ["Development length required / available",f"{result.values['ld_required_mm']:.1f} / {result.values['ld_available_mm']:.1f}","mm (INFO only)"],
            ["Column bearing demand / capacity",f"{result.values['column_bearing_demand_mpa']:.3f} / {result.values['column_bearing_capacity_mpa']:.3f}","N/mm²"],
        ],columns=["Parameter","Value","Unit"]),hide_index=True,use_container_width=True)
    with tabs[1]:
        st.dataframe(pd.DataFrame([
            ["Flexural Ast required",result.values['ast_x_req_mm2'],result.values['ast_y_req_mm2'],"mm²"],
            ["Minimum Ast",result.values['ast_x_min_mm2'],result.values['ast_y_min_mm2'],"mm²"],
            ["Ast to provide",result.values['ast_x_to_provide_mm2'],result.values['ast_y_to_provide_mm2'],"mm²"],
            ["Ast provided",result.values['ast_x_provided_mm2'],result.values['ast_y_provided_mm2'],"mm²"],
            ["Spacing",result.values['spacing_x_mm'],result.values['spacing_y_mm'],"mm c/c"],
            ["Maximum spacing",result.values['max_spacing_x_mm'],result.values['max_spacing_y_mm'],"mm c/c"],
        ],columns=["Parameter","X","Y","Unit"]),hide_index=True,use_container_width=True)
        st.info(f"Minimum reinforcement method in force: {inp.min_reinf_method}")
    with tabs[2]:
        st.dataframe(pd.DataFrame([asdict(c) for c in result.checks]),hide_index=True,use_container_width=True)
    with tabs[3]:
        st.caption("The optimiser changes only variables you explicitly authorize below; every candidate reruns the full central engine.")
        c1,c2,c3,c4=st.columns(4)
        allow={"footing_length":c1.checkbox("Allow footing length"),"footing_depth":c2.checkbox("Allow footing depth"),"bars_x":c3.checkbox("Allow X bars"),"bars_y":c4.checkbox("Allow Y bars")}
        if st.button("Find Recommended SAFE Design"):
            cand,res,hist=recommend_safe_design(inp,allow)
            st.session_state.recommended=(cand,res,hist)
            if cand and res:
                st.session_state.optimizer_original=inp
        if st.session_state.recommended:
            cand,res,hist=st.session_state.recommended
            if cand and res:
                st.success("Verified SAFE candidate found using the full engine.")
                st.json({"footing_length_mm":cand.footing_length_mm,"footing_width_mm":res.values['footing_width_mm'],"footing_depth_mm":cand.footing_depth_mm,"bars_x":cand.bars_x,"bars_y":cand.bars_y,"status":res.overall_status})
                if st.button("Apply Recommended SAFE Design", type="primary", use_container_width=True, key="apply_safe_design"):
                    # Do not mutate active widget keys here. Queue the candidate and rerun.
                    _queue_design_update(
                        cand,
                        result=res,
                        calc_signature=sig(cand),
                        optimizer_original=st.session_state.optimizer_original,
                    )
                    st.rerun()
            else:
                st.warning("No SAFE candidate was found using the variables you authorised.")
        if st.session_state.optimizer_original is not None and inp != st.session_state.optimizer_original:
            if st.button("Restore Original Inputs", use_container_width=True):
                orig=st.session_state.optimizer_original
                restored_result = calculate(orig)
                _queue_design_update(
                    orig,
                    result=restored_result,
                    calc_signature=sig(orig),
                    optimizer_original=None,
                )
                st.rerun()
    with tabs[4]:
        pdf=build_pdf(project,inp,result)
        xlsx=build_excel(project,inp,result)
        c1,c2=st.columns(2)
        c1.download_button("Download PDF Report",pdf,file_name="COLUMN_FOOTING_Design_Report.pdf",mime="application/pdf",use_container_width=True,type="primary")
        c2.download_button("Download Excel Report",xlsx,file_name="COLUMN_FOOTING_Design_Report.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True,type="primary")
else:
    st.markdown(
        '<div class="status-stale">Set the inputs and click <b>Run Design Calculation</b>. Reports and SAFE/UNSAFE status remain unavailable until a current calculation exists.</div>',
        unsafe_allow_html=True,
    )
