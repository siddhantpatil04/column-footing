from __future__ import annotations

from io import BytesIO
from typing import Iterable

import xlsxwriter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from models import DesignInputs, DesignResult, ProjectInfo


def _fmt(v):
    if isinstance(v, float):
        return f"{v:,.3f}"
    return str(v)


def build_pdf(project: ProjectInfo, inp: DesignInputs, result: DesignResult) -> bytes:
    """Build a compact A4 submission-style PDF.

    Layout intentionally follows the user's C.V.Patil & Associates sample-report
    language: blue header/footer rules, red numbered headings, pale-blue compact
    tables, green SAFE/PASS text, and a dark-blue final summary.  Engineering
    calculations remain sourced only from ``result`` / the central engine.
    """
    from reportlab.platypus import KeepTogether, PageBreak

    buf = BytesIO()
    BLUE = colors.HexColor("#2F67A3")
    PALE = colors.HexColor("#EAF2F9")
    PALE2 = colors.HexColor("#F6F9FC")
    GRID = colors.HexColor("#C7D0DA")
    RED = colors.HexColor("#A53B32")
    GREEN = colors.HexColor("#38A169")
    DARK = colors.HexColor("#253746")
    GREY = colors.HexColor("#59636E")

    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=9*mm,
        rightMargin=9*mm,
        topMargin=17*mm,
        bottomMargin=15*mm,
        title="COLUMN FOOTING - DESIGN CALCULATION",
        author="C.V.Patil & Associates",
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "CFTitle", parent=styles["Heading1"], fontName="Helvetica-Bold",
        fontSize=12.4, leading=13.8, textColor=RED, alignment=TA_CENTER,
        spaceBefore=0, spaceAfter=2.2*mm,
    )
    section_style = ParagraphStyle(
        "CFSection", parent=styles["Heading2"], fontName="Helvetica-Bold",
        fontSize=9.6, leading=10.5, textColor=RED,
        spaceBefore=1.5*mm, spaceAfter=0.9*mm,
    )
    tiny = ParagraphStyle(
        "CFTiny", parent=styles["BodyText"], fontName="Helvetica",
        fontSize=6.35, leading=7.3, textColor=DARK, spaceAfter=0,
    )
    tiny_bold = ParagraphStyle(
        "CFTinyBold", parent=tiny, fontName="Helvetica-Bold",
    )
    note_style = ParagraphStyle(
        "CFNote", parent=tiny, fontSize=6.1, leading=7.0, textColor=GREY,
        spaceBefore=0.4*mm, spaceAfter=0.4*mm,
    )
    status_safe = ParagraphStyle(
        "CFSafe", parent=tiny_bold, textColor=GREEN, alignment=TA_CENTER,
    )
    status_fail = ParagraphStyle(
        "CFFail", parent=tiny_bold, textColor=RED, alignment=TA_CENTER,
    )

    def P(txt, style=tiny):
        return Paragraph(str(txt), style)

    def fmt(v, nd=2):
        if isinstance(v, (int, float)):
            return f"{float(v):,.{nd}f}"
        return str(v)

    def status_p(s):
        return Paragraph(str(s), status_safe if str(s).upper() in {"SAFE", "PASS"} else status_fail)

    def calc_cell(formula, substitution, source=""):
        # Keep the calculation trace in the table but move Excel-source references
        # to one compact line below each section to conserve vertical space.
        return P(f"{formula}<br/><font color=\"#59636E\">{substitution}</font>")

    def compact_table(data, widths, header=True, alignments=None, font_size=6.25, row_bgs=True):
        tbl = Table(data, colWidths=widths, repeatRows=1 if header else 0, hAlign="LEFT")
        cmds = [
            ("GRID", (0,0), (-1,-1), 0.32, GRID),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("FONTSIZE", (0,0), (-1,-1), font_size),
            ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
            ("TEXTCOLOR", (0,0), (-1,-1), DARK),
            ("LEFTPADDING", (0,0), (-1,-1), 2.0),
            ("RIGHTPADDING", (0,0), (-1,-1), 2.0),
            ("TOPPADDING", (0,0), (-1,-1), 1.25),
            ("BOTTOMPADDING", (0,0), (-1,-1), 1.25),
        ]
        if header:
            cmds += [
                ("BACKGROUND", (0,0), (-1,0), PALE),
                ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
                ("ALIGN", (0,0), (-1,0), "CENTER"),
            ]
        if row_bgs:
            startrow = 1 if header else 0
            cmds.append(("ROWBACKGROUNDS", (0,startrow), (-1,-1), [colors.white, PALE2]))
        if alignments:
            for col, align in alignments.items():
                cmds.append(("ALIGN", (col,1 if header else 0), (col,-1), align))
        tbl.setStyle(TableStyle(cmds))
        return tbl

    def paired_data(rows):
        """rows = [(label,value,label,value), ...] compact paired table."""
        return [[P(a, tiny_bold), P(b), P(c, tiny_bold), P(d)] for a,b,c,d in rows]

    def footer_header(canvas, doc_):
        canvas.saveState()
        w, h = A4
        y_top = h - 8.2*mm
        canvas.setStrokeColor(BLUE)
        canvas.setLineWidth(0.55)
        canvas.line(9*mm, y_top-2.0*mm, w-9*mm, y_top-2.0*mm)
        canvas.setFont("Helvetica-Bold", 7.1)
        canvas.setFillColor(BLUE)
        canvas.drawString(9*mm, y_top, "C.V.Patil & Associates")
        canvas.drawRightString(w-9*mm, y_top, "Design of COLUMN FOOTING")

        y_bot = 8.3*mm
        canvas.line(9*mm, y_bot+3.8*mm, w-9*mm, y_bot+3.8*mm)
        canvas.setFont("Helvetica", 5.0)
        canvas.setFillColor(colors.HexColor("#6B7280"))
        canvas.drawString(9*mm, y_bot+5.0*mm, "Status applies to mandatory checks implemented within the visible FOOTING-sheet scope.")
        canvas.setFont("Helvetica", 5.8)
        canvas.setFillColor(BLUE)
        project_txt = project.project or "-"
        client_txt = project.client or "-"
        canvas.drawString(9*mm, y_bot+1.1*mm, f"Project: {project_txt}")
        canvas.drawString(9*mm, y_bot-1.3*mm, f"Client: {client_txt}")
        canvas.setFont("Helvetica-Bold", 6.2)
        canvas.drawRightString(w-9*mm, y_bot-0.5*mm, str(doc_.page))
        canvas.restoreState()

    story = [Paragraph("COLUMN FOOTING - DESIGN CALCULATION", title_style)]

    # Submission metadata - compact, same visual hierarchy as reference report.
    meta = paired_data([
        ("Design Standard", "= IS 456:2000", "Document", f"= {project.document_no or '-'}"),
        ("Project", f"= {project.project or '-'}", "Status", f"= {result.overall_status}"),
        ("Structure", f"= {project.structure or 'COLUMN FOOTING'}", "Revision", f"= {project.revision or '-'}"),
        ("Client", f"= {project.client or '-'}", "Min. steel method", f"= {inp.min_reinf_method}"),
    ])
    meta_t = compact_table(meta, [31*mm, 68*mm, 31*mm, 59*mm], header=False, row_bgs=False)
    meta_t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (0,-1), PALE), ("BACKGROUND", (2,0), (2,-1), PALE),
        ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"), ("FONTNAME", (2,0), (2,-1), "Helvetica-Bold"),
        ("TEXTCOLOR", (3,1), (3,1), GREEN if result.overall_status == "SAFE" else RED),
        ("FONTNAME", (3,1), (3,1), "Helvetica-Bold"),
    ]))
    story += [meta_t]

    # 1. Design data - 2-up rows minimize vertical height.
    story += [Paragraph("1. DESIGN DATA", section_style)]
    design_rows = paired_data([
        ("Safe bearing capacity", f"= {fmt(inp.sbc_kn_m2)} kN/m<super>2</super>", "Design load, Pu", f"= {fmt(inp.design_load_kn)} kN"),
        ("Mux", f"= {fmt(inp.mux_knm)} kNm", "Muy", f"= {fmt(inp.muy_knm)} kNm"),
        ("Column width, b", f"= {fmt(inp.column_width_mm,0)} mm", "Column depth, D", f"= {fmt(inp.column_depth_mm,0)} mm"),
        ("Footing length, Lf", f"= {fmt(inp.footing_length_mm,0)} mm", "Footing width, Bf", f"= {fmt(result.values['footing_width_mm'],0)} mm"),
        ("Overall footing depth", f"= {fmt(inp.footing_depth_mm,0)} mm", "Clear cover", f"= {fmt(inp.cover_mm,0)} mm"),
        ("Concrete grade", f"= M{fmt(inp.fck_mpa,0)}", "Steel grade", f"= Fe{fmt(inp.fy_mpa,0)}"),
        ("Bottom bars X", f"= {inp.bars_x} nos. - {fmt(inp.bar_dia_mm,0)} mm", "Bottom bars Y", f"= {inp.bars_y} nos. - {fmt(inp.bar_dia_mm,0)} mm"),
        ("Ru,max", f"= {fmt(result.values['ru_max_mpa'],3)} N/mm<super>2</super>", "Bar type", f"= {inp.bar_type}"),
    ])
    dt = compact_table(design_rows, [39*mm, 52*mm, 39*mm, 59*mm], header=False, row_bgs=True)
    dt.setStyle(TableStyle([("BACKGROUND", (0,0), (0,-1), PALE), ("BACKGROUND", (2,0), (2,-1), PALE), ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"), ("FONTNAME", (2,0), (2,-1), "Helvetica-Bold")]))
    story += [dt]

    # 2. Footing sizing and bearing pressure.
    story += [Paragraph("2. FOOTING SIZING & BEARING PRESSURE", section_style)]
    sizing = [
        [P("Item", tiny_bold), P("Calculation", tiny_bold), P("Result", tiny_bold), P("Status", tiny_bold)],
        [P("Footing self weight"), calc_cell("Wf = 0.10 Pu", f"0.10 x {fmt(inp.design_load_kn)}", "FOOTING!E18"), P(f"= {fmt(result.values['self_weight_kn'])} kN", tiny_bold), P("Retained")],
        [P("Required area"), calc_cell("Areq = Ptotal / (1.5 SBC)", f"{fmt(result.values['total_load_kn'])} / (1.5 x {fmt(inp.sbc_kn_m2)})", "FOOTING!E22"), P(f"= {fmt(result.values['required_area_m2'],3)} m<super>2</super>", tiny_bold), status_p(result.check_map()['CHK-AREA'].status)],
        [P("Provided area"), calc_cell("Af = Lf x Bf", f"{fmt(inp.footing_length_mm/1000,3)} x {fmt(result.values['footing_width_mm']/1000,3)}", "FOOTING!E34"), P(f"= {fmt(result.values['provided_area_m2'],3)} m<super>2</super>", tiny_bold), status_p(result.check_map()['CHK-AREA'].status)],
        [P("Factored pressure - X"), calc_cell("pmax/min = pd +/- pbx", f"{fmt(result.values['pmax_x_kn_m2'])} / {fmt(result.values['pmin_x_kn_m2'])}", "FOOTING!E43:E44"), P(f"= {fmt(result.values['pmax_x_kn_m2'])} / {fmt(result.values['pmin_x_kn_m2'])} kN/m<super>2</super>", tiny_bold), status_p(result.check_map()['CHK-BEAR-X-U'].status)],
        [P("Factored pressure - Y"), calc_cell("pmax/min = pd +/- pby", f"{fmt(result.values['pmax_y_kn_m2'])} / {fmt(result.values['pmin_y_kn_m2'])}", "FOOTING!E55:E56"), P(f"= {fmt(result.values['pmax_y_kn_m2'])} / {fmt(result.values['pmin_y_kn_m2'])} kN/m<super>2</super>", tiny_bold), status_p(result.check_map()['CHK-BEAR-Y-U'].status)],
        [P("Working pmax - X / Y"), P("Factored pressure / 1.50"), P(f"= {fmt(result.values['pmax_x_working_kn_m2'])} / {fmt(result.values['pmax_y_working_kn_m2'])} kN/m<super>2</super>", tiny_bold), P(f"{result.check_map()['CHK-BEAR-X-W'].status} / {result.check_map()['CHK-BEAR-Y-W'].status}", status_safe if result.check_map()['CHK-BEAR-X-W'].status == 'SAFE' and result.check_map()['CHK-BEAR-Y-W'].status == 'SAFE' else status_fail)],
    ]
    story += [compact_table(sizing, [37*mm, 80*mm, 48*mm, 24*mm], header=True, alignments={3:"CENTER"})]
    story += [P("Excel trace: self weight E18; required area E22; provided area E34; X pressure E43:E49; Y pressure E55:E61.", note_style)]

    # Contact + flexure compact side-by-side.
    story += [Paragraph("3. CONTACT, FLEXURE & EFFECTIVE DEPTH", section_style)]
    cf = [
        [P("Direction", tiny_bold), P("Eccentricity / limit", tiny_bold), P("Moment", tiny_bold), P("Req. eff. depth", tiny_bold), P("Prov. eff. depth", tiny_bold), P("Status", tiny_bold)],
        [P("X"), P(f"{fmt(result.values['eccentricity_x_mm'])} / {fmt(result.values['eccentricity_limit_x_mm'])} mm"), P(f"{fmt(result.values['moment_x_knm'])} kNm"), P(f"{fmt(result.values['required_eff_depth_x_mm'])} mm"), P(f"{fmt(result.values['effective_depth_x_mm'])} mm"), status_p("SAFE" if result.check_map()['CHK-CONTACT-X'].status=='SAFE' and result.check_map()['CHK-DEPTH'].status=='SAFE' else "UNSAFE")],
        [P("Y"), P(f"{fmt(result.values['eccentricity_y_mm'])} / {fmt(result.values['eccentricity_limit_y_mm'])} mm"), P(f"{fmt(result.values['moment_y_knm'])} kNm"), P(f"{fmt(result.values['required_eff_depth_y_mm'])} mm"), P(f"{fmt(result.values['effective_depth_y_mm'])} mm"), status_p("SAFE" if result.check_map()['CHK-CONTACT-Y'].status=='SAFE' and result.check_map()['CHK-DEPTH'].status=='SAFE' else "UNSAFE")],
    ]
    story += [compact_table(cf, [20*mm, 42*mm, 31*mm, 38*mm, 37*mm, 21*mm], header=True, alignments={0:"CENTER",5:"CENTER"})]
    story += [P("Required effective depth: dreq = sqrt(M / (Ru x b)). QD-003 changes nomenclature only; workbook formula basis retained.", note_style)]

    # 4. Punching shear and reinforcement designed for one page continuity.
    story += [Paragraph("4. TWO-WAY (PUNCHING) SHEAR", section_style)]
    ps = [
        [P("Calculation", tiny_bold), P("Demand", tiny_bold), P("Capacity", tiny_bold), P("Status", tiny_bold)],
        [calc_cell("Vu = Ptotal - pd x Ainside", f"{fmt(result.values['total_load_kn'])} - {fmt(result.values['pd_kn_m2'])} x {fmt(result.values['punching_inside_reaction_kn']/result.values['pd_kn_m2'],6)}", "FOOTING!E163"), P(f"{fmt(result.values['punching_demand_kn'])} kN", tiny_bold), P(f"{fmt(result.values['punching_capacity_kn'])} kN", tiny_bold), status_p(result.check_map()['CHK-PUNCH'].status)],
    ]
    story += [compact_table(ps, [93*mm, 34*mm, 38*mm, 24*mm], header=True, alignments={1:"CENTER",2:"CENTER",3:"CENTER"})]
    story += [P("Excel trace: punching shear FOOTING!E163; capacity FOOTING!E161.", note_style)]

    story += [Paragraph("5. BOTTOM REINFORCEMENT", section_style)]
    reinf = [
        [P("Direction", tiny_bold), P("Flexural Ast", tiny_bold), P("Minimum Ast", tiny_bold), P("Required Ast", tiny_bold), P("Provided Ast", tiny_bold), P("Spacing / max", tiny_bold), P("Status", tiny_bold)],
        [P("X"), P(f"{fmt(result.values['ast_x_req_mm2'])} mm<super>2</super>"), P(f"{fmt(result.values['ast_x_min_mm2'])} mm<super>2</super>"), P(f"{fmt(result.values['ast_x_to_provide_mm2'])} mm<super>2</super>"), P(f"{fmt(result.values['ast_x_provided_mm2'])} mm<super>2</super>"), P(f"{fmt(result.values['spacing_x_mm'])} / {fmt(result.values['max_spacing_x_mm'])} mm"), status_p("SAFE" if result.check_map()['CHK-AST-X'].status=='SAFE' and result.check_map()['CHK-SP-X'].status=='SAFE' else "UNSAFE")],
        [P("Y"), P(f"{fmt(result.values['ast_y_req_mm2'])} mm<super>2</super>"), P(f"{fmt(result.values['ast_y_min_mm2'])} mm<super>2</super>"), P(f"{fmt(result.values['ast_y_to_provide_mm2'])} mm<super>2</super>"), P(f"{fmt(result.values['ast_y_provided_mm2'])} mm<super>2</super>"), P(f"{fmt(result.values['spacing_y_mm'])} / {fmt(result.values['max_spacing_y_mm'])} mm"), status_p("SAFE" if result.check_map()['CHK-AST-Y'].status=='SAFE' and result.check_map()['CHK-SP-Y'].status=='SAFE' else "UNSAFE")],
    ]
    story += [compact_table(reinf, [18*mm, 27*mm, 27*mm, 28*mm, 29*mm, 39*mm, 21*mm], header=True, alignments={0:"CENTER",6:"CENTER"})]
    story += [P(f"Minimum reinforcement method in force: {inp.min_reinf_method}. Provided bars: X = {inp.bars_x} nos. x {fmt(inp.bar_dia_mm,0)} mm; Y = {inp.bars_y} nos. x {fmt(inp.bar_dia_mm,0)} mm.", note_style)]

    # Page 2 begins at logical detailing/check block to keep report readable.
    story += [Paragraph("6. ONE-WAY SHEAR", section_style)]
    sh = [
        [P("Direction", tiny_bold), P("Demand", tiny_bold), P("Capacity", tiny_bold), P("pt", tiny_bold), P("tau_c", tiny_bold), P("Status", tiny_bold)],
        [P("Y"), P(f"{fmt(result.values['one_way_y_demand_kn'])} kN"), P(f"{fmt(result.values['one_way_y_capacity_kn'])} kN"), P(f"{fmt(result.values['pt_y_percent'],3)} %"), P(f"{fmt(result.values['tau_c_y_mpa'],3)} N/mm<super>2</super>"), status_p(result.check_map()['CHK-1WAY-Y'].status)],
        [P("X"), P(f"{fmt(result.values['one_way_x_demand_kn'])} kN"), P(f"{fmt(result.values['one_way_x_capacity_kn'])} kN"), P(f"{fmt(result.values['pt_x_percent'],3)} %"), P(f"{fmt(result.values['tau_c_x_mpa'],3)} N/mm<super>2</super>"), status_p(result.check_map()['CHK-1WAY-X'].status)],
    ]
    story += [compact_table(sh, [23*mm, 36*mm, 36*mm, 26*mm, 43*mm, 25*mm], header=True, alignments={0:"CENTER",5:"CENTER"})]

    story += [Paragraph("7. DEVELOPMENT LENGTH & COLUMN-FOOTING BEARING", section_style)]
    dc = [
        [P("Check", tiny_bold), P("Calculation / basis", tiny_bold), P("Demand", tiny_bold), P("Capacity / available", tiny_bold), P("Status", tiny_bold)],
        [P("Development length"), calc_cell("Ld = 0.87 fy dia / (4 tau_bd)", f"tau_bd = {fmt(result.values['bond_stress_mpa'],3)} N/mm2", "FOOTING!E195:E196"), P(f"{fmt(result.values['ld_required_mm'])} mm"), P(f"{fmt(result.values['ld_available_mm'])} mm"), P("INFO", tiny_bold)],
        [P("Column-footing bearing"), calc_cell("sigma = Ptotal / (bD); sigma_br = 0.45 fck min(sqrt(A1/A2),2)", f"enhancement = {fmt(result.values['bearing_enhancement'],3)}", "FOOTING!E236:E245"), P(f"{fmt(result.values['column_bearing_demand_mpa'],3)} N/mm<super>2</super>"), P(f"{fmt(result.values['column_bearing_capacity_mpa'],3)} N/mm<super>2</super>"), status_p(result.check_map()['CHK-COL-BEAR'].status)],
    ]
    story += [compact_table(dc, [34*mm, 76*mm, 29*mm, 32*mm, 18*mm], header=True, alignments={4:"CENTER"})]
    story += [P("Excel trace: development length E195:E196; column-footing bearing E236:E245.", note_style)]

    story += [Paragraph("8. MANDATORY CHECK SUMMARY", section_style)]
    checks = result.checks
    # Pair two checks per row to halve vertical height.
    pair_rows = [[P("Check", tiny_bold), P("Status", tiny_bold), P("Check", tiny_bold), P("Status", tiny_bold)]]
    for i in range(0, len(checks), 2):
        c1 = checks[i]
        c2 = checks[i+1] if i+1 < len(checks) else None
        left = P(c1.name)
        lstat = status_p(c1.status)
        if c2:
            right, rstat = P(c2.name), status_p(c2.status)
        else:
            right, rstat = P(""), P("")
        pair_rows.append([left, lstat, right, rstat])
    story += [compact_table(pair_rows, [70*mm, 23*mm, 73*mm, 23*mm], header=True, alignments={1:"CENTER",3:"CENTER"})]

    story += [Paragraph("9. FINAL DESIGN SUMMARY", section_style)]
    final_rows = [
        [P("FINAL DESIGN SUMMARY", ParagraphStyle("SumHdr", parent=tiny_bold, textColor=colors.white)), "", "", ""],
        [P("Overall Status", tiny_bold), status_p(result.overall_status), P("Provided footing", tiny_bold), P(f"{fmt(inp.footing_length_mm,0)} x {fmt(result.values['footing_width_mm'],0)} x {fmt(inp.footing_depth_mm,0)} mm")],
        [P("Bottom reinforcement - X", tiny_bold), P(f"{inp.bars_x} nos. - {fmt(inp.bar_dia_mm,0)} mm @ {fmt(result.values['spacing_x_mm'])} mm c/c"), P("Bottom reinforcement - Y", tiny_bold), P(f"{inp.bars_y} nos. - {fmt(inp.bar_dia_mm,0)} mm @ {fmt(result.values['spacing_y_mm'])} mm c/c")],
        [P("Governing", tiny_bold), P(f"dreq = {fmt(result.values['required_effective_depth_mm'])} mm; punching = {fmt(result.values['punching_demand_kn'])}/{fmt(result.values['punching_capacity_kn'])} kN; bearing = {fmt(result.values['column_bearing_demand_mpa'],3)}/{fmt(result.values['column_bearing_capacity_mpa'],3)} N/mm<super>2</super>"), "", ""],
    ]
    ft = Table(final_rows, colWidths=[40*mm, 51*mm, 40*mm, 58*mm], hAlign="LEFT")
    ft.setStyle(TableStyle([
        ("SPAN", (0,0), (3,0)), ("SPAN", (1,3), (3,3)),
        ("BACKGROUND", (0,0), (3,0), BLUE),
        ("TEXTCOLOR", (0,0), (3,0), colors.white), ("FONTNAME", (0,0), (3,0), "Helvetica-Bold"),
        ("BACKGROUND", (0,1), (0,-1), PALE), ("BACKGROUND", (2,1), (2,2), PALE),
        ("GRID", (0,0), (-1,-1), 0.35, GRID), ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 2.2), ("RIGHTPADDING", (0,0), (-1,-1), 2.2),
        ("TOPPADDING", (0,0), (-1,-1), 1.35), ("BOTTOMPADDING", (0,0), (-1,-1), 1.35),
    ]))
    story += [ft]



    doc.build(story, onFirstPage=footer_header, onLaterPages=footer_header)
    return buf.getvalue()

def _table_style(font_size=7.5):
    return TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1F2937")), ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTNAME",(0,1),(-1,-1),"Helvetica"),
        ("FONTSIZE",(0,0),(-1,-1),font_size),("GRID",(0,0),(-1,-1),0.35,colors.HexColor("#CBD5E1")),
        ("VALIGN",(0,0),(-1,-1),"TOP"),("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F8FAFC")]),
        ("LEFTPADDING",(0,0),(-1,-1),3),("RIGHTPADDING",(0,0),(-1,-1),3),("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3),
    ])


def build_excel(project: ProjectInfo, inp: DesignInputs, result: DesignResult) -> bytes:
    buf = BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    title = wb.add_format({"bold":True,"font_size":16,"font_color":"#FFFFFF","bg_color":"#1F2937","align":"center","valign":"vcenter"})
    hdr = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#B42318","border":1,"align":"center"})
    cell = wb.add_format({"border":1,"valign":"top"})
    num = wb.add_format({"border":1,"num_format":"0.000","valign":"top"})
    safe = wb.add_format({"border":1,"bold":True,"font_color":"#166534","bg_color":"#DCFCE7","align":"center"})
    unsafe = wb.add_format({"border":1,"bold":True,"font_color":"#991B1B","bg_color":"#FEE2E2","align":"center"})

    def setup(ws, widths=(30,22,18,18,18)):
        ws.set_landscape(); ws.set_paper(9); ws.fit_to_pages(1, 0); ws.set_margins(0.3,0.3,0.5,0.5)
        for i,w in enumerate(widths): ws.set_column(i,i,w)

    ws = wb.add_worksheet("SUMMARY"); setup(ws)
    ws.merge_range("A1:E2","COLUMN FOOTING — DESIGN REPORT",title)
    rows = [("Project",project.project),("Client",project.client),("Structure",project.structure),("Revision",project.revision),("Design code","IS 456:2000"),("Minimum steel method",inp.min_reinf_method),("Overall status",result.overall_status)]
    ws.write_row(3,0,["Field","Value"],hdr)
    for r,(k,v) in enumerate(rows,4): ws.write(r,0,k,cell); ws.write(r,1,v, safe if k=="Overall status" and v=="SAFE" else unsafe if k=="Overall status" else cell)

    wi = wb.add_worksheet("INPUTS"); setup(wi)
    wi.write_row(0,0,["Parameter","Value","Unit","Excel source"],hdr)
    input_rows = [
        ("SBC",inp.sbc_kn_m2,"kN/m²","E8"),("Design load",inp.design_load_kn,"kN","E9"),("Mux",inp.mux_knm,"kNm","E10"),("Muy",inp.muy_knm,"kNm","E11"),
        ("Column width",inp.column_width_mm,"mm","E12"),("Column depth",inp.column_depth_mm,"mm","E13"),("fck",inp.fck_mpa,"N/mm²","E14"),("fy",inp.fy_mpa,"N/mm²","E15"),
        ("Provided footing length",inp.footing_length_mm,"mm","E30"),("Derived footing width",result.values["footing_width_mm"],"mm","E32"),("Overall footing depth",inp.footing_depth_mm,"mm","E122"),
        ("Cover",inp.cover_mm,"mm","E110"),("Bar dia",inp.bar_dia_mm,"mm","E181"),("Bars X",inp.bars_x,"Nos.","E182"),("Bars Y",inp.bars_y,"Nos.","E186"),
    ]
    for r,row in enumerate(input_rows,1):
        wi.write(r,0,row[0],cell); wi.write(r,1,row[1],num if isinstance(row[1],(int,float)) else cell); wi.write(r,2,row[2],cell); wi.write(r,3,"FOOTING!"+row[3],cell)

    wr = wb.add_worksheet("FOOTING DESIGN"); setup(wr,(34,20,18,25,22))
    wr.write_row(0,0,["Result","Value","Unit","Status/Note","Excel source"],hdr)
    key_rows = [
        ("Required area","required_area_m2","m²","E22"),("Required length","required_length_mm","mm","E26"),("Required width","required_width_mm","mm","E28"),
        ("Provided area","provided_area_m2","m²","E34"),("pmax X","pmax_x_kn_m2","kN/m²","E43"),("pmin X","pmin_x_kn_m2","kN/m²","E44"),
        ("pmax Y","pmax_y_kn_m2","kN/m²","E55"),("pmin Y","pmin_y_kn_m2","kN/m²","E56"),("Req effective depth X","required_eff_depth_x_mm","mm","E98/E117"),
        ("Req effective depth Y","required_eff_depth_y_mm","mm","E101/E119"),("Punching demand","punching_demand_kn","kN","E163"),("Punching capacity","punching_capacity_kn","kN","E161"),
        ("One-way shear demand Y","one_way_y_demand_kn","kN","E212"),("One-way shear capacity Y","one_way_y_capacity_kn","kN","E210"),
        ("One-way shear demand X","one_way_x_demand_kn","kN","E232"),("One-way shear capacity X","one_way_x_capacity_kn","kN","E230"),
        ("Column bearing demand","column_bearing_demand_mpa","N/mm²","E236"),("Column bearing capacity","column_bearing_capacity_mpa","N/mm²","E245"),
    ]
    for r,(label,key,unit,src) in enumerate(key_rows,1): wr.write(r,0,label,cell); wr.write(r,1,result.values[key],num); wr.write(r,2,unit,cell); wr.write(r,4,"FOOTING!"+src,cell)

    ws2 = wb.add_worksheet("REINFORCEMENT"); setup(ws2)
    ws2.write_row(0,0,["Parameter","X direction","Y direction","Unit","Method"],hdr)
    reinf = [
        ("Flexural Ast required",result.values["ast_x_req_mm2"],result.values["ast_y_req_mm2"],"mm²","LSM"),
        ("Minimum Ast",result.values["ast_x_min_mm2"],result.values["ast_y_min_mm2"],"mm²",inp.min_reinf_method),
        ("Ast to provide",result.values["ast_x_to_provide_mm2"],result.values["ast_y_to_provide_mm2"],"mm²","max(required, minimum)"),
        ("Ast provided",result.values["ast_x_provided_mm2"],result.values["ast_y_provided_mm2"],"mm²","bar count × bar area"),
        ("Spacing",result.values["spacing_x_mm"],result.values["spacing_y_mm"],"mm c/c","QD-015 corrected"),
        ("Maximum spacing",result.values["max_spacing_x_mm"],result.values["max_spacing_y_mm"],"mm c/c","IS 456 slab spacing"),
    ]
    for r,row in enumerate(reinf,1):
        ws2.write(r,0,row[0],cell); ws2.write(r,1,row[1],num); ws2.write(r,2,row[2],num); ws2.write(r,3,row[3],cell); ws2.write(r,4,row[4],cell)

    wc = wb.add_worksheet("CHECKS"); setup(wc,(38,14,18,18,18))
    wc.write_row(0,0,["Check","Status","Demand","Capacity","Unit"],hdr)
    for r,c in enumerate(result.checks,1):
        wc.write(r,0,c.name,cell); wc.write(r,1,c.status,safe if c.status=="SAFE" else unsafe); wc.write(r,2,c.demand if c.demand is not None else "",num); wc.write(r,3,c.capacity if c.capacity is not None else "",num); wc.write(r,4,c.unit,cell)

    wf = wb.add_worksheet("FORMULA TRACE"); setup(wf,(34,38,45,18,26))
    wf.write_row(0,0,["Calculation","Formula","Substitution","Result","Excel source"],hdr)
    for r,f in enumerate(result.trace,1): wf.write(r,0,f.name,cell); wf.write(r,1,f.formula,cell); wf.write(r,2,f.substitution,cell); wf.write(r,3,f"{_fmt(f.result)} {f.unit}",cell); wf.write(r,4,f.excel_source,cell)

    wn = wb.add_worksheet("RECONCILIATION NOTES"); setup(wn,(14,18,78,20,20))
    wn.write_row(0,0,["QD","Decision","Implementation","Status"],hdr)
    qds = [
        ("QD-001","IGNORE","Separate X/Y bearing checks retained; no combined biaxial pressure.","Closed"),
        ("QD-002","CORRECT","Working minimum pressures made live.","Closed"),("QD-003","RELABEL","Required total depth renamed required effective depth; formula retained.","Closed"),
        ("QD-004","CORRECT","Minimum steel references corrected by direction.","Closed"),("QD-005","CORRECT","One-way shear pressure axes corrected.","Closed"),
        ("QD-006","IGNORE","Development length remains information-only.","Closed"),("QD-007","CORRECT","Bond stress linked to concrete grade and bar type.","Closed"),
        ("QD-008","CORRECT","Bearing enhancement factor capped at 2.0.","Closed"),("QD-009","CORRECT","Second 1.5 factor removed from E236 equivalent.","Closed"),
        ("QD-010","IGNORE","10% self-weight assumption retained.","Closed"),("QD-011","CORRECT","Separate no-tension/eccentricity checks added.","Closed"),
        ("QD-012","CORRECT","Ru.max derived from fck/fy.","Closed"),("QD-013","SELECTABLE","Existing Excel method or IS 456 footing/solid-slab method.","Closed"),
        ("QD-014","CORRECT","Explicit Ast and spacing checks added.","Closed"),("QD-015","CORRECT","Bar c/c spacing corrected.","Closed"),
        ("QD-016","CORRECT","Equal-projection width linkage corrected for rectangular columns.","Closed"),("QD-017","CORRECT","Legacy external/broken names excluded from engine.","Closed"),
        ("QD-018","CORRECT","Punching demand reconciled with reaction inside critical perimeter.","Closed"),("QD-019","IGNORE","Pressure-at-face treatment retained.","Closed"),
        ("QD-020","CORRECT","Blank pt.min removed as an input.","Closed"),
    ]
    for r,row in enumerate(qds,1):
        for c,v in enumerate(row): wn.write(r,c,v,cell)

    wb.close(); return buf.getvalue()
