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
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=14*mm, leftMargin=14*mm, topMargin=14*mm, bottomMargin=14*mm)
    styles = getSampleStyleSheet()
    title = ParagraphStyle("Title2", parent=styles["Title"], fontSize=16, leading=19, textColor=colors.HexColor("#1F2937"), alignment=TA_CENTER)
    h2 = ParagraphStyle("H2", parent=styles["Heading2"], fontSize=11, leading=14, textColor=colors.HexColor("#B42318"), spaceBefore=6, spaceAfter=5)
    body = styles["BodyText"]
    story = [Paragraph("COLUMN FOOTING — DESIGN CALCULATION REPORT", title), Spacer(1, 5*mm)]

    meta = [
        ["Project", project.project or "—", "Client", project.client or "—"],
        ["Structure", project.structure, "Document", project.document_no or "—"],
        ["Revision", project.revision, "Design Code", "IS 456:2000"],
        ["Minimum steel method", inp.min_reinf_method, "Overall status", result.overall_status],
    ]
    t = Table(meta, colWidths=[29*mm, 55*mm, 29*mm, 55*mm])
    t.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#CBD5E1")),("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#F8FAFC")),("FONTNAME",(0,0),(-1,-1),"Helvetica"),("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),("FONTNAME",(2,0),(2,-1),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),8),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4)]))
    story += [t, Spacer(1, 4*mm), Paragraph("1. Design Inputs", h2)]

    input_rows = [["Parameter","Value","Unit"],
        ["Safe bearing capacity", _fmt(inp.sbc_kn_m2), "kN/m²"],
        ["Design load (already factored)", _fmt(inp.design_load_kn), "kN"],
        ["Mux", _fmt(inp.mux_knm), "kNm"], ["Muy", _fmt(inp.muy_knm), "kNm"],
        ["Column b × D", f"{inp.column_width_mm:.0f} × {inp.column_depth_mm:.0f}", "mm"],
        ["Concrete / Steel", f"M{inp.fck_mpa:.0f} / Fe{inp.fy_mpa:.0f}", "—"],
        ["Footing L × B", f"{inp.footing_length_mm:.0f} × {float(result.values['footing_width_mm']):.0f}", "mm"],
        ["Overall depth", _fmt(inp.footing_depth_mm), "mm"],
        ["Cover / bar dia", f"{inp.cover_mm:.0f} / {inp.bar_dia_mm:.0f}", "mm"],
        ["Bottom bars X / Y", f"{inp.bars_x} / {inp.bars_y}", "Nos."],
    ]
    ti = Table(input_rows, colWidths=[85*mm, 55*mm, 25*mm], repeatRows=1)
    ti.setStyle(_table_style())
    story += [ti, Paragraph("2. Key Results", h2)]

    keys = [
        ("Required area", "required_area_m2", "m²"), ("Provided area", "provided_area_m2", "m²"),
        ("Factored pmax X", "pmax_x_kn_m2", "kN/m²"), ("Factored pmax Y", "pmax_y_kn_m2", "kN/m²"),
        ("Required effective depth X", "required_eff_depth_x_mm", "mm"), ("Required effective depth Y", "required_eff_depth_y_mm", "mm"),
        ("Punching shear demand", "punching_demand_kn", "kN"), ("Punching shear capacity", "punching_capacity_kn", "kN"),
        ("Ast X to provide", "ast_x_to_provide_mm2", "mm²"), ("Ast X provided", "ast_x_provided_mm2", "mm²"),
        ("Ast Y to provide", "ast_y_to_provide_mm2", "mm²"), ("Ast Y provided", "ast_y_provided_mm2", "mm²"),
        ("Bar spacing X", "spacing_x_mm", "mm c/c"), ("Bar spacing Y", "spacing_y_mm", "mm c/c"),
        ("Column bearing demand", "column_bearing_demand_mpa", "N/mm²"), ("Column bearing capacity", "column_bearing_capacity_mpa", "N/mm²"),
    ]
    rr = [["Result","Value","Unit"]] + [[label,_fmt(result.values[k]),unit] for label,k,unit in keys]
    tr = Table(rr, colWidths=[85*mm, 55*mm, 25*mm], repeatRows=1)
    tr.setStyle(_table_style())
    story += [tr, Paragraph("3. Mandatory Checks", h2)]

    cr = [["Check","Status","Demand","Capacity","Unit"]]
    for c in result.checks:
        cr.append([c.name,c.status,_fmt(c.demand) if c.demand is not None else "—",_fmt(c.capacity) if c.capacity is not None else "—",c.unit])
    tc = Table(cr, colWidths=[69*mm,22*mm,30*mm,30*mm,22*mm], repeatRows=1)
    tc.setStyle(_table_style())
    story += [tc, Paragraph("4. Formula Trace", h2)]

    fr = [["Calculation","Formula / substitution","Result","Excel source"]]
    for f in result.trace:
        fr.append([f.name, f"{f.formula}\n{f.substitution}", f"{_fmt(f.result)} {f.unit}", f.excel_source])
    tf = Table(fr, colWidths=[40*mm,73*mm,28*mm,32*mm], repeatRows=1)
    tf.setStyle(_table_style(font_size=6.7))
    story += [tf, Paragraph("5. Assumptions / Reconciliation Notes", h2)]
    for a in result.assumptions:
        story.append(Paragraph("• " + a, body))
    story.append(Spacer(1,2*mm))
    story.append(Paragraph("SAFE means safe for all mandatory checks implemented in this application and within the visible workbook scope.", body))

    doc.build(story)
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
