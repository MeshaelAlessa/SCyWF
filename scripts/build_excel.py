#!/usr/bin/env python3
"""Build the SCyWF Training Priority Model workbook (excel/SCyWF_Training_Priority_Model.xlsx).

The workbook is formula driven: change a weight on the Weights sheet, add rows to a data
sheet, or type another entity on the Overview, and every score and ranking recalculates.

Model (all in formulas):
  record score  = severity weight x source-specific factors (recency, status, exposure, relevance, confidence)
  source share  = entity's record scores per gap category / entity's total score for that source
  gap index     = sum over sources of (source weight x source share) / sum of weights of sources that have data
  program score = sum over categories of (gap index x program coverage of that category)
"""
import csv
import json
from collections import Counter
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "excel" / "SCyWF_Training_Priority_Model.xlsx"
MAXR = 10000  # data rows covered by every formula range (add data below the last row, up to this)

# ---------- styles ----------
FONT = "Arial"
GREEN, DARK, MUTED, LINE = "0B6E4F", "14201B", "56655E", "D5DDD9"
f_base = Font(name=FONT, size=10, color=DARK)
f_bold = Font(name=FONT, size=10, bold=True, color=DARK)
f_head = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_title = Font(name=FONT, size=18, bold=True, color=DARK)
f_h2 = Font(name=FONT, size=12, bold=True, color=GREEN)
f_note = Font(name=FONT, size=9, italic=True, color=MUTED)
f_input = Font(name=FONT, size=11, bold=True, color="0000FF")
f_link = Font(name=FONT, size=10, color="008000")
f_kpi = Font(name=FONT, size=20, bold=True, color=DARK)
f_kpi_red = Font(name=FONT, size=20, bold=True, color="B3203A")
f_label = Font(name=FONT, size=8, bold=True, color=MUTED)
f_entity = Font(name=FONT, size=16, bold=True, color=DARK)
fill_head = PatternFill("solid", fgColor=GREEN)
fill_input = PatternFill("solid", fgColor="FFF59D")
fill_card = PatternFill("solid", fgColor="EEF4F1")
fill_sub = PatternFill("solid", fgColor="D8EDE4")
thin = Side(style="thin", color=LINE)
box = Border(left=thin, right=thin, top=thin, bottom=thin)
bottom = Border(bottom=thin)
wrap = Alignment(wrap_text=True, vertical="top")
center = Alignment(horizontal="center", vertical="center", wrap_text=True)


def load_csv(p):
    with open(p, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def style_range(ws, ref, font=None, fill=None, border=None, align=None, fmt=None):
    for row in ws[ref]:
        for c in row:
            if font: c.font = font
            if fill: c.fill = fill
            if border: c.border = border
            if align: c.alignment = align
            if fmt: c.number_format = fmt


def header(ws, row, col, labels, widths=None):
    for i, h in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=h)
        c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, box
    if widths:
        for i, w in enumerate(widths):
            ws.column_dimensions[L(col + i)].width = w


def name(wb, nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


def main():
    tasks = json.loads((ROOT / "data/scywf/tasks.json").read_text(encoding="utf-8"))
    roles = json.loads((ROOT / "data/scywf/job_roles.json").read_text(encoding="utf-8"))
    mapping = json.loads((ROOT / "data/mapping/vulnerability_task_mapping.json").read_text(encoding="utf-8"))
    catalog = json.loads((ROOT / "data/mapping/source_catalog.json").read_text(encoding="utf-8"))
    entities = load_csv(ROOT / "data/generated/entities.csv")
    incidents = load_csv(ROOT / "data/generated/incidents.csv")
    vulns = load_csv(ROOT / "data/generated/vulnerabilities.csv")
    threats = load_csv(ROOT / "data/generated/threats.csv")

    cats = list(mapping["categories"])
    programs = catalog["programs"]
    role_name = {r["id"]: r["name"] for r in roles}
    owners = {}
    for r in roles:
        for t in r["tasks"]:
            owners.setdefault(t, []).append(r["id"])

    # Catalog rows: (source, title, category, task ids)
    cat_rows = [("Incident", v["title"], v["category"], v["primary"]) for v in mapping["vulnerabilities"]]
    cat_rows += [("Vulnerability", v["title"], v["category"], v["tasks"]) for v in catalog["vulnerability_findings"]]
    cat_rows += [("Threat", v["title"], v["category"], v["tasks"]) for v in catalog["threats"]]

    # Key tasks / roles per category (derived from the catalog)
    cat_tasks, cat_roles = {}, {}
    for c in cats:
        cnt = Counter(t for (_, _, cc, ts) in cat_rows if cc == c for t in ts)
        cat_tasks[c] = [t for t, _ in cnt.most_common(4)]
        rc = Counter(r for t in cat_tasks[c] for r in owners.get(t, []))
        cat_roles[c] = [r for r, _ in rc.most_common(3)]

    wb = Workbook()
    ws_ov = wb.active
    ws_ov.title = "Overview"
    ws_ed = wb.create_sheet("Entity Detail")
    ws_nat = wb.create_sheet("National View")
    ws_how = wb.create_sheet("How It Works")
    ws_w = wb.create_sheet("Weights")
    ws_p = wb.create_sheet("Programs")
    ws_c = wb.create_sheet("Categories")
    ws_i = wb.create_sheet("Incidents")
    ws_v = wb.create_sheet("Vulnerabilities")
    ws_t = wb.create_sheet("Threats")
    ws_cat = wb.create_sheet("Catalog")
    ws_e = wb.create_sheet("Entities")
    ws_st = wb.create_sheet("SCyWF Tasks")
    ws_sr = wb.create_sheet("SCyWF Roles")
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
    for ws in (ws_ov, ws_ed, ws_nat, ws_how):
        ws.sheet_properties.tabColor = GREEN

    # ================= Weights =================
    w = ws_w
    w["B1"], w["B1"].font = "Weights and factors", f_title
    w["B2"] = "Blue text on yellow = editable inputs. Every score in the workbook recalculates from these cells."
    w["B2"].font = f_note
    header(w, 4, 2, ["Group", "Label", "Value", "Why this value"], [24, 22, 10, 80])
    rows = [
        # (row, group, label, value, note, name)
        (5, "Data source weight", "Incidents", 0.5, "Realized harm: an incident proves the gap exists in this entity. Highest priority.", "W_INC"),
        (6, "Data source weight", "Vulnerabilities", 0.3, "Current exposure: weaknesses that exist now and can still be exploited.", "W_VUL"),
        (7, "Data source weight", "Threats", 0.2, "Likelihood: what attackers are doing; often sector-wide, so least entity-specific.", "W_THR"),
        (10, "Severity weight", "Low", 1, "Applied to incident severity, CVSS rating and threat level alike.", None),
        (11, "Severity weight", "Medium", 2, "", None),
        (12, "Severity weight", "High", 3, "", None),
        (13, "Severity weight", "Critical", 4, "", None),
        (16, "Recency", "As-of date", date(2026, 10, 31), "Age of incidents and threats is measured to this date.", "AsOfDate"),
        (17, "Recency", "Recent window (days)", 90, "Records newer than this get the full recency weight.", "Rec_D1"),
        (18, "Recency", "Weight: recent", 1.0, "", "Rec_W1"),
        (19, "Recency", "Mid window (days)", 365, "Records up to this age get the mid weight.", "Rec_D2"),
        (20, "Recency", "Weight: mid", 0.7, "", "Rec_W2"),
        (21, "Recency", "Weight: older", 0.4, "Old records still signal a gap, but a weaker one.", "Rec_W3"),
        (24, "Incident status", "Open", 1.0, "Unresolved incidents count fully.", None),
        (25, "Incident status", "Contained", 0.9, "", None),
        (26, "Incident status", "Closed", 0.75, "A closed incident still shows the skill was missing when it happened.", None),
        (29, "Vulnerability exposure", "Internet-facing", 1.5, "Reachable by any attacker.", None),
        (30, "Vulnerability exposure", "Internal", 1.0, "", None),
        (31, "Vulnerability status", "Open", 1.0, "", None),
        (32, "Vulnerability status", "Remediated", 0.2, "Fixed findings keep a small weight: the gap recurred and needed fixing.", None),
        (33, "CVSS threshold", "Critical (>=)", 9.0, "CVSS v3.x qualitative rating scale (FIRST).", "CVSS_C"),
        (34, "CVSS threshold", "High (>=)", 7.0, "", "CVSS_H"),
        (35, "CVSS threshold", "Medium (>=)", 4.0, "Below this the finding is Low.", "CVSS_M"),
        (38, "Threat relevance", "Entity-specific", 1.5, "Intelligence naming this entity.", None),
        (39, "Threat relevance", "Sector-targeted", 1.2, "Campaigns targeting the entity's sector.", None),
        (40, "Threat relevance", "General", 0.8, "Broad, untargeted activity.", None),
        (41, "Threat confidence", "High", 1.0, "Analyst confidence in the intelligence.", None),
        (42, "Threat confidence", "Medium", 0.7, "", None),
        (43, "Threat confidence", "Low", 0.4, "", None),
    ]
    for r, g, lab, val, note, nm in rows:
        w.cell(r, 2, g).font = f_base
        w.cell(r, 3, lab).font = f_base
        c = w.cell(r, 4, val)
        c.font, c.fill, c.border = f_input, fill_input, box
        c.number_format = "yyyy-mm-dd" if isinstance(val, date) else ("0" if isinstance(val, int) and val > 10 else "0.00")
        w.cell(r, 5, note).font = f_note
        if nm:
            name(wb, nm, f"Weights!$D${r}")
    w["C8"], w["C8"].font = "Total", f_bold
    w["D8"] = "=SUM(D5:D7)"
    w["D8"].number_format, w["D8"].font = "0.00", f_bold
    w["E8"] = "Need not equal 1: gap scores divide by the weights of the sources that have data for the entity."
    w["E8"].font = f_note
    for nm, ref in [("SevLabels", "$C$10:$C$13"), ("SevW", "$D$10:$D$13"), ("IStatLabels", "$C$24:$C$26"),
                    ("IStatW", "$D$24:$D$26"), ("ExpLabels", "$C$29:$C$30"), ("ExpW", "$D$29:$D$30"),
                    ("VStatLabels", "$C$31:$C$32"), ("VStatW", "$D$31:$D$32"), ("RelLabels", "$C$38:$C$40"),
                    ("RelW", "$D$38:$D$40"), ("ConfLabels", "$C$41:$C$43"), ("ConfW", "$D$41:$D$43")]:
        name(wb, nm, f"Weights!{ref}")
    w["B45"] = "Program coverage (on the Programs sheet): 1.0 = the program's primary focus, 0.5 = secondary benefit, blank = none."
    w["B45"].font = f_note
    w.freeze_panes = "A5"

    # ================= Categories =================
    c = ws_c
    header(c, 1, 1, ["#", "Gap category", "Key SCyWF tasks", "SCyWF job roles that own these tasks", "Recommended practices"], [5, 30, 60, 50, 70])
    for i, cat in enumerate(cats):
        r = i + 2
        c.cell(r, 1, i + 1)
        c.cell(r, 2, cat).font = f_bold
        c.cell(r, 3, "\n".join(f"{t} {tasks[t]}" for t in cat_tasks[cat]))
        c.cell(r, 4, "\n".join(f"{role_name[x]} ({x})" for x in cat_roles[cat]))
        c.cell(r, 5, "\n".join("• " + p for p in mapping["categories"][cat]["practices"]))
        for col in range(1, 6):
            c.cell(r, col).alignment, c.cell(r, col).border = wrap, box
            if col != 2: c.cell(r, col).font = f_base
    c.freeze_panes = "C2"
    NC = len(cats)
    CAT_RANGE = f"Categories!$B$2:$B${NC + 1}"

    # ================= Programs =================
    p = ws_p
    p["A1"] = None
    cols = ["ID", "Training program", "Level", "Format / duration", "Audience"] + cats + ["Target SCyWF job roles", "Key SCyWF tasks addressed", "Primary focus"]
    header(p, 1, 1, cols, [6, 40, 12, 26, 24] + [11] * NC + [46, 60, 26])
    p.row_dimensions[1].height = 60
    PCOL0 = 6  # first coverage column (F)
    for i, prog in enumerate(programs):
        r = i + 2
        vals = [prog["id"], prog["name"], prog["level"], prog["format"], prog["audience"]]
        for j, v in enumerate(vals):
            p.cell(r, j + 1, v).font = f_bold if j == 1 else f_base
        for j, cat in enumerate(cats):
            cell = p.cell(r, PCOL0 + j, prog["coverage"].get(cat))
            cell.font, cell.fill, cell.alignment, cell.number_format = f_input, fill_input, center, "0.0"
        prim = [k for k, v in prog["coverage"].items() if v == max(prog["coverage"].values())]
        ktasks = []
        for k in prim:
            ktasks += [t for t in cat_tasks[k] if t not in ktasks]
        p.cell(r, PCOL0 + NC, "\n".join(f"{role_name[x]} ({x})" for x in prog["roles"]))
        p.cell(r, PCOL0 + NC + 1, "\n".join(f"{t} {tasks[t]}" for t in ktasks[:4]))
        p.cell(r, PCOL0 + NC + 2, ", ".join(prim))
        for col in range(1, PCOL0 + NC + 3):
            p.cell(r, col).border = box
            if col < PCOL0 or col >= PCOL0 + NC:
                p.cell(r, col).alignment = wrap
                if col >= PCOL0 + NC: p.cell(r, col).font = f_base
    NP = len(programs)
    p.freeze_panes = "C2"
    p.cell(NP + 3, 2, "Coverage cells (yellow) are inputs: 1.0 = primary focus, 0.5 = secondary. Add a program by adding a row and extending the ranges on Entity Detail and National View.").font = f_note
    COV = lambda r: f"Programs!${L(PCOL0)}${r}:${L(PCOL0 + NC - 1)}${r}"

    # ================= Catalog =================
    k = ws_cat
    header(k, 1, 1, ["Source", "Title (incident / finding / threat type)", "Gap category", "Most-needed SCyWF tasks"], [14, 52, 30, 90])
    for i, (src, title, cat, ts) in enumerate(cat_rows):
        r = i + 2
        for j, v in enumerate([src, title, cat, "; ".join(f"{t} {tasks[t]}" for t in ts)]):
            cell = k.cell(r, j + 1, v)
            cell.font, cell.border, cell.alignment = f_base, box, Alignment(vertical="top", wrap_text=(j == 3))
    k.freeze_panes = "A2"
    k.auto_filter.ref = f"A1:D{len(cat_rows) + 1}"
    CATN = len(cat_rows) + 1
    CAT_TITLE, CAT_CAT = f"Catalog!$B$2:$B${CATN}", f"Catalog!$C$2:$C${CATN}"

    # ================= Entities =================
    e = ws_e
    header(e, 1, 1, ["Entity code", "Entity name", "Sector", "Region", "Maturity", "Search match #"], [12, 40, 20, 18, 12, 14])
    for i, en in enumerate(entities):
        r = i + 2
        for j, v in enumerate([en["entity_code"], en["entity_name"], en["sector"], en["region"], en["maturity_level"]]):
            e.cell(r, j + 1, v).font = f_base
        e.cell(r, 6, f'=IF(ISNUMBER(SEARCH(Overview!$C$5,A{r}&" "&B{r}&" "&C{r})),MAX(F$1:F{r - 1})+1,"")').font = f_note
    NE = len(entities) + 1
    e.freeze_panes = "A2"
    name(wb, "EntCodes", f"Entities!$A$2:$A${NE}")
    name(wb, "EntNames", f"Entities!$B$2:$B${NE}")

    # ================= Data sheets =================
    def rng(sheet, col):
        return f"'{sheet}'!${col}$2:${col}${MAXR}" if " " in sheet else f"{sheet}!${col}$2:${col}${MAXR}"

    recency = lambda dcell: f"IF(AsOfDate-{dcell}<=Rec_D1,Rec_W1,IF(AsOfDate-{dcell}<=Rec_D2,Rec_W2,Rec_W3))"
    catlook = lambda tcell: f'=IFERROR(INDEX({CAT_CAT},MATCH({tcell},{CAT_TITLE},0)),"Unmapped")'

    # Incidents: A id B code C name D title E severity F date G status | H category I sev J recency K status L score M rank helper
    ws = ws_i
    header(ws, 1, 1, ["Incident ID", "Entity code", "Entity name", "Vulnerability title", "Severity", "Reported date", "Status",
                      "Gap category", "Severity weight", "Recency factor", "Status factor", "Record score", "Selected-entity sort key"],
           [16, 11, 32, 50, 10, 13, 11, 28, 10, 10, 10, 11, 13])
    for i, x in enumerate(incidents):
        r = i + 2
        ws.append([x["incident_id"], x["entity_code"], x["entity_name"], x["vulnerability_title"], x["severity"],
                   date.fromisoformat(x["reported_date"]), x["status"],
                   catlook(f"D{r}"), f"=INDEX(SevW,MATCH(E{r},SevLabels,0))", "=" + recency(f"F{r}"),
                   f"=INDEX(IStatW,MATCH(G{r},IStatLabels,0))", f"=I{r}*J{r}*K{r}", f"=IF(B{r}=SelCode,L{r}+ROW()/10000000,0)"])
        ws.cell(r, 6).number_format = "yyyy-mm-dd"
        for col in (10, 11, 12): ws.cell(r, col).number_format = "0.00"
    # Vulnerabilities: A id B code C title D cvss E severity F exposure G status H date | I category J sev K exposure L status M score N helper
    ws = ws_v
    header(ws, 1, 1, ["Finding ID", "Entity code", "Finding title", "CVSS score", "Severity (from CVSS)", "Exposure", "Status",
                      "Detected date", "Gap category", "Severity weight", "Exposure factor", "Status factor", "Record score",
                      "Selected-entity sort key"], [16, 11, 46, 9, 12, 15, 11, 13, 28, 10, 10, 10, 11, 13])
    for i, x in enumerate(vulns):
        r = i + 2
        ws.append([x["finding_id"], x["entity_code"], x["finding_title"], float(x["cvss_score"]),
                   f'=IF(D{r}>=CVSS_C,"Critical",IF(D{r}>=CVSS_H,"High",IF(D{r}>=CVSS_M,"Medium","Low")))',
                   x["exposure"], x["status"], date.fromisoformat(x["detected_date"]),
                   catlook(f"C{r}"), f"=INDEX(SevW,MATCH(E{r},SevLabels,0))", f"=INDEX(ExpW,MATCH(F{r},ExpLabels,0))",
                   f"=INDEX(VStatW,MATCH(G{r},VStatLabels,0))", f"=J{r}*K{r}*L{r}", f"=IF(B{r}=SelCode,M{r}+ROW()/10000000,0)"])
        ws.cell(r, 8).number_format = "yyyy-mm-dd"
        ws.cell(r, 4).number_format = "0.0"
        for col in (11, 12, 13): ws.cell(r, col).number_format = "0.00"
    # Threats: A id B code C title D level E relevance F confidence G date | H category I level J relevance K confidence L recency M score N helper
    ws = ws_t
    header(ws, 1, 1, ["Threat ID", "Entity code", "Threat title", "Threat level", "Relevance", "Confidence", "Reported date",
                      "Gap category", "Level weight", "Relevance factor", "Confidence factor", "Recency factor", "Record score",
                      "Selected-entity sort key"], [16, 11, 50, 11, 15, 11, 13, 28, 10, 10, 10, 10, 11, 13])
    for i, x in enumerate(threats):
        r = i + 2
        ws.append([x["threat_id"], x["entity_code"], x["threat_title"], x["threat_level"], x["relevance"], x["confidence"],
                   date.fromisoformat(x["reported_date"]), catlook(f"C{r}"), f"=INDEX(SevW,MATCH(D{r},SevLabels,0))",
                   f"=INDEX(RelW,MATCH(E{r},RelLabels,0))", f"=INDEX(ConfW,MATCH(F{r},ConfLabels,0))", "=" + recency(f"G{r}"),
                   f"=I{r}*J{r}*K{r}*L{r}", f"=IF(B{r}=SelCode,M{r}+ROW()/10000000,0)"])
        ws.cell(r, 7).number_format = "yyyy-mm-dd"
        for col in (10, 11, 12, 13): ws.cell(r, col).number_format = "0.00"
    for ws, n, last in ((ws_i, len(incidents), "M"), (ws_v, len(vulns), "N"), (ws_t, len(threats), "N")):
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{last}{n + 1}"
        for row in ws.iter_rows(min_row=2, max_row=n + 1):
            for cell in row:
                cell.font = f_base
        ws.row_dimensions[1].height = 32

    SRC = {  # source: (sheet, entity col, category col, score col)
        "inc": ("Incidents", "B", "H", "L"), "vul": ("Vulnerabilities", "B", "I", "M"), "thr": ("Threats", "B", "H", "M")}
    sumifs = lambda s, ent, cat=None: (f"SUMIFS({rng(SRC[s][0], SRC[s][3])},{rng(SRC[s][0], SRC[s][1])},{ent}"
                                       + (f",{rng(SRC[s][0], SRC[s][2])},{cat})" if cat else ")"))

    # ================= Overview (part 1: selector, so SelCode exists) =================
    o = ws_ov
    for col, wd in zip("ABCDEFGHIJKLMN", [2, 14, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 14, 2]):
        o.column_dimensions[col].width = wd
    o["B2"], o["B2"].font = "SCyWF Training Priority Model", f_title
    o["B3"] = ("Ranks training programs for an entity from three data sources (incidents, vulnerabilities, threats), "
               "mapped to the Saudi Cybersecurity Workforce Framework (SCyWF 1.5 : 2026). DUMMY DATA for a proof of concept.")
    o["B3"].font = f_note
    o["B5"], o["B5"].font = "Find entity", f_bold
    o["C5"] = "HLT-003"
    o.merge_cells("C5:F5")
    style_range(o, "C5:F5", font=f_input, fill=fill_input, border=box)
    o["C5"].comment = Comment("Type an entity code (e.g. FIN-001) or part of its name (e.g. Rimal), or pick a code from the list.", "Model")
    dv = DataValidation(type="list", formula1="=EntCodes", allow_blank=True, showErrorMessage=False)
    o.add_data_validation(dv)
    dv.add("C5")
    o["B6"], o["B6"].font = "Selected", f_bold
    o["C6"] = ('=IFERROR(INDEX(EntCodes,MATCH(C5,EntCodes,0)),IFERROR(INDEX(EntCodes,MATCH("*"&C5&"*",EntNames,0)),""))')
    o["C6"].font = Font(name=FONT, size=11, bold=True, color=GREEN)
    name(wb, "SelCode", "Overview!$C$6")
    o["D6"] = '=IF(C6="","No entity matches. Check the search results on the right.","")'
    o["D6"].font = Font(name=FONT, size=9, bold=True, color="B3203A")
    o["O4"], o["O4"].font = "SEARCH RESULTS (code, name or sector)", f_label
    for col, wd in zip("OPQR", [12, 20, 20, 18]):
        o.column_dimensions[col].width = wd
    for k_ in range(8):
        r = 5 + k_
        o.cell(r, 15, f'=IFERROR(INDEX(EntCodes,MATCH({k_ + 1},Entities!$F$2:$F${NE},0)),"")').font = Font(name=FONT, size=9, bold=True, color=GREEN)
        o.cell(r, 16, f'=IFERROR(INDEX(EntNames,MATCH({k_ + 1},Entities!$F$2:$F${NE},0)),"")').font = Font(name=FONT, size=9, color=DARK)
        o.merge_cells(start_row=r, start_column=16, end_row=r, end_column=17)
        o.cell(r, 18, f'=IFERROR(INDEX(Entities!$C$2:$C${NE},MATCH({k_ + 1},Entities!$F$2:$F${NE},0)),"")').font = Font(name=FONT, size=9, color=MUTED)
    o["O13"] = '=IF(COUNT(Entities!$F$2:$F$%d)>8,"+ "&(COUNT(Entities!$F$2:$F$%d)-8)&" more: type more letters to narrow","")' % (NE, NE)
    o["O13"].font = f_note
    style_range(o, "O4:R13", fill=fill_card)

    # ================= Entity Detail =================
    d = ws_ed
    d.column_dimensions["A"].width = 30
    for col in range(2, 17):
        d.column_dimensions[L(col)].width = 13
    d.column_dimensions["N"].width = 48
    d.column_dimensions["O"].width = 40
    d.column_dimensions["P"].width = 48
    d["A1"], d["A1"].font = "Entity detail", f_title
    d["A2"], d["A2"].font = "Entity", f_bold
    d["B2"] = "=SelCode"
    d["B2"].font = f_link
    d["C2"] = '=IFERROR(INDEX(EntNames,MATCH(SelCode,EntCodes,0)),"")'
    d["C2"].font = f_bold
    d["A3"] = "Change the entity on the Overview sheet."
    d["A3"].font = f_note

    CT0, CT1 = 13, 13 + NC - 1  # category table rows
    # helper rows 5-8 (horizontal copy used for program scoring)
    d["A4"], d["A4"].font = "Category vector (helper for program scoring)", f_label
    labels = ["Incident contribution", "Vulnerability contribution", "Threat contribution", "Gap index"]
    srccol = ["H", "I", "J", "K"]
    for j, cat in enumerate(cats):
        col = 2 + j
        d.cell(5, col, f"=INDEX($A${CT0}:$A${CT1},{j + 1})").font = f_label
        d.cell(5, col).alignment = center
        for i_, sc in enumerate(srccol):
            cell = d.cell(6 + i_, col, f"=INDEX(${sc}${CT0}:${sc}${CT1},{j + 1})")
            cell.number_format, cell.font = "0.0%", f_base
    for i_, lab in enumerate(labels):
        d.cell(6 + i_, 1, lab).font = f_base
    d.row_dimensions[5].height = 42

    d["A11"], d["A11"].font = "Gap by category", f_h2
    header(d, 12, 1, ["Gap category", "Incident score", "Vulnerability score", "Threat score", "Incident share",
                      "Vulnerability share", "Threat share", "Incident contribution", "Vulnerability contribution",
                      "Threat contribution", "Gap index", "Sort key", "Rank", "Key SCyWF tasks", "SCyWF job roles",
                      "Recommended practices"])
    d.row_dimensions[12].height = 32
    for j, cat in enumerate(cats):
        r = CT0 + j
        cr = j + 2
        d.cell(r, 1, f"=Categories!B{cr}")
        d.cell(r, 2, "=" + sumifs("inc", "SelCode", f"$A{r}"))
        d.cell(r, 3, "=" + sumifs("vul", "SelCode", f"$A{r}"))
        d.cell(r, 4, "=" + sumifs("thr", "SelCode", f"$A{r}"))
        for col, src in ((5, "B"), (6, "C"), (7, "D")):
            d.cell(r, col, f"=IFERROR({src}{r}/{src}${CT1 + 1},0)")
        for col, (sh, wn) in zip((8, 9, 10), (("E", "W_INC"), ("F", "W_VUL"), ("G", "W_THR"))):
            d.cell(r, col, f"=IFERROR({wn}*{sh}{r}/$B${CT1 + 3},0)")
        d.cell(r, 11, f"=H{r}+I{r}+J{r}")
        d.cell(r, 12, f"=K{r}+ROW()/10000000")
        d.cell(r, 13, f"=RANK(L{r},$L${CT0}:$L${CT1})")
        d.cell(r, 14, f"=Categories!C{cr}")
        d.cell(r, 15, f"=Categories!D{cr}")
        d.cell(r, 16, f"=Categories!E{cr}")
        for col in range(1, 17):
            cell = d.cell(r, col)
            cell.border, cell.font = box, f_base
            cell.alignment = wrap if col >= 14 else Alignment(vertical="top")
            if 2 <= col <= 4: cell.number_format = "0.0"
            elif 5 <= col <= 11: cell.number_format = "0.0%"
            elif col == 12: cell.number_format = "0.0000000"
        d.cell(r, 1).font = f_bold
        d.cell(r, 11).font = f_bold
        d.cell(r, 12).font = f_note
    tr = CT1 + 1
    d.cell(tr, 1, "Total").font = f_bold
    for col in range(2, 12):
        if col in (2, 3, 4, 8, 9, 10, 11):
            cl = L(col)
            cell = d.cell(tr, col, f"=SUM({cl}{CT0}:{cl}{CT1})")
            cell.font, cell.number_format, cell.border = f_bold, ("0.0" if col <= 4 else "0.0%"), box
    d.cell(tr + 2, 1, "Weights of sources with data").font = f_bold
    d.cell(tr + 2, 2, f"=W_INC*(B{tr}>0)+W_VUL*(C{tr}>0)+W_THR*(D{tr}>0)").number_format = "0.00"
    d.cell(tr + 2, 3, "A source with no records for this entity drops out, and the other weights are scaled up.").font = f_note
    d.conditional_formatting.add(f"K{CT0}:K{CT1}", DataBarRule(start_type="num", start_value=0, end_type="max", color="3FBF8F"))
    EFFW_CELL = f"$B${tr + 2}"

    # programs
    PT0 = tr + 5
    d.cell(PT0 - 1, 1, "Training program ranking").font = f_h2
    d.cell(PT0, 1, "Priority score = sum over categories of (gap index x program coverage). The source columns split that score by data source.").font = f_note
    header(d, PT0 + 1, 1, ["Program", "ID", "Incident-driven", "Vulnerability-driven", "Threat-driven", "Priority score",
                           "Sort key", "Rank", "Level", "Primary focus", "Format / duration", "Audience", "",
                           "Target SCyWF job roles", "Key SCyWF tasks addressed"])
    d.row_dimensions[PT0 + 1].height = 32
    P0, P1 = PT0 + 2, PT0 + 2 + NP - 1
    span = f"$B$%d:${L(1 + NC)}$%d"
    for i_ in range(NP):
        r, pr = P0 + i_, i_ + 2
        d.cell(r, 1, f"=Programs!B{pr}")
        d.cell(r, 2, f"=Programs!A{pr}")
        for col, hr in ((3, 6), (4, 7), (5, 8)):
            d.cell(r, col, f"=SUMPRODUCT({COV(pr)},{span % (hr, hr)})")
        d.cell(r, 6, f"=C{r}+D{r}+E{r}")
        d.cell(r, 7, f"=F{r}+ROW()/10000000")
        d.cell(r, 8, f"=RANK(G{r},$G${P0}:$G${P1})")
        d.cell(r, 9, f"=Programs!C{pr}")
        d.cell(r, 10, f"=Programs!{L(PCOL0 + NC + 2)}{pr}")
        d.cell(r, 11, f"=Programs!D{pr}")
        d.cell(r, 12, f"=Programs!E{pr}")
        d.cell(r, 14, f"=Programs!{L(PCOL0 + NC)}{pr}")
        d.cell(r, 15, f"=Programs!{L(PCOL0 + NC + 1)}{pr}")
        for col in list(range(1, 13)) + [14, 15]:
            cell = d.cell(r, col)
            cell.border, cell.font = box, f_base
            cell.alignment = wrap
            if 3 <= col <= 6: cell.number_format = "0.000"
            if col == 7: cell.number_format = "0.0000000"
        d.cell(r, 1).font = f_bold
        d.cell(r, 6).font = f_bold
        d.cell(r, 7).font = f_note
    d.conditional_formatting.add(f"F{P0}:F{P1}", DataBarRule(start_type="num", start_value=0, end_type="max", color="3FBF8F"))

    # top records per source
    R0 = P1 + 3
    d.cell(R0, 1, "Top records behind the scores").font = f_h2
    blocks = [
        ("Incidents", "Incidents", "M", ["Incident ID", "Vulnerability title", "Gap category", "Severity", "Reported date", "Status", "Record score"],
         ["A", "D", "H", "E", "F", "G", "L"]),
        ("Vulnerabilities", "Vulnerabilities", "N", ["Finding ID", "Finding title", "Gap category", "Severity", "Detected date", "Exposure / status", "Record score"],
         ["A", "C", "I", "E", "H", None, "M"]),
        ("Threats", "Threats", "N", ["Threat ID", "Threat title", "Gap category", "Threat level", "Reported date", "Relevance", "Record score"],
         ["A", "C", "H", "D", "G", "E", "M"]),
    ]
    r = R0 + 1
    for title, sh, keycol, heads, srccols in blocks:
        d.cell(r, 1, f"{title}: top 10 by record score").font = f_bold
        header(d, r + 1, 1, ["#"] + heads[:1] + ["", ""] + heads[1:], None)
        # layout: A #, B id, C-D (spare, merged into title), E title... keep simple: A #, B id, C title (merge C:E), F category, G sev, H date, I status, J score
        for cc in range(1, 16):
            d.cell(r + 1, cc).value = None
            d.cell(r + 1, cc).fill = PatternFill()
            d.cell(r + 1, cc).font = f_base
            d.cell(r + 1, cc).border = Border()
        header(d, r + 1, 1, ["#", heads[0], heads[1], "", "", heads[2], heads[3], heads[4], heads[5], heads[6], "Row (helper)"])
        d.merge_cells(start_row=r + 1, start_column=3, end_row=r + 1, end_column=5)
        for kk in range(10):
            rr = r + 2 + kk
            key = f"{sh}!${keycol}$2:${keycol}${MAXR}"
            d.cell(rr, 11, f'=IF(LARGE({key},{kk + 1})<=0,"",MATCH(LARGE({key},{kk + 1}),{key},0))').font = f_note
            d.cell(rr, 1, kk + 1).font = f_base
            vals = srccols
            tgt = [2, 3, 6, 7, 8, 9, 10]
            for tcol, scol in zip(tgt, vals):
                if scol is None:  # exposure / status combined
                    f = f'=IF($K{rr}="","",INDEX({sh}!$F$2:$F${MAXR},$K{rr})&" / "&INDEX({sh}!$G$2:$G${MAXR},$K{rr}))'
                else:
                    f = f'=IF($K{rr}="","",INDEX({sh}!${scol}$2:${scol}${MAXR},$K{rr}))'
                cell = d.cell(rr, tcol, f)
                cell.font = f_base
                if tcol == 8: cell.number_format = "yyyy-mm-dd"
                if tcol == 10: cell.number_format = "0.00"
            d.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=5)
            for cc in range(1, 11):
                d.cell(rr, cc).border = bottom
        r += 14
    d.freeze_panes = "B4"

    # ================= Overview (part 2: card, KPIs, rankings, charts) =================
    o.merge_cells("B8:M8")
    o["B8"] = '=IFERROR(INDEX(EntNames,MATCH(SelCode,EntCodes,0)),"")'
    o["B8"].font = f_entity
    o.merge_cells("B9:M9")
    o["B9"] = ('=IFERROR(SelCode&"   |   "&INDEX(Entities!$C$2:$C$%d,MATCH(SelCode,EntCodes,0))&"   |   "&'
               'INDEX(Entities!$D$2:$D$%d,MATCH(SelCode,EntCodes,0))&"   |   Maturity: "&INDEX(Entities!$E$2:$E$%d,MATCH(SelCode,EntCodes,0)),"")') % (NE, NE, NE)
    o["B9"].font = Font(name=FONT, size=10, color=MUTED)
    style_range(o, "B7:M10", fill=fill_card)
    kpis = [
        ("INCIDENTS", f"=COUNTIF({rng('Incidents', 'B')},SelCode)", False),
        ("CRITICAL INCIDENTS", f'=COUNTIFS({rng("Incidents", "B")},SelCode,{rng("Incidents", "E")},"Critical")', True),
        ("OPEN VULNERABILITIES", f'=COUNTIFS({rng("Vulnerabilities", "B")},SelCode,{rng("Vulnerabilities", "G")},"Open")', False),
        ("OPEN CRITICAL, INTERNET-FACING", f'=COUNTIFS({rng("Vulnerabilities", "B")},SelCode,{rng("Vulnerabilities", "G")},"Open",{rng("Vulnerabilities", "E")},"Critical",{rng("Vulnerabilities", "F")},"Internet-facing")', True),
        ("THREAT RECORDS", f"=COUNTIF({rng('Threats', 'B')},SelCode)", False),
        ("TARGETED THREATS", f'=COUNTIFS({rng("Threats", "B")},SelCode,{rng("Threats", "E")},"<>General")', True),
    ]
    for i_, (lab, f, red) in enumerate(kpis):
        c0 = 2 + i_ * 2
        o.merge_cells(start_row=11, start_column=c0, end_row=11, end_column=c0 + 1)
        o.merge_cells(start_row=12, start_column=c0, end_row=12, end_column=c0 + 1)
        o.cell(11, c0, lab).font = f_label
        o.cell(11, c0).alignment = Alignment(horizontal="left", vertical="bottom", wrap_text=True)
        o.cell(12, c0, f).font = f_kpi_red if red else f_kpi
        o.cell(12, c0).alignment = Alignment(horizontal="left")
        o.cell(12, c0).number_format = "#,##0"
    o.row_dimensions[11].height = 24
    o.row_dimensions[12].height = 30

    o["B14"], o["B14"].font = "Where the gap signal comes from", f_h2
    header(o, 15, 2, ["Data source", "Weight set", "Effective weight", "Records", "Weighted score", "Top gap category from this source"])
    o.merge_cells("G15:M15")
    style_range(o, "G15:M15", font=f_head, fill=fill_head, border=box)
    for i_, (lab, wn, sh, scol, cnt) in enumerate([
            ("Incidents", "W_INC", "B", "E", rng("Incidents", "B")),
            ("Vulnerabilities", "W_VUL", "C", "F", rng("Vulnerabilities", "B")),
            ("Threats", "W_THR", "D", "G", rng("Threats", "B"))]):
        r = 16 + i_
        o.cell(r, 2, lab).font = f_bold
        o.cell(r, 3, f"={wn}").number_format = "0.00"
        o.cell(r, 4, f"=IFERROR({wn}*('Entity Detail'!{sh}{tr}>0)/'Entity Detail'!{EFFW_CELL},0)").number_format = "0%"
        o.cell(r, 5, f"=COUNTIF({cnt},SelCode)").number_format = "#,##0"
        o.cell(r, 6, f"='Entity Detail'!{sh}{tr}").number_format = "0.0"
        o.cell(r, 7, f"=IF('Entity Detail'!{sh}{tr}=0,\"No records\",INDEX('Entity Detail'!$A${CT0}:$A${CT1},MATCH(MAX('Entity Detail'!{scol}{CT0}:{scol}{CT1}),'Entity Detail'!{scol}{CT0}:{scol}{CT1},0))&\"  (\"&TEXT(MAX('Entity Detail'!{scol}{CT0}:{scol}{CT1}),\"0%\")&\" of this source)\")")
        o.merge_cells(start_row=r, start_column=7, end_row=r, end_column=13)
        for col in range(2, 14):
            o.cell(r, col).border = box
            if col != 2: o.cell(r, col).font = f_base

    o["B20"], o["B20"].font = "Top skill-gap categories", f_h2
    o["B21"] = "Gap index = weighted share of the entity's risk signal in each category (all categories sum to 100%). Source columns split it."
    o["B21"].font = f_note
    header(o, 22, 2, ["Rank", "Gap category", "", "", "Gap index", "From incidents", "From vulns", "From threats", "Main driver", "Key SCyWF tasks", "", ""])
    o.merge_cells("C22:E22")
    o.merge_cells("K22:M22")
    o.row_dimensions[22].height = 28
    for k_ in range(5):
        r = 23 + k_
        m = f"MATCH({k_ + 1},'Entity Detail'!$M${CT0}:$M${CT1},0)"
        o.cell(r, 2, k_ + 1)
        o.cell(r, 3, f"=INDEX('Entity Detail'!$A${CT0}:$A${CT1},{m})")
        o.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        for col, sc in ((6, "K"), (7, "H"), (8, "I"), (9, "J")):
            o.cell(r, col, f"=INDEX('Entity Detail'!${sc}${CT0}:${sc}${CT1},{m})").number_format = "0.0%"
        o.cell(r, 10, f'=IF(F{r}=0,"-",CHOOSE(MATCH(MAX(G{r}:I{r}),G{r}:I{r},0),"Incidents","Vulnerabilities","Threats"))')
        o.cell(r, 11, f"=INDEX('Entity Detail'!$N${CT0}:$N${CT1},{m})")
        o.merge_cells(start_row=r, start_column=11, end_row=r, end_column=13)
        o.row_dimensions[r].height = 66
        for col in range(2, 14):
            cell = o.cell(r, col)
            cell.border, cell.font = box, f_base
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        o.cell(r, 2).alignment = Alignment(horizontal="center", vertical="top")
        o.cell(r, 3).font = f_bold
        o.cell(r, 6).font = f_bold
        o.cell(r, 11).font = Font(name=FONT, size=8, color=DARK)
    o.conditional_formatting.add("F23:F27", DataBarRule(start_type="num", start_value=0, end_type="max", color="3FBF8F"))

    o["B29"], o["B29"].font = "Recommended training programs", f_h2
    o["B30"] = "Priority score = sum over categories of (gap index x program coverage). Full ranking on Entity Detail."
    o["B30"].font = f_note
    header(o, 31, 2, ["Rank", "Training program", "", "", "", "Priority score", "Level", "Primary focus", "", "Target SCyWF job roles", "", ""])
    o.merge_cells("C31:F31")
    o.merge_cells("I31:J31")
    o.merge_cells("K31:M31")
    for k_ in range(5):
        r = 32 + k_
        m = f"MATCH({k_ + 1},'Entity Detail'!$H${P0}:$H${P1},0)"
        o.cell(r, 2, k_ + 1)
        o.cell(r, 3, f"=INDEX('Entity Detail'!$A${P0}:$A${P1},{m})")
        o.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
        o.cell(r, 7, f"=INDEX('Entity Detail'!$F${P0}:$F${P1},{m})").number_format = "0.000"
        o.cell(r, 8, f"=INDEX('Entity Detail'!$I${P0}:$I${P1},{m})")
        o.cell(r, 9, f"=INDEX('Entity Detail'!$J${P0}:$J${P1},{m})")
        o.merge_cells(start_row=r, start_column=9, end_row=r, end_column=10)
        o.cell(r, 11, f"=INDEX('Entity Detail'!$N${P0}:$N${P1},{m})")
        o.merge_cells(start_row=r, start_column=11, end_row=r, end_column=13)
        o.row_dimensions[r].height = 42
        for col in range(2, 14):
            cell = o.cell(r, col)
            cell.border, cell.font = box, f_base
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        o.cell(r, 2).alignment = Alignment(horizontal="center", vertical="top")
        o.cell(r, 3).font = f_bold
        o.cell(r, 7).font = f_bold
        o.cell(r, 11).font = Font(name=FONT, size=8, color=DARK)
    o.conditional_formatting.add("G32:G36", DataBarRule(start_type="num", start_value=0, end_type="max", color="3FBF8F"))
    o["B38"] = "Colour key: blue on yellow = input you can change. Weights are on the Weights sheet; program coverage on the Programs sheet."
    o["B38"].font = f_note

    def bar(title, cats_ref, data_ref, anchor, fmt):
        ch = BarChart()
        ch.type, ch.style, ch.title = "bar", 10, title
        ch.add_data(data_ref, titles_from_data=False)
        ch.set_categories(cats_ref)
        ch.legend = None
        ch.y_axis.numFmt = fmt
        ch.y_axis.majorGridlines = None
        ch.x_axis.scaling.orientation = "maxMin"
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        ch.roundedCorners = False
        ch.title.overlay = False
        ch.series[0].graphicalProperties.solidFill = GREEN
        ch.series[0].graphicalProperties.line.solidFill = GREEN
        ch.height, ch.width = 7.5, 15
        o.add_chart(ch, anchor)

    bar("Gap index, top 5 categories", Reference(o, min_col=3, min_row=23, max_row=27),
        Reference(o, min_col=6, min_row=23, max_row=27), "O15", "0%")
    bar("Priority score, top 5 programs", Reference(o, min_col=3, min_row=32, max_row=36),
        Reference(o, min_col=7, min_row=32, max_row=36), "O29", "0.00")
    o.freeze_panes = "A7"

    # ================= National View =================
    n = ws_nat
    n["A1"], n["A1"].font = "National view: gaps and program priorities across all entities", f_title
    n["A2"] = "Same model as the Overview, computed for every entity. Darker cells show bigger gaps or higher program priority."
    n["A2"].font = f_note
    G0 = 9  # first gap column (I)
    PG0 = G0 + NC + 4  # first program score column
    heads = ["Entity code", "Entity name", "Sector", "Maturity", "Incident score", "Vulnerability score", "Threat score",
             "Weights with data"] + cats + ["Top gap category", "Top program", "Top program score", ""]
    header(n, 4, 1, heads)
    header(n, 4, PG0, [p_["id"] for p_ in programs])
    for i_, p_ in enumerate(programs):
        cell = n.cell(3, PG0 + i_, p_["name"])
        cell.font, cell.alignment = f_label, Alignment(wrap_text=True, vertical="bottom")
    n.cell(4, PG0 - 1).fill = PatternFill()
    n.cell(4, PG0 - 1).border = Border()
    n.row_dimensions[3].height = 64
    n.row_dimensions[4].height = 64
    for col, wd in zip(range(1, 9), [11, 34, 16, 11, 10, 10, 10, 9]):
        n.column_dimensions[L(col)].width = wd
    for j in range(NC):
        n.column_dimensions[L(G0 + j)].width = 10
    n.column_dimensions[L(G0 + NC)].width = 28
    n.column_dimensions[L(G0 + NC + 1)].width = 40
    n.column_dimensions[L(G0 + NC + 2)].width = 10
    n.column_dimensions[L(G0 + NC + 3)].width = 3
    for i_ in range(NP):
        n.column_dimensions[L(PG0 + i_)].width = 10
    gl, gr = L(G0), L(G0 + NC - 1)
    pl, pr_ = L(PG0), L(PG0 + NP - 1)
    E0, E1 = 5, 5 + len(entities) - 1
    for i_ in range(len(entities)):
        r, er = E0 + i_, i_ + 2
        n.cell(r, 1, f"=Entities!A{er}")
        n.cell(r, 2, f"=Entities!B{er}")
        n.cell(r, 3, f"=Entities!C{er}")
        n.cell(r, 4, f"=Entities!E{er}")
        n.cell(r, 5, "=" + sumifs("inc", f"$A{r}"))
        n.cell(r, 6, "=" + sumifs("vul", f"$A{r}"))
        n.cell(r, 7, "=" + sumifs("thr", f"$A{r}"))
        n.cell(r, 8, f"=W_INC*(E{r}>0)+W_VUL*(F{r}>0)+W_THR*(G{r}>0)")
        for j in range(NC):
            hc = f"{L(G0 + j)}$4"
            f = (f"=IF($H{r}=0,0,(W_INC*IFERROR({sumifs('inc', f'$A{r}', hc)}/$E{r},0)"
                 f"+W_VUL*IFERROR({sumifs('vul', f'$A{r}', hc)}/$F{r},0)"
                 f"+W_THR*IFERROR({sumifs('thr', f'$A{r}', hc)}/$G{r},0))/$H{r})")
            n.cell(r, G0 + j, f).number_format = "0%"
        n.cell(r, G0 + NC, f"=INDEX(${gl}$4:${gr}$4,MATCH(MAX({gl}{r}:{gr}{r}),{gl}{r}:{gr}{r},0))")
        n.cell(r, G0 + NC + 1, f"=INDEX(${pl}$3:${pr_}$3,MATCH(MAX({pl}{r}:{pr_}{r}),{pl}{r}:{pr_}{r},0))")
        n.cell(r, G0 + NC + 2, f"=MAX({pl}{r}:{pr_}{r})").number_format = "0.000"
        for i2 in range(NP):
            n.cell(r, PG0 + i2, f"=SUMPRODUCT({COV(i2 + 2)},${gl}{r}:${gr}{r})").number_format = "0.000"
        for col in list(range(1, G0 + NC + 3)) + list(range(PG0, PG0 + NP)):
            n.cell(r, col).font = f_base
            n.cell(r, col).border = box
        for col in (5, 6, 7):
            n.cell(r, col).number_format = "0.0"
        n.cell(r, 8).number_format = "0.00"
    ar = E1 + 2
    n.cell(ar, 2, "National average").font = f_bold
    n.cell(ar + 1, 2, "Entities where this is the top item").font = f_bold
    for j in range(NC):
        cl = L(G0 + j)
        n.cell(ar, G0 + j, f"=AVERAGE({cl}{E0}:{cl}{E1})").number_format = "0%"
        n.cell(ar + 1, G0 + j, f"=COUNTIF(${L(G0 + NC)}${E0}:${L(G0 + NC)}${E1},{cl}$4)")
    for i2 in range(NP):
        cl = L(PG0 + i2)
        n.cell(ar, PG0 + i2, f"=AVERAGE({cl}{E0}:{cl}{E1})").number_format = "0.000"
        n.cell(ar + 1, PG0 + i2, f"=COUNTIF(${L(G0 + NC + 1)}${E0}:${L(G0 + NC + 1)}${E1},{cl}$3)")
    for rr in (ar, ar + 1):
        for col in list(range(G0, G0 + NC)) + list(range(PG0, PG0 + NP)):
            n.cell(rr, col).font, n.cell(rr, col).fill, n.cell(rr, col).border = f_bold, fill_sub, box
    # sector summary
    sectors = sorted({x["sector"] for x in entities})
    s0 = ar + 4
    n.cell(s0 - 1, 1, "Sector averages").font = f_h2
    header(n, s0, 1, ["", "Sector", "Entities", "", "", "", "", ""] + cats + ["Top gap category", "Top program", "Top program score"])
    header(n, s0, PG0, [p_["id"] for p_ in programs])
    for i_, sec in enumerate(sectors):
        r = s0 + 1 + i_
        n.cell(r, 2, sec).font = f_bold
        n.cell(r, 3, f'=COUNTIF($C${E0}:$C${E1},$B{r})')
        for j in range(NC):
            cl = L(G0 + j)
            n.cell(r, G0 + j, f"=AVERAGEIFS({cl}${E0}:{cl}${E1},$C${E0}:$C${E1},$B{r})").number_format = "0%"
        for i2 in range(NP):
            cl = L(PG0 + i2)
            n.cell(r, PG0 + i2, f"=AVERAGEIFS({cl}${E0}:{cl}${E1},$C${E0}:$C${E1},$B{r})").number_format = "0.000"
        n.cell(r, G0 + NC, f"=INDEX(${gl}$4:${gr}$4,MATCH(MAX({gl}{r}:{gr}{r}),{gl}{r}:{gr}{r},0))")
        n.cell(r, G0 + NC + 1, f"=INDEX(${pl}$3:${pr_}$3,MATCH(MAX({pl}{r}:{pr_}{r}),{pl}{r}:{pr_}{r},0))")
        n.cell(r, G0 + NC + 2, f"=MAX({pl}{r}:{pr_}{r})").number_format = "0.000"
        for col in list(range(2, G0 + NC + 3)) + list(range(PG0, PG0 + NP)):
            n.cell(r, col).border = box
            if col != 2: n.cell(r, col).font = f_base
    S1 = s0 + len(sectors)
    scale = ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF", end_type="max", end_color="3FBF8F")
    for ref in (f"{gl}{E0}:{gr}{E1}", f"{pl}{E0}:{pr_}{E1}", f"{gl}{s0 + 1}:{gr}{S1}", f"{pl}{s0 + 1}:{pr_}{S1}"):
        n.conditional_formatting.add(ref, scale)
    n.freeze_panes = "C5"
    n.auto_filter.ref = f"A4:{L(G0 + NC + 2)}{E1}"

    # ================= How It Works =================
    h = ws_how
    h.column_dimensions["B"].width = 120
    lines = [
        ("How the model works", f_title),
        ("", None),
        ("1. Three data sources per entity", f_h2),
        ("Incidents (what went wrong), vulnerability findings (what is exposed now) and threat intelligence (what attackers are doing). Each record is mapped to one of 14 gap categories through the Catalog sheet.", f_base),
        ("", None),
        ("2. Record score: how strong a signal each record is", f_h2),
        ("Incident score = severity weight x recency factor x status factor", f_bold),
        ("Vulnerability score = severity weight (from CVSS) x exposure factor x status factor", f_bold),
        ("Threat score = threat-level weight x relevance factor x confidence factor x recency factor", f_bold),
        ("Severity weights: Low 1, Medium 2, High 3, Critical 4. All factors are on the Weights sheet.", f_base),
        ("", None),
        ("3. Source share: normalise each source before combining", f_h2),
        ("For one entity, source share(category) = record scores in that category / all record scores from that source.", f_bold),
        ("Why: sources have very different volumes (a scan produces far more findings than there are incidents). Shares make them comparable, so the source weights decide their influence, not the record count.", f_base),
        ("", None),
        ("4. Gap index: weighted blend of the three sources", f_h2),
        ("Gap index(category) = [ W_inc x incident share + W_vul x vulnerability share + W_thr x threat share ] / (sum of weights of sources that have data)", f_bold),
        ("Default weights: Incidents 0.5 (proven harm), Vulnerabilities 0.3 (current exposure), Threats 0.2 (likelihood, least entity-specific). A source with no records drops out and the others are scaled up. Gap indexes across all categories add up to 100%.", f_base),
        ("", None),
        ("5. Program priority: rank training programs", f_h2),
        ("Program priority score = sum over categories of [ gap index(category) x program coverage(category) ]", f_bold),
        ("Coverage is 1.0 for a program's primary focus and 0.5 for a secondary benefit (Programs sheet). The score is split into incident-, vulnerability- and threat-driven parts so you can see why a program ranks where it does.", f_base),
        ("", None),
        ("6. Link to SCyWF", f_h2),
        ("Each catalog title lists the SCyWF tasks most needed to address it. Each category lists its key tasks and the SCyWF job roles that own them (from Appendix A of the framework). Each program lists its target SCyWF job roles.", f_base),
        ("", None),
        ("Sheets", f_h2),
        ("Overview: search an entity, see its summary card, top gaps and top 5 programs.", f_base),
        ("Entity Detail: the full calculation for the selected entity, all programs ranked, and the top records behind the scores.", f_base),
        ("National View: every entity and sector side by side, with national averages.", f_base),
        ("Weights, Programs: inputs you can change. Categories, Catalog: mappings. Incidents, Vulnerabilities, Threats: data. SCyWF Tasks, SCyWF Roles: framework reference.", f_base),
        ("", None),
        ("Adding real data", f_h2),
        (f"Paste records under the existing rows of Incidents, Vulnerabilities or Threats and copy the grey calculated columns down. Formulas read up to row {MAXR}. Every title must exist in the Catalog sheet; otherwise its category shows 'Unmapped' and it is ignored.", f_base),
        ("Source: task IDs, task descriptions and job roles are from the Saudi Cybersecurity Workforce Framework (SCyWF 1.5 : 2026), National Cybersecurity Authority, extracted by OCR. The catalog mapping, weights and programs are proposals for review. Entities and records are fictional.", f_note),
    ]
    for i_, (txt, fnt) in enumerate(lines):
        cell = h.cell(i_ + 1, 2, txt)
        cell.font = fnt or f_base
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    # ================= SCyWF reference =================
    header(ws_st, 1, 1, ["Task ID", "Task description", "Specialty area (by ID range)", "Job roles that list this task"], [10, 90, 34, 60])
    ranges = [(0, 499, "General"), (500, 999, "Cybersecurity Architecture"), (1000, 1499, "Cybersecurity Research and Development"),
              (1500, 1999, "Leadership"), (2000, 2499, "Workforce Development"), (2500, 2999, "Governance, Risk and Compliance"),
              (3000, 3499, "Laws and Data Protection"), (3500, 3999, "Defense"), (4000, 4499, "Protection"),
              (4500, 4999, "Vulnerability Assessment"), (5000, 5499, "Incident Response"), (5500, 5999, "Threat Management"),
              (6000, 6499, "ICS/OT")]
    for tid, desc in tasks.items():
        num = int(tid[1:])
        ws_st.append([tid, desc, next(nm for lo, hi, nm in ranges if lo <= num <= hi), "; ".join(owners.get(tid, []))])
    header(ws_sr, 1, 1, ["Job role ID", "Job role", "Category", "Specialty area", "Number of tasks", "Task IDs"], [18, 44, 44, 34, 10, 90])
    for r_ in roles:
        ws_sr.append([r_["id"], r_["name"], r_["category"], r_["specialty_area"], len(r_["tasks"]), ", ".join(r_["tasks"])])
    for ws in (ws_st, ws_sr):
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = f"A1:{L(ws.max_column)}{ws.max_row}"
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.font = f_base
                cell.alignment = Alignment(vertical="top", wrap_text=cell.column >= 2)

    # Grey "calculated" columns in data sheets
    for ws, first in ((ws_i, 8), (ws_v, 9), (ws_t, 8)):
        for col in range(first, ws.max_column + 1):
            ws.cell(1, col).fill = PatternFill("solid", fgColor="56655E")
    for ws, col in ((ws_v, 5),):
        ws.cell(1, col).fill = PatternFill("solid", fgColor="56655E")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"Saved {OUT}")


if __name__ == "__main__":
    main()
