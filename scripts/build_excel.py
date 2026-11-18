#!/usr/bin/env python3
"""Build the SCyWF Training Priority Model workbook (excel/SCyWF_Training_Priority_Model.xlsx).

Three data sources: incidents, vulnerability findings, NCA ECC compliance results.
  record weight  = severity weight (or ECC status weight); frequency = records add up
  task score     = record weight x the task's allocation (task relevance / sum of relevances of the title)
  gap index      = sum over sources of source weight x (task score / source total)   -> sums to 100% over tasks
  program score  = sum of the gap index of the tasks the program teaches
All live Excel formulas. Selection = an entity, a sector, or All.
"""
import csv
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "excel" / "SCyWF_Training_Priority_Model.xlsx"
MAXR = 10000  # formula ranges on the data sheets cover rows 2..MAXR

FONT, GREEN, DARK, MUTED = "Arial", "0B6E4F", "14201B", "56655E"
F = lambda **k: Font(name=FONT, **{"size": 10, "color": DARK, **k})
f_base, f_bold, f_small = F(), F(bold=True), F(size=8)
f_head, f_title, f_h2 = F(bold=True, color="FFFFFF"), F(size=18, bold=True), F(size=12, bold=True, color=GREEN)
f_note, f_input, f_id = F(size=9, italic=True, color=MUTED), F(size=11, bold=True, color="0000FF"), F(bold=True, color=GREEN)
f_label, f_kpi, f_kpi_red = F(size=8, bold=True, color=MUTED), F(size=20, bold=True), F(size=20, bold=True, color="B3203A")
fill_head, fill_calc = PatternFill("solid", fgColor=GREEN), PatternFill("solid", fgColor="56655E")
fill_input, fill_card = PatternFill("solid", fgColor="FFF59D"), PatternFill("solid", fgColor="EEF4F1")
thin = Side(style="thin", color="D5DDD9")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
center, top_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True), Alignment(vertical="top", wrap_text=True)

SPECIALTIES = [(0, 499, "General"), (500, 999, "Cybersecurity Architecture"),
               (1000, 1499, "Cybersecurity Research and Development"), (1500, 1999, "Leadership"),
               (2000, 2499, "Workforce Development"), (2500, 2999, "Governance, Risk and Compliance"),
               (3000, 3499, "Laws and Data Protection"), (3500, 3999, "Defense"), (4000, 4499, "Protection"),
               (4500, 4999, "Vulnerability Assessment"), (5000, 5499, "Incident Response"),
               (5500, 5999, "Threat Management"), (6000, 6499, "ICS/OT")]  # SCyWF Table 8


def specialty(tid):
    return next(nm for lo, hi, nm in SPECIALTIES if lo <= int(tid[1:]) <= hi)


def header(ws, row, col, labels, widths=None, height=None):
    for i, h in enumerate(labels):
        c = ws.cell(row, col + i, h)
        c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, box
    for i, w in enumerate(widths or []):
        if w: ws.column_dimensions[L(col + i)].width = w
    if height: ws.row_dimensions[row].height = height


def put(ws, r, c, v, font=f_base, fmt=None, align=None, fill=None, border=box):
    cell = ws.cell(r, c, v)
    cell.font, cell.border = font, border
    if fmt: cell.number_format = fmt
    if align: cell.alignment = align
    if fill: cell.fill = fill
    return cell


def merge(ws, r, c1, c2):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)


def bars(ws, ref):
    ws.conditional_formatting.add(ref, DataBarRule(start_type="num", start_value=0, end_type="max", color="3FBF8F"))


def main():
    rd = lambda p: json.loads((ROOT / p).read_text(encoding="utf-8"))
    tasks, roles = rd("data/scywf/tasks.json"), rd("data/scywf/job_roles.json")
    mapping, catalog = rd("data/mapping/vulnerability_task_mapping.json"), rd("data/mapping/source_catalog.json")
    csvs = {n: list(csv.DictReader(open(ROOT / f"data/generated/{n}.csv", encoding="utf-8")))
            for n in ("entities", "incidents", "vulnerabilities", "compliance")}
    entities, programs = csvs["entities"], catalog["programs"]
    role_name = {r["id"]: r["name"] for r in roles}

    titles = [("Incident", v) for v in mapping["vulnerabilities"]] + \
             [("Vulnerability", v) for v in catalog["vulnerability_findings"]] + \
             [("Compliance", v) for v in catalog["ecc_controls"]]
    task_map = [(s, v["title"], t, rel) for s, v in titles for rel, key in (("Primary", "primary"), ("Supporting", "supporting")) for t in v[key]]
    task_ids = sorted({t for *_, t, _ in task_map} | {t for p in programs for t in p["tasks"]})
    assert all(t in tasks for t in task_ids)
    owners = {t: [r["name"] for r in roles if t in r["tasks"]] for t in task_ids}

    wb = Workbook()
    names = ["Overview", "Details", "How It Works", "Weights", "Task Map", "Programs",
             "Incidents", "Vulnerabilities", "Compliance", "Entities", "SCyWF Tasks"]
    ws = {n: (wb.active if i == 0 else wb.create_sheet(n)) for i, n in enumerate(names)}
    ws["Overview"].title = "Overview"
    for n, s in ws.items():
        s.sheet_view.showGridLines = False
        s.sheet_properties.tabColor = {"Weights": "FFC000", "Task Map": "FFC000", "Programs": "FFC000"}.get(n, GREEN if n in names[:3] else None)

    def name(nm, ref):
        wb.defined_names[nm] = DefinedName(nm, attr_text=ref)

    # ------------------------------------------------------------ Weights
    w = ws["Weights"]
    w["B1"], w["B1"].font = "Weights", f_title
    w["B2"], w["B2"].font = "Blue on yellow = input. Change any value and every score recalculates.", f_note
    header(w, 4, 2, ["Group", "Item", "Weight", "Why"], [24, 24, 10, 100])
    rows = [
        (5, "Data source weight", "Incidents", 0.5, "W_INC", "Strongest evidence: the gap was exploited in this entity."),
        (6, "Data source weight", "Vulnerabilities", 0.3, "W_VUL", "The weakness exists in this entity now, found by scanning, but has not been exploited."),
        (7, "Data source weight", "ECC compliance", 0.2, "W_ECC", "A required control is missing or partial. A process gap that makes incidents likely, but less direct evidence."),
        (10, "Severity weight", "Low", 1, None, "Each severity level doubles the weight, so one Critical record counts as much as eight Low ones."),
        (11, "Severity weight", "Medium", 2, None, "Used for incident severity and vulnerability severity."),
        (12, "Severity weight", "High", 4, None, ""),
        (13, "Severity weight", "Critical", 8, None, ""),
        (14, "ECC status weight", "Not compliant", 4, None, "A missing control counts like a High finding."),
        (15, "ECC status weight", "Partially compliant", 2, None, "A partly implemented control counts like a Medium finding."),
        (16, "ECC status weight", "Compliant", 0, None, "No gap."),
        (17, "ECC status weight", "Not applicable", 0, None, "No gap."),
        (20, "Task relevance", "Primary", 1.0, None, "Doing this task well directly prevents the incident or vulnerability."),
        (21, "Task relevance", "Supporting", 0.5, None, "Helps prevent it, but would not stop it on its own."),
    ]
    for r, g, item, val, nm, why in rows:
        put(w, r, 2, g)
        put(w, r, 3, item)
        put(w, r, 4, val, font=f_input, fill=fill_input, fmt="0%" if r <= 7 else "0.0")
        put(w, r, 5, why, font=f_note, border=None)
        if nm: name(nm, f"Weights!$D${r}")
    put(w, 8, 3, "Total", font=f_bold)
    put(w, 8, 4, "=SUM(D5:D7)", font=f_bold, fmt="0%")
    w["E8"], w["E8"].font = "Should be 100%. If a source has no records for the selection, its weight is shared among the others.", f_note
    for nm, ref in (("SevLabels", "$C$10:$C$17"), ("SevW", "$D$10:$D$17"), ("RelLabels", "$C$20:$C$21"), ("RelW", "$D$20:$D$21")):
        name(nm, f"Weights!{ref}")

    # ------------------------------------------------------------ SCyWF Tasks, Entities
    st = ws["SCyWF Tasks"]
    header(st, 1, 1, ["Task ID", "Task description", "Specialty area (ID range, SCyWF Table 8)", "SCyWF job roles that list this task"], [10, 90, 34, 80])
    for tid, desc in tasks.items():
        st.append([tid, desc, specialty(tid), "; ".join(r["name"] for r in roles if tid in r["tasks"])])
    st.freeze_panes, st.auto_filter.ref = "B2", f"A1:D{len(tasks) + 1}"
    name("TaskIDs", f"'SCyWF Tasks'!$A$2:$A${len(tasks) + 1}")
    name("TaskDesc", f"'SCyWF Tasks'!$B$2:$B${len(tasks) + 1}")
    name("TaskSpec", f"'SCyWF Tasks'!$C$2:$C${len(tasks) + 1}")

    e = ws["Entities"]
    NE = len(entities) + 1
    header(e, 1, 1, ["Entity code", "Entity name", "Sector", "Region", "Maturity", "Search match #", "", "Sectors"], [12, 40, 20, 18, 12, 14, 3, 20])
    for i, x in enumerate(entities):
        r = i + 2
        for j, k in enumerate(("entity_code", "entity_name", "sector", "region", "maturity_level")):
            put(e, r, j + 1, x[k])
        put(e, r, 6, f'=IF(ISNUMBER(SEARCH(Overview!$C$5,A{r}&" "&B{r}&" "&C{r})),MAX(F$1:F{r - 1})+1,"")', font=f_note)
    sectors = sorted({x["sector"] for x in entities})
    for i, s in enumerate(sectors):
        put(e, i + 2, 8, s)
    e.freeze_panes = "A2"
    name("EntCodes", f"Entities!$A$2:$A${NE}")
    name("EntNames", f"Entities!$B$2:$B${NE}")
    name("EntSectors", f"Entities!$C$2:$C${NE}")
    name("SectorList", f"Entities!$H$2:$H${len(sectors) + 1}")

    # ------------------------------------------------------------ Data sheets
    # source -> sheet, raw headers, row builder, widths, (entity col, title col, label col)
    DS = {
        "Incident": ("Incidents", ["Incident ID", "Entity code", "Entity name", "Vulnerability title", "Severity", "Reported date"],
                     lambda x: [x["incident_id"], x["entity_code"], x["entity_name"], x["vulnerability_title"], x["severity"], x["reported_date"]],
                     [16, 11, 32, 52, 10, 13], "incidents"),
        "Vulnerability": ("Vulnerabilities", ["Finding ID", "Entity code", "Finding title", "Severity", "Detected date"],
                          lambda x: [x["finding_id"], x["entity_code"], x["finding_title"], x["severity"], x["detected_date"]],
                          [16, 11, 50, 10, 13], "vulnerabilities"),
        "Compliance": ("Compliance", ["Entity code", "ECC control (subdomain)", "Compliance status", "Assessment date"],
                       lambda x: [x["entity_code"], x["ecc_control"], x["compliance_status"], x["assessment_date"]],
                       [11, 62, 20, 15], "compliance"),
    }
    COLS = {}  # source -> (entity, title, weight, in-selection) column letters
    for src, (sh, heads, row, widths, key) in DS.items():
        s = ws[sh]
        n = len(heads)
        ent, title, lab = {"Incident": ("B", "D", "E"), "Vulnerability": ("B", "C", "D"), "Compliance": ("A", "B", "C")}[src]
        sec, wt, sel = L(n + 1), L(n + 2), L(n + 3)
        COLS[src] = (ent, title, wt, sel)
        header(s, 1, 1, heads + ["Sector", "Weight", "In selection"], widths + [18, 8, 10], height=30)
        for c in range(n + 1, n + 4):
            s.cell(1, c).fill = fill_calc
        for i, x in enumerate(csvs[key]):
            r = i + 2
            s.append(row(x) + [f'=IFERROR(INDEX(EntSectors,MATCH({ent}{r},EntCodes,0)),"")',
                               f"=IFERROR(INDEX(SevW,MATCH({lab}{r},SevLabels,0)),0)",
                               f'=IF(OR(Sel="All",{ent}{r}=Sel,{sec}{r}=Sel),1,0)'])
            for c in s[r]:
                c.font = f_base
        s.freeze_panes, s.auto_filter.ref = "A2", f"A1:{sel}{len(csvs[key]) + 1}"

    def rng(src, which):
        col = COLS[src][which]
        return f"'{DS[src][0]}'!${col}$2:${col}${MAXR}" if " " in DS[src][0] else f"{DS[src][0]}!${col}$2:${col}${MAXR}"

    # ------------------------------------------------------------ Overview selector (defines Sel)
    o = ws["Overview"]
    for col, wd in zip("ABCDEFGHIJKLMN", [2, 9, 9, 12, 12, 12, 12, 11, 10, 10, 10, 10, 26, 2]):
        o.column_dimensions[col].width = wd
    o["B2"], o["B2"].font = "SCyWF Training Priority Model", f_title
    o["B3"] = ("Which SCyWF tasks an entity most needs, and which training programs teach them, from its incidents, "
               "vulnerability findings and NCA ECC compliance results. DUMMY DATA for a proof of concept.")
    o["B3"].font = f_note
    o["B5"], o["B5"].font = "Find", f_bold
    o["C5"] = "HLT-003"
    o.merge_cells("C5:F5")
    for c in o["C5:F5"][0]:
        c.font, c.fill, c.border = f_input, fill_input, box
    dv = DataValidation(type="list", formula1="=EntCodes", allow_blank=True, showErrorMessage=False)
    o.add_data_validation(dv)
    dv.add("C5")
    o["G5"], o["G5"].font = "Entity code or name, a sector (e.g. Health), or All for the national view.", f_note
    o["B6"], o["B6"].font = "Showing", f_bold
    o["C6"] = ('=IF(OR(C5="",LOWER(C5)="all"),"All",IFERROR(INDEX(EntCodes,MATCH(C5,EntCodes,0)),'
               'IFERROR(INDEX(SectorList,MATCH(C5,SectorList,0)),IFERROR(INDEX(EntCodes,MATCH("*"&C5&"*",EntNames,0)),""))))')
    o["C6"].font = F(size=11, bold=True, color="008000")
    name("Sel", "Overview!$C$6")
    o["E6"] = '=IF(C6="","Nothing matches. See the search results on the right.","")'
    o["E6"].font = F(size=9, bold=True, color="B3203A")
    for col, wd in zip("OPQR", [12, 20, 20, 18]):
        o.column_dimensions[col].width = wd
    o["O4"], o["O4"].font = "SEARCH RESULTS", f_label
    for k in range(8):
        r, m = 5 + k, f"MATCH({k + 1},Entities!$F$2:$F${NE},0)"
        o.cell(r, 15, f'=IFERROR(INDEX(EntCodes,{m}),"")').font = F(size=9, bold=True, color=GREEN)
        o.cell(r, 16, f'=IFERROR(INDEX(EntNames,{m}),"")').font = F(size=9)
        merge(o, r, 16, 17)
        o.cell(r, 18, f'=IFERROR(INDEX(EntSectors,{m}),"")').font = F(size=9, color=MUTED)
    for row in o["O4:R12"]:
        for c in row:
            c.fill = fill_card

    # ------------------------------------------------------------ Task Map
    tm = ws["Task Map"]
    header(tm, 1, 1, ["Source", "Title", "Task ID", "Relevance", "Relevance weight", "Allocation", "Task description",
                      "Records (selection)", "Title score (selection)", "Task score"], [13, 52, 9, 11, 10, 10, 64, 11, 11, 10], height=30)
    for c in range(5, 11):
        tm.cell(1, c).fill = fill_calc
    pick = lambda r, f: f'IF(A{r}="Incident",{f("Incident")},IF(A{r}="Vulnerability",{f("Vulnerability")},{f("Compliance")}))'
    for i, (src, title, tid, rel) in enumerate(task_map):
        r = i + 2
        for c, v in enumerate((src, title, tid, rel), 1):
            put(tm, r, c, v, font=f_input, fill=fill_input)
        put(tm, r, 5, f"=INDEX(RelW,MATCH(D{r},RelLabels,0))", fmt="0.0")
        put(tm, r, 6, f"=E{r}/SUMIFS($E$2:$E$3000,$B$2:$B$3000,B{r})", fmt="0%")
        put(tm, r, 7, f'=IFERROR(INDEX(TaskDesc,MATCH(C{r},TaskIDs,0)),"Unknown task ID")')
        put(tm, r, 8, "=" + pick(r, lambda s: f'COUNTIFS({rng(s, 1)},$B{r},{rng(s, 3)},1,{rng(s, 2)},">0")'))
        put(tm, r, 9, "=" + pick(r, lambda s: f"SUMIFS({rng(s, 2)},{rng(s, 1)},$B{r},{rng(s, 3)},1)"), fmt="0")
        put(tm, r, 10, f"=F{r}*I{r}", fmt="0.00")
    tm.freeze_panes, tm.auto_filter.ref = "C2", f"A1:J{len(task_map) + 1}"
    TM = lambda col: f"'Task Map'!${col}$2:${col}$3000"

    # ------------------------------------------------------------ Details (all calculations for the selection)
    d = ws["Details"]
    NT, NP, NS = len(task_ids), len(programs), len(SPECIALTIES)
    PRG0 = 13; PRG1 = PRG0 + NP - 1
    SPC0 = PRG1 + 4; SPC1 = SPC0 + NS - 1
    TSK0 = SPC1 + 4; TSK1 = TSK0 + NT - 1
    T = lambda col: f"Details!${col}${TSK0}:${col}${TSK1}"
    for col, wd in zip(range(1, 14), [44, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 11, 8]):
        d.column_dimensions[L(col)].width = wd
    d["A1"], d["A1"].font = "Details", f_title
    d["A2"] = '="Showing: "&Sel&"   (change it on the Overview sheet)"'
    d["A2"].font = F(bold=True, color="008000")

    header(d, 4, 1, ["Data source", "Weight set", "Records with a gap", "Total score", "Weight used", "Top task from this source"], height=30)
    merge(d, 4, 6, 11)
    for i, (src, lab, wn) in enumerate((("Incident", "Incidents", "W_INC"), ("Vulnerability", "Vulnerabilities", "W_VUL"),
                                         ("Compliance", "ECC compliance", "W_ECC"))):
        r, sc = 5 + i, "FGH"[i]
        put(d, r, 1, lab, font=f_bold)
        put(d, r, 2, f"={wn}", fmt="0%")
        put(d, r, 3, f'=COUNTIFS({rng(src, 3)},1,{rng(src, 2)},">0")', fmt="#,##0")
        put(d, r, 4, f'=SUMIFS({TM("J")},{TM("A")},"{src}")', fmt="#,##0")
        put(d, r, 5, f"=IFERROR(IF(D{r}>0,B{r},0)/SUMPRODUCT(($D$5:$D$7>0)*$B$5:$B$7),0)", font=f_bold, fmt="0%")
        put(d, r, 6, f'=IF(D{r}=0,"No records",INDEX({T("A")},MATCH(MAX({T(sc)}),{T(sc)},0))&"  "&INDEX({T("B")},MATCH(MAX({T(sc)}),{T(sc)},0)))')
        merge(d, r, 6, 11)
        name({"Incident": "NW_INC", "Vulnerability": "NW_VUL", "Compliance": "NW_ECC"}[src], f"Details!$E${r}")
        name({"Incident": "TOT_INC", "Vulnerability": "TOT_VUL", "Compliance": "TOT_ECC"}[src], f"Details!$D${r}")
    d["A8"], d["A8"].font = "Weight used = weight set, shared out again if a source has no records for the selection.", f_note

    d.cell(PRG0 - 2, 1, "Training programs").font = f_h2
    header(d, PRG0 - 1, 1, ["Training program", "Gap covered", "From incidents", "From vulns", "From ECC", "Tasks taught",
                            "Top-10 tasks taught", "Rank"], height=30)
    P = lambda col: f"Programs!${col}$2:${col}$2000"
    for i, p in enumerate(programs):
        r = PRG0 + i
        put(d, r, 1, p["name"], font=f_bold)
        for c, pc in ((3, "E"), (4, "F"), (5, "G")):
            put(d, r, c, f"=SUMIFS({P(pc)},{P('A')},$A{r})", fmt="0.0%")
        put(d, r, 2, f"=C{r}+D{r}+E{r}", font=f_bold, fmt="0.0%")
        put(d, r, 6, f"=COUNTIF({P('A')},$A{r})")
        put(d, r, 7, f'=COUNTIFS({P("A")},$A{r},{P("H")},"<=10")')
        put(d, r, 8, f"=RANK(B{r},$B${PRG0}:$B${PRG1})+COUNTIF($B${PRG0 - 1}:B{r - 1},B{r})", font=f_bold)
    bars(d, f"B{PRG0}:B{PRG1}")
    d.cell(PRG1 + 1, 1, "Gap covered = sum of the gap index of the tasks the program teaches (Programs sheet).").font = f_note

    d.cell(SPC0 - 2, 1, "SCyWF specialty areas").font = f_h2
    header(d, SPC0 - 1, 1, ["Specialty area", "Gap index", "Rank"], height=30)
    for i, (lo, hi, nm) in enumerate(SPECIALTIES):
        r = SPC0 + i
        put(d, r, 1, nm, font=f_bold)
        put(d, r, 2, f"=SUMIFS({T('L')},{T('C')},A{r})", fmt="0.0%")
        put(d, r, 3, f"=RANK(B{r},$B${SPC0}:$B${SPC1})+COUNTIF($B${SPC0 - 1}:B{r - 1},B{r})", font=f_bold)
    bars(d, f"B{SPC0}:B{SPC1}")

    d.cell(TSK0 - 2, 1, "SCyWF tasks").font = f_h2
    header(d, TSK0 - 1, 1, ["Task ID", "Task description", "Specialty area", "SCyWF job roles", "Records", "Incident score",
                            "Vuln score", "ECC score", "From incidents", "From vulns", "From ECC", "Gap index", "Rank"], height=30)
    for i, t in enumerate(task_ids):
        r = TSK0 + i
        put(d, r, 1, t, font=f_id, align=top_wrap)
        put(d, r, 2, tasks[t], align=top_wrap)
        put(d, r, 3, specialty(t), font=f_small, align=top_wrap)
        put(d, r, 4, "; ".join(owners[t]), font=f_small, align=top_wrap)
        put(d, r, 5, f"=SUMIFS({TM('H')},{TM('C')},$A{r})", align=top_wrap)
        for c, src in ((6, "Incident"), (7, "Vulnerability"), (8, "Compliance")):
            put(d, r, c, f'=SUMIFS({TM("J")},{TM("C")},$A{r},{TM("A")},"{src}")', fmt="0.0", align=top_wrap)
        for c, sc, nw, tot in ((9, "F", "NW_INC", "TOT_INC"), (10, "G", "NW_VUL", "TOT_VUL"), (11, "H", "NW_ECC", "TOT_ECC")):
            put(d, r, c, f"=IF({tot}=0,0,{nw}*{sc}{r}/{tot})", fmt="0.00%", align=top_wrap)
        put(d, r, 12, f"=I{r}+J{r}+K{r}", font=f_bold, fmt="0.00%", align=top_wrap)
        put(d, r, 13, f"=RANK(L{r},$L${TSK0}:$L${TSK1})+COUNTIF($L${TSK0 - 1}:L{r - 1},L{r})", font=f_bold, align=top_wrap)
    put(d, TSK1 + 1, 1, "Total", font=f_bold)
    for c in range(9, 13):
        put(d, TSK1 + 1, c, f"=SUM({L(c)}{TSK0}:{L(c)}{TSK1})", font=f_bold, fmt="0.0%")
    bars(d, f"L{TSK0}:L{TSK1}")
    d.cell(TSK1 + 2, 1, "Records = records pointing to the task. Score = severity-weighted records x allocation. "
                        "From <source> = weight used x task score / source total. Gap index sums to 100%.").font = f_note
    d.freeze_panes = "A4"

    # ------------------------------------------------------------ Programs (long: one row per program-task)
    pg = ws["Programs"]
    header(pg, 1, 1, ["Training program", "Task ID", "Task description", "Specialty area", "From incidents", "From vulns",
                      "From ECC", "Task rank"], [44, 9, 70, 28, 11, 11, 11, 9], height=30)
    for c in range(3, 9):
        pg.cell(1, c).fill = fill_calc
    r = 2
    for p in programs:
        for t in p["tasks"]:
            put(pg, r, 1, p["name"], font=f_input, fill=fill_input)
            put(pg, r, 2, t, font=f_input, fill=fill_input)
            put(pg, r, 3, f'=IFERROR(INDEX(TaskDesc,MATCH(B{r},TaskIDs,0)),"Unknown task ID")')
            put(pg, r, 4, f'=IFERROR(INDEX(TaskSpec,MATCH(B{r},TaskIDs,0)),"")')
            for c, sc in ((5, "I"), (6, "J"), (7, "K")):
                put(pg, r, c, f"=IFERROR(INDEX({T(sc)},MATCH($B{r},{T('A')},0)),0)", fmt="0.0%")
            put(pg, r, 8, f'=IFERROR(INDEX({T("M")},MATCH($B{r},{T("A")},0)),"")')
            r += 1
    pg.freeze_panes, pg.auto_filter.ref = "C2", f"A1:H{r - 1}"
    pg.cell(r + 1, 1, "To change what a program teaches, add or delete rows (yellow columns). Use the exact program name.").font = f_note

    # ------------------------------------------------------------ Overview body
    o.merge_cells("B8:M8")
    o["B8"] = ('=IF(Sel="All","All entities (national view)",IFERROR(INDEX(EntNames,MATCH(Sel,EntCodes,0)),'
               'IF(Sel="","",Sel&" sector")))')
    o["B8"].font = F(size=16, bold=True)
    o.merge_cells("B9:M9")
    o["B9"] = (f'=IFERROR(Sel&"   |   "&INDEX(EntSectors,MATCH(Sel,EntCodes,0))&"   |   "&INDEX(Entities!$D$2:$D${NE},MATCH(Sel,EntCodes,0))'
               f'&"   |   Maturity: "&INDEX(Entities!$E$2:$E${NE},MATCH(Sel,EntCodes,0)),'
               f'IF(Sel="","",COUNTIF(EntSectors,IF(Sel="All","*",Sel))&" entities"))')
    o["B9"].font = F(color=MUTED)
    for row in o["B7:M10"]:
        for c in row:
            c.fill = fill_card
    kpis = [("INCIDENTS", "=Details!C5", False),
            ("CRITICAL INCIDENTS", f'=COUNTIFS({rng("Incident", 3)},1,Incidents!$E$2:$E${MAXR},"Critical")', True),
            ("VULNERABILITY FINDINGS", "=Details!C6", False),
            ("CRITICAL FINDINGS", f'=COUNTIFS({rng("Vulnerability", 3)},1,Vulnerabilities!$D$2:$D${MAXR},"Critical")', True),
            ("ECC CONTROLS NOT COMPLIANT", f'=COUNTIFS({rng("Compliance", 3)},1,Compliance!$C$2:$C${MAXR},"Not compliant")', True),
            ("ECC PARTIALLY COMPLIANT", f'=COUNTIFS({rng("Compliance", 3)},1,Compliance!$C$2:$C${MAXR},"Partially compliant")', False)]
    for i, (lab, f, red) in enumerate(kpis):
        c0 = 2 + i * 2 if i < 5 else 12
        c1 = c0 + 1 if i < 5 else 13
        merge(o, 11, c0, c1)
        merge(o, 12, c0, c1)
        o.cell(11, c0, lab).font = f_label
        o.cell(11, c0).alignment = Alignment(vertical="bottom", wrap_text=True)
        o.cell(12, c0, f).font = f_kpi_red if red else f_kpi
        o.cell(12, c0).alignment = Alignment(horizontal="left")
        o.cell(12, c0).number_format = "#,##0"
    o.row_dimensions[11].height = 24
    o.row_dimensions[12].height = 30

    o["B14"], o["B14"].font = "Data sources", f_h2
    header(o, 15, 2, ["Source", "", "", "Weight set", "Records with a gap", "Weight used", "Top task from this source", "", "", "", "", ""])
    merge(o, 15, 2, 4)
    merge(o, 15, 8, 13)
    for i in range(3):
        r, dr = 16 + i, 5 + i
        put(o, r, 2, f"=Details!A{dr}", font=f_bold)
        for c in (3, 4):
            put(o, r, c, None)
        merge(o, r, 2, 4)
        put(o, r, 5, f"=Details!B{dr}", fmt="0%")
        put(o, r, 6, f"=Details!C{dr}", fmt="#,##0")
        put(o, r, 7, f"=Details!E{dr}", font=f_bold, fmt="0%")
        put(o, r, 8, f"=Details!F{dr}", font=f_small)
        for c in range(9, 14):
            put(o, r, c, None)
        merge(o, r, 8, 13)

    o["B20"], o["B20"].font = "Top 10 SCyWF tasks needed", f_h2
    o["B21"], o["B21"].font = "Gap index = share of all the evidence (by severity and frequency) that points to the task. All tasks add up to 100%.", f_note
    header(o, 22, 2, ["Rank", "Task", "Task description", "", "", "", "Gap index", "Records", "From incidents", "From vulns", "From ECC", "Specialty area"], height=30)
    merge(o, 22, 4, 7)
    for k in range(10):
        r, m = 23 + k, f"MATCH({k + 1},{T('M')},0)"
        put(o, r, 2, k + 1, align=Alignment(horizontal="center", vertical="top"))
        put(o, r, 3, f"=INDEX({T('A')},{m})", font=f_id, align=top_wrap)
        put(o, r, 4, f"=INDEX({T('B')},{m})", align=top_wrap)
        for c in (5, 6, 7):
            put(o, r, c, None)
        merge(o, r, 4, 7)
        for c, sc in ((8, "L"), (9, "E"), (10, "I"), (11, "J"), (12, "K")):
            put(o, r, c, f"=INDEX({T(sc)},{m})", font=f_bold if c == 8 else f_base, fmt="0" if c == 9 else "0.0%", align=top_wrap)
        put(o, r, 13, f"=INDEX({T('C')},{m})", font=f_small, align=top_wrap)
        o.row_dimensions[r].height = 34
    bars(o, "H23:H32")

    o["B34"], o["B34"].font = "Recommended training programs", f_h2
    o["B35"], o["B35"].font = "Ranked by gap covered = sum of the gap index of the SCyWF tasks each program teaches.", f_note
    header(o, 36, 2, ["Rank", "Training program", "", "", "", "", "Gap covered", "Top-10 tasks taught", "Tasks taught", "", "", ""], height=30)
    merge(o, 36, 3, 7)
    merge(o, 36, 10, 13)
    for k in range(5):
        r, m = 37 + k, f"MATCH({k + 1},Details!$H${PRG0}:$H${PRG1},0)"
        put(o, r, 2, k + 1, align=Alignment(horizontal="center"))
        put(o, r, 3, f"=INDEX(Details!$A${PRG0}:$A${PRG1},{m})", font=f_bold)
        for c in (4, 5, 6, 7):
            put(o, r, c, None)
        merge(o, r, 3, 7)
        put(o, r, 8, f"=INDEX(Details!$B${PRG0}:$B${PRG1},{m})", font=f_bold, fmt="0.0%")
        put(o, r, 9, f'=INDEX(Details!$G${PRG0}:$G${PRG1},{m})&" of 10"')
        put(o, r, 10, f"=INDEX(Details!$F${PRG0}:$F${PRG1},{m})")
        for c in (11, 12, 13):
            put(o, r, c, None)
        merge(o, r, 10, 13)
    bars(o, "H37:H41")

    o["B43"], o["B43"].font = "SCyWF specialty areas (top 6)", f_h2
    header(o, 44, 2, ["Rank", "Specialty area", "", "", "", "", "Gap index"])
    merge(o, 44, 3, 7)
    for k in range(6):
        r, m = 45 + k, f"MATCH({k + 1},Details!$C${SPC0}:$C${SPC1},0)"
        put(o, r, 2, k + 1, align=Alignment(horizontal="center"))
        put(o, r, 3, f"=INDEX(Details!$A${SPC0}:$A${SPC1},{m})", font=f_bold)
        for c in (4, 5, 6, 7):
            put(o, r, c, None)
        merge(o, r, 3, 7)
        put(o, r, 8, f"=INDEX(Details!$B${SPC0}:$B${SPC1},{m})", font=f_bold, fmt="0.0%")
    bars(o, "H45:H50")

    for title, crow, anchor in (("Gap covered, top 5 programs", 37, "O20"), ("Gap by specialty area", 45, "O36")):
        ch = BarChart()
        ch.type, ch.style, ch.title, ch.legend = "bar", 10, title, None
        n = 5 if crow == 37 else 6
        ch.add_data(Reference(o, min_col=8, min_row=crow, max_row=crow + n - 1), titles_from_data=False)
        ch.set_categories(Reference(o, min_col=3, min_row=crow, max_row=crow + n - 1))
        ch.y_axis.numFmt, ch.y_axis.majorGridlines = "0%", None
        ch.x_axis.scaling.orientation = "maxMin"
        ch.x_axis.delete = ch.y_axis.delete = ch.roundedCorners = ch.title.overlay = False
        ch.series[0].graphicalProperties.solidFill = ch.series[0].graphicalProperties.line.solidFill = GREEN
        ch.height, ch.width = 7.5, 16
        o.add_chart(ch, anchor)
    o.freeze_panes = "A7"

    # ------------------------------------------------------------ How It Works
    h = ws["How It Works"]
    h.column_dimensions["B"].width = 130
    lines = [
        ("How the model works", f_title), ("", None),
        ("Three data sources, four fields each: entity code, title, severity (or ECC compliance status), date.", f_bold),
        ("Incidents and vulnerability findings come with a severity. ECC compliance results come with a status per ECC subdomain; "
         "Not compliant and Partially compliant mean the entity is not doing the tasks that control requires.", f_base),
        ("Filter the data to the period you want before pasting it in. The model uses every row on the data sheets.", f_base), ("", None),
        ("Step 1. Weigh each record by severity", f_h2),
        ("Low 1, Medium 2, High 4, Critical 8. ECC: Not compliant 4, Partially compliant 2, Compliant 0.", f_base),
        ("Frequency counts automatically: every record adds its weight, so 5 Medium incidents (5 x 2 = 10) outweigh 1 Critical one (8).", f_base), ("", None),
        ("Step 2. Pass each record's weight to the tasks that would have prevented it", f_h2),
        ("The Task Map lists, for each title, the SCyWF tasks that prevent it: Primary (1.0) prevents it directly, Supporting (0.5) helps prevent it.", f_base),
        ("The weight is split over the tasks: allocation = task relevance / sum of relevances for that title. Each record counts once in total.", f_base),
        ("Example: 'Phishing email led to credential compromise' has 4 Primary and 3 Supporting tasks: 4 x 1.0 + 3 x 0.5 = 5.5.", f_base),
        ("      Each Primary task gets 1.0 / 5.5 = 18%, each Supporting task 0.5 / 5.5 = 9%. A High incident (weight 4) gives each Primary task 0.73.", f_base), ("", None),
        ("Step 3. Gap index", f_h2),
        ("For each source: task share = task score / total score of that source. Then: gap index = 50% x incident share + 30% x vulnerability share + 20% x ECC share.", f_bold),
        ("Example: a task with 20% of the incident evidence, 10% of the vulnerability evidence and 0% of the ECC evidence: 50% x 20% + 30% x 10% + 20% x 0% = 13%.", f_base),
        ("Read it as: 13% of the selection's gap points to this task. All tasks add up to 100%. Source weights are on the Weights sheet.", f_base), ("", None),
        ("Step 4. Rank programs", f_h2),
        ("Gap covered = sum of the gap index of the tasks a program teaches (Programs sheet). Specialty areas are ranked the same way, using SCyWF task ID ranges.", f_base), ("", None),
        ("Sheets", f_h2),
        ("Overview: pick an entity, a sector or All. Details: every task, program and specialty area ranked. Weights, Task Map, Programs: inputs (yellow tabs).", f_base),
        ("Incidents, Vulnerabilities, Compliance, Entities: data. SCyWF Tasks: framework reference.", f_base), ("", None),
        ("Adding data: paste rows under the existing ones on a data sheet and copy the grey columns down. Every title must be on the Task Map.", f_base),
        ("Source: SCyWF 1.5 : 2026 (NCA), tasks and job roles extracted by OCR. ECC titles are the 28 subdomains of NCA ECC-2:2024. "
         "Task mapping, weights and programs are proposals for review. Entities and records are fictional.", f_note),
    ]
    for i, (txt, fnt) in enumerate(lines):
        c = h.cell(i + 1, 2, txt)
        c.font, c.alignment = fnt or f_base, Alignment(wrap_text=True, vertical="top")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"Saved {OUT}: {len(task_map)} task-map rows, {NT} tasks, {len(wb.sheetnames)} sheets")


if __name__ == "__main__":
    main()
