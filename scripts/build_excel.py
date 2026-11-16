#!/usr/bin/env python3
"""Build the SCyWF Training Priority Model workbook (excel/SCyWF_Training_Priority_Model.xlsx).

Task-based model. Every record maps directly to SCyWF tasks; there are no intermediate categories.
Inputs per record are only: entity, title, severity, date.

  record score      = severity weight x recency weight (half-life decay per source)
  task score        = sum of record scores x allocation (each record's score is split over its mapped
                      tasks by relevance: primary 1.0, supporting 0.5, so every record counts once)
  source share      = task score / entity's total score from that source
  effective weight  = source weight x data sufficiency (records / minimum records, capped at 1)
  task gap index    = sum over sources of effective weight x source share / sum of effective weights
  program priority  = sum of the gap index of the tasks the program teaches (share of the entity's gap covered)
  role / specialty  = sum of the gap index of the tasks that belong to the role / specialty area

Everything is a live formula: change a weight, edit the task map or a program's task list, add records,
or type another entity on the Overview, and all scores and rankings recalculate.
"""
import csv
import json
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
MAXR = 10000  # data rows covered by formula ranges on the data sheets

FONT = "Arial"
GREEN, DARK, MUTED, LINE = "0B6E4F", "14201B", "56655E", "D5DDD9"
f_base = Font(name=FONT, size=10, color=DARK)
f_small = Font(name=FONT, size=8, color=DARK)
f_bold = Font(name=FONT, size=10, bold=True, color=DARK)
f_head = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_title = Font(name=FONT, size=18, bold=True, color=DARK)
f_h2 = Font(name=FONT, size=12, bold=True, color=GREEN)
f_note = Font(name=FONT, size=9, italic=True, color=MUTED)
f_input = Font(name=FONT, size=11, bold=True, color="0000FF")
f_link = Font(name=FONT, size=10, bold=True, color="008000")
f_kpi = Font(name=FONT, size=20, bold=True, color=DARK)
f_kpi_red = Font(name=FONT, size=20, bold=True, color="B3203A")
f_label = Font(name=FONT, size=8, bold=True, color=MUTED)
f_entity = Font(name=FONT, size=16, bold=True, color=DARK)
f_id = Font(name=FONT, size=10, bold=True, color=GREEN)
fill_head = PatternFill("solid", fgColor=GREEN)
fill_calc = PatternFill("solid", fgColor="56655E")
fill_input = PatternFill("solid", fgColor="FFF59D")
fill_card = PatternFill("solid", fgColor="EEF4F1")
fill_sub = PatternFill("solid", fgColor="D8EDE4")
thin = Side(style="thin", color=LINE)
box = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
top_wrap = Alignment(vertical="top", wrap_text=True)
BAR = "3FBF8F"

SPECIALTIES = [(0, 499, "General"), (500, 999, "Cybersecurity Architecture"),
               (1000, 1499, "Cybersecurity Research and Development"), (1500, 1999, "Leadership"),
               (2000, 2499, "Workforce Development"), (2500, 2999, "Governance, Risk and Compliance"),
               (3000, 3499, "Laws and Data Protection"), (3500, 3999, "Defense"), (4000, 4499, "Protection"),
               (4500, 4999, "Vulnerability Assessment"), (5000, 5499, "Incident Response"),
               (5500, 5999, "Threat Management"), (6000, 6499, "ICS/OT")]


def specialty(tid):
    n = int(tid[1:])
    return next(nm for lo, hi, nm in SPECIALTIES if lo <= n <= hi)


def load_csv(p):
    with open(p, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def header(ws, row, col, labels, widths=None, height=None):
    for i, h in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=h)
        c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, box
    if widths:
        for i, w in enumerate(widths):
            if w: ws.column_dimensions[L(col + i)].width = w
    if height:
        ws.row_dimensions[row].height = height


def put(ws, r, c, v, font=f_base, fmt=None, align=None, border=box, fill=None):
    cell = ws.cell(r, c, v)
    cell.font = font
    if fmt: cell.number_format = fmt
    if align: cell.alignment = align
    if border: cell.border = border
    if fill: cell.fill = fill
    return cell


def name(wb, nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


def q(sheet):
    return f"'{sheet}'" if " " in sheet else sheet


def main():
    tasks = json.loads((ROOT / "data/scywf/tasks.json").read_text(encoding="utf-8"))
    roles = json.loads((ROOT / "data/scywf/job_roles.json").read_text(encoding="utf-8"))
    mapping = json.loads((ROOT / "data/mapping/vulnerability_task_mapping.json").read_text(encoding="utf-8"))
    catalog = json.loads((ROOT / "data/mapping/source_catalog.json").read_text(encoding="utf-8"))
    entities = load_csv(ROOT / "data/generated/entities.csv")
    incidents = load_csv(ROOT / "data/generated/incidents.csv")
    vulns = load_csv(ROOT / "data/generated/vulnerabilities.csv")
    threats = load_csv(ROOT / "data/generated/threats.csv")
    programs = catalog["programs"]
    role_name = {r["id"]: r["name"] for r in roles}

    # Catalog of record types and the long task map
    cat_rows = [("Incident", v["title"], v["primary"], v["supporting"]) for v in mapping["vulnerabilities"]]
    cat_rows += [("Vulnerability", v["title"], v["primary"], v["supporting"]) for v in catalog["vulnerability_findings"]]
    cat_rows += [("Threat", v["title"], v["primary"], v["supporting"]) for v in catalog["threats"]]
    assert len({r[1] for r in cat_rows}) == len(cat_rows), "catalog titles must be unique"
    task_map = [(s, t, tid, "Primary") for s, t, p, _ in cat_rows for tid in p]
    task_map += [(s, t, tid, "Supporting") for s, t, _, su in cat_rows for tid in su]
    task_map.sort(key=lambda x: [r[1] for r in cat_rows].index(x[1]))
    task_ids = sorted({x[2] for x in task_map} | {t for p in programs for t in p["tasks"]})
    for t in task_ids:
        assert t in tasks, t
    owners = {t: [r["id"] for r in roles if t in r["tasks"]] for t in task_ids}

    wb = Workbook()
    sheets = ["Overview", "Entity Detail", "National View", "How It Works", "Weights", "Programs", "Program Tasks",
              "Catalog", "Task Map", "Incidents", "Vulnerabilities", "Threats", "Entities", "Title Matrix",
              "Role Tasks", "SCyWF Tasks", "SCyWF Roles"]
    ws = {}
    ws["Overview"] = wb.active
    ws["Overview"].title = "Overview"
    for s in sheets[1:]:
        ws[s] = wb.create_sheet(s)
    for s in ws.values():
        s.sheet_view.showGridLines = False
    for s in sheets[:4]:
        ws[s].sheet_properties.tabColor = GREEN
    for s in ("Weights", "Programs", "Program Tasks", "Task Map"):
        ws[s].sheet_properties.tabColor = "FFC000"

    # ---------------------------------------------------------------- Weights
    w = ws["Weights"]
    w["B1"], w["B1"].font = "Weights", f_title
    w["B2"] = "Blue on yellow = input. Only fields every source has are used: severity and date (plus how many records exist)."
    w["B2"].font = f_note
    header(w, 4, 2, ["Group", "Item", "Input", "Result", "Why"], [22, 30, 10, 10, 92])
    W_ROWS = [
        (5, "Source evidence strength", "Incidents", 3, "=D5/SUM($D$5:$D$7)", "W_INC",
         "Strongest evidence: the gap was exploited in this entity. Strength 3 of 3."),
        (6, "Source evidence strength", "Vulnerabilities", 2, "=D6/SUM($D$5:$D$7)", "W_VUL",
         "The gap exists in this entity now but has not (yet) caused harm. Strength 2 of 3."),
        (7, "Source evidence strength", "Threats", 1, "=D7/SUM($D$5:$D$7)", "W_THR",
         "Attacker activity is real but external and often sector-wide, so it says least about this entity. Strength 1 of 3."),
        (8, "Data sufficiency", "Minimum records for full weight", 10, None, "MinRec",
         "A source with fewer records than this gets proportionally less weight (5 records = half weight), so one or two records cannot dominate."),
        (11, "Severity weight", "Low", 1, None, None, "Each step up doubles the weight. Severity scales are ordinal and impact grows faster than linearly,"),
        (12, "Severity weight", "Medium", 2, None, None, "so one Critical record counts as much as eight Low ones. The same scale is used for incident severity,"),
        (13, "Severity weight", "High", 4, None, None, "vulnerability severity and threat level, so all three sources are measured alike."),
        (14, "Severity weight", "Critical", 8, None, None, ""),
        (17, "Recency", "As-of date", date(2026, 10, 31), None, "AsOfDate", "Ages are measured to this date. Set it to the date of your data extract."),
        (18, "Recency half-life (days)", "Incidents", 365, None, "HL_INC", "An incident a year old counts half as much as one today: the gap may since have been addressed."),
        (19, "Recency half-life (days)", "Vulnerabilities", 0, None, "HL_VUL", "0 = no decay. A finding in the current scan is current exposure, whatever its detection date."),
        (20, "Recency half-life (days)", "Threats", 180, None, "HL_THR", "Threat intelligence goes stale faster: a six-month-old campaign counts half."),
        (23, "Task relevance", "Primary", 1.0, None, None, "The task most needed to prevent or handle the record."),
        (24, "Task relevance", "Supporting", 0.5, None, None, "Tasks that close the rest of the gap. A record's score is split over its tasks in proportion to relevance."),
    ]
    for r, g, item, val, res, nm, why in W_ROWS:
        put(w, r, 2, g)
        put(w, r, 3, item)
        put(w, r, 4, val, font=f_input, fill=fill_input,
            fmt="yyyy-mm-dd" if isinstance(val, date) else ("0" if isinstance(val, int) else "0.0"))
        if res:
            put(w, r, 5, res, font=f_bold, fmt="0%")
        put(w, r, 6, why, font=f_note, border=None, align=Alignment(wrap_text=False))
        if nm:
            name(wb, nm, f"Weights!$E${r}" if res else f"Weights!$D${r}")
    for nm, ref in [("SevLabels", "$C$11:$C$14"), ("SevW", "$D$11:$D$14"), ("RelLabels", "$C$23:$C$24"), ("RelW", "$D$23:$D$24")]:
        name(wb, nm, f"Weights!{ref}")
    w["B26"] = "Record score = severity weight x 0.5 ^ (age in days / half-life). The Result column shows each source's base weight."
    w["B26"].font = f_note
    w.freeze_panes = "A5"

    # ---------------------------------------------------------------- SCyWF reference
    st = ws["SCyWF Tasks"]
    header(st, 1, 1, ["Task ID", "Task description", "Specialty area (by ID range)", "Job roles that list this task"], [10, 90, 34, 70])
    for tid, desc in tasks.items():
        st.append([tid, desc, specialty(tid), "; ".join(r["id"] for r in roles if tid in r["tasks"])])
    sr = ws["SCyWF Roles"]
    header(sr, 1, 1, ["Job role ID", "Job role", "Category", "Specialty area", "Number of tasks", "Task IDs"], [18, 44, 44, 34, 10, 90])
    for r in roles:
        sr.append([r["id"], r["name"], r["category"], r["specialty_area"], len(r["tasks"]), ", ".join(r["tasks"])])
    for s in (st, sr):
        s.freeze_panes = "B2"
        s.auto_filter.ref = f"A1:{L(s.max_column)}{s.max_row}"
        for row in s.iter_rows(min_row=2):
            for c in row:
                c.font, c.alignment = f_base, top_wrap
    NST = len(tasks) + 1
    name(wb, "TaskIDs", f"'SCyWF Tasks'!$A$2:$A${NST}")
    name(wb, "TaskDesc", f"'SCyWF Tasks'!$B$2:$B${NST}")
    name(wb, "TaskSpec", f"'SCyWF Tasks'!$C$2:$C${NST}")

    # ---------------------------------------------------------------- Entities
    e = ws["Entities"]
    header(e, 1, 1, ["Entity code", "Entity name", "Sector", "Region", "Maturity", "Search match #"], [12, 40, 20, 18, 12, 14])
    NE = len(entities) + 1
    for i, en in enumerate(entities):
        r = i + 2
        for j, v in enumerate([en["entity_code"], en["entity_name"], en["sector"], en["region"], en["maturity_level"]]):
            put(e, r, j + 1, v)
        put(e, r, 6, f'=IF(ISNUMBER(SEARCH(Overview!$C$5,A{r}&" "&B{r}&" "&C{r})),MAX(F$1:F{r - 1})+1,"")', font=f_note)
    e.freeze_panes = "A2"
    name(wb, "EntCodes", f"Entities!$A$2:$A${NE}")
    name(wb, "EntNames", f"Entities!$B$2:$B${NE}")

    # ---------------------------------------------------------------- Data sheets
    def rng(sheet, col):
        return f"{q(sheet)}!${col}$2:${col}${MAXR}"

    def recency(hl, dcell):
        return f"IF({hl}<=0,1,0.5^(MAX(0,AsOfDate-{dcell})/{hl}))"

    DS = {  # source -> sheet, entity col, title col, score col, sort-key col, severity col, date col
        "Incident": ("Incidents", "B", "D", "I", "J", "E", "F"),
        "Vulnerability": ("Vulnerabilities", "B", "C", "H", "I", "D", "E"),
        "Threat": ("Threats", "B", "C", "H", "I", "D", "E"),
    }
    specs = [
        ("Incident", incidents, ["Incident ID", "Entity code", "Entity name", "Vulnerability title", "Severity", "Reported date"],
         lambda x: [x["incident_id"], x["entity_code"], x["entity_name"], x["vulnerability_title"], x["severity"], date.fromisoformat(x["reported_date"])],
         [16, 11, 32, 52, 10, 13], "HL_INC"),
        ("Vulnerability", vulns, ["Finding ID", "Entity code", "Finding title", "Severity", "Detected date"],
         lambda x: [x["finding_id"], x["entity_code"], x["finding_title"], x["severity"], date.fromisoformat(x["detected_date"])],
         [16, 11, 50, 10, 13], "HL_VUL"),
        ("Threat", threats, ["Threat ID", "Entity code", "Threat title", "Threat level", "Reported date"],
         lambda x: [x["threat_id"], x["entity_code"], x["threat_title"], x["threat_level"], date.fromisoformat(x["reported_date"])],
         [16, 11, 54, 11, 13], "HL_THR"),
    ]
    for src, rows, heads, conv, widths, hl in specs:
        sh, ec, tc, sc, kc, vc, dc = DS[src]
        s = ws[sh]
        nraw = len(heads)
        header(s, 1, 1, heads + ["Severity weight", "Recency weight", "Record score", "Selected-entity sort key"],
               widths + [10, 10, 10, 13], height=32)
        for col in range(nraw + 1, nraw + 5):
            s.cell(1, col).fill = fill_calc
        for i, x in enumerate(rows):
            r = i + 2
            s.append(conv(x) + [f"=INDEX(SevW,MATCH({vc}{r},SevLabels,0))", "=" + recency(hl, f"{dc}{r}"),
                                f"={L(nraw + 1)}{r}*{L(nraw + 2)}{r}", f"=IF({ec}{r}=SelCode,{sc}{r}+ROW()/10000000,0)"])
            s.cell(r, nraw).number_format = "yyyy-mm-dd"
            s.cell(r, nraw + 2).number_format = "0.000"
            s.cell(r, nraw + 3).number_format = "0.000"
            for c in s[r]:
                c.font = f_base
        s.freeze_panes = "A2"
        s.auto_filter.ref = f"A1:{L(nraw + 4)}{len(rows) + 1}"

    def sumifs(src, ent, title=None):
        sh, ec, tc, sc = DS[src][:4]
        f = f"SUMIFS({rng(sh, sc)},{rng(sh, ec)},{ent}"
        return f + (f",{rng(sh, tc)},{title})" if title else ")")

    # ---------------------------------------------------------------- Catalog (one row per record type)
    k = ws["Catalog"]
    header(k, 1, 1, ["Source", "Title", "Primary SCyWF tasks", "Supporting SCyWF tasks", "Tasks mapped",
                     "Selected entity: records", "Selected entity: title score"], [13, 52, 56, 56, 9, 11, 12], height=32)
    for col in (5, 6, 7):
        k.cell(1, col).fill = fill_calc
    NCAT = len(cat_rows) + 1
    for i, (src, title, prim, sup) in enumerate(cat_rows):
        r = i + 2
        sh, ec, tc = DS[src][:3]
        put(k, r, 1, src, align=top_wrap)
        put(k, r, 2, title, font=f_bold, align=top_wrap)
        put(k, r, 3, "\n".join(f"{t} {tasks[t]}" for t in prim), font=f_small, align=top_wrap)
        put(k, r, 4, "\n".join(f"{t} {tasks[t]}" for t in sup), font=f_small, align=top_wrap)
        put(k, r, 5, f"=COUNTIF('Task Map'!$B$2:$B$2000,B{r})", align=top_wrap)
        put(k, r, 6, f"=COUNTIFS({rng(sh, ec)},SelCode,{rng(sh, tc)},B{r})", align=top_wrap)
        put(k, r, 7, f'=IF(A{r}="Incident",{sumifs("Incident", "SelCode", f"B{r}")},IF(A{r}="Vulnerability",'
                     f'{sumifs("Vulnerability", "SelCode", f"B{r}")},{sumifs("Threat", "SelCode", f"B{r}")}))',
            fmt="0.000", align=top_wrap)
    k.freeze_panes = "C2"
    k.auto_filter.ref = f"A1:G{NCAT}"
    name(wb, "CatSrc", f"Catalog!$A$2:$A${NCAT}")
    name(wb, "CatTitle", f"Catalog!$B$2:$B${NCAT}")
    name(wb, "CatScore", f"Catalog!$G$2:$G${NCAT}")
    k.cell(NCAT + 2, 2, "To map a new title: add it here (source + title), then add one row per task on the Task Map sheet.").font = f_note

    # ---------------------------------------------------------------- Programs + Program Tasks
    p = ws["Programs"]
    header(p, 1, 1, ["ID", "Training program", "Level", "Format / duration", "Audience", "Focus (label)",
                     "Target SCyWF job roles", "SCyWF tasks taught"], [6, 44, 12, 28, 26, 30, 52, 10], height=32)
    p.cell(1, 8).fill = fill_calc
    NP = len(programs)
    for i, pr in enumerate(programs):
        r = i + 2
        vals = [pr["id"], pr["name"], pr["level"], pr["format"], pr["audience"], ", ".join(pr["focus"]),
                "\n".join(f"{role_name[x]} ({x})" for x in pr["roles"])]
        for j, v in enumerate(vals):
            put(p, r, j + 1, v, font=f_bold if j == 1 else f_base, align=top_wrap)
        put(p, r, 8, f"=COUNTIF('Program Tasks'!$A$2:$A$2000,A{r})", align=top_wrap)
    p.freeze_panes = "C2"
    p.cell(NP + 3, 2, "A program is defined by the SCyWF tasks it teaches (Program Tasks sheet). Its priority for an entity = share of the entity's task gap those tasks cover.").font = f_note
    name(wb, "ProgIDs", f"Programs!$A$2:$A${NP + 1}")
    name(wb, "ProgNames", f"Programs!$B$2:$B${NP + 1}")

    # Entity Detail row layout (needed by lookups below)
    ED = "'Entity Detail'"
    NT = len(task_ids)
    SPN = len(SPECIALTIES)
    NR = len(roles)
    SRC_R0 = 6                       # sources rows 6-8, total row 9
    PRG_H = 13; PRG0 = PRG_H + 1; PRG1 = PRG0 + NP - 1
    SPC_H = PRG1 + 4; SPC0 = SPC_H + 1; SPC1 = SPC0 + SPN - 1
    ROL_H = SPC1 + 4; ROL0 = ROL_H + 1; ROL1 = ROL0 + NR - 1
    TSK_H = ROL1 + 4; TSK0 = TSK_H + 1; TSK1 = TSK0 + NT - 1
    REC0 = TSK1 + 4
    T_ID = f"{ED}!$A${TSK0}:$A${TSK1}"
    T_GAP = f"{ED}!$K${TSK0}:$K${TSK1}"
    T_RANK = f"{ED}!$M${TSK0}:$M${TSK1}"
    T_C = {s: f"{ED}!${c}${TSK0}:${c}${TSK1}" for s, c in (("Incident", "H"), ("Vulnerability", "I"), ("Threat", "J"))}

    pt = ws["Program Tasks"]
    header(pt, 1, 1, ["Program ID", "Task ID", "Program", "Task description", "Specialty area",
                      "Gap index (selected entity)", "From incidents", "From vulnerabilities", "From threats", "Task rank"],
           [10, 9, 40, 70, 28, 11, 10, 10, 10, 8], height=32)
    for col in range(3, 11):
        pt.cell(1, col).fill = fill_calc
    r = 2
    for pr in programs:
        for t in pr["tasks"]:
            put(pt, r, 1, pr["id"], font=f_input, fill=fill_input)
            put(pt, r, 2, t, font=f_input, fill=fill_input)
            put(pt, r, 3, f'=IFERROR(INDEX(ProgNames,MATCH(A{r},ProgIDs,0)),"Unknown program")')
            put(pt, r, 4, f'=IFERROR(INDEX(TaskDesc,MATCH(B{r},TaskIDs,0)),"Unknown task ID")')
            put(pt, r, 5, f'=IFERROR(INDEX(TaskSpec,MATCH(B{r},TaskIDs,0)),"")')
            put(pt, r, 6, f"=IFERROR(INDEX({T_GAP},MATCH(B{r},{T_ID},0)),0)", fmt="0.00%")
            for col, s in ((7, "Incident"), (8, "Vulnerability"), (9, "Threat")):
                put(pt, r, col, f"=IFERROR(INDEX({T_C[s]},MATCH(B{r},{T_ID},0)),0)", fmt="0.00%")
            put(pt, r, 10, f'=IFERROR(INDEX({T_RANK},MATCH(B{r},{T_ID},0)),"")')
            r += 1
    NPT = r - 1
    pt.freeze_panes = "C2"
    pt.auto_filter.ref = f"A1:J{NPT}"
    pt.cell(NPT + 2, 1, "Add or remove rows to change what a program teaches. Formulas on other sheets read rows 2-2000.").font = f_note
    PT = lambda col: f"'Program Tasks'!${col}$2:${col}$2000"

    # ---------------------------------------------------------------- Task Map
    tm = ws["Task Map"]
    heads = ["Source", "Title", "Task ID", "Relevance", "Relevance weight", "Allocation", "Task description",
             "Specialty area", "Selected entity: task score"] + [pr["id"] for pr in programs]
    header(tm, 1, 1, heads, [12, 46, 9, 11, 9, 10, 60, 26, 12] + [7] * NP, height=32)
    for col in range(5, len(heads) + 1):
        tm.cell(1, col).fill = fill_calc
    NTM = len(task_map) + 1
    for i, (src, title, tid, rel) in enumerate(task_map):
        r = i + 2
        put(tm, r, 1, src, font=f_input, fill=fill_input)
        put(tm, r, 2, title, font=f_input, fill=fill_input)
        put(tm, r, 3, tid, font=f_input, fill=fill_input)
        put(tm, r, 4, rel, font=f_input, fill=fill_input)
        put(tm, r, 5, f"=INDEX(RelW,MATCH(D{r},RelLabels,0))", fmt="0.0")
        put(tm, r, 6, f"=E{r}/SUMIFS($E$2:$E$2000,$B$2:$B$2000,B{r})", fmt="0.000")
        put(tm, r, 7, f'=IFERROR(INDEX(TaskDesc,MATCH(C{r},TaskIDs,0)),"Unknown task ID")')
        put(tm, r, 8, f'=IFERROR(INDEX(TaskSpec,MATCH(C{r},TaskIDs,0)),"")')
        put(tm, r, 9, f"=F{r}*IFERROR(INDEX(CatScore,MATCH(B{r},CatTitle,0)),0)", fmt="0.000")
        for j in range(NP):
            put(tm, r, 10 + j, f"=IF(COUNTIFS({PT('A')},{L(10 + j)}$1,{PT('B')},$C{r})>0,$F{r},0)", fmt="0.00;-0.00;")
    tm.freeze_panes = "D2"
    tm.auto_filter.ref = f"A1:{L(9 + NP)}{NTM}"
    tm.cell(NTM + 2, 1, "Inputs: source, title, task ID, relevance. Allocation = relevance weight / sum of relevance weights for that title, so each record is counted once in total. "
                        "Program columns = allocation if the program teaches the task (used by National View).").font = f_note
    TM = lambda col: f"'Task Map'!${col}$2:${col}$2000"

    # ---------------------------------------------------------------- Role Tasks
    rt = ws["Role Tasks"]
    header(rt, 1, 1, ["Job role ID", "Task ID", "Gap index (selected entity)", "Task rank"], [18, 9, 12, 9], height=32)
    r = 2
    for ro in roles:
        for t in ro["tasks"]:
            if t in owners:
                put(rt, r, 1, ro["id"])
                put(rt, r, 2, t)
                put(rt, r, 3, f"=IFERROR(INDEX({T_GAP},MATCH(B{r},{T_ID},0)),0)", fmt="0.00%")
                put(rt, r, 4, f'=IFERROR(INDEX({T_RANK},MATCH(B{r},{T_ID},0)),"")')
                r += 1
    NRT = r - 1
    rt.freeze_panes = "A2"
    rt.cell(NRT + 2, 1, "SCyWF job role to task links (Appendix A), limited to tasks that appear in the Task Map or a program.").font = f_note

    # ---------------------------------------------------------------- Overview: selector first (SelCode)
    o = ws["Overview"]
    for col, wd in zip("ABCDEFGHIJKLMN", [2, 9, 12, 12, 12, 12, 11, 11, 11, 11, 11, 13, 13, 2]):
        o.column_dimensions[col].width = wd
    o["B2"], o["B2"].font = "SCyWF Training Priority Model", f_title
    o["B3"] = ("Ranks the SCyWF tasks an entity most needs and the training programs that teach them, from incidents, "
               "vulnerabilities and threats. DUMMY DATA for a proof of concept.")
    o["B3"].font = f_note
    o["B5"], o["B5"].font = "Find entity", f_bold
    o["C5"] = "HLT-003"
    o.merge_cells("C5:F5")
    for c in o["C5:F5"][0]:
        c.font, c.fill, c.border = f_input, fill_input, box
    o["C5"].comment = Comment("Type an entity code (e.g. FIN-001) or part of its name (e.g. Rimal), or pick a code from the list.", "Model")
    dv = DataValidation(type="list", formula1="=EntCodes", allow_blank=True, showErrorMessage=False)
    o.add_data_validation(dv)
    dv.add("C5")
    o["B6"], o["B6"].font = "Selected", f_bold
    o["C6"] = '=IFERROR(INDEX(EntCodes,MATCH(C5,EntCodes,0)),IFERROR(INDEX(EntCodes,MATCH("*"&C5&"*",EntNames,0)),""))'
    o["C6"].font = f_link
    name(wb, "SelCode", "Overview!$C$6")
    o["D6"] = '=IF(C6="","No entity matches. Check the search results on the right.","")'
    o["D6"].font = Font(name=FONT, size=9, bold=True, color="B3203A")
    for col, wd in zip("OPQR", [12, 20, 20, 18]):
        o.column_dimensions[col].width = wd
    o["O4"], o["O4"].font = "SEARCH RESULTS (code, name or sector)", f_label
    for kk in range(8):
        r = 5 + kk
        o.cell(r, 15, f'=IFERROR(INDEX(EntCodes,MATCH({kk + 1},Entities!$F$2:$F${NE},0)),"")').font = Font(name=FONT, size=9, bold=True, color=GREEN)
        o.cell(r, 16, f'=IFERROR(INDEX(EntNames,MATCH({kk + 1},Entities!$F$2:$F${NE},0)),"")').font = Font(name=FONT, size=9, color=DARK)
        o.merge_cells(start_row=r, start_column=16, end_row=r, end_column=17)
        o.cell(r, 18, f'=IFERROR(INDEX(Entities!$C$2:$C${NE},MATCH({kk + 1},Entities!$F$2:$F${NE},0)),"")').font = Font(name=FONT, size=9, color=MUTED)
    o["O13"] = f'=IF(COUNT(Entities!$F$2:$F${NE})>8,"+ "&(COUNT(Entities!$F$2:$F${NE})-8)&" more: type more letters to narrow","")'
    o["O13"].font = f_note
    for row in o["O4:R13"]:
        for c in row:
            c.fill = fill_card

    # ---------------------------------------------------------------- Entity Detail
    d = ws["Entity Detail"]
    for col, wd in zip(range(1, 17), [14, 50, 26, 40, 11, 11, 11, 11, 11, 11, 11, 12, 8, 10, 10, 30]):
        d.column_dimensions[L(col)].width = wd
    d["A1"], d["A1"].font = "Entity detail", f_title
    d["A2"], d["A2"].font = "Entity", f_bold
    d["B2"] = '=SelCode&"  "&IFERROR(INDEX(EntNames,MATCH(SelCode,EntCodes,0)),"")'
    d["B2"].font = f_link
    d["A3"] = "Change the entity on the Overview sheet. Every table below is for that entity."
    d["A3"].font = f_note

    # sources
    header(d, 5, 1, ["Source", "Base weight", "Records", "Data sufficiency", "Effective weight", "Normalized weight",
                     "Total score", "Top task from this source"], height=32)
    d.merge_cells(start_row=5, start_column=8, end_row=5, end_column=12)
    for i, (src, sh, wn) in enumerate((("Incident", "Incidents", "W_INC"), ("Vulnerability", "Vulnerabilities", "W_VUL"),
                                       ("Threat", "Threats", "W_THR"))):
        r = SRC_R0 + i
        col_c = {"Incident": "E", "Vulnerability": "F", "Threat": "G"}[src]
        put(d, r, 1, sh, font=f_bold)
        put(d, r, 2, f"={wn}", fmt="0%")
        put(d, r, 3, f"=COUNTIF({rng(sh, 'B')},SelCode)", fmt="#,##0")
        put(d, r, 4, f"=MIN(1,C{r}/MinRec)", fmt="0%")
        put(d, r, 5, f"=B{r}*D{r}", fmt="0.000")
        put(d, r, 6, f"=IFERROR(E{r}/$E${SRC_R0 + 3},0)", fmt="0%")
        put(d, r, 7, f'=SUMIFS(CatScore,CatSrc,"{src}")', fmt="0.00")
        put(d, r, 8, f'=IF(G{r}=0,"No records",INDEX($A${TSK0}:$A${TSK1},MATCH(MAX(${col_c}${TSK0}:${col_c}${TSK1}),${col_c}${TSK0}:${col_c}${TSK1},0))'
                     f'&" "&INDEX($B${TSK0}:$B${TSK1},MATCH(MAX(${col_c}${TSK0}:${col_c}${TSK1}),${col_c}${TSK0}:${col_c}${TSK1},0)))')
        d.merge_cells(start_row=r, start_column=8, end_row=r, end_column=12)
        name(wb, {"Incident": "NW_INC", "Vulnerability": "NW_VUL", "Threat": "NW_THR"}[src], f"{ED}!$F${r}")
        name(wb, {"Incident": "TOT_INC", "Vulnerability": "TOT_VUL", "Threat": "TOT_THR"}[src], f"{ED}!$G${r}")
    rt_ = SRC_R0 + 3
    put(d, rt_, 1, "Total", font=f_bold)
    for col, fmt in ((3, "#,##0"), (5, "0.000"), (6, "0%"), (7, "0.00")):
        put(d, rt_, col, f"=SUM({L(col)}{SRC_R0}:{L(col)}{SRC_R0 + 2})", font=f_bold, fmt=fmt)
    d.cell(rt_ + 1, 1, "Effective weight = base weight x data sufficiency (records / minimum records, capped at 100%). Normalized weight = effective weight / total.").font = f_note

    # programs
    d.cell(PRG_H - 1, 1, "Training program ranking").font = f_h2
    header(d, PRG_H, 1, ["ID", "Training program", "Level", "Target SCyWF job roles", "From incidents", "From vulns",
                         "From threats", "Gap covered", "Tasks taught", "Top-10 tasks covered", "Sort key", "Rank"],
           height=32)
    for i in range(NP):
        r, pr_ = PRG0 + i, i + 2
        put(d, r, 1, f"=Programs!A{pr_}", font=f_id, align=top_wrap)
        put(d, r, 2, f"=Programs!B{pr_}", font=f_bold, align=top_wrap)
        put(d, r, 3, f"=Programs!C{pr_}", align=top_wrap)
        put(d, r, 4, f"=Programs!G{pr_}", font=f_small, align=top_wrap)
        for col, pc in ((5, "G"), (6, "H"), (7, "I")):
            put(d, r, col, f"=SUMIFS({PT(pc)},{PT('A')},$A{r})", fmt="0.0%", align=top_wrap)
        put(d, r, 8, f"=E{r}+F{r}+G{r}", font=f_bold, fmt="0.0%", align=top_wrap)
        put(d, r, 9, f"=Programs!H{pr_}", align=top_wrap)
        put(d, r, 10, f'=COUNTIFS({PT("A")},$A{r},{PT("J")},"<=10")', align=top_wrap)
        put(d, r, 11, f"=H{r}+ROW()/10000000000", font=f_note, fmt="0.0000", align=top_wrap)
        put(d, r, 12, f"=RANK(K{r},$K${PRG0}:$K${PRG1})", font=f_bold, align=top_wrap)
    d.cell(PRG1 + 1, 1, "Gap covered = sum of the gap index of every task the program teaches (Program Tasks sheet). A task taught by two programs counts for both.").font = f_note
    d.conditional_formatting.add(f"H{PRG0}:H{PRG1}", DataBarRule(start_type="num", start_value=0, end_type="max", color=BAR))

    # specialty areas
    d.cell(SPC_H - 1, 1, "Gap by SCyWF specialty area").font = f_h2
    header(d, SPC_H, 1, ["#", "Specialty area (SCyWF Table 8 ID range)", "ID range", "", "Gap index", "Sort key", "Rank"], height=32)
    for i, (lo, hi, nm) in enumerate(SPECIALTIES):
        r = SPC0 + i
        put(d, r, 1, i + 1)
        put(d, r, 2, nm, font=f_bold)
        put(d, r, 3, f"T{lo:04d}-T{hi:04d}")
        put(d, r, 5, f"=SUMIFS($K${TSK0}:$K${TSK1},$C${TSK0}:$C${TSK1},B{r})", font=f_bold, fmt="0.0%")
        put(d, r, 6, f"=E{r}+ROW()/10000000000", font=f_note, fmt="0.0000")
        put(d, r, 7, f"=RANK(F{r},$F${SPC0}:$F${SPC1})", font=f_bold)
    d.conditional_formatting.add(f"E{SPC0}:E{SPC1}", DataBarRule(start_type="num", start_value=0, end_type="max", color=BAR))

    # job roles
    d.cell(ROL_H - 1, 1, "SCyWF job roles needed").font = f_h2
    header(d, ROL_H, 1, ["Job role ID", "Job role", "Specialty area", "Category", "Gap covered", "Top-10 tasks owned",
                         "Sort key", "Rank"], height=32)
    for i, ro in enumerate(roles):
        r = ROL0 + i
        put(d, r, 1, ro["id"], font=f_id)
        put(d, r, 2, ro["name"], font=f_bold)
        put(d, r, 3, ro["specialty_area"])
        put(d, r, 4, ro["category"], font=f_small)
        put(d, r, 5, f"=SUMIFS('Role Tasks'!$C$2:$C${NRT},'Role Tasks'!$A$2:$A${NRT},A{r})", font=f_bold, fmt="0.0%")
        put(d, r, 6, f"=COUNTIFS('Role Tasks'!$A$2:$A${NRT},A{r},'Role Tasks'!$D$2:$D${NRT},\"<=10\")")
        put(d, r, 7, f"=E{r}+ROW()/10000000000", font=f_note, fmt="0.0000")
        put(d, r, 8, f"=RANK(G{r},$G${ROL0}:$G${ROL1})", font=f_bold)
    d.cell(ROL1 + 1, 1, "Gap covered = sum of the gap index of the tasks SCyWF assigns to the role (Appendix A).").font = f_note
    d.conditional_formatting.add(f"E{ROL0}:E{ROL1}", DataBarRule(start_type="num", start_value=0, end_type="max", color=BAR))

    # tasks
    d.cell(TSK_H - 1, 1, "Task gap (all mapped SCyWF tasks)").font = f_h2
    header(d, TSK_H, 1, ["Task ID", "Task description", "Specialty area", "SCyWF job roles", "Incident score",
                         "Vulnerability score", "Threat score", "From incidents", "From vulns", "From threats",
                         "Gap index", "Sort key", "Rank"], height=32)
    for i, t in enumerate(task_ids):
        r = TSK0 + i
        put(d, r, 1, t, font=f_id, align=top_wrap)
        put(d, r, 2, tasks[t], align=top_wrap)
        put(d, r, 3, specialty(t), align=top_wrap)
        put(d, r, 4, "; ".join(role_name[x] for x in owners[t]), font=f_small, align=top_wrap)
        for col, s in ((5, "Incident"), (6, "Vulnerability"), (7, "Threat")):
            put(d, r, col, f'=SUMIFS({TM("I")},{TM("C")},$A{r},{TM("A")},"{s}")', fmt="0.000", align=top_wrap)
        for col, sc, nw, tot in ((8, "E", "NW_INC", "TOT_INC"), (9, "F", "NW_VUL", "TOT_VUL"), (10, "G", "NW_THR", "TOT_THR")):
            put(d, r, col, f"=IF({tot}=0,0,{nw}*{sc}{r}/{tot})", fmt="0.00%", align=top_wrap)
        put(d, r, 11, f"=H{r}+I{r}+J{r}", font=f_bold, fmt="0.00%", align=top_wrap)
        put(d, r, 12, f"=K{r}+ROW()/10000000000", font=f_note, fmt="0.0000", align=top_wrap)
        put(d, r, 13, f"=RANK(L{r},$L${TSK0}:$L${TSK1})", font=f_bold, align=top_wrap)
    put(d, TSK1 + 1, 1, "Total", font=f_bold)
    for col in range(5, 12):
        put(d, TSK1 + 1, col, f"=SUM({L(col)}{TSK0}:{L(col)}{TSK1})", font=f_bold, fmt="0.000" if col < 8 else "0.0%")
    d.cell(TSK1 + 2, 1, "Task score = record score x allocation, summed over the entity's records. From <source> = normalized source weight x task score / source total. Gap index sums to 100%.").font = f_note
    d.conditional_formatting.add(f"K{TSK0}:K{TSK1}", DataBarRule(start_type="num", start_value=0, end_type="max", color=BAR))

    # top records
    d.cell(REC0, 1, "Top records behind the scores").font = f_h2
    r = REC0 + 1
    for src, title, lab in (("Incident", "Incidents", "Severity"), ("Vulnerability", "Vulnerabilities", "Severity"), ("Threat", "Threats", "Threat level")):
        sh, ec, tc, sc, kc, vc, dc = DS[src]
        d.cell(r, 1, f"{title}: top 10 by record score").font = f_bold
        header(d, r + 1, 1, ["#", "Record ID", "Title", lab, "Date", "Severity weight", "Recency weight", "Record score", "Row"])
        key = rng(sh, kc)
        for kk in range(10):
            rr = r + 2 + kk
            put(d, rr, 9, f'=IF(LARGE({key},{kk + 1})<=0,"",MATCH(LARGE({key},{kk + 1}),{key},0))', font=f_note)
            put(d, rr, 1, kk + 1)
            nraw = {"Incident": 6, "Vulnerability": 5, "Threat": 5}[src]
            for col, scol, fmt in ((2, "A", None), (3, tc, None), (4, vc, None), (5, dc, "yyyy-mm-dd"),
                                   (6, L(nraw + 1), "0"), (7, L(nraw + 2), "0.00"), (8, sc, "0.00")):
                put(d, rr, col, f'=IF($I{rr}="","",INDEX({q(sh)}!${scol}$2:${scol}${MAXR},$I{rr}))', fmt=fmt)
        r += 14
    d.freeze_panes = "A5"

    # ---------------------------------------------------------------- Overview: card, KPIs, rankings, charts
    o.merge_cells("B8:M8")
    o["B8"] = '=IFERROR(INDEX(EntNames,MATCH(SelCode,EntCodes,0)),"")'
    o["B8"].font = f_entity
    o.merge_cells("B9:M9")
    o["B9"] = (f'=IFERROR(SelCode&"   |   "&INDEX(Entities!$C$2:$C${NE},MATCH(SelCode,EntCodes,0))&"   |   "&'
               f'INDEX(Entities!$D$2:$D${NE},MATCH(SelCode,EntCodes,0))&"   |   Maturity: "&INDEX(Entities!$E$2:$E${NE},MATCH(SelCode,EntCodes,0)),"")')
    o["B9"].font = Font(name=FONT, size=10, color=MUTED)
    for row in o["B7:M10"]:
        for c in row:
            c.fill = fill_card
    kpis = [("INCIDENTS", f"=COUNTIF({rng('Incidents', 'B')},SelCode)", False),
            ("CRITICAL INCIDENTS", f'=COUNTIFS({rng("Incidents", "B")},SelCode,{rng("Incidents", "E")},"Critical")', True),
            ("VULNERABILITY FINDINGS", f"=COUNTIF({rng('Vulnerabilities', 'B')},SelCode)", False),
            ("CRITICAL FINDINGS", f'=COUNTIFS({rng("Vulnerabilities", "B")},SelCode,{rng("Vulnerabilities", "D")},"Critical")', True),
            ("THREAT RECORDS", f"=COUNTIF({rng('Threats', 'B')},SelCode)", False),
            ("CRITICAL THREATS", f'=COUNTIFS({rng("Threats", "B")},SelCode,{rng("Threats", "D")},"Critical")', True)]
    for i, (lab, f, red) in enumerate(kpis):
        c0 = 2 + i * 2
        o.merge_cells(start_row=11, start_column=c0, end_row=11, end_column=c0 + 1)
        o.merge_cells(start_row=12, start_column=c0, end_row=12, end_column=c0 + 1)
        o.cell(11, c0, lab).font = f_label
        o.cell(11, c0).alignment = Alignment(vertical="bottom")
        o.cell(12, c0, f).font = f_kpi_red if red else f_kpi
        o.cell(12, c0).alignment = Alignment(horizontal="left")
    o.row_dimensions[12].height = 30

    o["B14"], o["B14"].font = "How much each data source counts for this entity", f_h2
    header(o, 15, 2, ["Source", "", "Base weight", "Records", "Data sufficiency", "Weight used", "Top task from this source", "", "", "", "", ""])
    o.merge_cells("B15:C15")
    o.merge_cells("H15:M15")
    for i, lab in enumerate(("Incidents", "Vulnerabilities", "Threats")):
        r, dr = 16 + i, SRC_R0 + i
        put(o, r, 2, lab, font=f_bold)
        put(o, r, 3, None)
        o.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        put(o, r, 4, f"={ED}!B{dr}", fmt="0%")
        put(o, r, 5, f"={ED}!C{dr}", fmt="#,##0")
        put(o, r, 6, f"={ED}!D{dr}", fmt="0%")
        put(o, r, 7, f"={ED}!F{dr}", font=f_bold, fmt="0%")
        put(o, r, 8, f"={ED}!H{dr}", font=f_small)
        for cc in range(9, 14):
            put(o, r, cc, None)
        o.merge_cells(start_row=r, start_column=8, end_row=r, end_column=13)
    o["B19"] = "Weight used = base weight x data sufficiency, rescaled to 100%. Change base weights on the Weights sheet."
    o["B19"].font = f_note

    o["B21"], o["B21"].font = "Top 10 SCyWF tasks this entity needs", f_h2
    o["B22"] = "Gap index = share of the entity's weighted risk signal that falls on the task (all tasks sum to 100%)."
    o["B22"].font = f_note
    header(o, 23, 2, ["Rank", "Task", "Task description", "", "", "", "Gap index", "From incidents", "From vulns",
                      "From threats", "Specialty area", ""], height=30)
    o.merge_cells("D23:G23")
    o.merge_cells("L23:M23")
    for kk in range(10):
        r = 24 + kk
        m = f"MATCH({kk + 1},{ED}!$M${TSK0}:$M${TSK1},0)"
        put(o, r, 2, kk + 1, align=Alignment(horizontal="center", vertical="top"))
        put(o, r, 3, f"=INDEX({ED}!$A${TSK0}:$A${TSK1},{m})", font=f_id, align=Alignment(vertical="top"))
        put(o, r, 4, f"=INDEX({ED}!$B${TSK0}:$B${TSK1},{m})", align=top_wrap)
        for cc in (5, 6, 7):
            put(o, r, cc, None)
        o.merge_cells(start_row=r, start_column=4, end_row=r, end_column=7)
        for cc, sc in ((8, "K"), (9, "H"), (10, "I"), (11, "J")):
            put(o, r, cc, f"=INDEX({ED}!${sc}${TSK0}:${sc}${TSK1},{m})", font=f_bold if cc == 8 else f_base,
                fmt="0.0%", align=Alignment(vertical="top"))
        put(o, r, 12, f"=INDEX({ED}!$C${TSK0}:$C${TSK1},{m})", font=f_small, align=top_wrap)
        put(o, r, 13, None)
        o.merge_cells(start_row=r, start_column=12, end_row=r, end_column=13)
        o.row_dimensions[r].height = 28
    o.conditional_formatting.add("H24:H33", DataBarRule(start_type="num", start_value=0, end_type="max", color=BAR))

    o["B35"], o["B35"].font = "Recommended training programs", f_h2
    o["B36"] = "Ranked by gap covered = sum of the gap index of the SCyWF tasks each program teaches. Full ranking on Entity Detail."
    o["B36"].font = f_note
    header(o, 37, 2, ["Rank", "Training program", "", "", "", "Gap covered", "Top-10 tasks covered", "Level",
                      "Target SCyWF job roles", "", "", ""], height=30)
    o.merge_cells("C37:F37")
    o.merge_cells("J37:M37")
    for kk in range(5):
        r = 38 + kk
        m = f"MATCH({kk + 1},{ED}!$L${PRG0}:$L${PRG1},0)"
        put(o, r, 2, kk + 1, align=Alignment(horizontal="center", vertical="top"))
        put(o, r, 3, f"=INDEX({ED}!$B${PRG0}:$B${PRG1},{m})", font=f_bold, align=top_wrap)
        for cc in (4, 5, 6):
            put(o, r, cc, None)
        o.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
        put(o, r, 7, f"=INDEX({ED}!$H${PRG0}:$H${PRG1},{m})", font=f_bold, fmt="0.0%", align=Alignment(vertical="top"))
        put(o, r, 8, f'=INDEX({ED}!$J${PRG0}:$J${PRG1},{m})&" of 10"', align=Alignment(vertical="top"))
        put(o, r, 9, f"=INDEX({ED}!$C${PRG0}:$C${PRG1},{m})", align=Alignment(vertical="top"))
        put(o, r, 10, f"=INDEX({ED}!$D${PRG0}:$D${PRG1},{m})", font=f_small, align=top_wrap)
        for cc in (11, 12, 13):
            put(o, r, cc, None)
        o.merge_cells(start_row=r, start_column=10, end_row=r, end_column=13)
        o.row_dimensions[r].height = 40
    o.conditional_formatting.add("G38:G42", DataBarRule(start_type="num", start_value=0, end_type="max", color=BAR))

    o["B44"], o["B44"].font = "Gap by SCyWF specialty area (top 6)", f_h2
    header(o, 45, 2, ["Rank", "Specialty area", "", "", "", "Gap index"])
    o.merge_cells("C45:F45")
    for kk in range(6):
        r = 46 + kk
        m = f"MATCH({kk + 1},{ED}!$G${SPC0}:$G${SPC1},0)"
        put(o, r, 2, kk + 1, align=Alignment(horizontal="center"))
        put(o, r, 3, f"=INDEX({ED}!$B${SPC0}:$B${SPC1},{m})", font=f_bold)
        for cc in (4, 5, 6):
            put(o, r, cc, None)
        o.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
        put(o, r, 7, f"=INDEX({ED}!$E${SPC0}:$E${SPC1},{m})", font=f_bold, fmt="0.0%")
    o.conditional_formatting.add("G46:G51", DataBarRule(start_type="num", start_value=0, end_type="max", color=BAR))
    o["B53"] = "Blue on yellow = input. Weights: Weights sheet. What each program teaches: Program Tasks sheet. Record-to-task mapping: Task Map sheet."
    o["B53"].font = f_note

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
        ch.height, ch.width = 7.5, 16
        o.add_chart(ch, anchor)

    bar("Gap covered, top 5 programs", Reference(o, min_col=3, min_row=38, max_row=42),
        Reference(o, min_col=7, min_row=38, max_row=42), "O21", "0%")
    bar("Gap by specialty area, top 6", Reference(o, min_col=3, min_row=46, max_row=51),
        Reference(o, min_col=7, min_row=46, max_row=51), "O37", "0%")
    o.freeze_panes = "A7"

    # ---------------------------------------------------------------- Title Matrix (national helper)
    tx = ws["Title Matrix"]
    ent_codes = [x["entity_code"] for x in entities]
    PC0 = 3                       # program coverage columns
    SC0 = PC0 + NP                # specialty coverage columns
    EC0 = SC0 + SPN               # entity title-score columns
    header(tx, 1, 1, ["Source", "Title"] + [pr["id"] for pr in programs] + [nm for _, _, nm in SPECIALTIES] + ent_codes,
           [12, 46] + [7] * NP + [10] * SPN + [9] * len(ent_codes), height=60)
    tx["A1"].comment = Comment("Program and specialty columns: share of a title's score that lands on tasks taught by the program / in the specialty area. "
                               "Entity columns: each entity's total score for the title. Used by National View.", "Model")
    for i, (src, title, _, _) in enumerate(cat_rows):
        r = i + 2
        put(tx, r, 1, f"=Catalog!A{r}")
        put(tx, r, 2, f"=Catalog!B{r}")
        for j in range(NP):
            put(tx, r, PC0 + j, f"=SUMIFS({TM(L(10 + j))},{TM('B')},$B{r})", fmt="0.00;-0.00;")
        for j, (_, _, nm) in enumerate(SPECIALTIES):
            put(tx, r, SC0 + j, f'=SUMIFS({TM("F")},{TM("B")},$B{r},{TM("H")},"{nm}")', fmt="0.00;-0.00;")
        for j in range(len(ent_codes)):
            hc = f"{L(EC0 + j)}$1"
            put(tx, r, EC0 + j, f'=IF($A{r}="Incident",{sumifs("Incident", hc, f"$B{r}")},IF($A{r}="Vulnerability",'
                                f'{sumifs("Vulnerability", hc, f"$B{r}")},{sumifs("Threat", hc, f"$B{r}")}))', fmt="0.00;-0.00;")
    tx.freeze_panes = "C2"
    TXR = lambda col: f"'Title Matrix'!${col}$2:${col}${NCAT}"

    # ---------------------------------------------------------------- National View
    n = ws["National View"]
    n["A1"], n["A1"].font = "National view", f_title
    n["A2"] = "The same model for every entity: gap by SCyWF specialty area and gap covered by each training program. Darker = larger."
    n["A2"].font = f_note
    G0 = 15                       # first specialty column (O)
    TOPC = G0 + SPN               # top specialty, top program, top score
    P0C = TOPC + 4                # first program column
    heads = ["Entity code", "Entity name", "Sector", "Maturity", "Incidents", "Vulns", "Threats", "Eff. weight inc.",
             "Eff. weight vuln.", "Eff. weight threat", "Eff. weight total", "Incident total", "Vuln total", "Threat total"]
    header(n, 4, 1, heads + [nm for _, _, nm in SPECIALTIES] + ["Top specialty gap", "Top program", "Top program gap covered"], height=70)
    header(n, 4, P0C, [pr["id"] for pr in programs], height=70)
    for i, pr in enumerate(programs):
        c = n.cell(3, P0C + i, pr["name"])
        c.font, c.alignment = f_label, Alignment(wrap_text=True, vertical="bottom")
    n.row_dimensions[3].height = 70
    for col, wd in zip(range(1, 15), [11, 34, 16, 11, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9]):
        n.column_dimensions[L(col)].width = wd
    for j in range(SPN):
        n.column_dimensions[L(G0 + j)].width = 11
    n.column_dimensions[L(TOPC)].width = 26
    n.column_dimensions[L(TOPC + 1)].width = 40
    n.column_dimensions[L(TOPC + 2)].width = 11
    n.column_dimensions[L(TOPC + 3)].width = 3
    for j in range(NP):
        n.column_dimensions[L(P0C + j)].width = 10
    E0, E1 = 5, 5 + len(entities) - 1
    srcs = (("Incident", "H", "L"), ("Vulnerability", "I", "M"), ("Threat", "J", "N"))
    srcA = TXR("A")
    for i in range(len(entities)):
        r, er = E0 + i, i + 2
        ecol = L(EC0 + i)
        put(n, r, 1, f"=Entities!A{er}", font=f_id)
        put(n, r, 2, f"=Entities!B{er}")
        put(n, r, 3, f"=Entities!C{er}")
        put(n, r, 4, f"=Entities!E{er}")
        put(n, r, 5, f"=COUNTIF({rng('Incidents', 'B')},$A{r})")
        put(n, r, 6, f"=COUNTIF({rng('Vulnerabilities', 'B')},$A{r})")
        put(n, r, 7, f"=COUNTIF({rng('Threats', 'B')},$A{r})")
        put(n, r, 8, f"=W_INC*MIN(1,E{r}/MinRec)", fmt="0.000")
        put(n, r, 9, f"=W_VUL*MIN(1,F{r}/MinRec)", fmt="0.000")
        put(n, r, 10, f"=W_THR*MIN(1,G{r}/MinRec)", fmt="0.000")
        put(n, r, 11, f"=SUM(H{r}:J{r})", fmt="0.000")
        for col, (s, _, _) in zip((12, 13, 14), srcs):
            put(n, r, col, f'=SUMPRODUCT(({srcA}="{s}")*{TXR(ecol)})', fmt="0.0")

        def blend(covcol):
            parts = [f'{wc}{r}*IFERROR(SUMPRODUCT(({srcA}="{s}")*{TXR(ecol)}*{TXR(covcol)})/{tc}{r},0)' for s, wc, tc in srcs]
            return f"=IF($K{r}=0,0,({'+'.join(parts)})/$K{r})"

        for j in range(SPN):
            put(n, r, G0 + j, blend(L(SC0 + j)), fmt="0%")
        for j in range(NP):
            put(n, r, P0C + j, blend(L(PC0 + j)), fmt="0%")
        gl, gr, pl, pr2 = L(G0), L(G0 + SPN - 1), L(P0C), L(P0C + NP - 1)
        put(n, r, TOPC, f"=INDEX(${gl}$4:${gr}$4,MATCH(MAX({gl}{r}:{gr}{r}),{gl}{r}:{gr}{r},0))", font=f_bold)
        put(n, r, TOPC + 1, f"=INDEX(${pl}$3:${pr2}$3,MATCH(MAX({pl}{r}:{pr2}{r}),{pl}{r}:{pr2}{r},0))", font=f_bold)
        put(n, r, TOPC + 2, f"=MAX({pl}{r}:{pr2}{r})", fmt="0%")
    ar = E1 + 2
    put(n, ar, 2, "National average", font=f_bold, fill=fill_sub)
    put(n, ar + 1, 2, "Entities where this is the top item", font=f_bold, fill=fill_sub)
    for j in range(SPN):
        cl = L(G0 + j)
        put(n, ar, G0 + j, f"=AVERAGE({cl}{E0}:{cl}{E1})", font=f_bold, fmt="0%", fill=fill_sub)
        put(n, ar + 1, G0 + j, f"=COUNTIF(${L(TOPC)}${E0}:${L(TOPC)}${E1},{cl}$4)", font=f_bold, fill=fill_sub)
    for j in range(NP):
        cl = L(P0C + j)
        put(n, ar, P0C + j, f"=AVERAGE({cl}{E0}:{cl}{E1})", font=f_bold, fmt="0%", fill=fill_sub)
        put(n, ar + 1, P0C + j, f"=COUNTIF(${L(TOPC + 1)}${E0}:${L(TOPC + 1)}${E1},{cl}$3)", font=f_bold, fill=fill_sub)
    sectors = sorted({x["sector"] for x in entities})
    s0 = ar + 4
    n.cell(s0 - 1, 1, "Sector averages").font = f_h2
    header(n, s0, 1, ["", "Sector", "Entities"] + [""] * 11 + [nm for _, _, nm in SPECIALTIES] + ["Top specialty gap", "Top program", "Top program gap covered"])
    header(n, s0, P0C, [pr["id"] for pr in programs])
    gl, gr, pl, pr2 = L(G0), L(G0 + SPN - 1), L(P0C), L(P0C + NP - 1)
    for i, sec in enumerate(sectors):
        r = s0 + 1 + i
        put(n, r, 2, sec, font=f_bold)
        put(n, r, 3, f"=COUNTIF($C${E0}:$C${E1},$B{r})")
        for j in list(range(G0, G0 + SPN)) + list(range(P0C, P0C + NP)):
            cl = L(j)
            put(n, r, j, f"=AVERAGEIFS({cl}${E0}:{cl}${E1},$C${E0}:$C${E1},$B{r})", fmt="0%")
        put(n, r, TOPC, f"=INDEX(${gl}$4:${gr}$4,MATCH(MAX({gl}{r}:{gr}{r}),{gl}{r}:{gr}{r},0))", font=f_bold)
        put(n, r, TOPC + 1, f"=INDEX(${pl}$3:${pr2}$3,MATCH(MAX({pl}{r}:{pr2}{r}),{pl}{r}:{pr2}{r},0))", font=f_bold)
        put(n, r, TOPC + 2, f"=MAX({pl}{r}:{pr2}{r})", fmt="0%")
    S1 = s0 + len(sectors)
    scale = ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF", end_type="max", end_color=BAR)
    for ref in (f"{gl}{E0}:{gr}{E1}", f"{pl}{E0}:{pr2}{E1}", f"{gl}{s0 + 1}:{gr}{S1}", f"{pl}{s0 + 1}:{pr2}{S1}"):
        n.conditional_formatting.add(ref, scale)
    n.freeze_panes = "C5"
    n.auto_filter.ref = f"A4:{L(TOPC + 2)}{E1}"

    # ---------------------------------------------------------------- How It Works
    h = ws["How It Works"]
    h.column_dimensions["B"].width = 125
    lines = [
        ("How the model works", f_title), ("", None),
        ("The idea", f_h2),
        ("Every incident, vulnerability finding and threat record points to SCyWF tasks the entity needs to perform better. "
         "The model adds up that evidence per task, weighs the three sources, and ranks tasks, job roles, specialty areas and training programs.", f_base),
        ("Only four fields are needed per record: entity code, title, severity (or threat level) and date.", f_bold), ("", None),
        ("Step 1. Score each record", f_h2),
        ("Record score = severity weight x recency weight", f_bold),
        ("Severity weight doubles each level: Low 1, Medium 2, High 4, Critical 8, the same for all three sources.", f_base),
        ("Recency weight = 0.5 ^ (age in days / half-life). Half-life: incidents 365 days, threats 180 days, vulnerabilities no decay (0) because a finding in the current scan is current exposure.", f_base),
        ("", None),
        ("Step 2. Spread each record over its SCyWF tasks", f_h2),
        ("The Task Map lists, for every title, the tasks it maps to as Primary (1.0) or Supporting (0.5). A record's score is split over its tasks in proportion to those weights, "
         "so each record counts exactly once in total, however many tasks it maps to.", f_base),
        ("Task score (per source) = sum over the entity's records of record score x allocation to the task", f_bold), ("", None),
        ("Step 3. Weigh the sources", f_h2),
        ("Base weights come from evidence strength: incidents 3 (proven harm here), vulnerabilities 2 (exposure here, not yet exploited), threats 1 (external, often sector-wide). "
         "Normalized: 50% / 33% / 17%.", f_base),
        ("Data sufficiency: a source with fewer than 10 records gets proportionally less weight (records / 10), so a handful of records cannot swing the result. A source with no records drops out.", f_base),
        ("Source share = task score / entity's total score from that source. This stops a large scan report from outweighing incidents simply by having more rows.", f_base),
        ("Task gap index = sum over sources of (normalized weight x source share). All task gap indexes add up to 100%.", f_bold), ("", None),
        ("Step 4. Roll up", f_h2),
        ("Training program: gap covered = sum of the gap index of the tasks it teaches (Program Tasks sheet). Programs are ranked on this.", f_bold),
        ("Job role: sum of the gap index of the tasks SCyWF assigns to the role (Appendix A).", f_base),
        ("Specialty area: sum of the gap index of tasks in the area's ID range (SCyWF Table 8).", f_base),
        ("National View applies the same calculation to every entity (through the Title Matrix sheet) and averages by sector.", f_base), ("", None),
        ("Sheets", f_h2),
        ("Overview, Entity Detail, National View: results.   Weights, Task Map, Program Tasks, Programs: inputs you can edit (yellow tabs).", f_base),
        ("Incidents, Vulnerabilities, Threats: data (grey headers are calculated columns).   Catalog, Title Matrix, Role Tasks: helpers.   SCyWF Tasks, SCyWF Roles: framework reference.", f_base), ("", None),
        ("Adding real data", f_h2),
        (f"Paste records under the existing rows of Incidents, Vulnerabilities or Threats and copy the grey calculated columns down (formulas read up to row {MAXR}). "
         "Each title must exist in the Catalog and the Task Map. Add new titles to both: one Catalog row, and one Task Map row per task.", f_base),
        ("The task list on Entity Detail holds the tasks used when the workbook was built. If you map a title to a task that is not listed there, add a row for it on Entity Detail too, or rebuild the workbook with scripts/build_excel.py.", f_base),
        ("", None),
        ("Source: task IDs, descriptions, job roles and specialty ranges from the Saudi Cybersecurity Workforce Framework (SCyWF 1.5 : 2026), National Cybersecurity Authority, extracted by OCR. "
         "The task mapping, weights and programs are proposals for review. Entities and records are fictional.", f_note),
    ]
    for i, (txt, fnt) in enumerate(lines):
        c = h.cell(i + 1, 2, txt)
        c.font = fnt or f_base
        c.alignment = Alignment(wrap_text=True, vertical="top")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"Saved {OUT}: {len(task_map)} task-map rows, {NT} tasks, {NPT - 1} program-task rows, {NRT - 1} role-task rows")


if __name__ == "__main__":
    main()
