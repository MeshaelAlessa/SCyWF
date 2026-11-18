# SCyWF Gap Explorer (POC)

A proof of concept for turning cybersecurity incident data into **workforce gaps**.
It shows which skills Saudi entities lack, and which programs and practices would close those gaps,
using the **Saudi Cybersecurity Workforce Framework (SCyWF 1.5 : 2026)** published by the NCA.

```
incident ──► vulnerability title ──► SCyWF tasks ──► job roles / specialty areas ──► programs & practices
```

> All entities and incidents in this repo are **fictional dummy data**.

## What it does

Open `app/index.html` in a browser. It is a single self-contained file that works offline.

| View | What you get |
|---|---|
| **Entity lookup** | Search an entity by code, name or sector. You see its incidents grouped by vulnerability and ranked by **frequency** or **severity**. Click a row to see the SCyWF tasks it maps to. The view also lists the entity's **recommended SCyWF tasks** (scored), its **skill gap by specialty area**, the **job roles to build**, and suggested **programs and practices**. |
| **National gaps** | A heatmap of sectors × SCyWF specialty areas, plus the most frequent vulnerabilities, the most needed tasks and the job roles in demand. You can filter by sector. |
| **Task mapping** | The full vulnerability → task mapping, with a filter. |

### Scoring

- Severity weight: Low 1, Medium 2, High 3, Critical 4.
- **Task score** = Σ over incidents of (severity weight × relevance), where relevance is 1.0 for a *primary* task and 0.5 for a *supporting* task.
- **Specialty-area gap** = share of the task score in each SCyWF specialty area. The area comes from the task ID range in SCyWF Table 8; for example, T5000–T5499 is Incident Response.
- **Role score** = sum of the scores of the tasks that SCyWF assigns to that job role (Appendix A).
- **Ranking incidents:** *Frequency* sorts by count, then severity score. *Severity* sorts by highest severity, then the count at that severity, then severity score.

## Excel model: `excel/SCyWF_Training_Priority_Model.xlsx`

A formula-driven workbook. Three data sources: **incidents**, **vulnerability findings** and **NCA ECC compliance results**.
Each record maps to the SCyWF tasks that, done well, would have **prevented** it. The workbook ranks the tasks an entity,
a sector or the whole country (type `All`) most needs, and the training programs that teach them.

| Sheet | Purpose |
|---|---|
| **Overview** | Type an entity code or name, a sector, or `All`. Shows KPIs, source weights, the top 10 tasks, the top 5 programs and the top specialty areas. |
| **Details** | Every task, program and specialty area ranked for the selection. |
| **How It Works** | The method, with examples. |
| **Weights**, **Task Map**, **Programs** | Inputs: source and severity weights, title → tasks, program → tasks. |
| **Incidents**, **Vulnerabilities**, **Compliance**, **Entities** | Data. |
| **SCyWF Tasks** | Framework reference. |

### Weighting
1. **Severity × frequency.** Each record gets a severity weight: Low 1, Medium 2, High 4, Critical 8. ECC results use Not compliant 4 and Partially compliant 2. Frequency is built in, because every record adds its weight.
2. **Split over prevention tasks.** Primary tasks (1.0) prevent the record directly; Supporting tasks (0.5) help prevent it. Each record's weight is split in proportion, so it counts once.
3. **Source weights.** Incidents 50%, vulnerabilities 30%, ECC 20%. Gap index = Σ source weight × (task score ÷ source total). It sums to 100% across tasks.
4. **Programs** are ranked by the sum of the gap index of the tasks they teach.

There is no recency weighting: filter the data to the period you want before loading it. All weights are editable cells.

Rebuild with `python3 scripts/generate_dummy_data.py && python3 scripts/build_excel.py`. This needs `openpyxl`.
Open the file in Excel to calculate formulas. The committed copy already has its values calculated.

## Repository layout

```
data/
  scywf/tasks.json                       645 SCyWF tasks (ID -> description), from Appendix B, Table 9
  scywf/job_roles.json                   40 SCyWF job roles with their task IDs, from Appendix A
  mapping/vulnerability_task_mapping.json  34 vulnerability titles -> primary/supporting SCyWF tasks,
                                         grouped into 14 categories, each with programs and practices
  generated/entities.csv                 dummy entities (code, name, sector, region, maturity)
  generated/incidents.csv                dummy incidents (entity code/name, vulnerability title, severity, date)
  generated/vulnerabilities.csv          dummy vulnerability findings (entity code, finding title, severity, date)
  generated/compliance.csv               dummy NCA ECC compliance results (entity code, ECC subdomain, status, date)
  mapping/source_catalog.json            finding types and ECC subdomains -> SCyWF tasks; programs with the tasks they teach
  generated/vulnerability_task_mapping.csv flat mapping (one row per vulnerability-task pair) for BI tools
scripts/
  generate_dummy_data.py                 deterministic dummy data generator (--seed)
  build_app.py                           builds app/index.html and the flat CSV
  build_excel.py                         builds excel/SCyWF_Training_Priority_Model.xlsx
app/
  template.html                          UI source
  index.html                             built app with data embedded
```

## Running

```bash
python3 scripts/generate_dummy_data.py      # optional: regenerate dummy data (default seed 2026)
python3 scripts/build_app.py                # rebuild app/index.html
open app/index.html                         # or double-click it
```

Only the Python 3 standard library is needed.

### Using real data

Replace `data/generated/incidents.csv` with an export that has the same columns
(`incident_id, entity_code, entity_name, vulnerability_title, severity, reported_date, status`),
and `entities.csv` likewise, then run `build_app.py`. Every `vulnerability_title` must exist in
`vulnerability_task_mapping.json`; add new titles there with their primary and supporting task IDs.

## Notes on the SCyWF extraction

The tables in the NCA PDF are embedded as images, so tasks and job roles were extracted with OCR
(Tesseract). The extraction was cross-checked: every task ID referenced by the 40 job roles has a
description in `tasks.json`. OCR can still misread single characters, so validate against the
official document before any production use.

## Ideas for next steps

- Add SCyWF **knowledge and skills** (K/S IDs) per job role, for a finer gap profile.
- Map free-text incident titles to the catalog automatically (keyword or LLM classification), instead of using exact titles.
- Link each task or role to a catalog of real national training programs and certifications.
- Track gaps over time per entity, to measure whether programs reduce repeat incidents.
