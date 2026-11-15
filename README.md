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

A formula-driven workbook that combines **three data sources** per entity and ranks **training programs**.

| Sheet | Purpose |
|---|---|
| **Overview** | Type an entity code or part of its name. It shows a summary card, KPIs, how much each data source contributes, the top 5 gap categories and the top 5 programs, with charts. |
| **Entity Detail** | The full calculation for the selected entity: category gaps, all 15 programs ranked (split into incident-, vulnerability- and threat-driven parts), and the top 10 records from each source. |
| **National View** | Every entity and sector side by side: gap heatmap, program scores, top gap and top program, and national averages. |
| **How It Works** | Explains the method in plain language. |
| **Weights** | Source weights and scoring factors (inputs). |
| **Programs** | Program catalog and coverage matrix (inputs). |
| **Categories**, **Catalog** | Maps every incident, finding and threat title to a gap category and its SCyWF tasks. |
| **Incidents**, **Vulnerabilities**, **Threats** | Data, with calculated scoring columns. |
| **SCyWF Tasks**, **SCyWF Roles** | Framework reference. |

### Weighting system

1. **Record score** (signal strength of one record):
   - Incident = severity × recency × status
   - Vulnerability = severity from CVSS × exposure (internet-facing 1.5) × status (remediated 0.2)
   - Threat = level × relevance (entity-specific 1.5, sector 1.2, general 0.8) × confidence × recency
2. **Source share**: for each source, the share of an entity's record score that falls in each gap category. This makes sources with very different volumes comparable.
3. **Gap index** = (0.5 × incident share + 0.3 × vulnerability share + 0.2 × threat share) ÷ (sum of the weights of sources that have data).
   - Incidents weigh most because they are proven harm.
   - Vulnerabilities are current exposure.
   - Threats are likelihood and the least entity-specific.
4. **Program priority** = Σ gap index × program coverage (1.0 primary focus, 0.5 secondary).

Every weight and factor is an editable cell on the Weights sheet, and every score recalculates when one changes.

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
  generated/incidents.csv                dummy incidents (entity code/name, vulnerability title, severity, date, status)
  generated/vulnerabilities.csv          dummy vulnerability findings (title, CVSS, exposure, status, date)
  generated/threats.csv                  dummy threat intelligence (title, level, relevance, confidence, date)
  mapping/source_catalog.json            finding and threat types -> gap category + SCyWF tasks; training program catalog
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
