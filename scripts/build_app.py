#!/usr/bin/env python3
"""Build the single-file web app (app/index.html) and flat CSV exports.

Inputs:
  data/scywf/tasks.json                         SCyWF task catalog (extracted from the NCA PDF)
  data/scywf/job_roles.json                     SCyWF job roles with their task IDs
  data/mapping/vulnerability_task_mapping.json  vulnerability title -> SCyWF tasks
  data/generated/entities.csv, incidents.csv    dummy data (scripts/generate_dummy_data.py)
Outputs:
  app/index.html                                self-contained explorer (data embedded)
  data/generated/vulnerability_task_mapping.csv flat mapping for BI tools
"""
import argparse
import csv
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SEVERITIES = ["Low", "Medium", "High", "Critical"]

# Table 8 of SCyWF: TKS numbering ranges -> specialty area
SPECIALTY_RANGES = [
    (0, 499, "General"), (500, 999, "Cybersecurity Architecture"),
    (1000, 1499, "Cybersecurity Research and Development"), (1500, 1999, "Leadership"),
    (2000, 2499, "Workforce Development"), (2500, 2999, "Governance, Risk and Compliance"),
    (3000, 3499, "Laws and Data Protection"), (3500, 3999, "Defense"), (4000, 4499, "Protection"),
    (4500, 4999, "Vulnerability Assessment"), (5000, 5499, "Incident Response"),
    (5500, 5999, "Threat Management"), (6000, 6499, "ICS/OT"),
]


def specialty(task_id):
    n = int(task_id[1:])
    return next(name for lo, hi, name in SPECIALTY_RANGES if lo <= n <= hi)


def load_csv(p):
    with open(p, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--fragment", help="also write the page body without the <html>/<head> wrapper (for hosts that add their own)")
    args = ap.parse_args()
    tasks = json.loads((ROOT / "data/scywf/tasks.json").read_text(encoding="utf-8"))
    roles = json.loads((ROOT / "data/scywf/job_roles.json").read_text(encoding="utf-8"))
    mapping = json.loads((ROOT / "data/mapping/vulnerability_task_mapping.json").read_text(encoding="utf-8"))
    entities = load_csv(ROOT / "data/generated/entities.csv")
    incidents = load_csv(ROOT / "data/generated/incidents.csv")

    vulns = mapping["vulnerabilities"]
    vidx = {v["title"]: i for i, v in enumerate(vulns)}
    used = sorted({t for v in vulns for t in v["primary"] + v["supporting"]})
    owners = {t: [r["id"] for r in roles if t in r["tasks"]] for t in used}

    # Flat CSV of the mapping
    with open(ROOT / "data/generated/vulnerability_task_mapping.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["vulnerability_title", "category", "task_id", "relevance", "task_description",
                    "specialty_area", "scywf_job_roles"])
        for v in vulns:
            for rel in ("primary", "supporting"):
                for t in v[rel]:
                    w.writerow([v["title"], v["category"], t, rel, tasks[t], specialty(t), "; ".join(owners[t])])

    data = {
        "severities": SEVERITIES,
        "tasks": {t: {"d": tasks[t], "s": specialty(t), "r": owners[t]} for t in used},
        "roles": {r["id"]: {"n": r["name"], "s": r["specialty_area"], "c": r["category"],
                            "t": len(r["tasks"])} for r in roles},
        "categories": mapping["categories"],
        "vulns": [{"t": v["title"], "c": v["category"], "p": v["primary"], "s": v["supporting"]} for v in vulns],
        "entities": [{"code": e["entity_code"], "name": e["entity_name"], "sector": e["sector"],
                      "region": e["region"], "maturity": e["maturity_level"]} for e in entities],
        # compact incidents: [id, entity_code, vuln_index, severity_index, date]
        "incidents": [[i["incident_id"], i["entity_code"], vidx[i["vulnerability_title"]],
                       SEVERITIES.index(i["severity"]), i["reported_date"]] for i in incidents],
    }
    template = (ROOT / "app/template.html").read_text(encoding="utf-8")
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    out = template.replace("/*__DATA__*/null", payload)
    if args.fragment:
        Path(args.fragment).write_text(out, encoding="utf-8")
    doc = ('<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n'
           '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
           '<style>body{margin:0}[hidden]{display:none!important}</style>\n</head>\n<body>\n' + out + '\n</body>\n</html>\n')
    (ROOT / "app/index.html").write_text(doc, encoding="utf-8")
    print(f"Built app/index.html ({len(out)//1024} KB): {len(entities)} entities, {len(incidents)} incidents, "
          f"{len(vulns)} vulnerability types, {len(used)} mapped SCyWF tasks")


if __name__ == "__main__":
    main()
