#!/usr/bin/env python3
"""Generate deterministic DUMMY entities and incidents for the SCyWF gap-analysis POC.

All entity names are fictional. Incidents use the vulnerability titles defined in
data/mapping/vulnerability_task_mapping.json so every incident can be mapped to SCyWF tasks.

Usage: python3 scripts/generate_dummy_data.py [--seed 2026]
"""
import argparse
import csv
import json
import random
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
MAPPING = ROOT / "data" / "mapping" / "vulnerability_task_mapping.json"
OUT = ROOT / "data" / "generated"

SEVERITIES = ["Low", "Medium", "High", "Critical"]
STATUSES = [("Closed", 0.62), ("Contained", 0.23), ("Open", 0.15)]
REGIONS = ["Riyadh", "Makkah", "Eastern Province", "Madinah", "Qassim", "Asir",
           "Tabuk", "Hail", "Jazan", "Najran", "Al-Baha", "Al-Jouf", "Northern Borders"]

# Fictional entities: (code, name, sector)
ENTITIES = [
    ("GOV-001", "Rawdah Municipal Services Authority", "Government"),
    ("GOV-002", "Najmah Civil Registry Agency", "Government"),
    ("GOV-003", "Sadeem Public Procurement Office", "Government"),
    ("GOV-004", "Ufuq Housing Support Program", "Government"),
    ("GOV-005", "Masar Licensing and Permits Center", "Government"),
    ("GOV-006", "Bayan Statistics and Planning Unit", "Government"),
    ("HLT-001", "Shifa Northern Health Cluster", "Health"),
    ("HLT-002", "Rahma Specialist Hospital Group", "Health"),
    ("HLT-003", "Nabd Medical Laboratories", "Health"),
    ("HLT-004", "Afiya Primary Care Network", "Health"),
    ("FIN-001", "Rimal Digital Bank", "Finance"),
    ("FIN-002", "Sanad Microfinance Company", "Finance"),
    ("FIN-003", "Dirham Payments Gateway", "Finance"),
    ("FIN-004", "Waqar Takaful Insurance", "Finance"),
    ("ENR-001", "Dana Petrochemical Processing Co.", "Energy"),
    ("ENR-002", "Shuaa Solar Power Plant Operator", "Energy"),
    ("ENR-003", "Nawras Gas Distribution Co.", "Energy"),
    ("UTL-001", "Ghadeer Regional Water Utility", "Water & Utilities"),
    ("UTL-002", "Barq Electricity Distribution Co.", "Water & Utilities"),
    ("UTL-003", "Ain Desalination Operations", "Water & Utilities"),
    ("TEL-001", "Sahab Telecom Services", "Telecom & ICT"),
    ("TEL-002", "Majal Data Center Hosting", "Telecom & ICT"),
    ("TEL-003", "Raqmi Cloud Solutions", "Telecom & ICT"),
    ("EDU-001", "Qimam University", "Education"),
    ("EDU-002", "Manar Technical College", "Education"),
    ("EDU-003", "Nukhba International Schools", "Education"),
    ("EDU-004", "Ilm E-Learning Platform", "Education"),
    ("TRN-001", "Wasl Logistics and Freight", "Transport"),
    ("TRN-002", "Sikka Rail Operations", "Transport"),
    ("TRN-003", "Marsa Port Services", "Transport"),
    ("RTL-001", "Souq Al-Nakheel E-commerce", "Retail & Services"),
    ("RTL-002", "Zad Food Distribution Group", "Retail & Services"),
    ("RTL-003", "Rihla Travel and Booking", "Retail & Services"),
]

# Relative likelihood of each vulnerability category per sector (default 1.0)
SECTOR_PROFILE = {
    "Government": {"Social Engineering": 3, "Patch & Vulnerability Management": 2.5, "Governance & Compliance": 2.5,
                    "Network Security": 1.8, "ICS/OT Security": 0},
    "Health": {"Malware & Ransomware": 3.5, "Data Protection": 3, "Social Engineering": 2,
               "Patch & Vulnerability Management": 2, "ICS/OT Security": 0},
    "Finance": {"Social Engineering": 3, "Application Security": 3, "Identity & Access": 2.5,
                "Threat Intelligence": 2, "Third-Party & Supply Chain": 2, "ICS/OT Security": 0},
    "Energy": {"ICS/OT Security": 4, "Threat Intelligence": 2.5, "Malware & Ransomware": 1.5,
               "Application Security": 0.4},
    "Water & Utilities": {"ICS/OT Security": 4, "Patch & Vulnerability Management": 2,
                          "Network Security": 1.8, "Application Security": 0.3},
    "Telecom & ICT": {"Network Security": 3, "Cloud Security": 3, "Threat Intelligence": 2.5,
                      "Third-Party & Supply Chain": 2, "ICS/OT Security": 0.2},
    "Education": {"Social Engineering": 3.5, "Identity & Access": 3, "Patch & Vulnerability Management": 2,
                  "Incident Response Readiness": 2, "ICS/OT Security": 0},
    "Transport": {"Malware & Ransomware": 2.5, "ICS/OT Security": 2, "Third-Party & Supply Chain": 2.5,
                  "Social Engineering": 1.5},
    "Retail & Services": {"Application Security": 3, "Identity & Access": 2.5, "Cloud Security": 2,
                          "Data Protection": 2, "ICS/OT Security": 0},
}


def weighted_choice(rng, items, weights):
    return rng.choices(items, weights=weights, k=1)[0]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--seed", type=int, default=2026)
    args = ap.parse_args()
    rng = random.Random(args.seed)

    mapping = json.loads(MAPPING.read_text(encoding="utf-8"))
    vulns = mapping["vulnerabilities"]
    OUT.mkdir(parents=True, exist_ok=True)

    start, end = date(2025, 1, 1), date(2026, 10, 31)
    span = (end - start).days

    entity_rows, incident_rows = [], []
    inc_no = 1
    for code, name, sector in ENTITIES:
        region = rng.choice(REGIONS)
        maturity = rng.choice(["Initial", "Developing", "Defined", "Managed"])
        volume = {"Initial": (35, 70), "Developing": (22, 45), "Defined": (12, 28), "Managed": (6, 16)}[maturity]
        n_incidents = rng.randint(*volume)
        entity_rows.append({"entity_code": code, "entity_name": name, "sector": sector,
                            "region": region, "maturity_level": maturity})

        profile = SECTOR_PROFILE.get(sector, {})
        # Each entity has its own "weak spots": a few vulnerabilities it keeps repeating
        weights = [profile.get(v["category"], 1.0) * rng.uniform(0.3, 1.0) for v in vulns]
        weak = set()
        while len(weak) < 3:
            weak.add(weighted_choice(rng, range(len(vulns)), weights))
        for i in weak:
            weights[i] *= rng.uniform(2.5, 4)

        for _ in range(n_incidents):
            v = weighted_choice(rng, vulns, weights)
            base = SEVERITIES.index(v["base_severity"])
            shift = weighted_choice(rng, [-1, 0, 1], [0.25, 0.6, 0.15])
            sev = SEVERITIES[max(0, min(3, base + shift))]
            d = start + timedelta(days=rng.randint(0, span))
            status = weighted_choice(rng, [s for s, _ in STATUSES], [w for _, w in STATUSES])
            incident_rows.append({
                "incident_id": f"INC-{d.year}-{inc_no:05d}",
                "entity_code": code,
                "entity_name": name,
                "vulnerability_title": v["title"],
                "severity": sev,
                "reported_date": d.isoformat(),
                "status": status,
            })
            inc_no += 1

    incident_rows.sort(key=lambda r: (r["reported_date"], r["incident_id"]))
    for i, r in enumerate(incident_rows, 1):  # number incidents in date order
        r["incident_id"] = f"INC-{r['reported_date'][:4]}-{i:05d}"
    with open(OUT / "entities.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(entity_rows[0]))
        w.writeheader(); w.writerows(entity_rows)
    with open(OUT / "incidents.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(incident_rows[0]))
        w.writeheader(); w.writerows(incident_rows)
    print(f"Wrote {len(entity_rows)} entities and {len(incident_rows)} incidents to {OUT}")


if __name__ == "__main__":
    main()
