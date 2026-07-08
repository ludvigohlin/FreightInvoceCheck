"""One-off: invoice 903121050525 has 2 shipments to Spain (ES-28007) whose
kolli IDs carry a trailing letter suffix (e.g. "03015167848844C") — the
postal-code regex only recognized FI-/DK-/NO-/5-digit codes, and the kolli-id
regex required pure digits, so both lines fell through to a garbled "Unknown"
surcharge instead of being captured as proper Parcel shipments. Fixed in
postnord_parser.py:
  - postal code group generalized to any "XX-digits" country prefix (not just
    FI-/DK-/NO-)
  - kolli-id group now allows an optional trailing letter
  - infer_country_from_postal_code() (utils.py) now extracts the actual
    2-letter prefix (e.g. "ES") instead of bucketing every non-Nordic code as
    "Unknown" — the existing non-Nordic destination flagging
    (detect_non_nordic_destinations) already treats anything outside
    SE/NO/DK/FI as foreign, so this makes the flag show a real country code.

Verified via full regression test across all invoices in 00_Inbox: this
invoice still reconciles exactly (diff 0.00), the 2 Spain shipments are now
correctly captured and flagged as a non-Nordic destination anomaly, and no
other invoice's total changed.

Safe to delete after running once.
"""
import csv
import shutil
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DASHBOARD_DATA = PROJECT_ROOT / "02_Output" / "Dashboard_Data"
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

INV_NUM = "903121050525"


def _rewrite(path: Path, keep_row):
    backup = path.with_suffix(".csv.bak6")
    shutil.copy2(path, backup)
    with open(path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        fieldnames = reader.fieldnames
        rows = list(reader)
    kept = [r for r in rows if keep_row(r)]
    removed = len(rows) - len(kept)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=";")
        writer.writeheader()
        writer.writerows(kept)
    print(f"{path.name}: removed {removed} row(s), {len(kept)} remain (backup: {backup.name})")


def main():
    keep = lambda r: r.get("invoice_number") != INV_NUM
    _rewrite(DASHBOARD_DATA / "invoice_header.csv", keep)
    _rewrite(DASHBOARD_DATA / "invoice_lines.csv", keep)
    _rewrite(DASHBOARD_DATA / "surcharge_lines.csv", keep)
    _rewrite(DASHBOARD_DATA / "invoice_checks.csv", keep)
    _rewrite(DASHBOARD_DATA / "anomalies.csv", keep)
    print(f"\n{INV_NUM} removed from all CSVs — run main.py to reprocess fresh.")


if __name__ == "__main__":
    main()
