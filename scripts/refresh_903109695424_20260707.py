"""One-off: invoice 903109695424 had a 'PostNord Home' shipment line whose city
field ("JÄRFÄLLA SE-A1" — a city name with a trailing PostNord zone code) made
the amount regex misread 158.00 as 1158.00 (the zone code's digit was
swallowed into the amount instead of the city). Fixed in postnord_parser.py
by allowing the city group to optionally capture a trailing "XX-X9"-style
zone code. Verified via full regression test across all invoices in
00_Inbox: this invoice now reconciles exactly (diff 0.00) and no other
invoice's total changed.

Safe to delete after running once.
"""
import csv
import shutil
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DASHBOARD_DATA = PROJECT_ROOT / "02_Output" / "Dashboard_Data"
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

INV_NUM = "903109695424"


def _rewrite(path: Path, keep_row):
    backup = path.with_suffix(".csv.bak5")
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
