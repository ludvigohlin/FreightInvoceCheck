"""One-off: two PostNord parser gaps found and fixed in postnord_parser.py:
1. "PostNord Home" shipment lines weren't recognized by _PARCEL_RE/_PARCEL_SHORT_RE
   (service-name alternation only had Parcel/Parcel Locker/Service Point).
2. "Kredittillägg" (credit surcharge) had no entry in _INV_SURCHARGES, so it was
   silently dropped from every invoice that had one.

Confirmed by re-parsing against the current (fixed) parser that invoices
903110324329, 903116037628, and 903121236629 now reconcile exactly (diff
0.00) — this removes their stale pre-fix rows so main.py reprocesses them
fresh on the next run.

NOT included: 903109695424, which also has a "PostNord Home" line but hits a
separate, rare, pre-existing ambiguity (a city/zone code ending in a digit,
"JÄRFÄLLA SE-A1", causes the amount regex to misread 158.00 as 1158.00) —
deliberately left untouched to avoid forcing a risky broad regex change for
a single-line edge case. Flagged to the user separately.

Safe to delete after running once.
"""
import csv
import shutil
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DASHBOARD_DATA = PROJECT_ROOT / "02_Output" / "Dashboard_Data"
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

INV_NUMS = {"903110324329", "903116037628", "903121236629"}


def _rewrite(path: Path, keep_row):
    backup = path.with_suffix(".csv.bak4")
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
    keep = lambda r: r.get("invoice_number") not in INV_NUMS
    _rewrite(DASHBOARD_DATA / "invoice_header.csv", keep)
    _rewrite(DASHBOARD_DATA / "invoice_lines.csv", keep)
    _rewrite(DASHBOARD_DATA / "surcharge_lines.csv", keep)
    _rewrite(DASHBOARD_DATA / "invoice_checks.csv", keep)
    _rewrite(DASHBOARD_DATA / "anomalies.csv", keep)
    print(f"\n{INV_NUMS} removed from all CSVs — run main.py to reprocess fresh.")


if __name__ == "__main__":
    main()
