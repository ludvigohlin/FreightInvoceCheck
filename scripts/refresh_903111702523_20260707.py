"""One-off: invoice 903111702523 was first processed on 2026-06-06, before a
parser fix (dual 'Drivmedelstillägg Paket' lines on page 1 were only captured
once via re.search()). Re-parsing today with the current parser gives an
exact match (line_sum == header total_ex_vat, diff 0.00), but the stale
2026-06-06 rows (with 3 fewer lines than a correct parse) are still sitting in
every CSV, showing a stale "Error" status.

Rather than just patching the status field, this removes all rows tied to
this invoice_number so the pipeline treats it as unprocessed and reprocesses
it fresh on the next main.py run — the source PDF is still in 00_Inbox — so
line/surcharge/check/anomaly data is fully correct, not just the header status.

Safe to delete after running once.
"""
import csv
import shutil
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DASHBOARD_DATA = PROJECT_ROOT / "02_Output" / "Dashboard_Data"
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

INV_NUM = "903111702523"


def _rewrite(path: Path, keep_row):
    backup = path.with_suffix(".csv.bak3")
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
    print(f"\n{INV_NUM} removed from all CSVs — run main.py to reprocess it fresh.")


if __name__ == "__main__":
    main()
