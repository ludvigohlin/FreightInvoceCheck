"""One-off migration: add the new `weight_kg` column to invoice_lines.csv.

`INVOICE_LINES_FIELDS` in src/output_writer.py now includes weight_kg, so all
*future* runs emit it natively. This script backfills the existing file:
- For the 5 Bring invoices already in invoice_header.csv, re-parses their
  original Excel specs (still sitting in 00_Inbox/) and fills weight_kg by
  matching (invoice_number, line_no) — a safe join key since re-parsing the
  same file reproduces the same line_no sequence deterministically.
- All other existing rows (PostNord, anything else) get an empty weight_kg —
  that data was never captured for those runs and can't be recovered.

Safe to delete after running once.
"""
import csv
import shutil
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from src.bring_parser import parse_bring_excel_specification
from src.processing_logger import ProcessingLogger
from src.output_writer import INVOICE_LINES_FIELDS

CSV_PATH = PROJECT_ROOT / "02_Output" / "Dashboard_Data" / "invoice_lines.csv"
INBOX = PROJECT_ROOT / "00_Inbox"

# invoice_number -> source .xlsx filename (glob pattern)
BRING_INVOICES = [
    "4040266117",
    "4040267896",
    "4040271101",
    "4040271782",
    "4040273322",
]


def find_spec_file(invoice_number: str) -> Path | None:
    matches = list(INBOX.glob(f"*Specificeradfaktura*Fakturanummer_{invoice_number}.xlsx"))
    return matches[0] if matches else None


def main():
    backup_path = CSV_PATH.with_suffix(".csv.bak")
    shutil.copy2(CSV_PATH, backup_path)
    print(f"Backup written to {backup_path}")

    with open(CSV_PATH, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        rows = list(reader)
    print(f"Read {len(rows)} existing rows")

    logger = ProcessingLogger(run_id="migrate_weight_kg")

    # Build (invoice_number, line_no) -> weight_kg lookup from fresh parses
    weight_lookup: dict[tuple[str, str], float] = {}
    for inv_num in BRING_INVOICES:
        spec_path = find_spec_file(inv_num)
        if spec_path is None:
            print(f"WARNING: no source Excel found for invoice {inv_num}, skipping")
            continue
        _, lines = parse_bring_excel_specification(
            spec_path, run_id="migrate_weight_kg", logger=logger
        )
        for ln in lines:
            if ln.weight_kg is not None:
                weight_lookup[(inv_num, str(ln.line_no))] = ln.weight_kg
        print(f"  {inv_num}: parsed {len(lines)} lines from {spec_path.name}, "
              f"{sum(1 for l in lines if l.weight_kg is not None)} with weight")

    filled = 0
    for row in rows:
        key = (row.get("invoice_number", ""), row.get("line_no", ""))
        if key in weight_lookup:
            row["weight_kg"] = weight_lookup[key]
            filled += 1
        else:
            row.setdefault("weight_kg", "")

    with open(CSV_PATH, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=INVOICE_LINES_FIELDS, delimiter=";", extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)

    print(f"\nRows with weight_kg filled: {filled}")
    print(f"Total rows written: {len(rows)} (unchanged count expected)")


if __name__ == "__main__":
    main()
