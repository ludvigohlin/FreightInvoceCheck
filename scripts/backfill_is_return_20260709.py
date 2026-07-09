"""One-off migration: add the new `is_return` column to invoice_lines.csv.

`INVOICE_LINES_FIELDS` in src/output_writer.py now includes is_return, so all
*future* runs emit it natively for PostNord (see postnord_parser.py — a
BaseFreight line addressed back to our own warehouse postal code is a return
leg). This script backfills every existing PostNord invoice already in
invoice_header.csv by re-parsing its original source PDF (using the
source_file column to locate it — some are still in 00_Inbox/, the rest are
in the "Tillfälligt" folder where the 2026-07-08 historical backfill sourced
them from) and matching on (invoice_number, line_no), which is a safe join
key since re-parsing the same file reproduces the same line_no sequence
deterministically.

Bring rows are left untouched — Bring already flags returns as an
"Attempted Delivery Return" surcharge line (surcharge_category == "Return"
in surcharge_lines.csv), not a BaseFreight-line concept, so is_return stays
blank for Bring's invoice_lines.csv rows by design.

Safe to delete after running once.
"""
import csv
import shutil
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from src.postnord_parser import parse_postnord_pdf
from src.processing_logger import ProcessingLogger
from src.output_writer import INVOICE_LINES_FIELDS

CSV_PATH = PROJECT_ROOT / "02_Output" / "Dashboard_Data" / "invoice_lines.csv"
HEADER_CSV_PATH = PROJECT_ROOT / "02_Output" / "Dashboard_Data" / "invoice_header.csv"
INBOX = PROJECT_ROOT / "00_Inbox"
TILLFALLIGT = Path(r"C:\Users\LudvigOhlin\OneDrive - Isicom AB\Tillfälligt")


def find_source_pdf(source_file: str) -> Path | None:
    for base in (INBOX, TILLFALLIGT):
        candidate = base / source_file
        if candidate.exists():
            return candidate
    return None


def main():
    with open(HEADER_CSV_PATH, encoding="utf-8-sig", newline="") as f:
        headers = [r for r in csv.DictReader(f, delimiter=";") if r["carrier"] == "PostNord"]
    print(f"Found {len(headers)} PostNord invoice header(s)")

    logger = ProcessingLogger(run_id="backfill_is_return")

    # Build (invoice_number, line_no) -> is_return lookup from fresh parses
    return_lookup: dict[tuple[str, str], bool] = {}
    parsed, missing, failed = 0, [], []
    for h in headers:
        inv_num = h["invoice_number"]
        pdf_path = find_source_pdf(h["source_file"])
        if pdf_path is None:
            missing.append((inv_num, h["source_file"]))
            continue
        try:
            _, lines = parse_postnord_pdf(pdf_path, run_id="backfill_is_return", logger=logger)
        except Exception as e:
            failed.append((inv_num, str(e)))
            continue
        base_lines = [ln for ln in lines if ln.line_type == "BaseFreight"]
        for ln in base_lines:
            return_lookup[(inv_num, str(ln.line_no))] = ln.is_return
        parsed += 1

    print(f"Re-parsed {parsed} invoice(s)")
    if missing:
        print(f"WARNING: source PDF not found for {len(missing)} invoice(s):")
        for inv_num, src in missing:
            print(f"  {inv_num}: {src}")
    if failed:
        print(f"WARNING: re-parse failed for {len(failed)} invoice(s):")
        for inv_num, err in failed:
            print(f"  {inv_num}: {err}")

    backup_path = CSV_PATH.with_suffix(".csv.bak")
    shutil.copy2(CSV_PATH, backup_path)
    print(f"Backup written to {backup_path}")

    with open(CSV_PATH, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f, delimiter=";"))
    print(f"Read {len(rows)} existing rows")

    filled = 0
    for row in rows:
        key = (row.get("invoice_number", ""), row.get("line_no", ""))
        if key in return_lookup:
            row["is_return"] = return_lookup[key]
            filled += 1
        else:
            row.setdefault("is_return", "")

    with open(CSV_PATH, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=INVOICE_LINES_FIELDS, delimiter=";", extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)

    return_count = sum(1 for row in rows if str(row.get("is_return", "")) == "True")
    print(f"\nRows with is_return filled: {filled}")
    print(f"Rows flagged as returns (is_return=True): {return_count}")
    print(f"Total rows written: {len(rows)} (unchanged count expected)")


if __name__ == "__main__":
    main()
