"""One-off migration: add the new `chargeable_weight_kg` column to invoice_lines.csv.

`INVOICE_LINES_FIELDS` in src/output_writer.py now includes chargeable_weight_kg,
so all *future* runs emit it natively:
- Bring: same value as weight_kg (Bring bills by actual weight, no separate
  volumetric figure) — no re-parse needed, just copied from the existing column.
- PostNord: from `fraktdr_vikt`, the third weight column on Parcel/Service Point
  shipment lines (not currently persisted) — requires re-parsing the source
  PDFs, same approach as migrate_shipment_date_20260708.py.

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

CSV_LINES_PATH = PROJECT_ROOT / "02_Output" / "Dashboard_Data" / "invoice_lines.csv"
CSV_HEADER_PATH = PROJECT_ROOT / "02_Output" / "Dashboard_Data" / "invoice_header.csv"
INBOX = PROJECT_ROOT / "00_Inbox"


def main():
    logger = ProcessingLogger(run_id="migrate_chargeable_weight")

    with open(CSV_HEADER_PATH, encoding="utf-8-sig", newline="") as f:
        headers = list(csv.DictReader(f, delimiter=";"))
    postnord_invoices = [(h["invoice_number"], h["source_file"]) for h in headers if h["carrier"] == "PostNord"]
    print(f"Found {len(postnord_invoices)} PostNord invoices in invoice_header.csv")

    pn_lookup: dict[tuple[str, str], float] = {}
    missing_source = []
    for inv_num, source_file in postnord_invoices:
        pdf_path = INBOX / source_file
        if not pdf_path.exists():
            missing_source.append(inv_num)
            continue
        _, lines = parse_postnord_pdf(pdf_path, run_id="migrate_chargeable_weight", logger=logger)
        n_filled = 0
        for ln in lines:
            if ln.fraktdr_vikt is not None:
                pn_lookup[(inv_num, str(ln.line_no))] = ln.fraktdr_vikt
                n_filled += 1
        print(f"  PostNord {inv_num}: {len(lines)} lines from {pdf_path.name}, {n_filled} with fraktdr_vikt")

    if missing_source:
        print(f"\nWARNING: no source PDF found for {len(missing_source)} PostNord invoice(s):")
        for inv_num in missing_source:
            print(f"  {inv_num}")

    backup_path = CSV_LINES_PATH.with_suffix(".csv.bak2")
    shutil.copy2(CSV_LINES_PATH, backup_path)
    print(f"\nBackup written to {backup_path}")

    with open(CSV_LINES_PATH, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f, delimiter=";"))
    print(f"Read {len(rows)} existing rows")

    filled_bring, filled_postnord = 0, 0
    for row in rows:
        if row.get("carrier") == "Bring":
            wk = row.get("weight_kg", "")
            row["chargeable_weight_kg"] = wk
            if wk:
                filled_bring += 1
        else:
            key = (row.get("invoice_number", ""), row.get("line_no", ""))
            val = pn_lookup.get(key, "")
            row["chargeable_weight_kg"] = val
            if val != "":
                filled_postnord += 1

    with open(CSV_LINES_PATH, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=INVOICE_LINES_FIELDS, delimiter=";", extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)

    print(f"\nBring rows filled: {filled_bring}")
    print(f"PostNord rows filled: {filled_postnord}")
    print(f"Total rows written: {len(rows)}")


if __name__ == "__main__":
    main()
