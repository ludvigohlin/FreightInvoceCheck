"""One-off migration: add the new `shipment_date` column to invoice_lines.csv.

`INVOICE_LINES_FIELDS` in src/output_writer.py now includes shipment_date, so
all *future* runs emit it natively (from Bring's "Mottaget av Bring datum och
tid" per kolli row, and PostNord's "Inhämtningsdatum" per shipment section).
This script backfills the existing rows by re-parsing each invoice's original
source file (re-parsing is deterministic, so joining on (invoice_number,
line_no) reproduces the same rows):

- Bring: re-parses the Excel spec (found in 00_Inbox/, or in the historical
  bring6mo temp dump for the 12 pre-2026-06 SpecOnly invoices) and fills
  shipment_date from each kolli row's received_datetime.
- PostNord: re-parses the combined PDF (found in 00_Inbox/ via source_file
  from invoice_header.csv) and fills shipment_date from each shipment's
  pickup_date. Invoice-level surcharge rows (e.g. Drivmedelstillägg) have no
  natural shipment_date and are left blank — the dashboard falls back to
  invoice_date for those.

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
from src.postnord_parser import parse_postnord_pdf
from src.processing_logger import ProcessingLogger
from src.output_writer import INVOICE_LINES_FIELDS

CSV_LINES_PATH = PROJECT_ROOT / "02_Output" / "Dashboard_Data" / "invoice_lines.csv"
CSV_HEADER_PATH = PROJECT_ROOT / "02_Output" / "Dashboard_Data" / "invoice_header.csv"
INBOX = PROJECT_ROOT / "00_Inbox"
HISTORICAL_BRING_DIR = Path(r"C:\Users\LudvigOhlin\AppData\Local\Temp\bring6mo")


def find_bring_spec(invoice_number: str) -> Path | None:
    matches = list(INBOX.glob(f"*Specificeradfaktura*Fakturanummer_{invoice_number}.xlsx"))
    if matches:
        return matches[0]
    matches = list(HISTORICAL_BRING_DIR.glob(f"*Fakturanummer_{invoice_number}.xlsx"))
    return matches[0] if matches else None


def main():
    logger = ProcessingLogger(run_id="migrate_shipment_date")

    with open(CSV_HEADER_PATH, encoding="utf-8-sig", newline="") as f:
        headers = list(csv.DictReader(f, delimiter=";"))
    invoices = [(h["carrier"], h["invoice_number"], h["source_file"]) for h in headers]
    print(f"Found {len(invoices)} invoices in invoice_header.csv")

    date_lookup: dict[tuple[str, str], str] = {}
    missing_source = []

    for carrier, inv_num, source_file in invoices:
        if carrier == "Bring":
            spec_path = find_bring_spec(inv_num)
            if spec_path is None:
                missing_source.append((carrier, inv_num))
                continue
            _, lines = parse_bring_excel_specification(
                spec_path, run_id="migrate_shipment_date", logger=logger
            )
            n_filled = 0
            for ln in lines:
                if ln.received_datetime:
                    date_lookup[(inv_num, str(ln.line_no))] = ln.received_datetime[:10]
                    n_filled += 1
            print(f"  Bring {inv_num}: {len(lines)} lines from {spec_path.name}, {n_filled} with shipment_date")

        elif carrier == "PostNord":
            pdf_path = INBOX / source_file
            if not pdf_path.exists():
                missing_source.append((carrier, inv_num))
                continue
            _, lines = parse_postnord_pdf(pdf_path, run_id="migrate_shipment_date", logger=logger)
            n_filled = 0
            for ln in lines:
                if ln.pickup_date:
                    date_lookup[(inv_num, str(ln.line_no))] = ln.pickup_date
                    n_filled += 1
            print(f"  PostNord {inv_num}: {len(lines)} lines from {pdf_path.name}, {n_filled} with shipment_date")

    if missing_source:
        print(f"\nWARNING: no source file found for {len(missing_source)} invoice(s), shipment_date left blank:")
        for carrier, inv_num in missing_source:
            print(f"  {carrier} {inv_num}")

    backup_path = CSV_LINES_PATH.with_suffix(".csv.bak")
    shutil.copy2(CSV_LINES_PATH, backup_path)
    print(f"\nBackup written to {backup_path}")

    with open(CSV_LINES_PATH, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f, delimiter=";"))
    print(f"Read {len(rows)} existing rows")

    filled = 0
    for row in rows:
        key = (row.get("invoice_number", ""), row.get("line_no", ""))
        shipment_date = date_lookup.get(key, "")
        row["shipment_date"] = shipment_date
        if shipment_date:
            filled += 1

    with open(CSV_LINES_PATH, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=INVOICE_LINES_FIELDS, delimiter=";", extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)

    print(f"\nRows with shipment_date filled: {filled} / {len(rows)}")


if __name__ == "__main__":
    main()
