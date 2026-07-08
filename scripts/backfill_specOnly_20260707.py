"""One-off backfill of historical Bring invoices that only have an Excel
specification (no PDF — pulled from the Mybring mailbox, PDFs will never
arrive since they're 2-6 months old). Uses the already-implemented
merge_bring_headers(None, excel_header) fallback, which main.py's live
pipeline never exercises (it requires both PDF+Excel and skips otherwise).

Deliberately skips anomaly/check detection (PDF-vs-Excel checks are
meaningless with no PDF) and does not touch main.py's live gating or the
"pending files" alert logic at all.

Safe to delete after running once. Idempotent — re-running skips invoices
already present in invoice_header.csv.
"""
import sys
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from dateutil.relativedelta import relativedelta
from dateutil.parser import isoparse

from src import config
from src.bring_parser import parse_bring_excel_specification
from src.normalization import merge_bring_headers
from src.processing_logger import ProcessingLogger
from src.output_writer import (
    get_existing_invoice_keys,
    write_invoice_headers,
    write_invoice_lines,
    write_surcharge_lines,
)

SPEC_DIR = Path(r"C:\Users\LudvigOhlin\AppData\Local\Temp\bring6mo")


def main():
    logger = ProcessingLogger(run_id="backfill_specOnly")
    service_mapping = config.load_service_mapping()
    surcharge_mapping = config.load_surcharge_mapping()

    existing_keys = get_existing_invoice_keys()
    files = sorted(SPEC_DIR.glob("*.xlsx"))
    print(f"Found {len(files)} candidate files in {SPEC_DIR}")

    processed, skipped = [], []

    for path in files:
        excel_header, lines = parse_bring_excel_specification(
            path,
            run_id=f"pending_lookup_{path.stem}",
            logger=logger,
            service_mapping=service_mapping,
            surcharge_mapping=surcharge_mapping,
        )
        if excel_header is None or not excel_header.invoice_number:
            print(f"SKIP {path.name}: could not parse a header/invoice number")
            skipped.append(path.name)
            continue

        inv_num = excel_header.invoice_number
        key = ("Bring", inv_num)
        if key in existing_keys:
            print(f"SKIP {path.name}: invoice {inv_num} already in invoice_header.csv")
            skipped.append(path.name)
            continue

        if not excel_header.invoice_date:
            print(f"SKIP {path.name}: invoice {inv_num} has no invoice_date, cannot backfill safely")
            skipped.append(path.name)
            continue

        # Re-parse with the real run_id used for this backfill entry
        run_id = f"backfill_{inv_num}"
        excel_header, lines = parse_bring_excel_specification(
            path, run_id=run_id, logger=logger,
            service_mapping=service_mapping, surcharge_mapping=surcharge_mapping,
        )

        merged = merge_bring_headers(None, excel_header)
        merged.currency = "SEK"  # confirmed from live Bring PDF text ("Valuta SEK")
        merged.reconciliation_status = "SpecOnly"
        merged.processed_timestamp = datetime.now().isoformat(timespec="seconds")

        inv_date = isoparse(merged.invoice_date)
        merged.due_date = (inv_date + relativedelta(months=1)).strftime("%Y-%m-%d")

        for ln in lines:
            ln.run_id = run_id
            ln.processed_timestamp = merged.processed_timestamp

        written = write_invoice_headers([merged], logger, skip_keys=existing_keys)
        write_invoice_lines(lines, logger, skip_keys=existing_keys)
        write_surcharge_lines(lines, logger, skip_keys=existing_keys)

        if written:
            existing_keys.add(key)
            processed.append((inv_num, merged.invoice_date, merged.due_date, len(lines)))
            print(f"OK   {path.name}: invoice {inv_num}, date={merged.invoice_date}, "
                  f"due={merged.due_date}, {len(lines)} lines")
        else:
            skipped.append(path.name)

    print(f"\nProcessed: {len(processed)}  Skipped: {len(skipped)}")
    for inv_num, idate, ddate, n in processed:
        print(f"  {inv_num}  {idate} -> {ddate}  ({n} lines)")


if __name__ == "__main__":
    main()
