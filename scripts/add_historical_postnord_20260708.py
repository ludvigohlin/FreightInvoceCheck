"""One-off backfill: add historical PostNord PDF invoices (Jan-May 2026,
supplied by the user from an external "Tillfällligt" folder, not 00_Inbox).

Unlike the Bring spec-only backfill, PostNord's PDF is the combined
invoice+specification, so these get the *same* treatment as a live run:
full parse, full validation checks (run_postnord_checks), and
reconciliation_status derived from the LineSumVsHeaderTotal check — exactly
matching main.py's _apply_reconciliation_status logic.

Deliberately skipped (out of scope for a historical backfill, and these
require cross-invoice context that's easiest to get right in a live run):
duplicate-shipment detection, non-Nordic destination flagging, price-increase-
vs-history anomaly detection.

Safe to delete after running once.
"""
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from src import config
from src.postnord_parser import parse_postnord_pdf
from src.validation import run_postnord_checks
from src.processing_logger import ProcessingLogger
from src.output_writer import (
    get_existing_invoice_keys, write_invoice_headers, write_invoice_lines,
    write_surcharge_lines, write_invoice_checks,
)

SRC_DIR = Path(r"C:\Users\LudvigOhlin\OneDrive - Isicom AB\Tillfälligt")

FILES = [
    "invoice-20335260-903103539321.pdf",
    "invoice-20335260-940000138328.pdf",
    "invoice-20335260-903102776122.pdf",
    "invoice-20335260-903099339827.pdf",
    "invoice-20828950-903099401726.pdf",
    "invoice-20828950-903100702328.pdf",
    "invoice-20335260-903096568428.pdf",
    "invoice-20335260-903096773028.pdf",
    "invoice-20828950-903096692624.pdf",
    "invoice-20335260-903095169129.pdf",
    "invoice-20335260-940000104023.pdf",
    "invoice-20335260-903093614027.pdf",
    "invoice-20335260-903092477129.pdf",
    "invoice-20828950-903092528822.pdf",
    "invoice-20335260-903092287122.pdf",
    "invoice-20335260-903089314624.pdf",
    "invoice-20335260-903087332925.pdf",
    "invoice-20335260-903085846520.pdf",
    "invoice-20335260-903084306625.pdf",
    "invoice-20828950-903084535926.pdf",
    "invoice-20335260-903082800520.pdf",
    "invoice-20335260-903080599520.pdf",
    "invoice-20828950-903081236924.pdf",
    "invoice-20335260-903076389829.pdf",
    "invoice-20335260-903074892022 (1).pdf",
    "invoice-20335260-903073377520.pdf",
    "invoice-20828950-903073549722.pdf",
    "invoice-20335260-903070669424.pdf",
    "invoice-20828950-903069985328.pdf",
    "invoice-20828950-903071280528.pdf",
    "invoice-20335260-903066967428.pdf",
    "invoice-20335260-903065489820.pdf",
    "invoice-20335260-940000008224 (1).pdf",
    "invoice-20335260-903064064426 (1).pdf",
    "invoice-20335260-903062518324.pdf",
    "invoice-20828950-903063061829.pdf",
]


def main():
    logger = ProcessingLogger(run_id="add_historical_postnord")
    service_mapping = config.load_service_mapping()
    surcharge_mapping = config.load_surcharge_mapping()
    existing_keys = get_existing_invoice_keys()

    processed, skipped, failed = [], [], []

    for fname in FILES:
        path = SRC_DIR / fname
        if not path.exists():
            print(f"MISSING FILE: {fname}")
            failed.append((fname, "file not found"))
            continue

        header, lines = parse_postnord_pdf(
            path, run_id="probe", logger=logger,
            service_mapping=service_mapping, surcharge_mapping=surcharge_mapping,
        )
        if header is None or not header.invoice_number:
            print(f"FAILED to parse invoice number from {fname}")
            failed.append((fname, "no invoice number"))
            continue

        inv_num = header.invoice_number
        key = ("PostNord", inv_num)
        if key in existing_keys:
            print(f"SKIP {fname}: invoice {inv_num} already exists")
            skipped.append((fname, inv_num))
            continue

        run_id = f"histpn_{inv_num}"
        header, lines = parse_postnord_pdf(
            path, run_id=run_id, logger=logger,
            service_mapping=service_mapping, surcharge_mapping=surcharge_mapping,
        )

        checks = run_postnord_checks(run_id, header, lines, logger)
        recon = next((c.status for c in checks if c.check_name == "LineSumVsHeaderTotal"), None)
        header.reconciliation_status = recon or "NotChecked"

        written = write_invoice_headers([header], logger, skip_keys=existing_keys)
        write_invoice_lines(lines, logger, skip_keys=existing_keys)
        write_surcharge_lines(lines, logger, skip_keys=existing_keys)
        write_invoice_checks(checks, logger)

        if written:
            existing_keys.add(key)
            processed.append((fname, inv_num, header.invoice_date, header.total_ex_vat,
                               header.reconciliation_status, len(lines)))
            print(f"OK {fname}: {inv_num}, date={header.invoice_date}, "
                  f"total={header.total_ex_vat}, status={header.reconciliation_status}, {len(lines)} lines")
        else:
            failed.append((fname, "write returned no keys"))

    print(f"\nProcessed: {len(processed)}  Skipped (dup): {len(skipped)}  Failed: {len(failed)}")
    for fname, inv_num, date, total, status, n in sorted(processed, key=lambda x: x[2] or ""):
        print(f"  {date}  {inv_num}  {total:>10.2f} SEK  {status:<10}  {n} lines")
    if failed:
        print("\nFailed:")
        for fname, reason in failed:
            print(f"  {fname}: {reason}")


if __name__ == "__main__":
    main()
