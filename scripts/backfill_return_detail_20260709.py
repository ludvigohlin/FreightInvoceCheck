"""One-off migration: backfill `to_postal`/`to_city` (new columns) for all
PostNord invoice_lines.csv rows, and repair `service_code` (kolli_id) values
that were corrupted into float scientific-notation at some point in the past
(e.g. "5.7313253152586726e+17" instead of "00573132901718387124").

The corruption predates this script and is NOT a bug in postnord_parser.py —
freshly-parsed rows already carry the correct string kolli_id (verified
against the most recent invoice, 903121236629, before writing this script).
Every historical .bak/.bak2../.bak6 snapshot already shows the same
corrupted value for the same row, meaning it happened once, externally
(almost certainly the CSV being opened and re-saved in Excel, which
auto-converts long digit-only strings to numbers), and was then carried
forward untouched since old rows are never rewritten. This script repairs
it by re-parsing every PostNord invoice's original source PDF (same
source_file/(invoice_number, line_no) join technique as the other
backfill scripts in this folder) and overwriting service_code/to_postal/
to_city with values read straight from the PDF.

Needed for the returns-detail dashboard feature: knowing where a returned
kolli was originally headed, and being able to reliably group by kolli_id
to detect a shipment that's been returned more than once, both require
`service_code` to actually be the kolli_id (not a mangled float) and the
destination fields to exist at all.

Bring rows are left untouched — Bring's return signal is a flat surcharge
fee (surcharge_category=="Return"), not a BaseFreight-level concept, so
there's no equivalent kolli-chain to reconstruct for it.

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

    logger = ProcessingLogger(run_id="backfill_return_detail")

    # (invoice_number, line_no) -> (kolli_id, to_postal, to_city)
    lookup: dict[tuple[str, str], tuple[str, str, str]] = {}
    parsed, missing, failed = 0, [], []
    for h in headers:
        inv_num = h["invoice_number"]
        pdf_path = find_source_pdf(h["source_file"])
        if pdf_path is None:
            missing.append((inv_num, h["source_file"]))
            continue
        try:
            _, lines = parse_postnord_pdf(pdf_path, run_id="backfill_return_detail", logger=logger)
        except Exception as e:
            failed.append((inv_num, str(e)))
            continue
        for ln in lines:
            if ln.line_type != "BaseFreight":
                continue
            lookup[(inv_num, str(ln.line_no))] = (ln.kolli_id, ln.to_postal, ln.to_city)
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

    backup_path = CSV_PATH.with_suffix(".csv.bak_returndetail")
    shutil.copy2(CSV_PATH, backup_path)
    print(f"Backup written to {backup_path}")

    with open(CSV_PATH, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f, delimiter=";"))
    print(f"Read {len(rows)} existing rows")

    fixed_kolli, filled_dest = 0, 0
    for row in rows:
        key = (row.get("invoice_number", ""), row.get("line_no", ""))
        hit = lookup.get(key)
        if hit:
            kolli_id, to_postal, to_city = hit
            if row.get("service_code") != kolli_id:
                row["service_code"] = kolli_id
                fixed_kolli += 1
            row["to_postal"] = to_postal
            row["to_city"] = to_city
            filled_dest += 1
        else:
            row.setdefault("to_postal", "")
            row.setdefault("to_city", "")

    with open(CSV_PATH, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=INVOICE_LINES_FIELDS, delimiter=";", extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)

    print(f"\nservice_code repaired on {fixed_kolli} row(s)")
    print(f"to_postal/to_city filled on {filled_dest} row(s)")
    print(f"Total rows written: {len(rows)} (unchanged count expected)")


if __name__ == "__main__":
    main()
