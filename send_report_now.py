"""
One-shot report generator.
Parses the specified invoices fresh from 00_Inbox, generates the Excel,
and sends the email. Does NOT write to any output CSVs.

Run with the AppData venv to avoid OneDrive DLL loading issues:
  AppData\\Local\\FreightInvoiceControl\\venv\\Scripts\\python.exe send_report_now.py
"""

from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from src import config
from src.utils import generate_run_id
from src.processing_logger import ProcessingLogger
from src.postnord_parser import parse_postnord_pdf
from src.bring_parser import parse_bring_pdf_header, parse_bring_excel_specification
from src.normalization import merge_bring_headers
from src.validation import run_all_checks
from src.anomaly_detection import detect_bring_anomalies, detect_non_nordic_destinations
from src.summary_writer import build_summary_payload, write_deterministic_summary
from src.run_exporter import write_run_export
from src.email_sender import send_summary_email
from src.output_writer import get_existing_invoice_keys

# ── Which PostNord invoices to include ───────────────────────────────────────
POSTNORD_INVOICE_NUMBERS = ["903111702523"]

run_id = generate_run_id()
logger = ProcessingLogger(run_id)
config.ensure_all_directories()
config.FOR_EMAIL_DIR.mkdir(parents=True, exist_ok=True)

service_mapping  = config.load_service_mapping()
surcharge_mapping = config.load_surcharge_mapping()

inbox = config.INBOX_DIR
all_pdf_headers: dict  = {}
all_lines_dict:  dict  = {}
all_invoice_headers    = []
all_invoice_lines      = []
all_anomalies          = []

# ── Parse the requested PostNord invoices ────────────────────────────────────
for inv_num in POSTNORD_INVOICE_NUMBERS:
    match = next(
        (f for f in inbox.iterdir()
         if f.is_file() and inv_num in f.name and f.suffix.lower() == ".pdf"),
        None,
    )
    if not match:
        print(f"[WARN] No PDF found in 00_Inbox for PostNord {inv_num}")
        continue

    print(f"Parsing: {match.name}")
    h, lines = parse_postnord_pdf(
        match, run_id, logger,
        service_mapping=service_mapping,
        surcharge_mapping=surcharge_mapping,
    )
    if not h:
        print(f"[ERROR] Failed to parse {match.name}")
        continue

    h.invoice_number = h.invoice_number or inv_num
    all_pdf_headers[("PostNord", h.invoice_number)] = h
    all_lines_dict[("PostNord", h.invoice_number)]  = lines
    all_invoice_headers.append(h)
    all_invoice_lines.extend(lines)
    all_anomalies.extend(
        detect_non_nordic_destinations("PostNord", h.invoice_number, lines, logger)
    )

# ── Detect pending Bring invoices (skip any already processed) ───────────────
existing_keys = get_existing_invoice_keys()

inbox_files = [f for f in inbox.iterdir() if f.is_file()]
bring_pdfs  = {f for f in inbox_files if "Faktura" in f.name and f.suffix.lower() == ".pdf"
               and "FAKTURA_9" not in f.name}
bring_xls   = {f for f in inbox_files if f.suffix.lower() in (".xlsx", ".xls")}

bring_all_excel_headers: dict = {}
bring_all_pdf_headers:  dict  = {}
for xls_f in bring_xls:
    xh, _ = parse_bring_excel_specification(
        xls_f, run_id, logger,
        service_mapping=service_mapping,
        surcharge_mapping=surcharge_mapping,
    )
    if xh:
        bring_all_excel_headers[("Bring", xh.invoice_number)] = xh

for pdf_f in bring_pdfs:
    ph = parse_bring_pdf_header(pdf_f, run_id, logger)
    if ph:
        bring_all_pdf_headers[("Bring", ph.invoice_number)] = ph

bring_invoice_numbers = set(
    inv for (c, inv) in list(bring_all_pdf_headers.keys()) + list(bring_all_excel_headers.keys())
    if c == "Bring"
)

missing_bring = []
all_excel_headers: dict = {}

for inv_num in bring_invoice_numbers:
    if ("Bring", inv_num) in existing_keys:
        print(f"[SKIP] Bring {inv_num}: already processed — excluded from report")
        continue

    has_pdf = ("Bring", inv_num) in bring_all_pdf_headers
    has_xls = ("Bring", inv_num) in bring_all_excel_headers

    if has_pdf and has_xls:
        pdf_h = bring_all_pdf_headers[("Bring", inv_num)]
        xls_h = bring_all_excel_headers[("Bring", inv_num)]
        _, bring_lines = parse_bring_excel_specification(
            next(f for f in bring_xls if inv_num in f.name),
            run_id, logger,
            service_mapping=service_mapping,
            surcharge_mapping=surcharge_mapping,
        )
        try:
            merged = merge_bring_headers(pdf_h, xls_h)
            all_invoice_headers.append(merged)
            all_invoice_lines.extend(bring_lines)
            all_pdf_headers[("Bring", inv_num)]   = pdf_h
            all_excel_headers[("Bring", inv_num)] = xls_h
            all_lines_dict[("Bring", inv_num)]    = bring_lines
            all_anomalies.extend(detect_bring_anomalies(pdf_h, xls_h, bring_lines, logger))
        except ValueError:
            pass
    elif has_pdf and not has_xls:
        ph = bring_all_pdf_headers[("Bring", inv_num)]
        missing_bring.append({
            "invoice_number": inv_num,
            "missing_file":   "Excel specification (.xlsx)",
            "found_file":     "PDF invoice",
            "known_total_ex_vat": ph.total_ex_vat if ph else "",
            "source_file":    ph.source_file if ph else "",
            "message":        "PDF invoice received but Excel specification is missing.",
        })
        print(f"[INFO] Bring {inv_num}: PDF found, Excel missing — shown as pending")
    elif has_xls and not has_pdf:
        xh = bring_all_excel_headers[("Bring", inv_num)]
        missing_bring.append({
            "invoice_number": inv_num,
            "missing_file":   "PDF invoice",
            "found_file":     "Excel specification (.xlsx)",
            "known_total_ex_vat": xh.total_ex_vat if xh else "",
            "source_file":    xh.source_file if xh else "",
            "message":        "Excel specification received but PDF invoice is missing.",
        })
        print(f"[INFO] Bring {inv_num}: Excel found, PDF missing — shown as pending")

if not all_invoice_headers and not missing_bring:
    print("Nothing to report. Exiting.")
    sys.exit(0)

# ── Validation ───────────────────────────────────────────────────────────────
all_checks = run_all_checks(run_id, all_pdf_headers, all_excel_headers, all_lines_dict, logger)

# ── Build and write report ───────────────────────────────────────────────────
payload = build_summary_payload(
    run_id          = run_id,
    scan_timestamp  = datetime.now().isoformat(timespec="seconds"),
    file_records    = [],
    headers         = all_invoice_headers,
    lines           = all_invoice_lines,
    checks          = all_checks,
    anomalies       = all_anomalies,
    log_counts      = logger.get_counts(),
)

det_path = write_deterministic_summary(
    run_id, payload, [], all_invoice_headers,
    all_invoice_lines, all_checks, all_anomalies, logger,
)

write_run_export(
    run_id, payload,
    all_invoice_headers, all_invoice_lines,
    all_checks, logger,
    ai_summary    = None,
    anomalies     = all_anomalies,
    missing_bring = missing_bring,
    all_lines_dict= all_lines_dict,
)

# ── Send email ───────────────────────────────────────────────────────────────
xlsx_path = config.FOR_EMAIL_DIR / f"summary_{run_id}.xlsx"
ok = send_summary_email(
    run_id       = run_id,
    summary_md_path = det_path,
    xlsx_path    = xlsx_path if xlsx_path.exists() else None,
    logger       = logger,
    check_counts = payload.get("check_counts", {}),
    total_amount = sum(ln.amount or 0.0 for ln in all_invoice_lines),
)

print(f"\nReport: {xlsx_path}")
print(f"Email {'sent' if ok else 'FAILED — check Outlook is open'}")
