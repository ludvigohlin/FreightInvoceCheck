"""
Freight Invoice Control — MVP Entry Point

Usage:
    python main.py [--dry-run] [--use-claude] [--move-files]

Workflow:
    1. Scan 00_Inbox for supported files
    2. Classify each file (carrier, document type, invoice number)
    3. Parse Bring PDF invoices and Excel specifications
    4. Run reconciliation and validation checks
    5. Detect anomalies
    6. Write normalized output CSVs
    7. Generate deterministic Markdown summary
    8. Optionally generate AI summary if USE_CLAUDE_API=true
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

# ── Bootstrap: ensure src/ is importable when running from project root ───────
sys.path.insert(0, str(Path(__file__).parent))

# Reconfigure stdout/stderr to UTF-8 so Unicode log messages print correctly on Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from src import config
from src.utils import generate_run_id
from src.processing_logger import ProcessingLogger
from src.file_scanner import scan_inbox
from src.file_classifier import classify_all
from src.bring_parser import parse_bring_pdf_header, parse_bring_excel_specification
from src.postnord_parser import parse_postnord_pdf
from src.normalization import merge_bring_headers
from src.validation import run_all_checks
from src.anomaly_detection import detect_bring_anomalies
from src.unknown_carrier_parser import parse_unknown_carrier_file
from src.output_writer import (
    ensure_all_output_headers,
    get_existing_invoice_keys,
    write_file_inventory,
    write_invoice_headers,
    write_invoice_lines,
    write_surcharge_lines,
    write_invoice_checks,
    write_anomalies,
)
from src.summary_writer import (
    build_summary_payload,
    write_deterministic_summary,
    write_ai_summary,
)
from src.claude_client import (
    is_claude_enabled,
    classify_ambiguous_line,
    generate_management_summary,
    explain_anomalies,
)
from src.run_exporter import write_run_export, write_missing_file_alert
from src.dashboard_writer import write_html_dashboard


def parse_args():
    parser = argparse.ArgumentParser(description="Freight Invoice Control MVP")
    parser.add_argument("--dry-run", action="store_true",
                        help="Scan and classify only; do not write output files.")
    parser.add_argument("--use-claude", action="store_true",
                        help="Override config and enable Claude API for this run.")
    parser.add_argument("--move-files", action="store_true",
                        help="Move files after successful processing (default: keep in place).")
    parser.add_argument("--input-folder", type=str, default=None,
                        help="Override inbox folder path.")
    return parser.parse_args()


def main():
    args = parse_args()

    # Apply CLI overrides
    if args.use_claude:
        config.USE_CLAUDE_API = True
    if args.move_files:
        config.MOVE_FILES_AFTER_PROCESSING = True

    inbox_dir = Path(args.input_folder) if args.input_folder else config.INBOX_DIR

    # ── Initialise ────────────────────────────────────────────────────────────
    run_id = generate_run_id()
    scan_ts = datetime.now().isoformat(timespec="seconds")
    config.ensure_all_directories()

    logger = ProcessingLogger(run_id)
    logger.info("Main", f"=== Freight Invoice Control — Run {run_id} ===")
    logger.info("Main", f"Inbox: {inbox_dir}")
    logger.info("Main", f"Claude API: {'Enabled' if is_claude_enabled() else 'Disabled'}")

    if not args.dry_run:
        ensure_all_output_headers()

    # ── Step 1: Scan inbox ────────────────────────────────────────────────────
    logger.info("Main", "Step 1: Scanning inbox...")
    file_records = scan_inbox(run_id, logger, inbox_dir=inbox_dir)

    if not file_records:
        logger.info("Main", "No files found in inbox. Exiting.")
        if not args.dry_run:
            write_file_inventory([], logger)
        _print_summary(run_id, scan_ts, [], [], [], [], [], logger)
        return

    # ── Step 2: Classify files ────────────────────────────────────────────────
    logger.info("Main", "Step 2: Classifying files...")
    file_records = classify_all(file_records, logger)

    if args.dry_run:
        logger.info("Main", "Dry-run mode: classification complete, skipping parsing and output.")
        for r in file_records:
            print(f"  {r.file_name}: {r.detected_carrier} / {r.detected_document_type} / inv={r.detected_invoice_number}")
        return

    # ── Step 3: Parse files by carrier ───────────────────────────────────────
    logger.info("Main", "Step 3: Parsing files...")

    # Group classified files by carrier and document type
    bring_pdf_records = [
        r for r in file_records
        if r.detected_carrier == "Bring"
        and r.detected_document_type == "Invoice"
        and r.file_extension == ".pdf"
        and r.processing_status not in ("SkippedUnsupportedType", "Failed")
    ]
    bring_xls_records = [
        r for r in file_records
        if r.detected_carrier == "Bring"
        and r.detected_document_type == "Specification"
        and r.file_extension in (".xlsx", ".xls")
        and r.processing_status not in ("SkippedUnsupportedType", "Failed")
    ]
    postnord_pdf_records = [
        r for r in file_records
        if r.detected_carrier == "PostNord"
        and r.file_extension == ".pdf"
        and r.processing_status not in ("SkippedUnsupportedType", "Failed")
    ]
    unknown_carrier_records = [
        r for r in file_records
        if r.detected_carrier not in ("Bring", "PostNord")
        and r.processing_status not in ("SkippedUnsupportedType", "Failed")
    ]

    # Data stores keyed by (carrier, invoice_number)
    all_pdf_headers: dict = {}
    all_excel_headers: dict = {}
    all_lines: dict = {}

    # Collected for output
    all_invoice_headers = []
    all_invoice_lines = []

    service_mapping = config.load_service_mapping()
    surcharge_mapping = config.load_surcharge_mapping()

    # ── Parse Bring PDFs ──────────────────────────────────────────────────────
    for rec in bring_pdf_records:
        fp = Path(rec.file_path)
        try:
            pdf_h = parse_bring_pdf_header(fp, run_id, logger)
            if pdf_h:
                inv = pdf_h.invoice_number or rec.detected_invoice_number
                pdf_h.invoice_number = inv
                all_pdf_headers[("Bring", inv)] = pdf_h
                rec.detected_invoice_number = inv
                rec.processing_status = "Parsed"
            else:
                rec.processing_status = "Failed"
                rec.error_message = "PDF header extraction returned None"
        except Exception as e:
            rec.processing_status = "Failed"
            rec.error_message = str(e)
            logger.error("Main", f"Failed to parse Bring PDF: {e}", file_name=rec.file_name, error=e)

    # ── Parse Bring Excel specifications ─────────────────────────────────────
    for rec in bring_xls_records:
        fp = Path(rec.file_path)
        try:
            xls_h, lines = parse_bring_excel_specification(
                fp, run_id, logger,
                service_mapping=service_mapping,
                surcharge_mapping=surcharge_mapping,
            )
            if xls_h:
                inv = xls_h.invoice_number or rec.detected_invoice_number
                xls_h.invoice_number = inv
                all_excel_headers[("Bring", inv)] = xls_h
                all_lines[("Bring", inv)] = lines
                rec.detected_invoice_number = inv
                rec.processing_status = "Parsed"
            else:
                rec.processing_status = "Failed"
                rec.error_message = "Excel header extraction returned None"
        except Exception as e:
            rec.processing_status = "Failed"
            rec.error_message = str(e)
            logger.error("Main", f"Failed to parse Bring Excel: {e}", file_name=rec.file_name, error=e)

    # ── Parse PostNord PDFs (header + full per-shipment spec) ────────────────
    for rec in postnord_pdf_records:
        fp = Path(rec.file_path)
        try:
            pn_h, pn_lines = parse_postnord_pdf(
                fp, run_id, logger,
                service_mapping=service_mapping,
                surcharge_mapping=surcharge_mapping,
            )
            if pn_h:
                inv = pn_h.invoice_number or rec.detected_invoice_number
                pn_h.invoice_number = inv
                all_pdf_headers[("PostNord", inv)] = pn_h
                all_lines[("PostNord", inv)] = pn_lines
                rec.detected_invoice_number = inv
                rec.processing_status = "Parsed"
            else:
                rec.processing_status = "Failed"
                rec.error_message = "PostNord PDF header extraction returned None"
        except Exception as e:
            rec.processing_status = "Failed"
            rec.error_message = str(e)
            logger.error("Main", f"Failed to parse PostNord PDF: {e}", file_name=rec.file_name, error=e)

    # ── Collect merged Bring invoice headers for output ───────────────────────
    bring_invoice_numbers = set(
        inv for (c, inv) in list(all_pdf_headers.keys()) + list(all_excel_headers.keys())
        if c == "Bring"
    )
    all_anomalies = []

    for inv_num in bring_invoice_numbers:
        pdf_h = all_pdf_headers.get(("Bring", inv_num))
        xls_h = all_excel_headers.get(("Bring", inv_num))
        lines = all_lines.get(("Bring", inv_num), [])

        try:
            merged = merge_bring_headers(pdf_h, xls_h)
            all_invoice_headers.append(merged)
        except ValueError:
            pass

        all_invoice_lines.extend(lines)

        # Anomaly detection per invoice
        anomalies = detect_bring_anomalies(pdf_h, xls_h, lines, logger)
        all_anomalies.extend(anomalies)

    # Collect PostNord headers and lines
    for (carrier, inv_num), h in all_pdf_headers.items():
        if carrier == "PostNord":
            all_invoice_headers.append(h)
            all_invoice_lines.extend(all_lines.get(("PostNord", inv_num), []))

    # ── Step 3c: Unknown carrier — AI extraction ─────────────────────────────
    if unknown_carrier_records:
        logger.warning(
            "Main",
            f"{len(unknown_carrier_records)} file(s) from unknown carrier(s) — "
            f"attempting AI extraction.",
        )
        for rec in unknown_carrier_records:
            h, lns, anom = parse_unknown_carrier_file(rec, run_id, logger)
            if h:
                all_invoice_headers.append(h)
                all_invoice_lines.extend(lns)
                rec.processing_status = "ParsedByAI"
            else:
                rec.processing_status = "Failed"
            if anom:
                all_anomalies.append(anom)

    # ── Step 3b: Detect incomplete Bring invoices (missing PDF or Excel) ─────
    missing_bring = []
    for inv_num in bring_invoice_numbers:
        has_pdf = ("Bring", inv_num) in all_pdf_headers
        has_xls = ("Bring", inv_num) in all_excel_headers
        if has_pdf and not has_xls:
            missing_bring.append({
                "invoice_number": inv_num,
                "missing_file": "Excel specification (.xlsx)",
                "found_file": "PDF invoice",
                "message": "PDF invoice received but Excel specification is missing. "
                           "Cannot reconcile line items without the specification.",
            })
            logger.warning("Main",
                           f"Bring invoice {inv_num}: PDF received but Excel specification is missing.")
        elif has_xls and not has_pdf:
            missing_bring.append({
                "invoice_number": inv_num,
                "missing_file": "PDF invoice",
                "found_file": "Excel specification (.xlsx)",
                "message": "Excel specification received but PDF invoice is missing. "
                           "Cannot confirm invoice total without the PDF.",
            })
            logger.warning("Main",
                           f"Bring invoice {inv_num}: Excel specification received but PDF invoice is missing.")

    if missing_bring:
        write_missing_file_alert(run_id, missing_bring, logger)

    # ── Step 3d: AI classification of unresolved lines ────────────────────────
    if is_claude_enabled():
        unknown_lines = [
            ln for ln in all_invoice_lines
            if ln.service_category == "Unknown"
            or (ln.line_type == "Surcharge" and ln.surcharge_category in ("Unknown", ""))
        ]
        if unknown_lines:
            logger.info("Main", f"Step 3d: AI classifying {len(unknown_lines)} unknown line(s)...")
            for ln in unknown_lines:
                result = classify_ambiguous_line(run_id, ln.to_dict(), logger)
                if result.get("service_category") not in (None, "Unknown"):
                    ln.service_category = result["service_category"]
                if result.get("surcharge_category") not in (None, "Unknown", ""):
                    ln.surcharge_category = result["surcharge_category"]
                if result.get("line_type") not in (None, "Unknown"):
                    ln.line_type = result["line_type"]
                ln.classified_by = result.get("classified_by", "Claude")
                ln.classification_confidence = float(result.get("confidence", 0.0))
                ln.manual_review_required = bool(result.get("should_review_manually", True))
        else:
            logger.info("Main", "Step 3d: All lines classified by rules — no AI needed.")

    # ── Optional: Claude anomaly explanations ─────────────────────────────────
    if is_claude_enabled() and all_anomalies:
        logger.info("Main", "Requesting Claude anomaly explanations...")
        anomaly_dicts = [a.to_dict() for a in all_anomalies]
        explanations = explain_anomalies(run_id, anomaly_dicts, logger)
        # Attach explanations back to anomaly objects by type
        explanation_map = {e.get("anomaly_type"): e for e in explanations if isinstance(e, dict)}
        for a in all_anomalies:
            if a.anomaly_type in explanation_map:
                e = explanation_map[a.anomaly_type]
                a.claude_explanation = e.get("explanation", "")

    # ── Step 4: Validation checks ─────────────────────────────────────────────
    logger.info("Main", "Step 4: Running validation checks...")
    all_checks = run_all_checks(
        run_id, all_pdf_headers, all_excel_headers, all_lines, logger
    )

    # Apply reconciliation status before writing so CSVs have correct values
    _apply_reconciliation_status(all_invoice_headers, all_checks)

    # ── Step 5: Write output CSVs (with run-level deduplication) ─────────────
    logger.info("Main", "Step 5: Writing output files...")
    existing_keys = get_existing_invoice_keys()
    write_file_inventory(file_records, logger)
    written_keys = write_invoice_headers(all_invoice_headers, logger, skip_keys=existing_keys)
    write_invoice_lines(all_invoice_lines, logger, skip_keys=existing_keys)
    write_surcharge_lines(all_invoice_lines, logger, skip_keys=existing_keys)
    # Checks: skip for invoices already written in a previous run
    new_checks = [c for c in all_checks if (c.carrier, c.invoice_number) not in existing_keys]
    write_invoice_checks(new_checks, logger)
    write_anomalies(all_anomalies, logger, skip_keys=existing_keys)

    # ── Step 6: Generate summaries ────────────────────────────────────────────
    logger.info("Main", "Step 6: Generating summaries...")
    log_counts = logger.get_counts()
    payload = build_summary_payload(
        run_id=run_id,
        scan_timestamp=scan_ts,
        file_records=file_records,
        headers=all_invoice_headers,
        lines=all_invoice_lines,
        checks=all_checks,
        anomalies=all_anomalies,
        log_counts=log_counts,
    )
    det_path = write_deterministic_summary(
        run_id, payload, file_records, all_invoice_headers,
        all_invoice_lines, all_checks, all_anomalies, logger,
    )

    # ── Step 7: Optional AI summary (only when new invoices were written) ────
    ai_text: str | None = None
    if not written_keys:
        logger.info("Main", "Step 7: Skipped — no new invoices in this run.")
    elif is_claude_enabled():
        logger.info("Main", "Step 7: Generating AI management summary...")
        ai_text = generate_management_summary(run_id, payload, logger)
        if ai_text:
            ai_path = write_ai_summary(run_id, payload, ai_text, logger)
            logger.info("Main", f"AI summary: {ai_path.name}")
        else:
            logger.warning("Main", "Claude API did not return a summary.")
    else:
        logger.info("Main", "Step 7: Skipped (USE_CLAUDE_API=false).")

    # ── Step 8: Run export + HTML dashboard ──────────────────────────────────
    logger.info("Main", "Step 8: Generating run export and dashboard...")
    if written_keys:
        write_run_export(
            run_id, payload, all_invoice_headers, all_invoice_lines,
            all_checks, logger, ai_summary=ai_text, anomalies=all_anomalies,
        )
    else:
        logger.info("Main", "Step 8: Skipping For_Email export — no new invoices.")
    write_html_dashboard(logger)

    # ── File movement (if configured) ─────────────────────────────────────────
    if config.MOVE_FILES_AFTER_PROCESSING:
        _move_processed_files(file_records, logger)

    # ── Final console summary ─────────────────────────────────────────────────
    _print_summary(run_id, scan_ts, file_records, all_invoice_headers,
                   all_checks, all_anomalies, all_invoice_lines, logger)
    logger.info("Main", f"Run complete. Deterministic summary: {det_path}")


_RECON_CHECK_NAMES = {"PDFTotalVsExcelSummary", "LineSumVsHeaderTotal"}


def _apply_reconciliation_status(headers, checks):
    """Set reconciliation_status on headers based on check results."""
    recon_by_inv = {}
    for c in checks:
        if c.check_name in _RECON_CHECK_NAMES:
            recon_by_inv[c.invoice_number] = c.status
    for h in headers:
        if h.invoice_number in recon_by_inv:
            h.reconciliation_status = recon_by_inv[h.invoice_number]
        elif not h.reconciliation_status:
            h.reconciliation_status = "NotChecked"


def _move_processed_files(file_records, logger):
    """Copy/move files from inbox to 01_Raw after processing."""
    from shutil import copy2
    for rec in file_records:
        if rec.processing_status not in ("Parsed", "Classified"):
            continue
        fp = Path(rec.file_path)
        dt = datetime.now()
        carrier_folder = rec.detected_carrier or "Unknown"
        target_dir = config.RAW_DIR / dt.strftime("%Y") / dt.strftime("%m") / carrier_folder
        target_dir.mkdir(parents=True, exist_ok=True)
        target = target_dir / fp.name
        if target.exists():
            stem = fp.stem
            suffix = fp.suffix
            target = target_dir / f"{stem}_{rec.run_id}{suffix}"
        try:
            copy2(str(fp), str(target))
            logger.info("FileMove", f"Copied {fp.name} → {target}", file_name=fp.name)
        except Exception as e:
            logger.error("FileMove", f"Failed to copy {fp.name}: {e}", file_name=fp.name, error=e)


def _print_summary(run_id, scan_ts, file_records, headers, checks, anomalies, lines, logger):
    """Print a concise console summary at the end of the run."""
    counts = logger.get_counts()
    total_amount = sum(ln.amount or 0.0 for ln in lines)
    recon_results = [c for c in checks if c.check_name in _RECON_CHECK_NAMES]
    recon_str = ", ".join(f"{c.status} ({c.invoice_number})" for c in recon_results) or "N/A"

    print()
    print("=" * 60)
    print(f"  Freight Invoice Control — Run Summary")
    print(f"  Run ID:      {run_id}")
    print(f"  Files:       {len(file_records)} scanned, {sum(1 for r in file_records if r.processing_status == 'Parsed')} parsed")
    print(f"  Invoices:    {len(headers)} detected")
    print(f"  Lines:       {len(lines)}")
    print(f"  Total amt:   {total_amount:,.2f}")
    print(f"  Reconcil.:   {recon_str}")
    print(f"  Anomalies:   {len(anomalies)}")
    print(f"  Warnings:    {counts.get('WARNING', 0)}")
    print(f"  Errors:      {counts.get('ERROR', 0)}")
    print("=" * 60)
    print()


if __name__ == "__main__":
    main()
