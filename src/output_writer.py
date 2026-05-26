"""Write all output CSV files. Append-friendly with per-invoice deduplication."""

from __future__ import annotations

import csv
from pathlib import Path
from typing import List, Set, Tuple

from src import config
from src.bring_parser import BringInvoiceHeader, BringInvoiceLine
from src.file_scanner import FileRecord
from src.processing_logger import ProcessingLogger
from src.validation import CheckResult


def get_existing_invoice_keys(csv_path: Path = None) -> Set[Tuple[str, str]]:
    """
    Read invoice_header.csv and return all (carrier, invoice_number) pairs already stored.
    Used to skip re-writing invoices that were processed in a previous run.
    """
    path = csv_path or config.INVOICE_HEADER_CSV
    keys: Set[Tuple[str, str]] = set()
    if not path.exists():
        return keys
    try:
        with open(path, encoding=config.CSV_ENCODING, newline="") as f:
            reader = csv.DictReader(f, delimiter=config.CSV_DELIMITER)
            for row in reader:
                carrier = row.get("carrier", "").strip()
                inv = row.get("invoice_number", "").strip()
                if carrier and inv:
                    keys.add((carrier, inv))
    except Exception:
        pass
    return keys


# ── Field definitions (defines column order in each CSV) ─────────────────────

FILE_INVENTORY_FIELDS = [
    "run_id", "scan_timestamp", "file_name", "file_path", "file_extension",
    "file_size_bytes", "file_modified_timestamp", "detected_carrier",
    "detected_document_type", "detected_invoice_number", "processing_status", "error_message",
]

INVOICE_HEADER_FIELDS = [
    "run_id", "processed_timestamp", "carrier", "invoice_number", "invoice_date",
    "due_date", "customer_number", "customer_reference", "period_from", "period_to",
    "currency", "total_ex_vat", "vat_amount", "total_inc_vat", "source_file",
    "document_type", "reconciliation_status", "error_message",
]

INVOICE_LINES_FIELDS = [
    "run_id", "processed_timestamp", "carrier", "invoice_number", "source_file",
    "line_no", "article_number", "service_code", "service_name_raw", "service_category",
    "from_country", "to_country", "quantity", "unit", "unit_price", "discount_percent",
    "vat_type", "amount", "line_type", "classified_by", "classification_confidence",
    "manual_review_required",
]

SURCHARGE_LINES_FIELDS = [
    "run_id", "processed_timestamp", "carrier", "invoice_number", "source_file",
    "line_no", "surcharge_raw", "surcharge_category", "service_name_raw",
    "quantity", "unit_price", "amount", "related_service_category",
    "classified_by", "classification_confidence", "manual_review_required",
]

ANOMALY_FIELDS = [
    "run_id", "processed_timestamp", "carrier", "invoice_number",
    "anomaly_type", "severity", "description", "detail",
    "line_no", "value", "threshold", "suggested_action", "claude_explanation",
]

INVOICE_CHECKS_FIELDS = [
    "run_id", "processed_timestamp", "carrier", "invoice_number", "check_name",
    "expected_value", "actual_value", "difference", "status", "severity",
    "message", "source_files",
]


def _ensure_csv_header(path: Path, fieldnames: list[str]) -> None:
    """Write CSV header if the file does not yet exist."""
    if not path.exists():
        with open(path, "w", newline="", encoding=config.CSV_ENCODING) as f:
            w = csv.DictWriter(f, fieldnames=fieldnames, delimiter=config.CSV_DELIMITER,
                               extrasaction="ignore")
            w.writeheader()


def _append_rows(path: Path, fieldnames: list[str], rows: list[dict]) -> None:
    if not rows:
        return
    _ensure_csv_header(path, fieldnames)
    with open(path, "a", newline="", encoding=config.CSV_ENCODING) as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, delimiter=config.CSV_DELIMITER,
                           extrasaction="ignore")
        w.writerows(rows)


def write_file_inventory(records: List[FileRecord], logger: ProcessingLogger) -> None:
    rows = [r.to_dict() for r in records]
    _append_rows(config.FILE_INVENTORY_CSV, FILE_INVENTORY_FIELDS, rows)
    logger.info("OutputWriter", f"Wrote {len(rows)} record(s) to file_inventory.csv")


def write_invoice_headers(
    headers: list,
    logger: ProcessingLogger,
    skip_keys: Set[Tuple[str, str]] = None,
) -> Set[Tuple[str, str]]:
    """
    Write invoice headers, skipping any (carrier, invoice_number) already in skip_keys.
    Returns the set of keys actually written this run.
    """
    if skip_keys is None:
        skip_keys = set()
    written_keys: Set[Tuple[str, str]] = set()
    rows = []
    for h in headers:
        key = (h.carrier, h.invoice_number)
        if key in skip_keys:
            logger.info(
                "OutputWriter",
                f"Skipping duplicate invoice {h.invoice_number} ({h.carrier}) — already in output.",
            )
            continue
        rows.append(h.to_dict())
        written_keys.add(key)
    _append_rows(config.INVOICE_HEADER_CSV, INVOICE_HEADER_FIELDS, rows)
    logger.info("OutputWriter", f"Wrote {len(rows)} header(s) to invoice_header.csv")
    return written_keys


def write_invoice_lines(
    lines: list,
    logger: ProcessingLogger,
    skip_keys: Set[Tuple[str, str]] = None,
) -> None:
    """Write invoice lines, skipping any belonging to already-existing invoices."""
    if skip_keys is None:
        skip_keys = set()
    rows = [
        ln.to_dict() for ln in lines
        if (ln.carrier, ln.invoice_number) not in skip_keys
    ]
    _append_rows(config.INVOICE_LINES_CSV, INVOICE_LINES_FIELDS, rows)
    logger.info("OutputWriter", f"Wrote {len(rows)} line(s) to invoice_lines.csv")


def write_surcharge_lines(
    lines: list,
    logger: ProcessingLogger,
    skip_keys: Set[Tuple[str, str]] = None,
) -> None:
    """Write surcharge lines, skipping any belonging to already-existing invoices."""
    if skip_keys is None:
        skip_keys = set()
    surcharge_rows = [
        ln.to_surcharge_dict()
        for ln in lines
        if ln.line_type == "Surcharge"
        and (ln.carrier, ln.invoice_number) not in skip_keys
    ]
    _append_rows(config.SURCHARGE_LINES_CSV, SURCHARGE_LINES_FIELDS, surcharge_rows)
    logger.info("OutputWriter", f"Wrote {len(surcharge_rows)} surcharge line(s) to surcharge_lines.csv")


def write_invoice_checks(checks: List[CheckResult], logger: ProcessingLogger) -> None:
    rows = [c.to_dict() for c in checks]
    _append_rows(config.INVOICE_CHECKS_CSV, INVOICE_CHECKS_FIELDS, rows)
    logger.info("OutputWriter", f"Wrote {len(rows)} check result(s) to invoice_checks.csv")


def write_anomalies(
    anomalies: list,
    logger: ProcessingLogger,
    skip_keys: Set[Tuple[str, str]] = None,
) -> None:
    """Write anomalies, skipping those belonging to already-processed invoices."""
    from datetime import datetime
    if skip_keys is None:
        skip_keys = set()
    rows = []
    for a in anomalies:
        if (a.carrier, a.invoice_number) in skip_keys:
            continue
        d = a.to_dict()
        d["processed_timestamp"] = datetime.now().isoformat(timespec="seconds")
        rows.append(d)
    _append_rows(config.ANOMALIES_CSV, ANOMALY_FIELDS, rows)
    logger.info("OutputWriter", f"Wrote {len(rows)} anomaly record(s) to anomalies.csv")


PENDING_INVOICE_FIELDS = [
    "run_id", "processed_timestamp", "carrier", "invoice_number",
    "reconciliation_status", "known_total_ex_vat", "source_file", "note",
]


def write_pending_invoices(
    missing_bring: list,
    logger: ProcessingLogger,
    run_id: str = "",
) -> None:
    """
    Overwrite pending_invoices.csv with the current set of incomplete invoices.
    Called every run so the file always reflects what is currently unresolved.
    """
    from datetime import datetime
    config.ensure_all_directories()
    rows = []
    for m in missing_bring:
        rows.append({
            "run_id": run_id,
            "processed_timestamp": datetime.now().isoformat(timespec="seconds"),
            "carrier": "Bring",
            "invoice_number": m.get("invoice_number", ""),
            "reconciliation_status": "Pending",
            "known_total_ex_vat": m.get("known_total_ex_vat", ""),
            "source_file": m.get("source_file", ""),
            "note": m.get("message", ""),
        })
    path = config.PENDING_INVOICES_CSV
    with open(path, "w", newline="", encoding=config.CSV_ENCODING) as f:
        w = csv.DictWriter(f, fieldnames=PENDING_INVOICE_FIELDS, delimiter=config.CSV_DELIMITER)
        w.writeheader()
        w.writerows(rows)
    logger.info("OutputWriter", f"Wrote {len(rows)} pending invoice(s) to pending_invoices.csv")


def ensure_all_output_headers() -> None:
    """Pre-create all output CSV files with headers so they exist even on empty runs."""
    config.ensure_all_directories()
    _ensure_csv_header(config.FILE_INVENTORY_CSV, FILE_INVENTORY_FIELDS)
    _ensure_csv_header(config.INVOICE_HEADER_CSV, INVOICE_HEADER_FIELDS)
    _ensure_csv_header(config.INVOICE_LINES_CSV, INVOICE_LINES_FIELDS)
    _ensure_csv_header(config.SURCHARGE_LINES_CSV, SURCHARGE_LINES_FIELDS)
    _ensure_csv_header(config.INVOICE_CHECKS_CSV, INVOICE_CHECKS_FIELDS)
    _ensure_csv_header(config.ANOMALIES_CSV, ANOMALY_FIELDS)
