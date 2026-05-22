"""
Parser for invoices from unknown/unsupported carriers.

Uses Claude to extract whatever it can from the raw PDF text.
Every extracted value is flagged manual_review_required=True.
Also produces a code_recommendation so a developer knows what to build
to support this carrier properly in future runs.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple, List

from src.bring_parser import BringInvoiceHeader, BringInvoiceLine
from src.anomaly_detection import Anomaly
from src.file_scanner import FileRecord
from src.processing_logger import ProcessingLogger


def _extract_pdf_text(file_path: Path, logger: ProcessingLogger) -> str:
    try:
        import pdfplumber
        with pdfplumber.open(str(file_path)) as pdf:
            parts = [page.extract_text() or "" for page in pdf.pages]
        return "\n".join(parts)
    except Exception as e:
        logger.warning("UnknownCarrierParser", f"Could not extract PDF text: {e}", error=e)
        return ""


def parse_unknown_carrier_file(
    record: FileRecord,
    run_id: str,
    logger: ProcessingLogger,
) -> Tuple[Optional[BringInvoiceHeader], List[BringInvoiceLine], Optional[Anomaly]]:
    """
    Attempt to extract invoice data from an unknown carrier file using Claude.
    Returns (header, lines, anomaly). Any of these may be None/empty on failure.
    All returned data is flagged manual_review_required.
    """
    from src.claude_client import extract_unknown_carrier_invoice, is_claude_enabled

    file_path = Path(record.file_path)
    now_ts = datetime.now().isoformat(timespec="seconds")

    if not is_claude_enabled():
        logger.warning(
            "UnknownCarrierParser",
            f"{record.file_name}: Unknown carrier — Claude API disabled, cannot extract. "
            f"Enable USE_CLAUDE_API=true to process this file.",
            file_name=record.file_name,
        )
        anomaly = Anomaly(
            anomaly_type="UnknownCarrier",
            severity="Error",
            carrier="Unknown",
            invoice_number=record.detected_invoice_number or "?",
            description=f"Invoice from unknown carrier received: {record.file_name}. "
                        f"Claude API is disabled — file could not be extracted.",
            detail=f"File: {record.file_name} | Size: {record.file_size_bytes} bytes",
            suggested_action="Enable Claude API and re-run, or manually identify the carrier "
                             "and build a dedicated parser.",
        )
        return None, [], anomaly

    # Extract raw text from PDF
    raw_text = ""
    if record.file_extension.lower() == ".pdf":
        raw_text = _extract_pdf_text(file_path, logger)

    if not raw_text.strip():
        logger.warning(
            "UnknownCarrierParser",
            f"{record.file_name}: Could not extract text (possibly image-based PDF or unsupported format).",
            file_name=record.file_name,
        )
        anomaly = Anomaly(
            anomaly_type="UnknownCarrier",
            severity="Error",
            carrier="Unknown",
            invoice_number=record.detected_invoice_number or "?",
            description=f"Invoice from unknown carrier received: {record.file_name}. "
                        f"No text could be extracted — may be a scanned/image PDF.",
            detail=f"File: {record.file_name} | Size: {record.file_size_bytes} bytes",
            suggested_action="Open the file manually, identify the carrier, and either scan "
                             "with OCR or request a text-based PDF from the carrier.",
        )
        return None, [], anomaly

    logger.info(
        "UnknownCarrierParser",
        f"{record.file_name}: Sending to Claude for extraction ({len(raw_text)} chars)...",
        file_name=record.file_name,
    )

    result = extract_unknown_carrier_invoice(run_id, raw_text, record.file_name, logger)

    if not result:
        anomaly = Anomaly(
            anomaly_type="UnknownCarrier",
            severity="Error",
            carrier="Unknown",
            invoice_number=record.detected_invoice_number or "?",
            description=f"Invoice from unknown carrier received: {record.file_name}. "
                        f"Claude extraction failed — see logs.",
            detail=f"File: {record.file_name}",
            suggested_action="Check Claude API logs and retry, or process the file manually.",
        )
        return None, [], anomaly

    carrier_name = result.get("carrier_name") or "Unknown"
    invoice_number = result.get("invoice_number") or record.detected_invoice_number or "UNKNOWN"
    confidence = float(result.get("extraction_confidence", 0.0))
    code_rec = result.get("code_recommendation", "")

    logger.info(
        "UnknownCarrierParser",
        f"{record.file_name}: Claude identified carrier='{carrier_name}', "
        f"invoice={invoice_number}, confidence={confidence:.0%}",
        file_name=record.file_name,
    )

    # Build invoice header — all amounts null unless Claude is confident
    header = BringInvoiceHeader(
        run_id=run_id,
        processed_timestamp=now_ts,
        carrier=carrier_name,
        invoice_number=invoice_number,
        invoice_date=result.get("invoice_date") or "",
        due_date=result.get("due_date") or "",
        customer_number=result.get("customer_number") or "",
        currency=result.get("currency") or "SEK",
        total_ex_vat=result.get("total_ex_vat"),
        total_inc_vat=result.get("total_inc_vat"),
        source_file=record.file_name,
        document_type="Invoice",
        reconciliation_status="ManualReview",
        error_message=f"AI-extracted (confidence {confidence:.0%}) — manual verification required.",
    )

    # Build lines
    lines: List[BringInvoiceLine] = []
    for i, item in enumerate(result.get("line_items", []), start=1):
        ln = BringInvoiceLine(
            run_id=run_id,
            processed_timestamp=now_ts,
            invoice_number=invoice_number,
            carrier=carrier_name,
            source_file=record.file_name,
            line_no=i,
            service_name_raw=item.get("description", ""),
            base_service_name=item.get("description", ""),
            line_type=item.get("line_type", "Unknown"),
            service_category="Unknown",
            surcharge_category="",
            classified_by="Claude",
            classification_confidence=confidence,
            manual_review_required=True,
            quantity=float(item.get("quantity") or 1),
            amount=float(item.get("amount") or 0.0),
        )
        lines.append(ln)

    # Build anomaly that carries the code recommendation
    detail_parts = [
        f"File: {record.file_name}",
        f"AI confidence: {confidence:.0%}",
        f"Lines extracted: {len(lines)}",
    ]
    if code_rec:
        detail_parts.append(f"Code recommendation: {code_rec}")

    anomaly = Anomaly(
        anomaly_type="UnknownCarrier",
        severity="Error",
        carrier=carrier_name,
        invoice_number=invoice_number,
        description=(
            f"Invoice from new carrier '{carrier_name}' — AI-extracted at {confidence:.0%} confidence. "
            f"All values require manual verification. Build a dedicated parser to handle this carrier automatically."
        ),
        detail=" | ".join(detail_parts),
        suggested_action=(
            f"1. Verify all extracted amounts against the original file. "
            f"2. Ask your developer to build a '{carrier_name}' parser. "
            f"Recommendation: {code_rec}"
        ),
    )

    return header, lines, anomaly
