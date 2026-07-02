"""Classify files by carrier and document type using filename and content."""

from __future__ import annotations

import re
import shutil
import tempfile
from pathlib import Path
from typing import Optional, Tuple

from src import config
from src.file_scanner import FileRecord
from src.processing_logger import ProcessingLogger


def _filename_lower(file_name: str) -> str:
    return file_name.lower()


def _extract_invoice_number_from_filename(file_name: str) -> Optional[str]:
    """Try to extract invoice number from filename patterns."""
    patterns = [
        r"fakturanummer[_\s]?(\d+)",
        r"faktura[_\s]?(\d{8,13})",
        # Power Automate prefix: YYYY-MM-DD_HHMMSS_SHARED_originalname
        r"_(\d{10,13})(?:\.|_)",
        # Raw filename like FAKTURA_903103539321.pdf
        r"_(\d{10,13})\.",
        r"(\d{10,13})",
    ]
    name_lower = file_name.lower()
    for pat in patterns:
        m = re.search(pat, name_lower)
        if m:
            return m.group(1)
    return None


def _detect_bring_from_filename(name_lower: str) -> bool:
    keywords = ["bring", "specificerad", "specificeradfaktura"]
    return any(k in name_lower for k in keywords)


def _detect_postnord_from_filename(name_lower: str) -> bool:
    return "postnord" in name_lower


def _detect_bring_from_pdf_text(text: str) -> bool:
    keywords = ["bring e-commerce", "bring.com", "bring.se", "invoicing.se@bring.com",
                "bring e-commerce & logistics"]
    text_lower = text.lower()
    return any(k in text_lower for k in keywords)


def _detect_postnord_from_pdf_text(text: str) -> bool:
    keywords = ["postnord", "postnord.se", "postnord.com"]
    text_lower = text.lower()
    return any(k in text_lower for k in keywords)


def _extract_invoice_number_from_text(text: str) -> Optional[str]:
    """Try to extract invoice number from PDF text or Excel cell content."""
    patterns = [
        r"fakturanr\.?\s*(\d+)",
        r"fakturanummer\s*[:\s]?\s*(\d+)",
        r"faktura\s+(\d{8,12})\b",
        r"\b(40\d{8})\b",  # Bring-style: starts with 40
        r"\b(\d{10})\b",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return m.group(1)
    return None


def _is_bring_specification_excel(file_path: Path) -> Tuple[bool, Optional[str]]:
    """
    Check Excel workbook for Bring specification markers. Requires a plausible
    line-item header row (Artikelnummer/Beskrivning + Avtalspris/Bruttopris) —
    the same signal bring_parser.py itself relies on to find its data — instead
    of accepting any Excel file that merely opens successfully. Without this, a
    file from a different sender whose filename happened to contain "fakturanummer"
    or "specificerad" would previously always classify as a Bring specification.
    Returns (is_bring_spec, invoice_number).
    """
    try:
        import openpyxl
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp.close()
        shutil.copy2(str(file_path), tmp.name)
        wb = openpyxl.load_workbook(tmp.name, read_only=True, data_only=True)
        ws = wb.active
        inv_num = None
        has_header_row = False
        for i, row in enumerate(ws.iter_rows(max_row=20, values_only=True)):
            row_str = " ".join(str(c) for c in row if c is not None).lower()
            has_service_col = "artikelnummer" in row_str or "beskrivning" in row_str
            has_price_col = "avtalspris" in row_str or "bruttopris" in row_str
            if has_service_col and has_price_col:
                has_header_row = True
            for cell in row:
                if cell is None:
                    continue
                cell_str = str(cell)
                m = re.search(r"fakturanummer[:\s]+(\d+)", cell_str, re.IGNORECASE)
                if m:
                    inv_num = m.group(1)
                m2 = re.search(r"f.r fakturanummer[:\s]+(\d+)", cell_str, re.IGNORECASE)
                if m2:
                    inv_num = m2.group(1)
        wb.close()
        return has_header_row, inv_num
    except Exception:
        return False, None
    finally:
        try:
            Path(tmp.name).unlink(missing_ok=True)
        except Exception:
            pass


def classify_file(record: FileRecord, logger: ProcessingLogger) -> FileRecord:
    """
    Classify a file record into carrier + document type + invoice number.
    Updates record in-place and returns it.
    """
    file_path = Path(record.file_path)
    name_lower = _filename_lower(record.file_name)
    ext = record.file_extension.lower()

    carrier = "Unknown"
    doc_type = "Unknown"
    invoice_number = ""

    # ── Step 1: filename-based carrier detection ──────────────────────────────
    if _detect_bring_from_filename(name_lower):
        carrier = "Bring"
    elif _detect_postnord_from_filename(name_lower):
        carrier = "PostNord"

    # ── Step 2: invoice number from filename ──────────────────────────────────
    fn_inv = _extract_invoice_number_from_filename(record.file_name)
    if fn_inv:
        invoice_number = fn_inv

    # ── Step 3: document type and content-based refinement ───────────────────
    if ext == ".pdf":
        pdf_text = _safe_extract_pdf_text(file_path, logger)
        if pdf_text:
            if carrier == "Unknown" and _detect_bring_from_pdf_text(pdf_text):
                carrier = "Bring"
            if carrier == "Unknown" and _detect_postnord_from_pdf_text(pdf_text):
                carrier = "PostNord"

            if not invoice_number:
                inv = _extract_invoice_number_from_text(pdf_text)
                if inv:
                    invoice_number = inv

            if carrier == "Bring":
                doc_type = "Invoice"
            elif carrier == "PostNord":
                # PostNord PDFs often include spec within same document
                doc_type = "InvoiceAndSpecification"
            else:
                doc_type = "Unknown"

    elif ext in (".xlsx", ".xls"):
        if carrier == "Bring" or "fakturanummer" in name_lower or "specificerad" in name_lower:
            is_bring_spec, xl_inv = _is_bring_specification_excel(file_path)
            if is_bring_spec:
                carrier = "Bring"
                doc_type = "Specification"
                if xl_inv:
                    invoice_number = xl_inv  # content takes priority over filename
            elif "specificerad" in name_lower:
                # openpyxl couldn't read the file (e.g. auto-saved/locked format),
                # but the filename is an unambiguous Bring spec marker — trust it.
                carrier = "Bring"
                doc_type = "Specification"
                logger.warning(
                    "FileClassifier",
                    f"{record.file_name}: openpyxl could not read file; classified as "
                    f"Bring Specification from filename. Re-save the file from Excel if "
                    f"parsing fails.",
                    file_name=record.file_name,
                )
        else:
            doc_type = "Unknown"

    elif ext == ".csv":
        doc_type = "Unknown"

    record.detected_carrier = carrier
    record.detected_document_type = doc_type
    record.detected_invoice_number = invoice_number
    record.processing_status = "Classified"

    logger.info(
        "FileClassifier",
        f"{record.file_name} -> carrier={carrier}, type={doc_type}, invoice={invoice_number or 'N/A'}",
        file_name=record.file_name,
    )
    return record


def _safe_extract_pdf_text(file_path: Path, logger: ProcessingLogger) -> str:
    """Extract all text from a PDF file, returning empty string on failure."""
    try:
        import pdfplumber
        with pdfplumber.open(str(file_path)) as pdf:
            parts = []
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    parts.append(t)
            return "\n".join(parts)
    except Exception as e:
        logger.warning(
            "FileClassifier",
            f"Could not extract text from PDF: {file_path.name} — {e}",
            file_name=file_path.name,
            error=e,
        )
        return ""


def classify_all(records: list[FileRecord], logger: ProcessingLogger) -> list[FileRecord]:
    """Classify all records in place."""
    for record in records:
        if record.processing_status in ("SkippedUnsupportedType", "Failed"):
            continue
        try:
            classify_file(record, logger)
        except Exception as e:
            record.processing_status = "Failed"
            record.error_message = str(e)
            logger.error(
                "FileClassifier",
                f"Classification failed for {record.file_name}: {e}",
                file_name=record.file_name,
                error=e,
            )
    return records
