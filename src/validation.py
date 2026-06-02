"""Reconciliation checks and validation — all deterministic Python logic."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import List, Optional

from src import config
from src.bring_parser import BringInvoiceHeader, BringInvoiceLine
from src.processing_logger import ProcessingLogger
from src.postnord_parser import PostNordInvoiceHeader, PostNordInvoiceLine


@dataclass
class CheckResult:
    run_id: str
    processed_timestamp: str
    carrier: str
    invoice_number: str
    check_name: str
    expected_value: str
    actual_value: str
    difference: str
    status: str       # OK | Warning | Error | Info
    severity: str     # OK | Warning | Error
    message: str
    source_files: str

    def to_dict(self) -> dict:
        return {
            "run_id": self.run_id,
            "processed_timestamp": self.processed_timestamp,
            "carrier": self.carrier,
            "invoice_number": self.invoice_number,
            "check_name": self.check_name,
            "expected_value": self.expected_value,
            "actual_value": self.actual_value,
            "difference": self.difference,
            "status": self.status,
            "severity": self.severity,
            "message": self.message,
            "source_files": self.source_files,
        }


def _make_check(
    run_id: str,
    carrier: str,
    invoice_number: str,
    check_name: str,
    expected,
    actual,
    difference,
    status: str,
    severity: str,
    message: str,
    source_files: str,
) -> CheckResult:
    ts = datetime.now().isoformat(timespec="seconds")
    return CheckResult(
        run_id=run_id,
        processed_timestamp=ts,
        carrier=carrier,
        invoice_number=invoice_number,
        check_name=check_name,
        expected_value=str(expected) if expected is not None else "",
        actual_value=str(actual) if actual is not None else "",
        difference=str(difference) if difference is not None else "",
        status=status,
        severity=severity,
        message=message,
        source_files=source_files,
    )


def run_bring_checks(
    run_id: str,
    pdf_header: Optional[BringInvoiceHeader],
    excel_header: Optional[BringInvoiceHeader],
    lines: List[BringInvoiceLine],
    logger: ProcessingLogger,
) -> List[CheckResult]:
    """
    Run all reconciliation checks for a Bring invoice + specification pair.
    Returns list of CheckResult objects.
    """
    rules = config.load_validation_rules()
    tol_ok = rules.get("reconciliation", {}).get("total_tolerance_ok", 0.01)
    tol_warn = rules.get("reconciliation", {}).get("total_tolerance_warning", 1.00)

    checks: List[CheckResult] = []
    carrier = "Bring"
    inv_num = (pdf_header and pdf_header.invoice_number) or \
              (excel_header and excel_header.invoice_number) or "Unknown"
    pdf_file = pdf_header.source_file if pdf_header else ""
    xls_file = excel_header.source_file if excel_header else ""
    source = "; ".join(filter(None, [pdf_file, xls_file]))

    # ── Check 1: Invoice number present ──────────────────────────────────────
    if not inv_num or inv_num == "Unknown":
        checks.append(_make_check(
            run_id, carrier, inv_num, "MissingInvoiceNumber",
            "Non-empty invoice number", inv_num, "",
            "Error", "Error", "Invoice number could not be detected.", source,
        ))
    else:
        checks.append(_make_check(
            run_id, carrier, inv_num, "InvoiceNumberPresent",
            "Non-empty invoice number", inv_num, "",
            "OK", "OK", f"Invoice number detected: {inv_num}", source,
        ))

    # ── Check 2: PDF and Excel matched by same invoice number ─────────────────
    if pdf_header and excel_header:
        pdf_inv = pdf_header.invoice_number
        xls_inv = excel_header.invoice_number
        if pdf_inv and xls_inv and pdf_inv == xls_inv:
            checks.append(_make_check(
                run_id, carrier, inv_num, "PDFExcelInvoiceNumberMatch",
                pdf_inv, xls_inv, "",
                "OK", "OK", f"PDF and Excel share invoice number {inv_num}.", source,
            ))
        else:
            checks.append(_make_check(
                run_id, carrier, inv_num, "PDFExcelInvoiceNumberMatch",
                pdf_inv, xls_inv, "",
                "Warning", "Warning",
                f"Invoice number mismatch: PDF={pdf_inv}, Excel={xls_inv}", source,
            ))

    # ── Check 3: PDF total_ex_vat vs Excel specification sum ─────────────────
    if pdf_header and excel_header and pdf_header.total_ex_vat is not None and excel_header.total_ex_vat is not None:
        pdf_total = pdf_header.total_ex_vat
        xls_total = excel_header.total_ex_vat
        diff = round(xls_total - pdf_total, 2)
        abs_diff = abs(diff)

        if abs_diff <= tol_ok:
            status, severity = "OK", "OK"
            msg = f"PDF total ({pdf_total:.2f}) matches Excel summary ({xls_total:.2f})"
        elif abs_diff <= tol_warn:
            status, severity = "Warning", "Warning"
            msg = f"Small difference between PDF total ({pdf_total:.2f}) and Excel summary ({xls_total:.2f}): {diff:+.2f}"
        else:
            status, severity = "Error", "Error"
            msg = f"MISMATCH: PDF total ({pdf_total:.2f}) vs Excel summary ({xls_total:.2f}): {diff:+.2f}"

        checks.append(_make_check(
            run_id, carrier, inv_num, "PDFTotalVsExcelSummary",
            f"{pdf_total:.2f}", f"{xls_total:.2f}", f"{diff:+.2f}",
            status, severity, msg, source,
        ))

    # ── Check 4: Excel line sum vs Excel summary total ────────────────────────
    if lines and excel_header and excel_header.total_ex_vat is not None:
        line_sum = round(sum(ln.amount for ln in lines if ln.amount is not None), 2)
        xls_summary = excel_header.total_ex_vat
        diff = round(line_sum - xls_summary, 2)
        abs_diff = abs(diff)

        if abs_diff <= tol_ok:
            status, severity = "OK", "OK"
            msg = f"Excel line sum ({line_sum:.2f}) matches Excel summary ({xls_summary:.2f})"
        elif abs_diff <= tol_warn:
            status, severity = "Warning", "Warning"
            msg = f"Small difference: line sum ({line_sum:.2f}) vs Excel summary ({xls_summary:.2f}): {diff:+.2f}"
        else:
            status, severity = "Error", "Error"
            msg = f"MISMATCH: line sum ({line_sum:.2f}) vs Excel summary ({xls_summary:.2f}): {diff:+.2f}"

        checks.append(_make_check(
            run_id, carrier, inv_num, "ExcelLineSumVsExcelSummary",
            f"{xls_summary:.2f}", f"{line_sum:.2f}", f"{diff:+.2f}",
            status, severity, msg, source,
        ))

    # ── Check 5: Missing specification ───────────────────────────────────────
    if pdf_header and not excel_header:
        checks.append(_make_check(
            run_id, carrier, inv_num, "SpecificationPresent",
            "Excel specification file", "Not found", "",
            "Warning", "Warning",
            "Bring PDF invoice found but no matching Excel specification in inbox.", pdf_file,
        ))
    elif excel_header and not pdf_header:
        checks.append(_make_check(
            run_id, carrier, inv_num, "PDFInvoicePresent",
            "PDF invoice file", "Not found", "",
            "Warning", "Warning",
            "Bring Excel specification found but no matching PDF invoice in inbox.", xls_file,
        ))
    else:
        checks.append(_make_check(
            run_id, carrier, inv_num, "SpecificationPresent",
            "Excel specification file", xls_file, "",
            "OK", "OK", "Both PDF invoice and Excel specification present.", source,
        ))

    # ── Check 6: All lines classified ────────────────────────────────────────
    if lines:
        unknown_service = [ln for ln in lines if ln.service_category == "Unknown"]
        unknown_surcharge = [ln for ln in lines if ln.line_type == "Surcharge" and ln.surcharge_category == "Unknown"]

        if unknown_service:
            checks.append(_make_check(
                run_id, carrier, inv_num, "UnclassifiedServiceLines",
                "0 unclassified service lines",
                str(len(unknown_service)),
                str(len(unknown_service)),
                "Warning", "Warning",
                f"{len(unknown_service)} line(s) have Unknown service_category. Manual review recommended.",
                source,
            ))
        else:
            checks.append(_make_check(
                run_id, carrier, inv_num, "UnclassifiedServiceLines",
                "0", "0", "0",
                "OK", "OK", "All lines classified into a known service_category.", source,
            ))

        if unknown_surcharge:
            checks.append(_make_check(
                run_id, carrier, inv_num, "UnclassifiedSurchargeLines",
                "0 unclassified surcharge lines",
                str(len(unknown_surcharge)),
                str(len(unknown_surcharge)),
                "Warning", "Warning",
                f"{len(unknown_surcharge)} surcharge line(s) have Unknown surcharge_category.",
                source,
            ))
        else:
            checks.append(_make_check(
                run_id, carrier, inv_num, "UnclassifiedSurchargeLines",
                "0", "0", "0",
                "OK", "OK", "All surcharge lines classified into a known surcharge_category.", source,
            ))

    # ── Check 7: Mandatory header fields ─────────────────────────────────────
    header_to_check = pdf_header or excel_header
    if header_to_check:
        mandatory = rules.get("mandatory_header_fields", {})
        for field_name, severity in mandatory.items():
            val = getattr(header_to_check, field_name, None)
            has_val = val is not None and str(val).strip() not in ("", "None")
            checks.append(_make_check(
                run_id, carrier, inv_num, f"MandatoryField_{field_name}",
                "Present", "Present" if has_val else "Missing", "",
                "OK" if has_val else severity,
                "OK" if has_val else severity,
                f"Field '{field_name}' = {val!r}" if has_val else f"Mandatory field '{field_name}' is missing.",
                source,
            ))

    for c in checks:
        log_fn = logger.warning if c.severity == "Warning" else (
            logger.error if c.severity == "Error" else logger.info
        )
        log_fn("Validation", f"[{c.check_name}] {c.status}: {c.message}")

    return checks


def run_postnord_checks(
    run_id: str,
    header: PostNordInvoiceHeader,
    lines: List[PostNordInvoiceLine],
    logger: ProcessingLogger,
) -> List[CheckResult]:
    """Run validation checks for a PostNord combined PDF invoice."""
    rules = config.load_validation_rules()
    tol_ok = rules.get("reconciliation", {}).get("total_tolerance_ok", 0.01)
    tol_warn = rules.get("reconciliation", {}).get("total_tolerance_warning", 1.00)

    checks: List[CheckResult] = []
    carrier = "PostNord"
    inv_num = header.invoice_number or "Unknown"
    source = header.source_file

    # Check 1: Invoice number present
    if not inv_num or inv_num == "Unknown":
        checks.append(_make_check(
            run_id, carrier, inv_num, "InvoiceNumberPresent",
            "Non-empty invoice number", inv_num, "",
            "Error", "Error", "Invoice number could not be detected.", source,
        ))
    else:
        checks.append(_make_check(
            run_id, carrier, inv_num, "InvoiceNumberPresent",
            "Non-empty invoice number", inv_num, "",
            "OK", "OK", f"Invoice number detected: {inv_num}", source,
        ))

    # Check 2: Line sum vs header total_ex_vat
    if header.total_ex_vat is not None and lines:
        line_sum = round(sum(ln.amount for ln in lines if ln.amount is not None), 2)
        expected = header.total_ex_vat
        diff = round(line_sum - expected, 2)
        abs_diff = abs(diff)

        if abs_diff <= tol_ok:
            status, severity = "OK", "OK"
            msg = f"Line sum ({line_sum:.2f}) matches header total_ex_vat ({expected:.2f})"
        elif abs_diff <= tol_warn:
            status, severity = "Warning", "Warning"
            msg = f"Small difference: line sum ({line_sum:.2f}) vs header ({expected:.2f}): {diff:+.2f}"
        else:
            status, severity = "Error", "Error"
            msg = f"MISMATCH: line sum ({line_sum:.2f}) vs header ({expected:.2f}): {diff:+.2f}"

        checks.append(_make_check(
            run_id, carrier, inv_num, "LineSumVsHeaderTotal",
            f"{expected:.2f}", f"{line_sum:.2f}", f"{diff:+.2f}",
            status, severity, msg, source,
        ))
    elif header.total_ex_vat is not None and not lines and header.total_ex_vat > 0:
        checks.append(_make_check(
            run_id, carrier, inv_num, "LineSumVsHeaderTotal",
            f"{header.total_ex_vat:.2f}", "0.00", f"{-header.total_ex_vat:.2f}",
            "Warning", "Warning",
            f"No lines could be parsed from this invoice (header total: {header.total_ex_vat:.2f} SEK). "
            f"May be a supplement or adjustment invoice — verify manually.",
            source,
        ))

    # Check 3: Mandatory header fields
    mandatory = rules.get("mandatory_header_fields", {})
    for field_name, severity in mandatory.items():
        val = getattr(header, field_name, None)
        has_val = val is not None and str(val).strip() not in ("", "None")
        checks.append(_make_check(
            run_id, carrier, inv_num, f"MandatoryField_{field_name}",
            "Present", "Present" if has_val else "Missing", "",
            "OK" if has_val else severity,
            "OK" if has_val else severity,
            f"Field '{field_name}' = {val!r}" if has_val else f"Mandatory field '{field_name}' is missing.",
            source,
        ))

    # Check 4: Unclassified service lines
    if lines:
        base_lines = [ln for ln in lines if ln.line_type == "BaseFreight"]
        unknown_service = [ln for ln in base_lines if ln.service_category == "Unknown"]
        unknown_surcharge = [ln for ln in lines if ln.line_type == "Surcharge"
                             and ln.surcharge_category == "Unknown"]

        if unknown_service:
            checks.append(_make_check(
                run_id, carrier, inv_num, "UnclassifiedServiceLines",
                "0", str(len(unknown_service)), str(len(unknown_service)),
                "Warning", "Warning",
                f"{len(unknown_service)} line(s) have Unknown service_category.", source,
            ))
        else:
            checks.append(_make_check(
                run_id, carrier, inv_num, "UnclassifiedServiceLines",
                "0", "0", "0", "OK", "OK",
                "All service lines classified.", source,
            ))

        if unknown_surcharge:
            checks.append(_make_check(
                run_id, carrier, inv_num, "UnclassifiedSurchargeLines",
                "0", str(len(unknown_surcharge)), str(len(unknown_surcharge)),
                "Warning", "Warning",
                f"{len(unknown_surcharge)} surcharge line(s) have Unknown surcharge_category.", source,
            ))
        else:
            checks.append(_make_check(
                run_id, carrier, inv_num, "UnclassifiedSurchargeLines",
                "0", "0", "0", "OK", "OK",
                "All surcharge lines classified.", source,
            ))

    for c in checks:
        log_fn = logger.warning if c.severity == "Warning" else (
            logger.error if c.severity == "Error" else logger.info
        )
        log_fn("Validation", f"[{c.check_name}] {c.status}: {c.message}")

    return checks


def run_all_checks(
    run_id: str,
    all_pdf_headers: dict,    # carrier+invoice_number → header
    all_excel_headers: dict,  # carrier+invoice_number → header
    all_lines: dict,          # carrier+invoice_number → [lines]
    logger: ProcessingLogger,
) -> List[CheckResult]:
    """Coordinate validation across all parsed invoices in the run."""
    all_checks: List[CheckResult] = []

    # Bring invoices
    bring_invoice_numbers = set()
    for key in all_pdf_headers:
        carrier, inv_num = key
        if carrier == "Bring":
            bring_invoice_numbers.add(inv_num)
    for key in all_excel_headers:
        carrier, inv_num = key
        if carrier == "Bring":
            bring_invoice_numbers.add(inv_num)

    for inv_num in bring_invoice_numbers:
        pdf_h = all_pdf_headers.get(("Bring", inv_num))
        xls_h = all_excel_headers.get(("Bring", inv_num))
        lines = all_lines.get(("Bring", inv_num), [])
        all_checks.extend(run_bring_checks(run_id, pdf_h, xls_h, lines, logger))

    # PostNord invoices
    for key, header in all_pdf_headers.items():
        carrier, inv_num = key
        if carrier == "PostNord":
            lines = all_lines.get(("PostNord", inv_num), [])
            all_checks.extend(run_postnord_checks(run_id, header, lines, logger))

    return all_checks
