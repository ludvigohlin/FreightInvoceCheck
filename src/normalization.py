"""Normalization — merge PDF header + Excel header into unified invoice header."""

from __future__ import annotations

from typing import Optional

from src.bring_parser import BringInvoiceHeader


def merge_bring_headers(
    pdf_header: Optional[BringInvoiceHeader],
    excel_header: Optional[BringInvoiceHeader],
) -> BringInvoiceHeader:
    """
    Merge PDF invoice header and Excel specification header into one record.
    PDF is authoritative for financial totals; Excel provides supplementary fields.
    The PDF 'total_ex_vat' is the source of truth for the invoiced amount.
    """
    if pdf_header is None and excel_header is None:
        raise ValueError("Both pdf_header and excel_header are None")

    if pdf_header is None:
        excel_header.document_type = "Specification"
        return excel_header

    if excel_header is None:
        return pdf_header

    # Start from PDF header (authoritative for financial data)
    merged = pdf_header

    # Supplement missing fields from Excel
    if not merged.customer_number and excel_header.customer_number:
        merged.customer_number = excel_header.customer_number

    if not merged.invoice_number and excel_header.invoice_number:
        merged.invoice_number = excel_header.invoice_number

    # Excel-derived totals used for reconciliation, not to override PDF
    # (stored separately in validation module)
    merged.document_type = "Invoice"  # PDF header represents the invoice

    return merged
