"""Generate deterministic Markdown summaries and optional Claude AI summaries."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import List, Optional

from src import config
from src.anomaly_detection import Anomaly
from src.bring_parser import BringInvoiceHeader, BringInvoiceLine
from src.file_scanner import FileRecord
from src.processing_logger import ProcessingLogger
from src.validation import CheckResult


def _fmt(val) -> str:
    if val is None:
        return "N/A"
    if isinstance(val, float):
        return f"{val:,.2f}"
    return str(val)


def build_summary_payload(
    run_id: str,
    scan_timestamp: str,
    file_records: List[FileRecord],
    headers: List[BringInvoiceHeader],
    lines: List[BringInvoiceLine],
    checks: List[CheckResult],
    anomalies: List[Anomaly],
    log_counts: dict,
) -> dict:
    """
    Build a structured dict of pre-calculated data for the summary.
    All numbers here are already calculated by Python code.
    Claude only uses this to write prose — it cannot change the numbers.
    """
    total_files = len(file_records)
    processed_files = sum(1 for r in file_records if r.processing_status not in
                          ("SkippedUnsupportedType", "Failed", "Found"))
    skipped_files = sum(1 for r in file_records if r.processing_status == "SkippedUnsupportedType")
    failed_files = sum(1 for r in file_records if r.processing_status == "Failed")

    # Carrier summary
    carrier_totals = {}
    for h in headers:
        c = h.carrier
        if c not in carrier_totals:
            carrier_totals[c] = {"invoices": 0, "total_ex_vat": 0.0, "currency": h.currency or "SEK"}
        carrier_totals[c]["invoices"] += 1
        if h.total_ex_vat:
            carrier_totals[c]["total_ex_vat"] += h.total_ex_vat

    # Service category totals
    service_totals = {}
    for ln in lines:
        cat = ln.service_category or "Unknown"
        if cat not in service_totals:
            service_totals[cat] = 0.0
        service_totals[cat] += ln.amount or 0.0

    # Surcharge category totals
    surcharge_totals = {}
    for ln in lines:
        if ln.line_type == "Surcharge":
            cat = ln.surcharge_category or "Unknown"
            if cat not in surcharge_totals:
                surcharge_totals[cat] = 0.0
            surcharge_totals[cat] += ln.amount or 0.0

    # Check summary
    check_counts = {"OK": 0, "Warning": 0, "Error": 0}
    for c in checks:
        check_counts[c.status] = check_counts.get(c.status, 0) + 1

    anomaly_list = [
        {"type": a.anomaly_type, "severity": a.severity, "description": a.description}
        for a in anomalies
    ]

    invoices = [
        {
            "carrier": h.carrier,
            "invoice_number": h.invoice_number,
            "invoice_date": h.invoice_date,
            "total_ex_vat": h.total_ex_vat,
            "currency": h.currency,
            "document_type": h.document_type,
        }
        for h in headers
    ]

    return {
        "run_id": run_id,
        "scan_timestamp": scan_timestamp,
        "total_files_scanned": total_files,
        "files_processed": processed_files,
        "files_skipped": skipped_files,
        "files_failed": failed_files,
        "invoices": invoices,
        "carrier_totals": carrier_totals,
        "service_category_totals": {k: round(v, 2) for k, v in service_totals.items()},
        "surcharge_category_totals": {k: round(v, 2) for k, v in surcharge_totals.items()},
        "check_counts": check_counts,
        "anomaly_count": len(anomalies),
        "anomalies": anomaly_list,
        "log_counts": log_counts,
    }


def write_deterministic_summary(
    run_id: str,
    payload: dict,
    file_records: List[FileRecord],
    headers: List[BringInvoiceHeader],
    lines: List[BringInvoiceLine],
    checks: List[CheckResult],
    anomalies: List[Anomaly],
    logger: ProcessingLogger,
) -> Path:
    """Write deterministic Markdown summary to 02_Output/Summaries/."""
    out_path = config.SUMMARIES_DIR / f"summary_{run_id}_deterministic.md"

    lines_base = [ln for ln in lines if ln.line_type == "BaseFreight"]
    lines_surcharge = [ln for ln in lines if ln.line_type == "Surcharge"]
    total_all = sum(ln.amount or 0.0 for ln in lines)
    total_surcharge = sum(ln.amount or 0.0 for ln in lines_surcharge)
    surcharge_pct = (total_surcharge / total_all * 100) if total_all > 0 else 0.0

    error_checks = [c for c in checks if c.severity == "Error"]
    warn_checks = [c for c in checks if c.severity == "Warning"]

    md = []
    md.append(f"# Freight Invoice Reconciliation — Run {run_id}")
    md.append("")
    md.append(f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    md.append(f"**Scan timestamp:** {payload['scan_timestamp']}")
    md.append(f"**Claude API:** {'Enabled' if config.USE_CLAUDE_API else 'Disabled (deterministic only)'}")
    md.append("")

    md.append("## Files")
    md.append(f"- Total scanned: {payload['total_files_scanned']}")
    md.append(f"- Processed: {payload['files_processed']}")
    md.append(f"- Skipped (unsupported): {payload['files_skipped']}")
    md.append(f"- Failed: {payload['files_failed']}")
    md.append("")
    if file_records:
        md.append("| File | Carrier | Type | Invoice# | Status |")
        md.append("|------|---------|------|----------|--------|")
        for r in file_records:
            md.append(f"| {r.file_name} | {r.detected_carrier} | {r.detected_document_type} | {r.detected_invoice_number or '—'} | {r.processing_status} |")
        md.append("")

    md.append("## Carrier Summary")
    for carrier, info in payload["carrier_totals"].items():
        md.append(f"**{carrier}:** {info['invoices']} invoice(s), total ex VAT = {info['total_ex_vat']:,.2f} {info['currency']}")
    md.append("")

    md.append("## Invoice Line Summary")
    md.append(f"- Total lines: {len(lines)}")
    md.append(f"- Base freight lines: {len(lines_base)}")
    md.append(f"- Surcharge lines: {len(lines_surcharge)}")
    md.append(f"- Total amount (all lines): {total_all:,.2f}")
    md.append(f"- Total surcharges: {total_surcharge:,.2f} ({surcharge_pct:.1f}% of total)")
    md.append("")

    md.append("## Cost by Service Category")
    if payload["service_category_totals"]:
        md.append("| Service Category | Amount |")
        md.append("|-----------------|--------|")
        for cat, amt in sorted(payload["service_category_totals"].items(), key=lambda x: -x[1]):
            pct = (amt / total_all * 100) if total_all > 0 else 0
            md.append(f"| {cat} | {amt:,.2f} ({pct:.1f}%) |")
    else:
        md.append("_No service category data._")
    md.append("")

    md.append("## Cost by Surcharge Category")
    if payload["surcharge_category_totals"]:
        md.append("| Surcharge Category | Amount |")
        md.append("|-------------------|--------|")
        for cat, amt in sorted(payload["surcharge_category_totals"].items(), key=lambda x: -x[1]):
            pct = (amt / total_surcharge * 100) if total_surcharge > 0 else 0
            md.append(f"| {cat} | {amt:,.2f} ({pct:.1f}% of surcharges) |")
    else:
        md.append("_No surcharge data._")
    md.append("")

    # ── Per-invoice status table ──────────────────────────────────────────────
    md.append("## Invoice Status")
    ok_n = payload['check_counts'].get('OK', 0)
    warn_n = payload['check_counts'].get('Warning', 0)
    err_n = payload['check_counts'].get('Error', 0)
    md.append(f"Checks: {ok_n} OK · {warn_n} Warning · {err_n} Error")
    md.append("")
    if headers:
        # Group checks by invoice number
        from collections import defaultdict as _dd
        checks_by_inv = _dd(list)
        for c in checks:
            checks_by_inv[c.invoice_number].append(c)

        md.append("| Carrier | Invoice# | Date | Total ex VAT | Overall |")
        md.append("|---------|----------|------|-------------|---------|")
        for h in headers:
            inv_checks = checks_by_inv.get(h.invoice_number, [])
            if any(c.severity == "Error" for c in inv_checks):
                icon, overall = "✗", "Error"
            elif any(c.severity == "Warning" for c in inv_checks):
                icon, overall = "⚠", "Warning"
            else:
                icon, overall = "✓", "OK"
            md.append(
                f"| {h.carrier} | {h.invoice_number} | {h.invoice_date} "
                f"| {_fmt(h.total_ex_vat)} SEK | {icon} {overall} |"
            )
    md.append("")

    # ── Issues only (errors + warnings) ──────────────────────────────────────
    issue_checks = [c for c in checks if c.severity in ("Error", "Warning")]
    if issue_checks:
        md.append("## Issues Requiring Attention")
        from collections import defaultdict as _dd2
        by_inv = _dd2(list)
        for c in issue_checks:
            by_inv[c.invoice_number].append(c)
        for inv_num_key, inv_issues in sorted(by_inv.items()):
            carrier_name = next((c.carrier for c in inv_issues), "")
            md.append(f"### {carrier_name} — {inv_num_key}")
            for c in inv_issues:
                icon = "✗" if c.severity == "Error" else "⚠"
                md.append(f"- {icon} **{c.check_name}**: {c.message}")
                if c.claude_explanation:
                    md.append(f"  - _AI: {c.claude_explanation}_")
        md.append("")

    md.append("## Anomalies")
    if not anomalies:
        md.append("_No anomalies detected._")
    else:
        for a in anomalies:
            icon = {"Info": "ℹ", "Warning": "⚠", "Error": "✗"}.get(a.severity, "?")
            md.append(f"- {icon} **{a.anomaly_type}** ({a.severity}): {a.description}")
            if a.claude_explanation:
                md.append(f"  - **AI explanation:** {a.claude_explanation}")
            if a.suggested_action:
                md.append(f"  - _Action: {a.suggested_action}_")
    md.append("")

    md.append("## Recommended Actions")
    actions = []
    if error_checks:
        actions.append(f"Resolve {len(error_checks)} error check(s) before approving this invoice for payment.")
    if warn_checks:
        actions.append(f"Review {len(warn_checks)} warning check(s).")
    review_lines = [ln for ln in lines if ln.manual_review_required]
    if review_lines:
        actions.append(f"Manually review {len(review_lines)} invoice line(s) with unresolved classification.")
    if not actions:
        actions.append("No manual actions required — all checks passed.")
    for a in actions:
        md.append(f"- {a}")
    md.append("")

    md.append("---")
    md.append("_Generated by Freight Invoice Control MVP — deterministic output, no AI calculations._")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md))

    logger.info("SummaryWriter", f"Deterministic summary written: {out_path.name}")
    return out_path


def write_ai_summary(
    run_id: str,
    payload: dict,
    ai_text: str,
    logger: ProcessingLogger,
) -> Path:
    """Write the Claude-generated AI summary to 02_Output/Summaries/."""
    out_path = config.SUMMARIES_DIR / f"summary_{run_id}_ai.md"

    md = []
    md.append(f"# Freight Invoice — AI Management Summary — Run {run_id}")
    md.append("")
    md.append(f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    md.append(f"**Model:** {config.CLAUDE_MODEL}")
    md.append("")
    md.append("> **Note:** All figures in this summary were calculated by Python code.")
    md.append("> Claude was used only to write explanatory prose — it did not calculate any totals.")
    md.append("")
    md.append(ai_text)
    md.append("")
    md.append("---")
    md.append("_AI summary generated by Claude API. Financial figures are from deterministic Python calculations._")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md))

    logger.info("SummaryWriter", f"AI summary written: {out_path.name}")
    return out_path
