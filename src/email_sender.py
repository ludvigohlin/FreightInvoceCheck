"""Send run summary email via Outlook desktop app (COM automation)."""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from src import config
from src.processing_logger import ProcessingLogger


_RECIPIENT = "ludvig@isicom.se"

_STYLE = (
    'font-family:Segoe UI,Arial,sans-serif;font-size:14px;'
    'line-height:1.6;color:#1c1c1e;padding:24px;max-width:560px'
)


def _log_row(log_counts: dict | None) -> str:
    """Return an HTML table row for log warnings/errors, or empty string if none."""
    lc = log_counts or {}
    le, lw = lc.get("ERROR", 0), lc.get("WARNING", 0)
    if not le and not lw:
        return ""
    return (
        f'<tr><td style="padding:3px 16px 3px 0;color:#666;font-size:13px">Körningslogg</td>'
        f'<td style="padding:3px 0;color:#c62828;font-size:13px">'
        f'{le} fel &nbsp;·&nbsp; {lw} varningar — kontrollera 03_Logs/</td></tr>'
    )


def _pending_table(pending_items: list) -> str:
    if not pending_items:
        return ""
    rows = "".join(
        f"<tr><td style='padding:3px 8px'>{p.get('invoice_number','')}</td>"
        f"<td style='padding:3px 8px;color:#666'>{p.get('found_file','')}</td>"
        f"<td style='padding:3px 8px;color:#c62828'>{p.get('missing_file','')}</td></tr>"
        for p in pending_items
    )
    return (
        "<p style='font-size:13px;font-weight:700;color:#333;margin:14px 0 4px'>"
        "Inväntar dokument</p>"
        "<table style='border-collapse:collapse;font-size:12px;width:100%'>"
        "<tr style='background:#f5f5f5'>"
        "<th style='padding:3px 8px;text-align:left'>Faktura</th>"
        "<th style='padding:3px 8px;text-align:left'>Mottaget</th>"
        "<th style='padding:3px 8px;text-align:left'>Saknas</th></tr>"
        f"{rows}</table>"
    )


def send_idle_email(
    run_id: str,
    logger: ProcessingLogger,
    reason: str = "Inga filer i inkorg.",
    log_counts: dict | None = None,
    pending_items: list | None = None,
) -> bool:
    """
    Send a brief 'nothing new' status email so the user knows the job ran.
    Returns True on success, False on failure (non-fatal).
    """
    if not config.SEND_EMAIL:
        logger.info("EmailSender", "Email sending disabled (SEND_EMAIL=false).")
        return False

    try:
        import win32com.client  # type: ignore
    except ImportError:
        logger.warning("EmailSender", "pywin32 not installed — email skipped.")
        return False

    try:
        lc = log_counts or {}
        has_issues = lc.get("ERROR", 0) or lc.get("WARNING", 0)
        status_color = "#c62828" if lc.get("ERROR", 0) else "#e65100" if has_issues else "#546e7a"

        subject = f"Freight Invoice Control — {run_id[:10]} — Ingen ny faktura"

        html_body = (
            f'<html><body style="{_STYLE}">'
            f'<p style="font-size:17px;font-weight:700;color:{status_color};margin:0 0 10px">'
            f'Ingen ny faktura</p>'
            f'<p style="font-size:13px;color:#555;margin:0 0 8px">{reason}</p>'
            f'<table style="border-collapse:collapse;margin-bottom:16px">'
            f'{_log_row(log_counts)}'
            f'</table>'
            f'{_pending_table(pending_items or [])}'
            f'<p style="font-size:11px;color:#aaa;margin:20px 0 0">'
            f'Run {run_id} &nbsp;·&nbsp; Freight Invoice Control &nbsp;·&nbsp; Isicom AB</p>'
            f'</body></html>'
        )

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = _RECIPIENT
        mail.Subject = subject
        mail.HTMLBody = html_body
        mail.Send()
        logger.info("EmailSender", f"Idle status email sent to {_RECIPIENT}")
        return True

    except Exception as e:
        logger.warning("EmailSender", f"Email send failed: {e}", error=e)
        return False


def send_summary_email(
    run_id: str,
    summary_md_path: Path,
    xlsx_path: Optional[Path],
    logger: ProcessingLogger,
    check_counts: dict,
    total_amount: float,
    log_counts: dict | None = None,
) -> bool:
    """
    Send a brief email with the Excel approval report attached.
    Returns True on success, False on failure (non-fatal).
    """
    if not config.SEND_EMAIL:
        logger.info("EmailSender", "Email sending disabled (SEND_EMAIL=false).")
        return False

    try:
        import win32com.client  # type: ignore
    except ImportError:
        logger.warning("EmailSender", "pywin32 not installed — email skipped.")
        return False

    try:
        ok_n   = check_counts.get("OK", 0)
        warn_n = check_counts.get("Warning", 0)
        err_n  = check_counts.get("Error", 0)

        if err_n:
            status_label = "Action Required"
            status_color = "#c62828"
        elif warn_n:
            status_label = "Review Needed"
            status_color = "#e65100"
        else:
            status_label = "All Clear"
            status_color = "#2e7d32"

        issues = []
        if err_n:   issues.append(f"{err_n} error(s) requiring action")
        if warn_n:  issues.append(f"{warn_n} warning(s) to review")
        issue_line = ", ".join(issues) if issues else "all checks passed"

        subject = f"Freight Invoice Control — {run_id[:10]} — {status_label}"

        attach_note = (
            "The full invoice approval report is attached as an Excel file.<br>"
            "It contains one sheet per invoice with service breakdown, surcharges, "
            "validation results, and anomalies."
            if xlsx_path and xlsx_path.exists()
            else "No Excel report was generated for this run."
        )

        html_body = (
            f'<html><body style="{_STYLE}">'
            f'<p style="font-size:17px;font-weight:700;color:{status_color};margin:0 0 10px">'
            f'{status_label}</p>'
            f'<table style="border-collapse:collapse;margin-bottom:16px">'
            f'<tr><td style="padding:3px 16px 3px 0;color:#666;font-size:13px">Total ex VAT</td>'
            f'<td style="padding:3px 0;font-weight:700">{total_amount:,.0f} SEK</td></tr>'
            f'<tr><td style="padding:3px 16px 3px 0;color:#666;font-size:13px">Checks</td>'
            f'<td style="padding:3px 0">{ok_n} OK &nbsp;·&nbsp;'
            f'<span style="color:#e65100">{warn_n} Warning</span> &nbsp;·&nbsp;'
            f'<span style="color:#c62828">{err_n} Error</span></td></tr>'
            f'<tr><td style="padding:3px 16px 3px 0;color:#666;font-size:13px">Status</td>'
            f'<td style="padding:3px 0;color:{status_color};font-weight:600">{issue_line}</td></tr>'
            f'{_log_row(log_counts)}'
            f'</table>'
            f'<p style="font-size:13px;color:#444;margin:0 0 6px">{attach_note}</p>'
            f'<p style="font-size:11px;color:#aaa;margin:20px 0 0">'
            f'Run {run_id} &nbsp;·&nbsp; Freight Invoice Control &nbsp;·&nbsp; Isicom AB</p>'
            f'</body></html>'
        )

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = _RECIPIENT
        mail.Subject = subject
        mail.HTMLBody = html_body

        if xlsx_path and xlsx_path.exists():
            mail.Attachments.Add(str(xlsx_path))

        mail.Send()
        logger.info("EmailSender", f"Summary email sent to {_RECIPIENT}")
        return True

    except Exception as e:
        logger.warning("EmailSender", f"Email send failed: {e}", error=e)
        return False
