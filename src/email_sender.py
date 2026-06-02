"""Send run summary email via Outlook desktop app (COM automation)."""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from src import config
from src.processing_logger import ProcessingLogger


_RECIPIENT = "ludvig@isicom.se"


def send_summary_email(
    run_id: str,
    summary_md_path: Path,
    xlsx_path: Optional[Path],
    logger: ProcessingLogger,
    check_counts: dict,
    total_amount: float,
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
        ok_n  = check_counts.get("OK", 0)
        warn_n = check_counts.get("Warning", 0)
        err_n  = check_counts.get("Error", 0)

        if err_n:
            status_label = "✗ Action Required"
            status_color = "#c62828"
        elif warn_n:
            status_label = "⚠ Review Needed"
            status_color = "#e65100"
        else:
            status_label = "✓ All Clear"
            status_color = "#2e7d32"

        issues = []
        if err_n:   issues.append(f"{err_n} error(s) requiring action")
        if warn_n:  issues.append(f"{warn_n} warning(s) to review")
        issue_line = ", ".join(issues) if issues else "all checks passed"

        subject = (
            f"Freight Invoice Control — "
            f"{run_id[:10]} — {status_label}"
        )

        html_body = f"""<html><body style="font-family:Segoe UI,Arial,sans-serif;
font-size:14px;line-height:1.6;color:#1c1c1e;padding:24px;max-width:560px">
<p style="font-size:17px;font-weight:700;color:{status_color};margin:0 0 10px">
  {status_label}
</p>
<table style="border-collapse:collapse;margin-bottom:16px">
  <tr>
    <td style="padding:3px 16px 3px 0;color:#666;font-size:13px">Total ex VAT</td>
    <td style="padding:3px 0;font-weight:700">{total_amount:,.0f} SEK</td>
  </tr>
  <tr>
    <td style="padding:3px 16px 3px 0;color:#666;font-size:13px">Checks</td>
    <td style="padding:3px 0">{ok_n} OK &nbsp;·&nbsp;
      <span style="color:#e65100">{warn_n} Warning</span> &nbsp;·&nbsp;
      <span style="color:#c62828">{err_n} Error</span></td>
  </tr>
  <tr>
    <td style="padding:3px 16px 3px 0;color:#666;font-size:13px">Status</td>
    <td style="padding:3px 0;color:{status_color};font-weight:600">{issue_line}</td>
  </tr>
</table>
<p style="font-size:13px;color:#444;margin:0 0 6px">
  The full invoice approval report is attached as an Excel file.<br>
  It contains one sheet per invoice with service breakdown, surcharges,
  validation results, and anomalies.
</p>
<p style="font-size:11px;color:#aaa;margin:20px 0 0">
  Run {run_id} &nbsp;·&nbsp; Freight Invoice Control &nbsp;·&nbsp; Isicom AB
</p>
</body></html>"""

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
