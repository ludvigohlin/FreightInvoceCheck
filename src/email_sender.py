"""Send run summary via Outlook desktop app using COM automation."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional

from src import config
from src.processing_logger import ProcessingLogger


_RECIPIENT = "ludvig@isicom.se"


def _md_to_html(text: str) -> str:
    """Convert the summary markdown subset to clean email HTML."""
    lines_out: list[str] = []
    in_table = False
    in_ul = False

    def flush_list():
        nonlocal in_ul
        if in_ul:
            lines_out.append("</ul>")
            in_ul = False

    def flush_table():
        nonlocal in_table
        if in_table:
            lines_out.append("</tbody></table>")
            in_table = False

    header_style = "font-family:Segoe UI,Arial,sans-serif;margin:18px 0 4px"
    table_style = (
        "border-collapse:collapse;width:100%;font-family:Segoe UI,Arial,sans-serif;"
        "font-size:13px;margin-bottom:12px"
    )
    th_style = (
        "background:#1565C0;color:#fff;padding:6px 10px;text-align:left;"
        "font-weight:600;font-size:12px"
    )
    td_style = "padding:5px 10px;border-bottom:1px solid #e8e8e8"

    STATUS_COLORS = {
        "✓ OK": "#e8f5e9", "⚠ Warning": "#fff8e1", "✗ Error": "#ffebee",
    }

    for raw in text.splitlines():
        line = raw.strip()

        # Skip separator lines
        if line == "---":
            flush_list(); flush_table(); continue

        # H1
        if line.startswith("# ") and not line.startswith("##"):
            flush_list(); flush_table()
            lines_out.append(
                f'<h2 style="font-family:Segoe UI,Arial,sans-serif;color:#1565C0;'
                f'margin:0 0 4px;font-size:18px">{line[2:]}</h2>'
            )
            continue

        # H2
        if line.startswith("## ") and not line.startswith("###"):
            flush_list(); flush_table()
            lines_out.append(
                f'<h3 style="{header_style};color:#333;font-size:15px">{line[3:]}</h3>'
            )
            continue

        # H3
        if line.startswith("### "):
            flush_list(); flush_table()
            lines_out.append(
                f'<h4 style="{header_style};color:#555;font-size:13px">{line[4:]}</h4>'
            )
            continue

        # Markdown table header row
        if line.startswith("|") and "---" not in line:
            if not in_table:
                flush_list()
                in_table = True
                lines_out.append(f'<table style="{table_style}"><thead><tr>')
                cells = [c.strip() for c in line.strip("|").split("|")]
                for cell in cells:
                    lines_out.append(f'<th style="{th_style}">{cell}</th>')
                lines_out.append("</tr></thead><tbody>")
            else:
                cells = [c.strip() for c in line.strip("|").split("|")]
                # Detect status cell for row colour
                row_bg = ""
                for c in cells:
                    if c in STATUS_COLORS:
                        row_bg = f'background:{STATUS_COLORS[c]}'
                        break
                row_style = f' style="{row_bg}"' if row_bg else ""
                lines_out.append(f"<tr{row_style}>")
                for cell in cells:
                    # Bold the status cell
                    if cell in STATUS_COLORS:
                        lines_out.append(
                            f'<td style="{td_style};font-weight:700">{cell}</td>'
                        )
                    else:
                        lines_out.append(f'<td style="{td_style}">{cell}</td>')
                lines_out.append("</tr>")
            continue

        # Table separator — skip
        if line.startswith("|") and "---" in line:
            continue

        # Bullet
        if line.startswith("- ") or line.startswith("* "):
            flush_table()
            if not in_ul:
                lines_out.append(
                    '<ul style="font-family:Segoe UI,Arial,sans-serif;'
                    'font-size:13px;margin:4px 0;padding-left:18px">'
                )
                in_ul = True
            inner = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", line[2:])
            inner = re.sub(r"_(.+?)_", r"<em>\1</em>", inner)
            lines_out.append(f"<li>{inner}</li>")
            continue

        # Bold key:value (**key:** value)
        if line.startswith("**") and "**" in line[2:]:
            flush_list(); flush_table()
            inner = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", line)
            inner = re.sub(r"_(.+?)_", r"<em>\1</em>", inner)
            lines_out.append(
                f'<p style="font-family:Segoe UI,Arial,sans-serif;'
                f'font-size:13px;margin:2px 0">{inner}</p>'
            )
            continue

        # Blank line
        if not line:
            flush_list(); flush_table()
            continue

        # Default paragraph
        flush_table()
        inner = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", line)
        inner = re.sub(r"_(.+?)_", r"<em>\1</em>", inner)
        lines_out.append(
            f'<p style="font-family:Segoe UI,Arial,sans-serif;'
            f'font-size:13px;margin:2px 0;color:#444">{inner}</p>'
        )

    flush_list()
    flush_table()
    return "\n".join(lines_out)


def send_summary_email(
    run_id: str,
    summary_md_path: Path,
    xlsx_path: Optional[Path],
    logger: ProcessingLogger,
    check_counts: dict,
    total_amount: float,
) -> bool:
    """
    Send the run summary via Outlook (COM automation).
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
        md_text = summary_md_path.read_text(encoding="utf-8")

        ok_n = check_counts.get("OK", 0)
        warn_n = check_counts.get("Warning", 0)
        err_n = check_counts.get("Error", 0)

        if err_n:
            status_label = "✗ Action Required"
            status_color = "#c62828"
            status_bg = "#ffebee"
        elif warn_n:
            status_label = "⚠ Review Needed"
            status_color = "#e65100"
            status_bg = "#fff8e1"
        else:
            status_label = "✓ All Clear"
            status_color = "#2e7d32"
            status_bg = "#e8f5e9"

        subject = f"Freight Invoice Control — {run_id[:10]} — {status_label}"

        html_body = f"""<html><body style="margin:0;padding:0;background:#f5f5f5">
<div style="max-width:620px;margin:20px auto;background:#fff;border-radius:8px;
            border:1px solid #e0e0e0;overflow:hidden;font-family:Segoe UI,Arial,sans-serif">

  <!-- Header -->
  <div style="background:#1565C0;padding:16px 24px">
    <div style="color:#fff;font-size:17px;font-weight:700">Freight Invoice Control</div>
    <div style="color:#90caf9;font-size:12px;margin-top:2px">Run {run_id}</div>
  </div>

  <!-- Status bar -->
  <div style="background:{status_bg};padding:10px 24px;border-bottom:1px solid #e0e0e0;
              display:flex;align-items:center">
    <span style="font-size:15px;font-weight:700;color:{status_color}">{status_label}</span>
    <span style="margin-left:16px;font-size:12px;color:#666">
      {ok_n} OK &nbsp;·&nbsp; {warn_n} Warning &nbsp;·&nbsp; {err_n} Error
      &nbsp;·&nbsp; {total_amount:,.0f} SEK ex-VAT
    </span>
  </div>

  <!-- Body -->
  <div style="padding:20px 24px">
    {_md_to_html(md_text)}
  </div>

  <div style="padding:10px 24px;border-top:1px solid #f0f0f0;
              font-size:11px;color:#aaa;text-align:center">
    Freight Invoice Control · Isicom AB · Generated automatically
  </div>
</div>
</body></html>"""

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
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
