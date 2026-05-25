"""Generate per-run Excel export and HTML email body for Power Automate pickup."""

from __future__ import annotations

import html as html_lib
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, List

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from src import config
from src.processing_logger import ProcessingLogger


# ── Colour palette ─────────────────────────────────────────────────────────────
_WHITE     = "FFFFFFFF"
_GRAY_ROW  = "FFF5F7FB"   # alternating row tint
_GRAY_BORD = "FFE0E0E0"   # cell border

_HDR_FILL = PatternFill("solid", fgColor="FF1565C0")
_HDR_FONT = Font(bold=True, color=_WHITE, size=10)
_BOLD     = Font(bold=True)
_ALT_FILL = PatternFill("solid", fgColor=_GRAY_ROW)
_WFILL    = PatternFill("solid", fgColor=_WHITE)

# Subtle whole-row tints by status
_ROW_FILL = {
    "OK":           PatternFill("solid", fgColor="FFF0FAF2"),
    "Warning":      PatternFill("solid", fgColor="FFFFFCF0"),
    "Error":        PatternFill("solid", fgColor="FFFFF5F5"),
    "ManualReview": PatternFill("solid", fgColor="FFF8F0FF"),
}
# Stronger fills/fonts for the status cell itself
_STATUS_FILL = {
    "OK":           PatternFill("solid", fgColor="FFD4EDDA"),
    "Warning":      PatternFill("solid", fgColor="FFFFF3CD"),
    "Error":        PatternFill("solid", fgColor="FFF8D7DA"),
    "ManualReview": PatternFill("solid", fgColor="FFE2D9F3"),
}
_STATUS_FONT = {
    "OK":           Font(bold=True, color="FF155724"),
    "Warning":      Font(bold=True, color="FF856404"),
    "Error":        Font(bold=True, color="FF842029"),
    "ManualReview": Font(bold=True, color="FF432874"),
}
_STATUS_TEXT = {
    "OK": "✓ OK", "Warning": "⚠ Warning",
    "Error": "✗ Error", "ManualReview": "? Review",
}
_STATUS_ICON = {"OK": "✓", "Warning": "⚠", "Error": "✗"}

# Thin border applied to every data cell
_BSIDE  = Side(style="thin", color=_GRAY_BORD)
_BORDER = Border(left=_BSIDE, right=_BSIDE, top=_BSIDE, bottom=_BSIDE)


def _fmt(val: Any) -> str:
    if val is None:
        return ""
    if isinstance(val, float):
        return f"{val:,.2f}"
    return str(val)


def _e(val: Any) -> str:
    return html_lib.escape(str(val) if val is not None else "")


def _auto_width(ws, max_width: int = 55, padding: int = 3) -> None:
    for col_cells in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col_cells), default=10)
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(
            max_len + padding, max_width
        )


def _write_header_row(ws, row: int, values: list) -> None:
    for col, val in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = _HDR_FONT
        cell.fill = _HDR_FILL
        cell.border = _BORDER
        cell.alignment = Alignment(wrap_text=False)


def _style_data_rows(
    ws, first_row: int, last_row: int, num_cols: int,
    status_col: int | None = None,
) -> None:
    """Alternating row fills + borders + optional status colouring."""
    for i, r in enumerate(range(first_row, last_row + 1)):
        base = _ALT_FILL if i % 2 == 1 else _WFILL
        status_key = None
        if status_col:
            raw = str(ws.cell(row=r, column=status_col).value or "")
            for key in ("Error", "Warning", "ManualReview", "OK"):
                if key in raw:
                    status_key = key
                    break
        row_fill = _ROW_FILL.get(status_key, base) if status_key else base

        for col in range(1, num_cols + 1):
            cell = ws.cell(row=r, column=col)
            cell.fill = row_fill
            cell.border = _BORDER
            if status_col and col == status_col and status_key:
                cell.fill = _STATUS_FILL.get(status_key, row_fill)
                cell.font = _STATUS_FONT.get(status_key, _BOLD)


# ── Service breakdown helper ──────────────────────────────────────────────────

def _compute_service_breakdown(lines) -> list[dict]:
    """Group BaseFreight lines by service_category; compute count, total, avg, pct."""
    agg: dict = defaultdict(lambda: {"count": 0, "total": 0.0})
    for ln in lines:
        if getattr(ln, "line_type", "") == "BaseFreight":
            cat = getattr(ln, "service_category", "Unknown") or "Unknown"
            agg[cat]["count"] += 1
            agg[cat]["total"] += getattr(ln, "amount", 0.0) or 0.0
    grand_total = sum(v["total"] for v in agg.values())
    result = []
    for cat, v in sorted(agg.items(), key=lambda x: -x[1]["total"]):
        count, total = v["count"], v["total"]
        result.append({
            "category": cat,
            "count": count,
            "total": round(total, 2),
            "avg": round(total / count, 2) if count else 0.0,
            "pct": total / grand_total * 100 if grand_total else 0.0,
        })
    return result


# ── Sheet builders ─────────────────────────────────────────────────────────────

def _sheet_run_summary(ws, run_id: str, payload: dict, all_lines) -> None:
    total_amount = sum(ln.amount or 0.0 for ln in all_lines)
    surcharge_amount = sum(
        ln.amount or 0.0 for ln in all_lines if ln.line_type == "Surcharge"
    )
    surcharge_pct = surcharge_amount / total_amount * 100 if total_amount else 0.0
    c = payload.get("check_counts", {})
    ok_n, warn_n, err_n = c.get("OK", 0), c.get("Warning", 0), c.get("Error", 0)
    overall = "Error" if err_n else ("Warning" if warn_n else "OK")

    # Title banner row
    ws.merge_cells("A1:C1")
    title = ws["A1"]
    title.value = "Freight Invoice Control — Run Summary"
    title.font = Font(bold=True, size=13, color=_WHITE)
    title.fill = _HDR_FILL
    title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    title.border = _BORDER
    ws.row_dimensions[1].height = 26
    for col in (2, 3):
        ws.cell(row=1, column=col).fill = _HDR_FILL
        ws.cell(row=1, column=col).border = _BORDER

    sections = [
        ("Run Information", [
            ("Run ID",    run_id),
            ("Generated", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ]),
        ("Volume", [
            ("Files Scanned",     payload.get("total_files_scanned", 0)),
            ("Invoices Detected", len(payload.get("invoices", []))),
            ("Total Lines",       len(all_lines)),
        ]),
        ("Amounts (ex-VAT)", [
            ("Total Amount ex-VAT (SEK)",     round(total_amount, 2)),
            ("Total Surcharges ex-VAT (SEK)", round(surcharge_amount, 2)),
            ("Surcharge Share",               surcharge_pct / 100),
        ]),
        ("Validation", [
            ("Checks Passed", ok_n),
            ("Warnings",      warn_n),
            ("Errors",        err_n),
            ("Overall Status", _STATUS_TEXT.get(overall, overall)),
        ]),
    ]

    row = 3
    for sec_title, sec_rows in sections:
        ws.merge_cells(f"A{row}:C{row}")
        sh = ws.cell(row=row, column=1, value=sec_title)
        sh.font = Font(bold=True, size=9, color="FF1565C0")
        sh.fill = PatternFill("solid", fgColor="FFE3F0FF")
        sh.alignment = Alignment(indent=1)
        for col in range(1, 4):
            ws.cell(row=row, column=col).border = _BORDER
            ws.cell(row=row, column=col).fill = PatternFill("solid", fgColor="FFE3F0FF")
        row += 1

        for key, val in sec_rows:
            k = ws.cell(row=row, column=1, value=key)
            k.font = _BOLD
            k.fill = _WFILL
            k.border = _BORDER
            k.alignment = Alignment(indent=1)
            v = ws.cell(row=row, column=2, value=val)
            v.fill = _WFILL
            v.border = _BORDER
            ws.cell(row=row, column=3).border = _BORDER
            ws.cell(row=row, column=3).fill = _WFILL

            if key == "Overall Status":
                v.fill = _STATUS_FILL.get(overall, _WFILL)
                v.font = _STATUS_FONT.get(overall, _BOLD)
            elif "Amount" in key or "Surcharge" in key:
                if isinstance(val, (int, float)):
                    v.number_format = "#,##0.00"
                    v.alignment = Alignment(horizontal="right")
            elif key == "Surcharge Share":
                v.number_format = "0.0%"
                v.alignment = Alignment(horizontal="right")
            row += 1
        row += 1  # blank spacer

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 8
    ws.sheet_properties.tabColor = "1565C0"


def _sheet_invoice_details(ws, headers) -> None:
    cols = [
        "Carrier", "Invoice #", "Invoice Date", "Due Date",
        "Total ex-VAT (SEK)", "Currency", "Status", "Source File",
    ]
    _write_header_row(ws, 1, cols)
    ws.freeze_panes = "A2"

    STATUS_COL = 7
    for i, h in enumerate(headers, start=2):
        status_raw = getattr(h, "reconciliation_status", "") or ""
        row_vals = [
            h.carrier,
            h.invoice_number,
            str(h.invoice_date) if h.invoice_date else "",
            str(getattr(h, "due_date", "") or ""),
            h.total_ex_vat,
            h.currency,
            _STATUS_TEXT.get(status_raw, status_raw),
            h.source_file,
        ]
        for col, val in enumerate(row_vals, 1):
            cell = ws.cell(row=i, column=col, value=val)
            if col == 5 and isinstance(val, (int, float)):
                cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right")

    if headers:
        _style_data_rows(ws, 2, 1 + len(headers), len(cols), status_col=STATUS_COL)

    _auto_width(ws)
    ws.sheet_properties.tabColor = "0D47A1"


def _sheet_check_results(ws, checks) -> None:
    cols = ["Carrier", "Invoice #", "Check", "Status", "Message"]
    _write_header_row(ws, 1, cols)
    ws.freeze_panes = "A2"

    STATUS_COL = 4
    order = {"Error": 0, "Warning": 1, "OK": 2}
    sorted_checks = sorted(checks, key=lambda ch: order.get(ch.status, 9))

    for i, ch in enumerate(sorted_checks, start=2):
        row_vals = [
            ch.carrier, ch.invoice_number, ch.check_name,
            _STATUS_TEXT.get(ch.status, ch.status), ch.message,
        ]
        for col, val in enumerate(row_vals, 1):
            ws.cell(row=i, column=col, value=val)

    if checks:
        _style_data_rows(ws, 2, 1 + len(checks), len(cols), status_col=STATUS_COL)

    _auto_width(ws)
    worst = min(
        (ch.status for ch in checks),
        key=lambda s: {"Error": 0, "Warning": 1, "OK": 2}.get(s, 9),
        default="OK",
    )
    ws.sheet_properties.tabColor = {
        "OK": "2E7D32", "Warning": "E65100", "Error": "C62828",
    }.get(worst, "1565C0")


def _sheet_service_breakdown(ws, breakdown: list[dict]) -> None:
    cols = [
        "Service Type", "Shipments",
        "Total Cost ex-VAT (SEK)", "Avg / Shipment ex-VAT (SEK)", "% of Total",
    ]
    _write_header_row(ws, 1, cols)
    ws.freeze_panes = "A2"

    grand_count = sum(r["count"] for r in breakdown)
    grand_total = sum(r["total"] for r in breakdown)

    for i, r in enumerate(breakdown, start=2):
        ws.cell(row=i, column=1, value=r["category"])
        c2 = ws.cell(row=i, column=2, value=r["count"])
        c2.alignment = Alignment(horizontal="right")
        c3 = ws.cell(row=i, column=3, value=r["total"])
        c3.number_format = "#,##0.00"
        c3.alignment = Alignment(horizontal="right")
        c4 = ws.cell(row=i, column=4, value=r["avg"])
        c4.number_format = "#,##0.00"
        c4.alignment = Alignment(horizontal="right")
        c5 = ws.cell(row=i, column=5, value=r["pct"] / 100)
        c5.number_format = "0.0%"
        c5.alignment = Alignment(horizontal="right")

    if breakdown:
        _style_data_rows(ws, 2, 1 + len(breakdown), len(cols))

    # Total row
    last = 2 + len(breakdown)
    for col, val in enumerate(["TOTAL", grand_count, round(grand_total, 2), None, 1.0], 1):
        cell = ws.cell(row=last, column=col, value=val)
        cell.font = Font(bold=True, color=_WHITE)
        cell.fill = _HDR_FILL
        cell.border = _BORDER
        if col == 3:
            cell.number_format = "#,##0.00"
            cell.alignment = Alignment(horizontal="right")
        elif col == 5:
            cell.number_format = "0.0%"
            cell.alignment = Alignment(horizontal="right")
        elif col == 2:
            cell.alignment = Alignment(horizontal="right")

    _auto_width(ws)
    ws.sheet_properties.tabColor = "006064"


def _sheet_surcharge_breakdown(ws, payload: dict) -> None:
    surcharge_totals = payload.get("surcharge_category_totals", {})
    total = sum(surcharge_totals.values())

    cols = ["Surcharge Category", "Amount ex-VAT (SEK)", "% of Surcharges"]
    _write_header_row(ws, 1, cols)
    ws.freeze_panes = "A2"

    sorted_sc = sorted(surcharge_totals.items(), key=lambda x: -x[1])
    for i, (cat, amt) in enumerate(sorted_sc, start=2):
        ws.cell(row=i, column=1, value=cat)
        c2 = ws.cell(row=i, column=2, value=round(amt, 2))
        c2.number_format = "#,##0.00"
        c2.alignment = Alignment(horizontal="right")
        c3 = ws.cell(row=i, column=3, value=(amt / total) if total else 0.0)
        c3.number_format = "0.0%"
        c3.alignment = Alignment(horizontal="right")

    if surcharge_totals:
        _style_data_rows(ws, 2, 1 + len(surcharge_totals), len(cols))

    last = 2 + len(surcharge_totals)
    for col, val in enumerate(["TOTAL", round(total, 2), 1.0], 1):
        cell = ws.cell(row=last, column=col, value=val)
        cell.font = Font(bold=True, color=_WHITE)
        cell.fill = _HDR_FILL
        cell.border = _BORDER
        if col == 2:
            cell.number_format = "#,##0.00"
            cell.alignment = Alignment(horizontal="right")
        elif col == 3:
            cell.number_format = "0.0%"
            cell.alignment = Alignment(horizontal="right")

    _auto_width(ws)
    ws.sheet_properties.tabColor = "6A1B9A"


# ── Excel writer ──────────────────────────────────────────────────────────────

def _sheet_anomalies(ws, anomalies: list) -> None:
    """Anomalies sheet — one row per anomaly, colour-coded by severity."""
    cols = ["Severity", "Type", "Carrier", "Invoice #", "Description", "Detail", "Suggested Action"]
    _write_header_row(ws, 1, cols)
    ws.freeze_panes = "A2"

    sev_fill = {
        "Error":   PatternFill("solid", fgColor="FFF8D7DA"),
        "Warning": PatternFill("solid", fgColor="FFFFFDE7"),
        "Info":    PatternFill("solid", fgColor="FFE3F2FD"),
    }
    sev_font = {
        "Error":   Font(bold=True, color="FF842029"),
        "Warning": Font(bold=True, color="FF856404"),
        "Info":    Font(bold=True, color="FF1565C0"),
    }
    _sev_order = {"Error": 0, "Warning": 1, "Info": 2}
    sorted_anom = sorted(anomalies, key=lambda a: _sev_order.get(a.severity, 9))

    for i, a in enumerate(sorted_anom, start=2):
        row_fill = sev_fill.get(a.severity, _WFILL)
        vals = [a.severity, a.anomaly_type, a.carrier, a.invoice_number,
                a.description, a.detail or "", a.suggested_action or ""]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.fill = row_fill
            c.border = _BORDER
            c.alignment = Alignment(wrap_text=(j in (5, 6, 7)), vertical="top")
        ws.cell(row=i, column=1).font = sev_font.get(a.severity, _BOLD)

    worst = _sev_order.get(sorted_anom[0].severity, 9) if sorted_anom else 9
    ws.sheet_properties.tabColor = (
        "C62828" if worst == 0 else "F9A825" if worst == 1 else "1565C0"
    )
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 50
    ws.column_dimensions["F"].width = 35
    ws.column_dimensions["G"].width = 45
    for col_cells in ws.iter_cols(min_row=2):
        for c in col_cells:
            c.alignment = Alignment(
                wrap_text=(c.column in (5, 6, 7)), vertical="top"
            )


def _sheet_pending_files(ws, missing: list) -> None:
    """Pending Files sheet — compact note per incomplete Bring invoice."""
    _AMBER = PatternFill("solid", fgColor="FFF9A825")
    _AMBER_LIGHT = PatternFill("solid", fgColor="FFFFFDE7")

    ws.sheet_properties.tabColor = "F9A825"
    ws.merge_cells("A1:D1")
    title = ws["A1"]
    title.value = "⚠ Pending Files — action may be required"
    title.font = Font(bold=True, size=11, color=_WHITE)
    title.fill = _AMBER
    title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for col in range(1, 5):
        ws.cell(row=1, column=col).fill = _AMBER
        ws.cell(row=1, column=col).border = _BORDER
    ws.row_dimensions[1].height = 22

    cols = ["Invoice #", "Carrier", "Status", "Note"]
    _write_header_row(ws, 2, cols)
    ws.freeze_panes = "A3"

    for i, m in enumerate(missing, start=3):
        inv = m.get("invoice_number", "")
        found = m.get("found_file", "")
        missing_f = m.get("missing_file", "")
        note = f"{found} received — {missing_f} not yet found in inbox"
        vals = [inv, "Bring", "Incomplete", note]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.fill = _AMBER_LIGHT
            c.border = _BORDER
            c.alignment = Alignment(indent=1)
        ws.cell(row=i, column=3).font = Font(bold=True, color="FF856404")

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 60


def _write_excel(
    path: Path, run_id: str, payload: dict, headers, lines, checks,
    anomalies=None, missing_bring=None,
) -> None:
    wb = openpyxl.Workbook()

    ws1 = wb.active
    ws1.title = "Summary"
    _sheet_run_summary(ws1, run_id, payload, lines)

    ws2 = wb.create_sheet("Invoices")
    _sheet_invoice_details(ws2, headers)

    ws3 = wb.create_sheet("Check Results")
    _sheet_check_results(ws3, checks)

    ws4 = wb.create_sheet("Service Breakdown")
    _sheet_service_breakdown(ws4, _compute_service_breakdown(lines))

    ws5 = wb.create_sheet("Surcharge Breakdown")
    _sheet_surcharge_breakdown(ws5, payload)

    if anomalies:
        ws6 = wb.create_sheet("Anomalies")
        _sheet_anomalies(ws6, anomalies)

    if missing_bring:
        ws_p = wb.create_sheet("Pending Files")
        _sheet_pending_files(ws_p, missing_bring)

    wb.save(path)


# ── HTML email body ────────────────────────────────────────────────────────────

def _status_badge(status: str) -> str:
    cls = {
        "OK": "badge-ok", "Warning": "badge-warn",
        "Error": "badge-err", "ManualReview": "badge-review",
    }
    icon = {"OK": "✓", "Warning": "⚠", "Error": "✗", "ManualReview": "?"}
    c = cls.get(status, "badge-def")
    i = icon.get(status, "")
    label = "Review" if status == "ManualReview" else status
    return f'<span class="badge {c}">{i} {_e(label)}</span>'


def _md_to_html(text: str) -> str:
    """Convert the subset of markdown Claude produces into safe HTML."""
    import re
    lines_out = []
    in_ul = False
    for raw in text.splitlines():
        line = raw.strip()
        # Headers
        if line.startswith("### "):
            if in_ul:
                lines_out.append("</ul>"); in_ul = False
            lines_out.append(f"<h4>{_e(line[4:])}</h4>")
        elif line.startswith("## "):
            if in_ul:
                lines_out.append("</ul>"); in_ul = False
            lines_out.append(f"<h3>{_e(line[3:])}</h3>")
        elif line.startswith("# "):
            if in_ul:
                lines_out.append("</ul>"); in_ul = False
            lines_out.append(f"<h3>{_e(line[2:])}</h3>")
        # Bullet
        elif line.startswith("- ") or line.startswith("* "):
            if not in_ul:
                lines_out.append("<ul>"); in_ul = True
            inner = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", _e(line[2:]))
            lines_out.append(f"<li>{inner}</li>")
        elif line == "":
            if in_ul:
                lines_out.append("</ul>"); in_ul = False
            lines_out.append("")
        else:
            if in_ul:
                lines_out.append("</ul>"); in_ul = False
            inner = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", _e(line))
            lines_out.append(f"<p>{inner}</p>")
    if in_ul:
        lines_out.append("</ul>")
    return "\n".join(lines_out)


def _write_html(
    path: Path, run_id: str, payload: dict, headers, lines, checks,
    ai_summary: str | None = None,
    anomalies=None,
) -> None:
    from collections import defaultdict as _dd

    total_amount = sum(ln.amount or 0.0 for ln in lines)
    surcharge_amount = sum(
        ln.amount or 0.0 for ln in lines if ln.line_type == "Surcharge"
    )
    surcharge_pct = surcharge_amount / total_amount * 100 if total_amount else 0.0
    c = payload.get("check_counts", {})
    ok_n, warn_n, err_n = c.get("OK", 0), c.get("Warning", 0), c.get("Error", 0)
    overall = "Error" if err_n else ("Warning" if warn_n else "OK")
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    overall_colors = {
        "OK":      ("background:#e8f5e9", "#2e7d32", "✓ OK"),
        "Warning": ("background:#fff8e1", "#f57f17", "⚠ Warning"),
        "Error":   ("background:#ffebee", "#c62828", "✗ Error"),
    }
    ob_bg, ob_fg, ob_txt = overall_colors.get(overall, ("background:#f5f5f5", "#333", overall))

    # ── Invoices table ────────────────────────────────────────────────────────
    inv_rows = "".join(
        f"<tr>"
        f"<td>{_e(h.carrier)}</td>"
        f"<td>{_e(h.invoice_number)}</td>"
        f"<td>{_e(h.invoice_date)}</td>"
        f"<td class='ra'>{h.total_ex_vat:,.2f} {_e(h.currency)}</td>"
        f"<td>{_status_badge(getattr(h,'reconciliation_status',''))}</td>"
        f"</tr>\n"
        for h in headers
    )

    # ── Validation rows ───────────────────────────────────────────────────────
    checks_ok = sum(1 for c in checks if c.status == "OK")
    checks_flagged = [c for c in checks if c.status != "OK"]
    check_rows = "".join(
        f"<tr>"
        f"<td>{_e(c.invoice_number)}</td>"
        f"<td>{_e(c.check_name)}</td>"
        f"<td>{_status_badge(c.status)}</td>"
        f"<td>{_e(c.message)}</td>"
        f"</tr>\n"
        for c in checks_flagged
    )

    # ── Surcharge table ───────────────────────────────────────────────────────
    surcharge_totals = payload.get("surcharge_category_totals", {})
    total_sc = sum(surcharge_totals.values())
    sc_rows = "".join(
        f"<tr>"
        f"<td>{_e(cat)}</td>"
        f"<td class='ra'>{amt:,.2f}</td>"
        f"<td class='ra'>{amt/total_sc*100:.1f}%</td>"
        f"</tr>\n"
        for cat, amt in sorted(surcharge_totals.items(), key=lambda x: -x[1])
    )
    if sc_rows:
        sc_rows += (
            f"<tr class='tot'><td>TOTAL</td>"
            f"<td class='ra'>{total_sc:,.2f}</td>"
            f"<td class='ra'>100.0%</td></tr>\n"
        )

    # ── Service breakdown (fixed 5 columns — no dynamic columns) ─────────────
    svc_breakdown = _compute_service_breakdown(lines)
    sc_total_by_svc: dict = _dd(float)
    for ln in lines:
        if getattr(ln, "line_type", "") == "Surcharge":
            svc_cat = getattr(ln, "related_service_category", "") or "Unknown"
            sc_total_by_svc[svc_cat] += getattr(ln, "amount", 0.0) or 0.0

    svc_rows = ""
    for r in svc_breakdown:
        n = r["count"] or 1
        avg_base = r["total"] / n
        avg_sc = sc_total_by_svc.get(r["category"], 0.0) / n
        avg_total = avg_base + avg_sc
        svc_rows += (
            f"<tr><td>{_e(r['category'])}</td>"
            f"<td class='ra'>{r['count']:,}</td>"
            f"<td class='ra'>{r['total']:,.2f}</td>"
            f"<td class='ra'>{avg_base:,.2f}</td>"
            f"<td class='ra'>{avg_total:,.2f}</td></tr>\n"
        )

    # ── Unknown carrier alert (split out for prominent display) ──────────────
    anomalies = anomalies or []
    unknown_carrier_anomalies = [a for a in anomalies if a.anomaly_type == "UnknownCarrier"]
    other_anomalies = [a for a in anomalies if a.anomaly_type != "UnknownCarrier"]

    unknown_carrier_html = ""
    if unknown_carrier_anomalies:
        cards = []
        for a in unknown_carrier_anomalies:
            # Extract code recommendation from detail field
            code_rec = ""
            for part in a.detail.split(" | "):
                if part.startswith("Code recommendation:"):
                    code_rec = part[len("Code recommendation:"):].strip()
            rec_html = (
                f'<div style="margin-top:8px;padding:8px 10px;background:#fff3e0;'
                f'border-radius:4px;font-size:11px;color:#555">'
                f'<strong>Developer recommendation:</strong> {_e(code_rec)}</div>'
            ) if code_rec else ""
            cards.append(
                f'<div style="border-left:3px solid #c62828;background:#ffebee;'
                f'padding:12px 16px;border-radius:0 4px 4px 0;margin-bottom:10px">'
                f'<div style="font-size:12px;font-weight:700;color:#c62828">'
                f'✗ NEW CARRIER — {_e(a.carrier)} | Invoice {_e(a.invoice_number)}</div>'
                f'<div style="margin-top:6px;font-size:12px">{_e(a.description)}</div>'
                f'<div style="margin-top:6px;font-size:11px;color:#666">'
                f'<strong>Action:</strong> Verify all amounts against original file '
                f'before booking. AI-extracted data is not authoritative.</div>'
                f'{rec_html}'
                f'</div>'
            )
        unknown_carrier_html = (
            '<div style="background:#ffebee;border:2px solid #c62828;border-radius:6px;'
            'padding:14px 16px;margin-bottom:18px">'
            '<div style="font-size:13px;font-weight:700;color:#c62828;margin-bottom:10px">'
            '&#9888; ACTION REQUIRED — Invoice(s) from Unknown Carrier</div>'
            + "".join(cards)
            + '</div>'
        )

    # ── Anomalies ─────────────────────────────────────────────────────────────
    sev_icon_map = {"Warning": "⚠", "Error": "✗", "Info": "ℹ"}
    sev_colors = {
        "Warning": ("#fff8e1", "#f57f17"),
        "Error":   ("#ffebee", "#c62828"),
        "Info":    ("#e3f2fd", "#1565c0"),
    }

    anom_banner = ""
    anom_cards_html = ""
    if other_anomalies:
        sev_order = {"Error": 0, "Warning": 1, "Info": 2}
        worst = min(other_anomalies, key=lambda a: sev_order.get(a.severity, 9)).severity
        w_bg, w_fg = sev_colors.get(worst, ("#f5f5f5", "#333"))
        w_ic = sev_icon_map.get(worst, "?")
        n = len(other_anomalies)
        banner_cls = {"Warning": "anom-banner-warn", "Error": "anom-banner-err"}.get(worst, "anom-banner-warn")
        anom_banner = (
            f'<div class="anom-banner {banner_cls}">'
            f'{w_ic} {n} anomal{"y" if n == 1 else "ies"} detected — see details below'
            f'</div>'
        )
        anom_cls = {"Warning": "anom-warn", "Error": "anom-err", "Info": "anom-info"}
        cards = []
        for a in other_anomalies:
            ic = sev_icon_map.get(a.severity, "?")
            cls = anom_cls.get(a.severity, "anom-warn")
            exp_html = (
                f'<div class="anom-ai">{_e(a.claude_explanation)}</div>'
                if a.claude_explanation else ""
            )
            cards.append(
                f'<div class="anom-card {cls}">'
                f'<div class="anom-card-hdr">{ic} {_e(a.severity)} &middot; {_e(a.anomaly_type)}</div>'
                f'<div class="anom-card-body">{_e(a.description)}</div>'
                f'{exp_html}'
                f'</div>'
            )
        anom_cards_html = (
            f'<div class="anom-lbl">Anomalies ({n})</div>'
            + "".join(cards)
        )

    # ── Build HTML sections ───────────────────────────────────────────────────
    hdr_cls = {"OK": "status-ok", "Warning": "status-warn", "Error": "status-err"}
    hdr_txt = {"OK": "✓ All Clear", "Warning": "⚠ Needs Attention", "Error": "✗ Action Required"}
    hdr_status = (
        f'<span class="hdr-status {hdr_cls.get(overall, "status-warn")}">'
        f'{hdr_txt.get(overall, overall)}</span>'
    )

    svc_section = ""
    if svc_rows:
        svc_section = f"""
<div class="sec"><div class="sec-lbl">Cost by Service Type</div>
<div class="tbl-wrap"><table>
<thead><tr><th>Service</th><th class="ra">Shipments</th>
<th class="ra">Total ex-VAT</th><th class="ra">Avg Base</th><th class="ra">Avg Total</th></tr></thead>
<tbody>{svc_rows}</tbody></table></div></div>"""

    sc_section = ""
    if sc_rows:
        sc_section = f"""
<div class="sec"><div class="sec-lbl">Surcharge Breakdown</div>
<div class="tbl-wrap"><table>
<thead><tr><th>Category</th><th class="ra">Amount ex-VAT (SEK)</th><th class="ra">Share</th></tr></thead>
<tbody>{sc_rows}</tbody></table></div></div>"""

    checks_section = f"""
<div class="sec"><div class="sec-lbl">Validation</div>
<div class="val-ok">&#10003;&nbsp; {checks_ok} check{"s" if checks_ok != 1 else ""} passed</div>
{"<div class='tbl-wrap' style='margin-top:10px'><table><thead><tr><th>Invoice</th><th>Check</th><th>Status</th><th>Message</th></tr></thead><tbody>" + check_rows + "</tbody></table></div>" if check_rows else ""}
</div>"""

    anom_section = ""
    if anom_cards_html:
        anom_section = f'<div class="sec">{anom_cards_html}</div>'

    ai_section = ""
    if ai_summary:
        ai_section = f"""
<div class="sec"><div class="sec-lbl">AI Analysis</div>
<div class="ai-box">{_md_to_html(ai_summary)}</div></div>"""

    status_pill = {
        "OK":      ('<span style="background:#e6f9ee;color:#1a7a3a;padding:3px 10px;'
                    'border-radius:20px;font-size:11px;font-weight:700">&#10003; OK</span>'),
        "Warning": ('<span style="background:#fff8e0;color:#8a5e00;padding:3px 10px;'
                    'border-radius:20px;font-size:11px;font-weight:700">&#9888; Warning</span>'),
        "Error":   ('<span style="background:#fff0f0;color:#c0000a;padding:3px 10px;'
                    'border-radius:20px;font-size:11px;font-weight:700">&#10007; Action required</span>'),
    }.get(overall, "")

    content = f"""<!DOCTYPE html>
<html lang="sv">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,'Helvetica Neue',Arial,sans-serif;
     background:#f0f0f5;color:#1c1c1e;font-size:13px;-webkit-font-smoothing:antialiased}}
.shell{{padding:16px 8px}}
.card{{background:#fff;border-radius:14px;max-width:580px;margin:0 auto;overflow:hidden;
       border:1px solid #e5e5ea}}

/* Top bar */
.topbar{{padding:14px 20px;border-bottom:1px solid #f0f0f0;
         display:flex;align-items:center;justify-content:space-between}}
.topbar-left{{font-size:13px;font-weight:700;color:#1c1c1e;letter-spacing:-.2px}}
.topbar-left span{{font-size:11px;font-weight:400;color:#8e8e93;margin-left:6px}}

/* Stats row */
.stats{{display:table;width:100%;border-bottom:1px solid #f0f0f0}}
.stats-row{{display:table-row}}
.stat{{display:table-cell;text-align:center;padding:14px 6px;vertical-align:middle;
       border-right:1px solid #f0f0f0;width:25%}}
.stat:last-child{{border-right:none}}
.stat-num{{font-size:18px;font-weight:700;color:#111;letter-spacing:-.3px;line-height:1.1}}
.stat-sub{{font-size:10px;color:#aeaeb2;margin-top:2px;font-weight:500}}

/* Body */
.body{{padding:16px 20px}}
.sec{{margin-bottom:16px}}
.sec-lbl{{font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;
          color:#aeaeb2;margin-bottom:8px}}

/* Tables */
.tbl-wrap{{overflow-x:auto;-webkit-overflow-scrolling:touch;border-radius:8px;
           border:1px solid #ebebeb}}
table{{border-collapse:collapse;width:100%;font-size:12px}}
th{{text-align:left;padding:8px 12px;font-size:10px;font-weight:700;letter-spacing:.6px;
    text-transform:uppercase;color:#aeaeb2;background:#fafafa;
    border-bottom:1px solid #ebebeb;white-space:nowrap}}
td{{padding:8px 12px;border-bottom:1px solid #f5f5f5;color:#1c1c1e;vertical-align:middle}}
tr:last-child td{{border-bottom:none}}
.ra{{text-align:right}}
.tot td{{font-weight:700;background:#fafafa}}

/* Badges */
.badge{{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:600}}
.badge-ok{{background:#e6f9ee;color:#1a7a3a}}
.badge-warn{{background:#fff8e0;color:#8a5e00}}
.badge-err{{background:#fff0f0;color:#c0000a}}
.badge-review{{background:#f0eaff;color:#5e00c0}}
.badge-def{{background:#f2f2f7;color:#636366}}

/* Unknown carrier */
.unk-block{{background:#fff0f0;border:1px solid #fca5a5;border-radius:10px;
            padding:12px 14px;margin-bottom:14px}}
.unk-title{{font-size:12px;font-weight:700;color:#991b1b;margin-bottom:8px}}
.unk-card{{background:#fff;border-radius:6px;padding:10px 12px;margin-bottom:6px;
           border-left:3px solid #ef4444}}
.unk-card-hdr{{font-size:11px;font-weight:700;color:#991b1b;margin-bottom:3px}}
.unk-card-body{{font-size:11px;color:#555;line-height:1.4}}
.unk-rec{{background:#fff7ed;border-radius:6px;padding:6px 8px;margin-top:6px;
          font-size:10px;color:#78350f}}

/* Anomaly banner */
.anom-banner{{border-radius:8px;padding:8px 12px;margin-bottom:12px;font-size:12px;font-weight:600}}
.anom-banner-warn{{background:#fffbeb;border:1px solid #fde68a;color:#92400e}}
.anom-banner-err{{background:#fff1f1;border:1px solid #fecaca;color:#991b1b}}

/* Anomaly cards */
.anom-lbl{{font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;
           color:#c0000a;margin-bottom:8px}}
.anom-card{{border-radius:8px;padding:10px 12px;margin-bottom:6px}}
.anom-warn{{background:#fffbeb;border:1px solid #fde68a}}
.anom-err{{background:#fff1f1;border:1px solid #fecaca}}
.anom-info{{background:#eff6ff;border:1px solid #bfdbfe}}
.anom-card-hdr{{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-bottom:3px}}
.anom-warn .anom-card-hdr{{color:#92400e}}
.anom-err .anom-card-hdr{{color:#991b1b}}
.anom-info .anom-card-hdr{{color:#1e40af}}
.anom-card-body{{font-size:12px;color:#444;line-height:1.45}}
.anom-ai{{font-size:11px;color:#888;margin-top:5px;line-height:1.4}}

/* Validation */
.val-ok{{background:#e6f9ee;border-radius:8px;padding:8px 12px;
         font-size:12px;font-weight:600;color:#1a7a3a}}

/* AI box */
.ai-box{{background:#f9fdf0;border:1px solid #d4eaa0;border-radius:8px;
         padding:12px 14px;font-size:12px;line-height:1.6;color:#333}}
.ai-box h3,.ai-box h4{{font-size:12px;font-weight:700;color:#333;margin:10px 0 3px}}
.ai-box p,.ai-box ul{{margin:3px 0}}
.ai-box li{{margin:2px 0}}
.ai-box ul{{padding-left:16px}}

/* Footer */
.foot{{padding:10px 20px;border-top:1px solid #f0f0f0;text-align:center;
       font-size:10px;color:#c7c7cc}}
</style>
</head>
<body>
<div class="shell">
<div class="card">

<!-- Top bar (replaces dark header) -->
<div class="topbar">
  <div class="topbar-left">Freight Invoice Control <span>{_e(now_str)}</span></div>
  {status_pill}
</div>

<!-- Stats -->
<table class="stats"><tr class="stats-row">
  <td class="stat">
    <div class="stat-num">{total_amount:,.0f}</div>
    <div class="stat-sub">SEK ex-VAT</div>
    <div class="stat-sub">Total</div>
  </td>
  <td class="stat">
    <div class="stat-num">{len(payload.get('invoices', []))}</div>
    <div class="stat-sub">&nbsp;</div>
    <div class="stat-sub">Invoices</div>
  </td>
  <td class="stat">
    <div class="stat-num">{len(lines)}</div>
    <div class="stat-sub">&nbsp;</div>
    <div class="stat-sub">Lines</div>
  </td>
  <td class="stat">
    <div class="stat-num">{surcharge_pct:.1f}%</div>
    <div class="stat-sub">&nbsp;</div>
    <div class="stat-sub">Surcharges</div>
  </td>
</tr></table>

<!-- Body -->
<div class="body">

{unknown_carrier_html}
{anom_banner}

<div class="sec"><div class="sec-lbl">Invoices</div>
<div class="tbl-wrap"><table>
<thead><tr><th>Carrier</th><th>Invoice</th><th>Date</th>
<th class="ra">Amount ex-VAT (SEK)</th><th>Status</th></tr></thead>
<tbody>{inv_rows}</tbody>
</table></div></div>

{svc_section}
{sc_section}
{checks_section}
{anom_section}
{ai_section}

</div>

<div class="foot">Freight Invoice Control &middot; Isicom AB &middot; Python + Claude (Anthropic)</div>
</div>
</div>
</body>
</html>"""

    with open(path, "w", encoding="utf-8") as f:
        f.write(content)


# ── Public entry point ────────────────────────────────────────────────────────

def write_run_export(
    run_id: str,
    payload: dict,
    all_invoice_headers,
    all_invoice_lines,
    all_checks,
    logger: ProcessingLogger,
    ai_summary: str | None = None,
    anomalies=None,
    missing_bring: list | None = None,
) -> None:
    """Write XLSX into For_Email/ for Power Automate pickup. HTML saved to Summaries/ as archive."""
    config.FOR_EMAIL_DIR.mkdir(parents=True, exist_ok=True)
    xlsx_path = config.FOR_EMAIL_DIR / f"summary_{run_id}.xlsx"
    html_path = config.SUMMARIES_DIR / f"summary_{run_id}.html"

    _write_excel(xlsx_path, run_id, payload, all_invoice_headers, all_invoice_lines, all_checks,
                 anomalies=anomalies or [], missing_bring=missing_bring or [])
    _write_html(html_path, run_id, payload, all_invoice_headers, all_invoice_lines, all_checks,
                ai_summary=ai_summary, anomalies=anomalies or [])

    logger.info("RunExporter", f"Power Automate file: {xlsx_path.name} | HTML archive: {html_path.name}")


def write_missing_file_alert(
    run_id: str,
    missing: list[dict],
    logger: ProcessingLogger,
) -> None:
    """
    Write a For_Email XLSX alert when a Bring invoice is missing its PDF or Excel pair.
    Each entry in `missing` must have keys: invoice_number, missing_file, found_file, message.
    Power Automate picks up the XLSX and sends it as an email attachment.
    """
    if not missing:
        return

    config.FOR_EMAIL_DIR.mkdir(parents=True, exist_ok=True)
    path = config.FOR_EMAIL_DIR / f"alert_missing_files_{run_id}.xlsx"
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    n = len(missing)

    _RED_FILL = PatternFill("solid", fgColor="FFC62828")

    wb = openpyxl.Workbook()

    # Sheet 1 — Alert summary
    ws1 = wb.active
    ws1.title = "Alert"
    ws1.sheet_properties.tabColor = "C62828"

    ws1.merge_cells("A1:C1")
    title = ws1["A1"]
    title.value = "⚠ ACTION REQUIRED — Missing Invoice File(s)"
    title.font = Font(bold=True, size=13, color=_WHITE)
    title.fill = _RED_FILL
    title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    title.border = _BORDER
    for col in (2, 3):
        ws1.cell(row=1, column=col).fill = _RED_FILL
        ws1.cell(row=1, column=col).border = _BORDER
    ws1.row_dimensions[1].height = 26

    sections = [
        ("Run Information", [
            ("Run ID",    run_id),
            ("Generated", now_str),
        ]),
        ("Status", [
            ("Incomplete invoices", n),
        ]),
        ("Action Required", [
            ("Step 1", "Get the missing file from Bring Business or your account manager."),
            ("Step 2", "Place it in the 01_Inbox folder."),
            ("Step 3", "Re-run Freight Invoice Control — both files will be processed together."),
        ]),
    ]
    row = 3
    for sec_title, sec_rows in sections:
        ws1.merge_cells(f"A{row}:C{row}")
        sh = ws1.cell(row=row, column=1, value=sec_title)
        sh.font = Font(bold=True, size=9, color="FFC62828")
        sh.fill = PatternFill("solid", fgColor="FFFFF0F0")
        sh.alignment = Alignment(indent=1)
        for col in range(1, 4):
            ws1.cell(row=row, column=col).border = _BORDER
            ws1.cell(row=row, column=col).fill = PatternFill("solid", fgColor="FFFFF0F0")
        row += 1
        for key, val in sec_rows:
            k = ws1.cell(row=row, column=1, value=key)
            k.font = _BOLD; k.fill = _WFILL; k.border = _BORDER
            k.alignment = Alignment(indent=1)
            v = ws1.cell(row=row, column=2, value=val)
            v.fill = _WFILL; v.border = _BORDER
            ws1.cell(row=row, column=3).border = _BORDER
            ws1.cell(row=row, column=3).fill = _WFILL
            row += 1
        row += 1

    ws1.column_dimensions["A"].width = 22
    ws1.column_dimensions["B"].width = 65
    ws1.column_dimensions["C"].width = 8

    # Sheet 2 — Incomplete invoices detail
    ws2 = wb.create_sheet("Incomplete Invoices")
    ws2.sheet_properties.tabColor = "C62828"
    cols = ["Invoice #", "Missing File", "Received File", "Details"]
    _write_header_row(ws2, 1, cols)
    ws2.freeze_panes = "A2"

    for i, m in enumerate(missing, start=2):
        ws2.cell(row=i, column=1, value=m.get("invoice_number", ""))
        c2 = ws2.cell(row=i, column=2, value=m.get("missing_file", ""))
        c2.font = Font(color="FFC62828")
        c3 = ws2.cell(row=i, column=3, value=m.get("found_file", ""))
        c3.font = Font(color="FF2E7D32")
        ws2.cell(row=i, column=4, value=m.get("message", ""))

    if missing:
        _style_data_rows(ws2, 2, 1 + len(missing), len(cols))

    _auto_width(ws2)
    wb.save(path)

    logger.warning("RunExporter",
                   f"Missing-file alert written: {path.name} ({n} incomplete invoice(s))")
