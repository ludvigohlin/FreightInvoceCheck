"""Map run data to SummaryInput and generate the freight invoice Excel report."""

from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

from src import config
from src.processing_logger import ProcessingLogger
from src.freight_summary import (
    build_summary, SummaryInput,
    Invoice, Service, Surcharge, Anomaly as SummaryAnomaly, Unallocated, ServiceSurcharge,
)


# ── Swedish label maps ────────────────────────────────────────────────────────

_SVC_LABEL = {
    "Parcel":        "Kolli (parcel)",
    "Pallet":        "Pall (pallet)",
    "Service Point": "Utlämningsställe",
    "Parcel Locker": "Paketbox",
    "Pickup":        "Upphämtning",
    "Return":        "Retur",
    "Other":         "Övrigt",
    "Unknown":       "Oklassificerat",
}

_SC_LABEL = {
    "Fuel":             "Bränsle (fuel)",
    "Remote Area":      "Avlägset område",
    "City":             "City (storstadstillägg)",
    "Special Handling": "Specialhantering",
    "Delivery Attempt": "Leveransförsök",
    "Notification":     "Avisering",
    "Box Address":      "Boxadress",
    "Currency":         "Valuta",
    "Sulphur":          "Svavel",
    "Heavy":            "Tungt gods",
    "Return":           "Retur",
    "Other":            "Övrigt",
    "Unknown":          "Okänt tillägg",
}


# ── Data mapper ───────────────────────────────────────────────────────────────

def _build_summary_input(
    run_id: str,
    payload: dict,
    headers,
    lines,
    checks,
    anomalies,
    missing_bring: list,
    all_lines_dict: dict | None,
) -> SummaryInput:
    checks_by_inv: dict = defaultdict(list)
    for c in checks:
        checks_by_inv[c.invoice_number].append(c)
    anom_by_inv: dict = defaultdict(list)
    for a in (anomalies or []):
        anom_by_inv[a.invoice_number].append(a)

    # ── Invoices ──────────────────────────────────────────────────────────────
    invoices: list[Invoice] = []
    for h in headers:
        inv_checks = checks_by_inv.get(h.invoice_number, [])
        inv_anom   = anom_by_inv.get(h.invoice_number, [])

        recon_chk = next(
            (c for c in inv_checks
             if c.check_name in ("PDFTotalVsExcelSummary", "LineSumVsHeaderTotal")), None
        )
        recon_status = recon_chk.status if recon_chk else "OK"
        # Short "Att kolla" text — just the key fact, no AI prose
        if not recon_chk or recon_chk.status == "OK":
            recon_msg = ""
        elif "no lines" in recon_chk.message.lower() or "0.00" in recon_chk.actual_value:
            recon_msg = "Inga rader tolkade"
        elif recon_chk.difference:
            recon_msg = f"Differens: {recon_chk.difference} SEK"
        else:
            recon_msg = recon_chk.message[:60]

        n_err_check  = sum(1 for c in inv_checks if c.severity == "Error")
        n_warn_check = sum(1 for c in inv_checks if c.severity == "Warning")
        n_anom_warn  = sum(1 for a in inv_anom if a.severity in ("Error", "Warning"))
        inv_status   = (
            "Error"   if n_err_check else
            "Warning" if (n_warn_check or n_anom_warn) else
            "OK"
        )

        invoices.append(Invoice(
            carrier       = h.carrier,
            number        = h.invoice_number or "",
            date          = str(h.invoice_date or ""),
            total_ex_vat  = h.total_ex_vat or 0.0,
            recon_status  = recon_status,
            recon_message = recon_msg,
            n_anomalies   = len(inv_anom),
            status        = inv_status,
        ))

    # ── Services: aggregate across ALL invoices per (carrier, service_category) ─
    # Aggregating per-invoice creates duplicate rows when a carrier has multiple
    # invoices. Shipment IDs are prefixed with invoice number to avoid collisions.
    services: list[Service]        = []
    unallocated: list[Unallocated] = []

    svc_agg: dict = defaultdict(lambda: {"shipments": set(), "total": 0.0, "lines": 0})

    # Headers with 0 parseable lines → unallocated
    header_map = {(h.carrier, h.invoice_number): h for h in headers}
    line_source = all_lines_dict or {}

    for (carrier, inv_num), inv_lines in line_source.items():
        base_lines = [ln for ln in inv_lines if getattr(ln, "line_type", "") == "BaseFreight"]

        if not base_lines:
            h = header_map.get((carrier, inv_num))
            if h and (h.total_ex_vat or 0.0) > 0:
                unallocated.append(Unallocated(
                    carrier = carrier,
                    label   = f"Ej specificerat – faktura {inv_num}",
                    amount  = h.total_ex_vat or 0.0,
                ))
            continue

        for ln in base_lines:
            cat  = getattr(ln, "service_category", None) or "Unknown"
            # Bring uses shipment_number; PostNord uses kolli_id — try both
            ship = (getattr(ln, "shipment_number", None) or
                    getattr(ln, "kolli_id", None))
            if ship:
                svc_agg[(carrier, cat)]["shipments"].add(f"{inv_num}|{ship}")
            else:
                # No ID available — count each line as one sändning
                svc_agg[(carrier, cat)]["shipments"].add(f"{inv_num}|line_{id(ln)}")
            svc_agg[(carrier, cat)]["total"] += getattr(ln, "amount", 0.0) or 0.0
            svc_agg[(carrier, cat)]["lines"] += 1

    for (carrier, cat), d in sorted(svc_agg.items(), key=lambda x: -x[1]["total"]):
        services.append(Service(
            carrier      = carrier,
            service_name = _SVC_LABEL.get(cat, cat),
            shipments    = len(d["shipments"]),
            total_ex_vat = round(d["total"], 2),
            packages     = d["lines"],
        ))

    # Reconciliation gaps (invoices where line sum ≠ header total) → unallocated
    for h in headers:
        if not line_source.get((h.carrier, h.invoice_number)):
            continue  # already handled above
        inv_checks = checks_by_inv.get(h.invoice_number, [])
        recon_chk = next(
            (c for c in inv_checks
             if c.check_name in ("PDFTotalVsExcelSummary", "LineSumVsHeaderTotal")
             and c.status in ("Warning", "Error")), None
        )
        if recon_chk:
            try:
                gap = float(recon_chk.difference)
            except (TypeError, ValueError):
                gap = 0.0
            if abs(gap) > 0.005:
                unallocated.append(Unallocated(
                    carrier = h.carrier,
                    label   = f"Avstämningsdifferens – faktura {h.invoice_number}",
                    amount  = round(abs(gap), 2),
                ))

    # ── Surcharges (carrier-level for Tillägg tab) + per-service for Kostnad tab ──
    surcharges: list[Surcharge] = []
    service_surcharges: list[ServiceSurcharge] = []

    sc_carrier: dict = defaultdict(float)        # (carrier, sc_cat) -> total
    sc_service: dict = defaultdict(float)        # (carrier, svc_cat, sc_cat) -> total

    for (carrier, inv_num), inv_lines in (all_lines_dict or {}).items():
        for ln in inv_lines:
            if getattr(ln, "line_type", "") != "Surcharge":
                continue
            sc_cat  = getattr(ln, "surcharge_category", None) or "Unknown"
            svc_cat = getattr(ln, "service_category",   None) or "Unknown"
            amt     = getattr(ln, "amount", 0.0) or 0.0
            sc_carrier[(carrier, sc_cat)]           += amt
            sc_service[(carrier, svc_cat, sc_cat)]  += amt

    for (carrier, cat), total in sorted(sc_carrier.items(), key=lambda x: -x[1]):
        surcharges.append(Surcharge(
            carrier    = carrier,
            name       = _SC_LABEL.get(cat, cat),
            amount     = round(total, 2),
            is_fuel    = cat == "Fuel",
            is_unknown = cat == "Unknown",
        ))

    for (carrier, svc_cat, sc_cat), total in sorted(sc_service.items(), key=lambda x: -x[1]):
        service_surcharges.append(ServiceSurcharge(
            carrier       = carrier,
            service_name  = _SVC_LABEL.get(svc_cat, svc_cat),
            surcharge_name= _SC_LABEL.get(sc_cat, sc_cat),
            amount        = round(total, 2),
            is_fuel       = sc_cat == "Fuel",
            is_unknown    = sc_cat == "Unknown",
        ))

    # ── Anomalies ─────────────────────────────────────────────────────────────
    summary_anomalies: list[SummaryAnomaly] = []
    for a in (anomalies or []):
        source = "AI" if a.claude_explanation else "Regel"
        summary_anomalies.append(SummaryAnomaly(
            severity         = a.severity,
            carrier          = a.carrier,
            invoice_number   = a.invoice_number,
            anomaly_type     = a.anomaly_type,
            description      = a.description,
            suggested_action = a.suggested_action or "",
            ai_explanation   = a.claude_explanation or "",
            source           = source,
        ))

    # Validation issues as anomalies (warnings/errors not already covered by invoice recon)
    recon_check_names = {"PDFTotalVsExcelSummary", "LineSumVsHeaderTotal"}
    for c in checks:
        if c.severity not in ("Error", "Warning"):
            continue
        if c.check_name in recon_check_names:
            continue  # already shown in Invoice.recon_message
        source = "AI" if c.claude_explanation else "Regel"
        summary_anomalies.append(SummaryAnomaly(
            severity         = c.severity,
            carrier          = c.carrier,
            invoice_number   = c.invoice_number,
            anomaly_type     = c.check_name,
            description      = c.message,
            suggested_action = "",
            ai_explanation   = c.claude_explanation or "",
            source           = source,
        ))

    # Pending as anomalies
    for m in missing_bring:
        summary_anomalies.append(SummaryAnomaly(
            severity         = "Warning",
            carrier          = "Bring",
            invoice_number   = m.get("invoice_number", ""),
            anomaly_type     = "Ofullständig fakturauppsättning",
            description      = m.get("message", "PDF eller Excel saknas"),
            suggested_action = "Hämta saknad fil och lägg i 00_Inbox för ombearbetning.",
            ai_explanation   = "",
            source           = "Regel",
        ))

    return SummaryInput(
        run_id             = run_id,
        generated          = datetime.now().strftime("%Y-%m-%d %H:%M"),
        files_scanned      = payload.get("total_files_scanned", 0),
        invoices           = invoices,
        services           = services,
        surcharges         = surcharges,
        anomalies          = summary_anomalies,
        unallocated        = unallocated,
        service_surcharges = service_surcharges,
    )


# ── HTML archive (minimal, not emailed) ──────────────────────────────────────

def _write_html_archive(path: Path, run_id: str, payload: dict) -> None:
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    c   = payload.get("check_counts", {})
    ok_n, warn_n, err_n = c.get("OK", 0), c.get("Warning", 0), c.get("Error", 0)
    overall = "Error" if err_n else ("Warning" if warn_n else "OK")
    total = sum(v.get("total_ex_vat", 0) for v in payload.get("carrier_totals", {}).values())
    content = (
        f"<!DOCTYPE html><html><head><meta charset='utf-8'>"
        f"<title>Freight Invoice Control {run_id}</title></head><body>"
        f"<h2>Freight Invoice Control</h2>"
        f"<p>Run: {run_id} | Generated: {now} | Status: {overall}</p>"
        f"<p>Invoices: {len(payload.get('invoices', []))} | "
        f"Total: {total:,.2f} SEK | Checks: {ok_n} OK / {warn_n} Warning / {err_n} Error</p>"
        f"</body></html>"
    )
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
    all_lines_dict: dict | None = None,
) -> None:
    """Write the invoice approval Excel to For_Email/ and an HTML archive to Summaries/."""
    config.FOR_EMAIL_DIR.mkdir(parents=True, exist_ok=True)
    xlsx_path = config.FOR_EMAIL_DIR / f"summary_{run_id}.xlsx"
    html_path = config.SUMMARIES_DIR / f"summary_{run_id}.html"

    summary_input = _build_summary_input(
        run_id        = run_id,
        payload       = payload,
        headers       = all_invoice_headers,
        lines         = all_invoice_lines,
        checks        = all_checks,
        anomalies     = anomalies or [],
        missing_bring = missing_bring or [],
        all_lines_dict= all_lines_dict,
    )
    build_summary(summary_input, str(xlsx_path))
    _write_html_archive(html_path, run_id, payload)

    logger.info("RunExporter",
                f"Power Automate file: {xlsx_path.name} | HTML archive: {html_path.name}")


def write_missing_file_alert(
    run_id: str,
    missing: list[dict],
    logger: ProcessingLogger,
) -> None:
    """Write a For_Email XLSX alert when a Bring invoice is missing its PDF or Excel pair."""
    if not missing:
        return
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    config.FOR_EMAIL_DIR.mkdir(parents=True, exist_ok=True)
    path = config.FOR_EMAIL_DIR / f"alert_missing_files_{run_id}.xlsx"
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    n = len(missing)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Alert"
    ws.cell(row=1, column=1, value="ACTION REQUIRED — Missing Invoice File(s)").font = Font(bold=True, size=13)
    for row, (key, val) in enumerate([
        ("Run ID", run_id), ("Generated", now_str), ("Incomplete invoices", n),
        ("Action", "Retrieve missing file, place in 00_Inbox, re-run."),
    ], start=3):
        ws.cell(row=row, column=1, value=key).font = Font(bold=True)
        ws.cell(row=row, column=2, value=val)

    ws2 = wb.create_sheet("Details")
    for col, hdr in enumerate(["Invoice #", "Missing File", "Received File", "Details"], 1):
        ws2.cell(row=1, column=col, value=hdr).font = Font(bold=True)
    for i, m in enumerate(missing, start=2):
        ws2.cell(row=i, column=1, value=m.get("invoice_number", ""))
        ws2.cell(row=i, column=2, value=m.get("missing_file", ""))
        ws2.cell(row=i, column=3, value=m.get("found_file", ""))
        ws2.cell(row=i, column=4, value=m.get("message", ""))

    wb.save(path)
    logger.warning("RunExporter",
                   f"Missing-file alert written: {path.name} ({n} incomplete invoice(s))")
