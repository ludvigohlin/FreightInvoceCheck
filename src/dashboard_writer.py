"""Generate self-contained HTML dashboard from historical output CSV files."""

from __future__ import annotations

import csv
import html as html_lib
import json
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

from src import config
from src.processing_logger import ProcessingLogger


def _read_csv(path: Path) -> list:
    if not path.exists():
        return []
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f, delimiter=";"))


def _safe_float(val: Any) -> float:
    try:
        return float(val) if val else 0.0
    except (TypeError, ValueError):
        return 0.0


def _e(val: Any) -> str:
    return html_lib.escape(str(val) if val is not None else "")


# Service families used to prorate invoice-level fuel surcharges (mirrors run_exporter.py)
_PAKET_SVC_CATS = frozenset({"Parcel", "Service Point", "Parcel Locker"})
_PALL_SVC_CATS  = frozenset({"Pallet"})


def _fuel_svc_hint(surcharge_raw: str) -> str | None:
    n = (surcharge_raw or "").lower()
    if "pall"  in n: return "pall"
    if "paket" in n: return "paket"
    return None


def write_html_dashboard(logger: ProcessingLogger) -> None:
    """Read all output CSVs and write 02_Output/dashboard.html."""
    invoices = _read_csv(config.INVOICE_HEADER_CSV)
    lines = _read_csv(config.INVOICE_LINES_CSV)
    surcharges = _read_csv(config.SURCHARGE_LINES_CSV)
    checks = _read_csv(config.INVOICE_CHECKS_CSV)
    anomalies_raw = _read_csv(config.ANOMALIES_CSV)
    pending_raw = _read_csv(config.PENDING_INVOICES_CSV)

    inv_lookup = {inv["invoice_number"]: inv for inv in invoices}

    # ── Invoice JS data ───────────────────────────────────────────────────────
    inv_js = []
    for inv in sorted(invoices, key=lambda x: x.get("invoice_date", ""), reverse=True):
        date = inv.get("invoice_date", "") or ""
        inv_js.append({
            "date": date,
            "year": date[:4] if len(date) >= 4 else "",
            "month": date[:7] if len(date) >= 7 else "",
            "carrier": inv.get("carrier", ""),
            "invoice_no": inv.get("invoice_number", ""),
            "total": _safe_float(inv.get("total_ex_vat")),
            "currency": inv.get("currency", "SEK"),
            "status": inv.get("reconciliation_status", ""),
            "customer_number": inv.get("customer_number", ""),
            "due_date": inv.get("due_date", ""),
            "vat_amount": _safe_float(inv.get("vat_amount")),
            "total_inc_vat": _safe_float(inv.get("total_inc_vat")),
            "source_file": inv.get("source_file", ""),
        })

    existing_inv_nos = {i["invoice_no"] for i in inv_js}
    for p in pending_raw:
        inv_no = p.get("invoice_number", "")
        if inv_no in existing_inv_nos:
            continue
        inv_js.append({
            "date": "", "year": "", "month": "",
            "carrier": p.get("carrier", "Bring"),
            "invoice_no": inv_no,
            "total": _safe_float(p.get("known_total_ex_vat")),
            "currency": "SEK",
            "status": "Pending",
            "customer_number": "", "due_date": "",
            "vat_amount": 0.0, "total_inc_vat": 0.0,
            "source_file": p.get("source_file", ""),
            "pending_note": p.get("note", ""),
        })

    # ── Service aggregates ────────────────────────────────────────────────────
    svc_agg: dict = defaultdict(lambda: {"count": 0, "total": 0.0, "carrier": "", "date": "", "year": "", "month": ""})
    for line in lines:
        if line.get("line_type") == "BaseFreight":
            inv_no = line.get("invoice_number", "")
            cat = line.get("service_category", "Unknown") or "Unknown"
            key = (inv_no, cat)
            info = inv_lookup.get(inv_no, {})
            date = info.get("invoice_date", "") or ""
            svc_agg[key]["count"] += 1
            svc_agg[key]["total"] += _safe_float(line.get("amount"))
            svc_agg[key]["carrier"] = line.get("carrier", "") or info.get("carrier", "")
            svc_agg[key]["date"] = date
            svc_agg[key]["year"] = date[:4] if len(date) >= 4 else ""
            svc_agg[key]["month"] = date[:7] if len(date) >= 7 else ""

    svc_js = [
        {"invoice_no": k[0], "service_cat": k[1], "carrier": v["carrier"],
         "date": v["date"], "year": v["year"], "month": v["month"],
         "count": v["count"], "total": round(v["total"], 2)}
        for k, v in svc_agg.items()
    ]

    # ── Per-line cost data (shipment-date based) ──────────────────────────────
    # Feeds the Cost timeline chart: each line's cost is bucketed by the day its
    # shipment actually went out (shipment_date), not the invoice's single
    # invoice_date — an invoice can bundle shipments sent on many different
    # days. Falls back to invoice_date for the handful of lines with no
    # shipment identity of their own (PostNord's invoice-level prorated
    # Drivmedelstillägg surcharges).
    line_js = []
    for line in lines:
        inv_no = line.get("invoice_number", "")
        info = inv_lookup.get(inv_no, {})
        date = line.get("shipment_date", "") or info.get("invoice_date", "") or ""
        if not date:
            continue
        line_js.append({
            "invoice_no": inv_no,
            "carrier": line.get("carrier", "") or info.get("carrier", ""),
            "status": info.get("reconciliation_status", ""),
            "date": date,
            "year": date[:4] if len(date) >= 4 else "",
            "month": date[:7] if len(date) >= 7 else "",
            "amount": _safe_float(line.get("amount")),
        })

    # ── Cross-check: header total vs. sum of that invoice's own lines ──────────
    # invoice_header.csv, invoice_lines.csv, and the Excel/summary reports are three
    # independently-read/aggregated views of the same underlying data (see project
    # audit notes on triplicated aggregation). This doesn't unify them, but it does
    # give an early warning in the log if the dashboard's own two CSV reads (header
    # vs. line-level) ever silently drift apart for the same invoice — which would
    # otherwise only surface as a controller noticing a wrong number by eye.
    line_sum_by_inv: dict = defaultdict(float)
    for line in lines:
        line_sum_by_inv[line.get("invoice_number", "")] += _safe_float(line.get("amount"))
    for inv in invoices:
        inv_no = inv.get("invoice_number", "")
        header_total = _safe_float(inv.get("total_ex_vat"))
        line_total = line_sum_by_inv.get(inv_no, 0.0)
        if abs(header_total - line_total) > 1.0:
            logger.warning(
                "DashboardWriter",
                f"Invoice {inv_no} ({inv.get('carrier','')}): header total_ex_vat "
                f"({header_total:.2f}) does not match sum of its own invoice_lines.csv "
                f"rows ({line_total:.2f}) — dashboard/report totals for this invoice may disagree.",
            )

    # ── Surcharge aggregates ──────────────────────────────────────────────────
    sc_agg: dict = defaultdict(lambda: {"total": 0.0, "carrier": "", "year": "", "month": ""})
    for line in surcharges:
        inv_no = line.get("invoice_number", "")
        cat = line.get("surcharge_category", "Unknown") or "Unknown"
        key = (inv_no, cat)
        info = inv_lookup.get(inv_no, {})
        date = info.get("invoice_date", "") or ""
        sc_agg[key]["total"] += _safe_float(line.get("amount"))
        sc_agg[key]["carrier"] = line.get("carrier", "") or info.get("carrier", "")
        sc_agg[key]["year"] = date[:4] if len(date) >= 4 else ""
        sc_agg[key]["month"] = date[:7] if len(date) >= 7 else ""

    sc_js = [
        {"invoice_no": k[0], "surcharge_cat": k[1], "carrier": v["carrier"],
         "year": v["year"], "month": v["month"], "total": round(v["total"], 2)}
        for k, v in sc_agg.items()
    ]

    # ── Service cost breakdown ────────────────────────────────────────────────
    base_agg: dict = defaultdict(lambda: {"count": 0, "total": 0.0, "carrier": "", "year": "", "month": ""})
    for line in lines:
        if line.get("line_type") == "BaseFreight":
            inv_no = line.get("invoice_number", "")
            cat = line.get("service_category", "Unknown") or "Unknown"
            key = (inv_no, cat)
            info = inv_lookup.get(inv_no, {})
            date = info.get("invoice_date", "") or ""
            base_agg[key]["count"] += 1
            base_agg[key]["total"] += _safe_float(line.get("amount"))
            base_agg[key]["carrier"] = line.get("carrier", "") or info.get("carrier", "")
            base_agg[key]["year"] = date[:4] if len(date) >= 4 else ""
            base_agg[key]["month"] = date[:7] if len(date) >= 7 else ""

    sc_svc_agg: dict = defaultdict(float)
    for sc in surcharges:
        inv_no  = sc.get("invoice_number", "")
        svc_cat = sc.get("related_service_category", "") or "Unknown"
        sc_cat  = sc.get("surcharge_category", "Unknown") or "Unknown"
        amt     = _safe_float(sc.get("amount"))

        if svc_cat != "Unknown":
            sc_svc_agg[(inv_no, svc_cat, sc_cat)] += amt
        elif sc_cat == "Fuel":
            # Prorate invoice-level fuel to services by line count.
            # PostNord separates Drivmedelstillägg Paket (parcel family) and Pall.
            hint = _fuel_svc_hint(sc.get("surcharge_raw", "") or "")
            if hint == "pall":
                candidates = _PALL_SVC_CATS
            elif hint == "paket":
                candidates = _PAKET_SVC_CATS
            else:
                candidates = {cat for (i, cat) in base_agg if i == inv_no}
            counts = {cat: base_agg[(inv_no, cat)]["count"]
                      for cat in candidates if (inv_no, cat) in base_agg}
            total_cnt = sum(counts.values())
            if total_cnt > 0:
                for cat, cnt in counts.items():
                    sc_svc_agg[(inv_no, cat, sc_cat)] += amt * cnt / total_cnt
            else:
                sc_svc_agg[(inv_no, "Unknown", sc_cat)] += amt
        else:
            sc_svc_agg[(inv_no, svc_cat, sc_cat)] += amt

    all_sc_cats = sorted({k[2] for k in sc_svc_agg})

    svc_cost_js = []
    for (inv_no, svc_cat), v in base_agg.items():
        sc_by_cat = {
            sc_cat: round(sc_svc_agg.get((inv_no, svc_cat, sc_cat), 0.0), 2)
            for sc_cat in all_sc_cats
        }
        sc_grand = sum(sc_by_cat.values())
        svc_cost_js.append({
            "invoice_no": inv_no, "service_cat": svc_cat, "carrier": v["carrier"],
            "year": v["year"], "month": v["month"], "count": v["count"],
            "base_total": round(v["total"], 2),
            "sc_by_cat": {k: val for k, val in sc_by_cat.items() if val > 0},
            "sc_grand": round(sc_grand, 2),
            "grand_total": round(v["total"] + sc_grand, 2),
        })

    # ── Chargeable weight per Parcel-family service ───────────────────────────
    # "Fraktdragande vikt" — the weight actually used for billing (Bring bills
    # Parcel by physical weight; PostNord separately tracks a chargeable figure,
    # fraktdr_vikt, that can differ from the physical weight_kg). Pallet is
    # excluded — it's billed by pallet count, weight isn't the pricing driver.
    weight_agg: dict = defaultdict(lambda: {"weight_sum": 0.0, "count": 0, "carrier": "", "year": "", "month": ""})
    for line in lines:
        if line.get("line_type") != "BaseFreight":
            continue
        cat = line.get("service_category", "") or ""
        if cat not in _PAKET_SVC_CATS:
            continue
        w = _safe_float(line.get("chargeable_weight_kg"))
        if not w:
            continue
        inv_no = line.get("invoice_number", "")
        info = inv_lookup.get(inv_no, {})
        date = line.get("shipment_date", "") or info.get("invoice_date", "") or ""
        if not date:
            continue
        key = (inv_no, cat)
        weight_agg[key]["weight_sum"] += w
        weight_agg[key]["count"] += 1
        weight_agg[key]["carrier"] = line.get("carrier", "") or info.get("carrier", "")
        weight_agg[key]["year"] = date[:4] if len(date) >= 4 else ""
        weight_agg[key]["month"] = date[:7] if len(date) >= 7 else ""

    weight_js = [
        {"invoice_no": k[0], "service_cat": k[1], "carrier": v["carrier"],
         "year": v["year"], "month": v["month"],
         "weight_sum": round(v["weight_sum"], 3), "count": v["count"]}
        for k, v in weight_agg.items()
    ]

    # ── Returns ────────────────────────────────────────────────────────────────
    # Detected differently per carrier: PostNord bills a return leg as an
    # ordinary second BaseFreight charge addressed back to our own warehouse
    # (is_return flag, set in postnord_parser.py); Bring flags it as an
    # "Attempted Delivery Return" surcharge line (surcharge_category=="Return")
    # instead, since it doesn't bill a separate return freight charge.
    #
    # For PostNord we can also say more: `service_code` carries the PDF's own
    # kolli_id (a real, stable shipment identifier), so grouping all of a
    # kolli's BaseFreight lines — across every invoice it ever appears on —
    # lets us show where a returned parcel was originally headed, and detect
    # the rare case where the same kolli has bounced back more than once.
    kolli_history: dict = defaultdict(list)
    for line in lines:
        if line.get("carrier") != "PostNord" or line.get("line_type") != "BaseFreight":
            continue
        kolli_id = line.get("service_code", "")
        if not kolli_id:
            continue
        inv_no = line.get("invoice_number", "")
        info = inv_lookup.get(inv_no, {})
        date = line.get("shipment_date", "") or info.get("invoice_date", "") or ""
        kolli_history[kolli_id].append({
            "date": date,
            "postal": line.get("to_postal", ""),
            "city": line.get("to_city", ""),
            "is_return": line.get("is_return") == "True",
            "invoice_no": inv_no,
            "amount": _safe_float(line.get("amount")),
        })
    for hist in kolli_history.values():
        hist.sort(key=lambda h: h["date"])

    return_js = []
    for line in lines:
        if line.get("line_type") != "BaseFreight" or line.get("is_return") != "True":
            continue
        inv_no = line.get("invoice_number", "")
        info = inv_lookup.get(inv_no, {})
        date = line.get("shipment_date", "") or info.get("invoice_date", "") or ""
        if not date:
            continue
        kolli_id = line.get("service_code", "")
        hist = kolli_history.get(kolli_id, [])
        repeat_count = sum(1 for h in hist if h["is_return"]) or 1
        # Most recent non-return (outbound) destination strictly before this
        # return's own date — i.e. where it was actually being delivered.
        origin = None
        for h in hist:
            if not h["is_return"] and h["date"] and h["date"] < date:
                origin = h
        origin_label = f"{origin['postal']} {origin['city']}".strip() if origin else ""
        return_js.append({
            "invoice_no": inv_no,
            "carrier": line.get("carrier", "") or info.get("carrier", ""),
            "date": date, "year": date[:4] if len(date) >= 4 else "",
            "month": date[:7] if len(date) >= 7 else "",
            "amount": _safe_float(line.get("amount")),
            "kolli_id": kolli_id,
            "origin": origin_label,
            "repeat_count": repeat_count,
        })
    for sc in surcharges:
        if (sc.get("surcharge_category") or "") != "Return":
            continue
        inv_no = sc.get("invoice_number", "")
        info = inv_lookup.get(inv_no, {})
        date = info.get("invoice_date", "") or ""
        if not date:
            continue
        return_js.append({
            "invoice_no": inv_no,
            "carrier": sc.get("carrier", "") or info.get("carrier", ""),
            "date": date, "year": date[:4] if len(date) >= 4 else "",
            "month": date[:7] if len(date) >= 7 else "",
            "amount": _safe_float(sc.get("amount")),
            "kolli_id": "", "origin": "", "repeat_count": 1,
        })

    # Only export history for kolli_ids that actually show up in a return —
    # exporting all ~14k PostNord kolli would bloat the dashboard for no
    # benefit, since this is only ever looked up from a return-detail click.
    returned_kolli_ids = {r["kolli_id"] for r in return_js if r.get("kolli_id")}
    kolli_history_js = {k: v for k, v in kolli_history.items() if k in returned_kolli_ids}

    # ── Anomaly JS data ───────────────────────────────────────────────────────
    anomaly_js = [
        {"carrier": a.get("carrier", ""), "invoice_no": a.get("invoice_number", ""),
         "type": a.get("anomaly_type", ""), "severity": a.get("severity", ""),
         "description": a.get("description", ""), "detail": a.get("detail", ""),
         "value": _safe_float(a.get("value")), "threshold": _safe_float(a.get("threshold")),
         "suggested_action": a.get("suggested_action", ""),
         "explanation": a.get("claude_explanation", "")}
        for a in reversed(anomalies_raw[-200:] if len(anomalies_raw) > 200 else anomalies_raw)
    ]

    # ── Recent checks ─────────────────────────────────────────────────────────
    check_js = [
        {"carrier": c.get("carrier", ""), "invoice_no": c.get("invoice_number", ""),
         "check": c.get("check_name", ""), "status": c.get("status", ""),
         "message": c.get("message", "")}
        for c in reversed(checks[-60:] if len(checks) > 60 else checks)
    ]

    # ── Filter options ────────────────────────────────────────────────────────
    years = sorted({i["year"] for i in inv_js if i["year"]}, reverse=True)
    carriers = sorted({i["carrier"] for i in inv_js if i["carrier"]})
    service_cats = sorted({s["service_cat"] for s in svc_js if s["service_cat"]})
    months_all = sorted({i["month"] for i in inv_js if i["month"]}, reverse=True)

    year_opts = "".join(f'<option value="{_e(y)}">{_e(y)}</option>' for y in years)
    month_opts = "".join(f'<option value="{_e(m)}">{_e(m)}</option>' for m in months_all)
    carrier_opts = "".join(f'<option value="{_e(c)}">{_e(c)}</option>' for c in carriers)
    svc_opts = "".join(f'<option value="{_e(s)}">{_e(s)}</option>' for s in service_cats)

    last_updated = datetime.now().strftime("%Y-%m-%d %H:%M")

    inv_json       = json.dumps(inv_js)
    svc_json       = json.dumps(svc_js)
    line_json      = json.dumps(line_js)
    sc_json        = json.dumps(sc_js)
    check_json     = json.dumps(check_js)
    svc_cost_json  = json.dumps(svc_cost_js)
    anomaly_json   = json.dumps(anomaly_js)
    sc_cats_json   = json.dumps(all_sc_cats)
    weight_json    = json.dumps(weight_js)
    return_json    = json.dumps(return_js)
    kolli_history_json = json.dumps(kolli_history_js)

    html_content = f"""<!DOCTYPE html>
<html lang="sv">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Freight Invoice Control — Dashboard</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4/dist/chart.umd.min.js"></script>
  <style>
    *,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
    body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;
          font-size:14px;background:#f0f2f5;color:#333}}
    header{{background:#1565c0;color:#fff;padding:14px 24px;
            display:flex;align-items:center;justify-content:space-between;
            position:sticky;top:0;z-index:100;box-shadow:0 2px 6px rgba(0,0,0,.2)}}
    header h1{{font-size:18px;font-weight:600;letter-spacing:-.2px}}
    header .meta{{font-size:12px;opacity:.75}}
    .filter-bar{{background:#fff;border-bottom:1px solid #ddd;padding:9px 24px;
                 display:flex;flex-wrap:wrap;gap:8px;align-items:center;
                 position:sticky;top:49px;z-index:99;box-shadow:0 1px 4px rgba(0,0,0,.06)}}
    .filter-bar label{{font-size:11px;color:#666;font-weight:700;text-transform:uppercase;letter-spacing:.4px}}
    .filter-bar select,.filter-bar input{{
      padding:5px 8px;border:1px solid #ccc;border-radius:4px;
      font-size:13px;background:#fff;min-width:110px}}
    .filter-bar input{{min-width:150px}}
    .filter-bar button{{padding:5px 13px;background:#1565c0;color:#fff;
                        border:none;border-radius:4px;cursor:pointer;font-size:13px}}
    .filter-bar button:hover{{background:#0d47a1}}
    .filter-bar .reset{{background:#9e9e9e}}
    .filter-bar .reset:hover{{background:#757575}}
    .filter-sep{{width:1px;height:22px;background:#e0e0e0;margin:0 2px}}
    main{{max-width:1340px;margin:0 auto;padding:18px 16px}}

    /* KPI row */
    .kpi-row{{display:flex;gap:12px;margin-bottom:18px;flex-wrap:wrap}}
    .kpi{{background:#fff;border-radius:8px;padding:14px 18px;
           box-shadow:0 1px 3px rgba(0,0,0,.09);flex:1;min-width:150px;
           border-top:3px solid transparent}}
    .kpi.blue{{border-top-color:#1565c0}}
    .kpi.green{{border-top-color:#2e7d32}}
    .kpi.orange{{border-top-color:#e65100}}
    .kpi.red{{border-top-color:#c62828}}
    .kpi.yellow{{border-top-color:#f57f17}}
    .kpi.grey{{border-top-color:#9e9e9e}}
    .kpi .value{{font-size:24px;font-weight:700;color:#1a1a1a;line-height:1.1}}
    .kpi .label{{font-size:11px;color:#777;margin-top:4px;text-transform:uppercase;letter-spacing:.5px;font-weight:600}}
    .kpi .sub{{font-size:12px;color:#aaa;margin-top:3px}}

    /* Card */
    .card{{background:#fff;border-radius:8px;padding:16px 18px;
           box-shadow:0 1px 3px rgba(0,0,0,.09);margin-bottom:16px}}
    .card h2{{font-size:13px;font-weight:700;color:#1565c0;margin-bottom:12px;
              border-bottom:1px solid #eee;padding-bottom:8px;
              display:flex;align-items:center;justify-content:space-between;
              text-transform:uppercase;letter-spacing:.4px}}
    .card h2 .count{{font-size:12px;color:#bbb;font-weight:400;text-transform:none;letter-spacing:0}}

    /* Grid layouts */
    .grid-2-1{{display:grid;grid-template-columns:2fr 1fr;gap:16px;margin-bottom:16px}}
    .grid-3-2{{display:grid;grid-template-columns:3fr 2fr;gap:16px;margin-bottom:16px}}

    /* Tables */
    table{{width:100%;border-collapse:collapse;font-size:13px}}
    th{{background:#1565c0;color:#fff;text-align:left;padding:7px 10px;
        font-weight:600;white-space:nowrap}}
    td{{padding:6px 10px;border-bottom:1px solid #f0f0f0}}
    tr:last-child td{{border-bottom:none}}
    tr:hover td{{background:#f5f9ff}}
    td.num{{text-align:right;font-variant-numeric:tabular-nums}}
    .table-wrap{{overflow-x:auto;overflow-y:auto;border-radius:6px;border:1px solid #e8e8e8}}
    .table-wrap.fixed-h{{max-height:380px}}

    /* Badges */
    .badge{{padding:2px 8px;border-radius:4px;font-weight:700;font-size:11px;
            display:inline-block;letter-spacing:.2px}}
    .badge-ok{{background:#e8f5e9;color:#2e7d32}}
    .badge-warn{{background:#fff8e1;color:#f57f17}}
    .badge-err{{background:#ffebee;color:#c62828}}
    .badge-pending{{background:#fff3e0;color:#e65100}}
    .badge-speconly{{background:#ede7f6;color:#5e35b1}}
    .badge-nc{{background:#f5f5f5;color:#9e9e9e}}

    .no-data{{color:#bbb;font-style:italic;padding:18px;text-align:center;font-size:13px}}

    /* Toggle buttons */
    .toggle-grp{{display:flex;gap:3px}}
    .toggle-btn{{padding:3px 10px;border:1px solid #1565c0;border-radius:4px;
                 cursor:pointer;font-size:12px;background:#fff;color:#1565c0;font-weight:600}}
    .toggle-btn.active{{background:#1565c0;color:#fff}}
    .toggle-btn:hover:not(.active){{background:#e3f2fd}}

    /* Invoice row clickable */
    .inv-row{{cursor:pointer;transition:background .12s}}
    .inv-row:hover td{{background:#e8f4fd !important}}

    /* Recent invoices card */
    #recentInvBody tr.pending-row td{{opacity:.75}}

    /* Modal */
    .modal-overlay{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);
                    z-index:1000;overflow-y:auto;padding:32px 16px}}
    .modal-overlay.open{{display:flex;align-items:flex-start;justify-content:center}}
    .modal{{background:#fff;border-radius:10px;width:100%;max-width:920px;
            box-shadow:0 8px 32px rgba(0,0,0,.25);overflow:hidden}}
    .modal-header{{background:#1565c0;color:#fff;padding:14px 20px;
                   display:flex;align-items:center;justify-content:space-between}}
    .modal-header h3{{font-size:15px;font-weight:700}}
    .modal-close{{background:none;border:none;color:#fff;font-size:22px;
                  cursor:pointer;line-height:1;padding:0 4px}}
    .modal-close:hover{{opacity:.7}}
    .modal-back{{background:none;border:1px solid rgba(255,255,255,.5);color:#fff;
                 font-size:12px;font-weight:600;cursor:pointer;padding:4px 10px;
                 border-radius:5px;margin-right:10px}}
    .modal-back:hover{{background:rgba(255,255,255,.15)}}
    .copy-btn{{background:#eef4fb;border:1px solid #d0dced;color:#1565c0;font-size:10px;
               font-weight:700;cursor:pointer;padding:1px 7px;border-radius:4px;margin-left:6px}}
    .copy-btn:hover{{background:#dceafd}}
    .modal-body{{padding:20px;overflow-y:auto;max-height:76vh}}
    .modal-section{{margin-bottom:18px}}
    .modal-section h4{{font-size:11px;font-weight:700;color:#1565c0;
                       border-bottom:1px solid #eee;padding-bottom:5px;margin-bottom:10px;
                       text-transform:uppercase;letter-spacing:.5px}}
    .info-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:8px 16px}}
    .info-item .ilabel{{font-size:10px;color:#aaa;text-transform:uppercase;letter-spacing:.4px}}
    .info-item .ivalue{{font-size:13px;font-weight:600;color:#222;margin-top:1px}}

    @media(max-width:900px){{
      .grid-2-1,.grid-3-2{{grid-template-columns:1fr}}
      .kpi{{min-width:130px}}
    }}
  </style>
</head>
<body>
<header>
  <h1>Freight Invoice Control</h1>
  <div class="meta">Isicom AB &nbsp;·&nbsp; Uppdaterad {_e(last_updated)}</div>
</header>

<div class="filter-bar">
  <label>År</label>
  <select id="fYear" onchange="applyFilters()">
    <option value="">Alla</option>{year_opts}
  </select>
  <label>Månad</label>
  <select id="fMonth" onchange="applyFilters()">
    <option value="">Alla</option>{month_opts}
  </select>
  <label>Transportör</label>
  <select id="fCarrier" onchange="applyFilters()">
    <option value="">Alla</option>{carrier_opts}
  </select>
  <label>Tjänst</label>
  <select id="fService" onchange="applyFilters()">
    <option value="">Alla</option>{svc_opts}
  </select>
  <input id="fInvoice" type="text" placeholder="Sök fakturanr..." oninput="applyFilters()">
  <div class="filter-sep"></div>
  <div class="toggle-grp">
    <button class="toggle-btn active" id="btnAll" onclick="setRecFilter('all')">Alla fakturor</button>
    <button class="toggle-btn" id="btnRecOnly" onclick="setRecFilter('reconciled')">Fullt avstämda</button>
  </div>
  <button class="reset" onclick="resetFilters()">Återställ</button>
</div>

<main>

  <!-- KPI row -->
  <div class="kpi-row" id="kpiRow"></div>

  <!-- Charts + Recent invoices -->
  <div class="grid-2-1">
    <div class="card">
      <h2>Antal kolli per månad
        <span class="count" id="volumeNote"></span>
      </h2>
      <div style="height:320px"><canvas id="volumeChart"></canvas></div>
    </div>
    <div class="card">
      <h2>Senaste fakturor <span class="count" id="recentInvCount"></span></h2>
      <div style="height:320px;overflow-y:auto;border-radius:6px;border:1px solid #e8e8e8">
        <table id="recentInvTable" style="font-size:12px">
          <thead><tr>
            <th>Datum</th><th>Transp.</th><th>Faktura #</th>
            <th class="num">Belopp</th><th>Status</th>
          </tr></thead>
          <tbody id="recentInvBody"></tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- Cost timeline -->
  <div class="card">
    <h2>Kostnadsutveckling ex-moms (SEK)
      <span style="display:flex;align-items:center;gap:10px">
        <span class="count" id="timelineNote"></span>
        <span class="toggle-grp">
          <button class="toggle-btn" id="btnDaily" onclick="setGranularity('daily')">Dagsvis</button>
          <button class="toggle-btn active" id="btnMonthly" onclick="setGranularity('monthly')">Månadsvis</button>
          <button class="toggle-btn" id="btnYearly" onclick="setGranularity('yearly')">Årsvis</button>
        </span>
      </span>
    </h2>
    <canvas id="timelineChart" height="90"></canvas>
  </div>

  <!-- Avg cost trend + Surcharges side by side -->
  <div class="grid-3-2">
    <div class="card">
      <h2>Snittkostnad per tjänstetyp och månad (inkl. tillägg)
        <span class="count" style="font-size:11px;color:#bbb;font-weight:400">
          &nbsp;Bring &amp; PostNord = SEK
        </span>
      </h2>
      <canvas id="avgTrendChart" height="140"></canvas>
    </div>
    <div class="card">
      <h2>Tilläggsavgifter <span class="count" id="scCount"></span></h2>
      <canvas id="surchargeChart" height="220"></canvas>
    </div>
  </div>

  <!-- Chargeable weight per Parcel-family service -->
  <div class="card">
    <h2>Fraktdragande snittvikt (Parcel-tjänster)
      <span class="count" id="weightTrendNote"></span>
    </h2>
    <div class="grid-3-2" style="margin-bottom:0">
      <canvas id="weightTrendChart" height="180"></canvas>
      <div class="table-wrap">
        <table id="weightTrendTable">
          <thead><tr>
            <th>Tjänst</th><th class="num">Antal</th><th class="num">Snitt (totalt)</th>
          </tr></thead>
          <tbody id="weightTrendBody"></tbody>
        </table>
      </div>
    </div>
    <div style="font-size:11px;color:#999;margin-top:8px">Linjer visar 3-månaders rullande snitt. Pall exkluderad (vikt styr inte prissättningen där).</div>
  </div>

  <!-- Cost per service table -->
  <div class="card">
    <h2>Kostnad per tjänstetyp (inkl. tillägg) <span class="count" id="svcCostCount"></span></h2>
    <div class="table-wrap">
      <table id="svcCostTable">
        <thead id="svcCostHead"></thead>
        <tbody id="svcCostBody"></tbody>
      </table>
    </div>
  </div>

  <!-- Returns -->
  <div class="card">
    <h2>Returer <span class="count" id="returnsCount"></span></h2>
    <div class="table-wrap">
      <table id="returnsTable">
        <thead><tr>
          <th>Transportör</th><th>Månad</th><th class="num">Antal</th><th class="num">Kostnad</th>
        </tr></thead>
        <tbody id="returnsBody"></tbody>
      </table>
    </div>
    <div style="font-size:11px;color:#999;margin-top:8px">
      PostNord: sändning till egen lageradress (retur). Bring: avgift för misslyckat leveransförsök/retur.
    </div>
  </div>

  <!-- Anomalies -->
  <div class="card">
    <h2>Avvikelser <span class="count" id="anomalyCount"></span></h2>
    <div class="table-wrap fixed-h">
      <table id="anomalyTable">
        <thead><tr>
          <th>Transportör</th><th>Faktura#</th><th>Typ</th><th>Allvarlighet</th>
          <th>Beskrivning</th><th>AI-förklaring</th>
        </tr></thead>
        <tbody id="anomalyBody"></tbody>
      </table>
    </div>
  </div>

  <!-- Validation checks -->
  <div class="card">
    <h2>Kontroller</h2>
    <div class="table-wrap fixed-h">
      <table id="checkTable">
        <thead><tr>
          <th>Transportör</th><th>Faktura#</th><th>Kontroll</th>
          <th>Status</th><th>Meddelande</th>
        </tr></thead>
        <tbody id="checkBody"></tbody>
      </table>
    </div>
  </div>

</main>

<!-- Invoice detail modal -->
<div class="modal-overlay" id="modalOverlay" onclick="closeModalOnBg(event)">
  <div class="modal" id="modalBox">
    <div class="modal-header">
      <div style="display:flex;align-items:center">
        <button class="modal-back" id="modalBackBtn" onclick="modalGoBack()" style="display:none">&#x2190; Tillbaka</button>
        <h3 id="modalTitle">Fakturadetaljer</h3>
      </div>
      <button class="modal-close" onclick="closeModal()">&#x2715;</button>
    </div>
    <div class="modal-body" id="modalBody"></div>
  </div>
</div>

<script>
const INV_DATA      = {inv_json};
const SVC_DATA      = {svc_json};
const LINE_DATA     = {line_json};
const SC_DATA       = {sc_json};
const CHECK_DATA    = {check_json};
const SVC_COST_DATA = {svc_cost_json};
const ANOMALY_DATA  = {anomaly_json};
const ALL_SC_CATS   = {sc_cats_json};
const WEIGHT_DATA   = {weight_json};
const RETURN_DATA   = {return_json};
const KOLLI_HISTORY = {kolli_history_json};

const CC  = {{Bring:'#1565c0', PostNord:'#e65100'}};
const CCA = {{Bring:'rgba(21,101,192,.72)', PostNord:'rgba(230,81,0,.72)'}};
function cCol(c)  {{ return CC[c]  || '#546e7a'; }}
function cColA(c) {{ return CCA[c] || 'rgba(84,110,122,.72)'; }}

function badge(status) {{
  const map = {{OK:'badge-ok', Warning:'badge-warn', Error:'badge-err',
                Pending:'badge-pending', SpecOnly:'badge-speconly'}};
  const cls = map[status] || 'badge-nc';
  return `<span class="badge ${{cls}}">${{esc(status)}}</span>`;
}}

// Only "Pending" (no spec or PDF received at all) is treated as missing data.
// "SpecOnly" (Bring spec received, PDF invoice never arrived) is treated as a
// fully usable cost source — we don't wait for the PDF to trust the spec.
const UNCONFIRMED = new Set(['Pending']);
function esc(s) {{
  return String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}}
function fmt(n)    {{ return (+n).toLocaleString('sv-SE',{{minimumFractionDigits:2,maximumFractionDigits:2}}); }}
function fmtInt(n) {{ return Math.round(n).toLocaleString('sv-SE'); }}

let volumeChart, surchargeChart, timelineChart, avgTrendChart, weightTrendChart;
let _tlGran   = 'monthly';
let _recFilter = 'all';

// ── Reconciliation filter ─────────────────────────────────────────────────────
function setRecFilter(f) {{
  _recFilter = f;
  document.getElementById('btnAll').classList.toggle('active', f==='all');
  document.getElementById('btnRecOnly').classList.toggle('active', f==='reconciled');
  applyFilters();
}}

// ── Main filter ───────────────────────────────────────────────────────────────
function applyFilters() {{
  const year    = document.getElementById('fYear').value;
  const month   = document.getElementById('fMonth').value;
  const carrier = document.getElementById('fCarrier').value;
  const svc     = document.getElementById('fService').value;
  const invQ    = document.getElementById('fInvoice').value.toLowerCase().trim();

  const filtInv = INV_DATA.filter(i =>
    (!year    || i.year    === year)   &&
    (!month   || i.month   === month)  &&
    (!carrier || i.carrier === carrier)&&
    (!invQ    || i.invoice_no.toLowerCase().includes(invQ)) &&
    (_recFilter !== 'reconciled' || !UNCONFIRMED.has(i.status))
  );
  const filtNos = new Set(filtInv.map(i => i.invoice_no));

  const filtLine    = LINE_DATA.filter(l => filtNos.has(l.invoice_no));
  const filtSvc     = SVC_DATA.filter(s => filtNos.has(s.invoice_no) && (!svc||s.service_cat===svc));
  const filtSc      = SC_DATA.filter(s => filtNos.has(s.invoice_no));
  const filtSvcCost = SVC_COST_DATA.filter(s => filtNos.has(s.invoice_no) && (!svc||s.service_cat===svc));
  const filtWeight  = WEIGHT_DATA.filter(w => filtNos.has(w.invoice_no) && (!svc||w.service_cat===svc));
  const filtAnom    = ANOMALY_DATA.filter(a => filtNos.has(a.invoice_no));
  const filtReturn  = RETURN_DATA.filter(r => filtNos.has(r.invoice_no));

  renderKPIs(filtInv, filtSvc, filtAnom, filtReturn);
  renderRecentInvoices(filtInv);
  renderVolumeChart(filtSvc);
  renderTimelineChart(filtLine);
  renderAvgCostTrendChart(filtSvcCost);
  renderServiceCostTable(filtSvcCost);
  renderWeightTrendChart(filtWeight);
  renderSurchargeChart(filtSc);
  renderAnomalyTable(filtAnom);
  renderReturnsTable(filtReturn);
}}

function resetFilters() {{
  ['fYear','fMonth','fCarrier','fService'].forEach(id => document.getElementById(id).value='');
  document.getElementById('fInvoice').value='';
  _recFilter='all';
  document.getElementById('btnAll').classList.add('active');
  document.getElementById('btnRecOnly').classList.remove('active');
  applyFilters();
}}

// ── KPIs ──────────────────────────────────────────────────────────────────────
function renderKPIs(invData, svcData, anomData, returnData) {{
  const reconInv    = invData.filter(i => !UNCONFIRMED.has(i.status));
  const pendInv     = invData.filter(i => i.status === 'Pending');
  const totalShip = svcData.reduce((s,v) => s+v.count, 0);

  // Group totals by currency (both carriers currently invoice in SEK, but this
  // stays currency-grouped rather than a flat sum in case that ever changes)
  const byCurrency = {{}};
  reconInv.forEach(i => {{
    const ccy = i.currency || 'SEK';
    byCurrency[ccy] = (byCurrency[ccy] || 0) + i.total;
  }});
  const costLines = Object.entries(byCurrency)
    .sort((a,b) => b[1]-a[1])
    .map(([ccy,amt]) => `${{fmtInt(amt)}} ${{ccy}}`)
    .join(' &nbsp;·&nbsp; ') || '–';

  // Surcharge % per currency group
  const scNos = new Set(reconInv.map(i => i.invoice_no));
  const scByCurrency = {{}};
  SC_DATA.filter(s => scNos.has(s.invoice_no)).forEach(s => {{
    const inv = reconInv.find(i => i.invoice_no === s.invoice_no);
    const ccy = inv ? (inv.currency || 'SEK') : 'SEK';
    scByCurrency[ccy] = (scByCurrency[ccy] || 0) + s.total;
  }});
  const scPctLines = Object.entries(byCurrency)
    .map(([ccy,tot]) => {{
      const sc = scByCurrency[ccy] || 0;
      return tot > 0 ? `${{(sc/tot*100).toFixed(1)}}% (${{ccy}})` : null;
    }})
    .filter(Boolean).join(' · ') || '–';

  const errCnt  = anomData.filter(a => a.severity==='Error').length;
  const warnCnt = anomData.filter(a => a.severity==='Warning').length;
  const anomCls = errCnt>0 ? 'red' : warnCnt>0 ? 'yellow' : 'green';
  const anomSub = errCnt>0
    ? `${{errCnt}} fel, ${{warnCnt}} varningar`
    : warnCnt>0 ? `${{warnCnt}} varning${{warnCnt!==1?'ar':''}}`
    : 'Inga avvikelser';

  const pendCard = pendInv.length > 0 ? `
    <div class="kpi orange">
      <div class="value">${{pendInv.length}}</div>
      <div class="label">Oavstämda</div>
      <div class="sub">Inväntar dokument</div>
    </div>` : '';

  const returnCost = returnData.reduce((s,r) => s+r.amount, 0);

  document.getElementById('kpiRow').innerHTML = `
    <div class="kpi blue">
      <div class="value">${{reconInv.length}}</div>
      <div class="label">Fakturor</div>
      <div class="sub">Med kostnadsdata</div>
    </div>
    ${{pendCard}}
    <div class="kpi blue">
      <div class="value" style="font-size:18px">${{costLines}}</div>
      <div class="label">Total kostnad ex-moms</div>
      <div class="sub">spec eller faktura</div>
    </div>
    <div class="kpi grey">
      <div class="value">${{totalShip.toLocaleString('sv-SE')}}</div>
      <div class="label">Kolli</div>
      <div class="sub">alla tjänstetyper</div>
    </div>
    <div class="kpi ${{anomCls}}">
      <div class="value">${{anomData.length}}</div>
      <div class="label">Avvikelser</div>
      <div class="sub">${{anomSub}}</div>
    </div>
    <div class="kpi grey">
      <div class="value" style="font-size:16px">${{scPctLines}}</div>
      <div class="label">Tillägg av total</div>
      <div class="sub">bränsle + övriga</div>
    </div>
    <div class="kpi orange">
      <div class="value">${{returnData.length.toLocaleString('sv-SE')}}</div>
      <div class="label">Returer</div>
      <div class="sub">${{fmtInt(returnCost)}} SEK</div>
    </div>
  `;
}}

// ── Recent invoices (scrollable — shows all filtered invoices, newest first,
//    so the card fills its grid-row height instead of leaving dead space
//    below a handful of rows) ────────────────────────────────────────────────
function renderRecentInvoices(invData) {{
  const total = invData.length;

  document.getElementById('recentInvCount').textContent =
    `(${{total}} faktura${{total!==1?'r':''}})`;

  const body = document.getElementById('recentInvBody');
  if (!total) {{
    body.innerHTML = '<tr><td colspan="5" class="no-data">Inga fakturor.</td></tr>';
    return;
  }}

  body.innerHTML = invData.map(i => {{
    const pending = i.status === 'Pending';
    const unconfirmed = UNCONFIRMED.has(i.status);
    return `<tr class="inv-row${{unconfirmed?' pending-row':''}}" onclick="openModalView(showInvoiceDetail, ['${{esc(i.invoice_no)}}'])">
      <td style="white-space:nowrap">${{pending?'<em style="color:#ccc">—</em>':esc(i.date)}}</td>
      <td style="white-space:nowrap">${{esc(i.carrier)}}</td>
      <td><strong style="font-size:11px">${{esc(i.invoice_no)}}</strong></td>
      <td class="num" style="white-space:nowrap">${{fmtInt(i.total)}}${{unconfirmed?' <small style="color:#ccc">*</small>':''}}</td>
      <td>${{badge(i.status)}}</td>
    </tr>`;
  }}).join('');
}}

// ── Volume chart — kolli per month per carrier ─────────────────────────────────
// NOTE: this counts kolli (rows), not distinct sändningar/försändelser — the
// exported CSVs don't carry a shipment_number column, so there's no way to
// deduplicate multi-kolli shipments down to a true shipment count here (see
// src/run_exporter.py for where that distinction exists, in-memory only, for
// the Excel report's Sändningar column).
function renderVolumeChart(svcData) {{
  // Aggregate kolli counts by (month, carrier) from SVC_DATA
  const agg = {{}};
  svcData.forEach(s => {{
    const k = s.month + '||' + s.carrier;
    if (!agg[k]) agg[k] = {{month: s.month, carrier: s.carrier, count: 0}};
    agg[k].count += s.count;
  }});
  const months = [...new Set(Object.values(agg).map(v => v.month))].filter(Boolean).sort();
  const cars   = [...new Set(Object.values(agg).map(v => v.carrier))].sort();
  const datasets = cars.map(c => ({{
    label: c,
    data: months.map(m => (agg[m+'||'+c] || {{count:0}}).count),
    backgroundColor: cColA(c), borderColor: cCol(c), borderWidth:1,
  }}));

  const total = Object.values(agg).reduce((s,v) => s+v.count, 0);
  document.getElementById('volumeNote').textContent = total ? `(${{total.toLocaleString('sv-SE')}} tot.)` : '(ingen data)';

  if (volumeChart) {{
    volumeChart.data = {{labels:months, datasets}};
    volumeChart.update('active');
  }} else {{
    volumeChart = new Chart(document.getElementById('volumeChart'), {{
      type: 'bar',
      data: {{labels:months, datasets}},
      options: {{
        responsive:true,
        maintainAspectRatio:false,
        plugins: {{legend:{{position:'bottom'}},
          tooltip:{{callbacks:{{label:ctx=>' '+ctx.parsed.y.toLocaleString('sv-SE')+' st'}}}}}},
        scales: {{
          x:{{grid:{{display:false}}}},
          y:{{beginAtZero:true, ticks:{{callback:v=>v.toLocaleString('sv-SE')}}, grid:{{color:'#f5f5f5'}}}},
        }},
      }},
    }});
  }}
}}

// ── Timeline chart (grouped bar) ─────────────────────────────────────────────
function setGranularity(g) {{
  _tlGran = g;
  document.getElementById('btnDaily').classList.toggle('active',   g==='daily');
  document.getElementById('btnMonthly').classList.toggle('active', g==='monthly');
  document.getElementById('btnYearly').classList.toggle('active',  g==='yearly');
  applyFilters();
}}

function renderTimelineChart(lineData) {{
  const recon  = lineData.filter(i => !UNCONFIRMED.has(i.status));
  const getKey = i => _tlGran==='yearly' ? i.year : (_tlGran==='daily' ? i.date : i.month);
  const labels = [...new Set(recon.map(getKey))].filter(Boolean).sort();
  const cars   = [...new Set(recon.map(i => i.carrier))].sort();
  const datasets = cars.map(c => ({{
    label:c,
    data: labels.map(lbl=>recon.filter(i=>getKey(i)===lbl&&i.carrier===c).reduce((s,i)=>s+i.amount,0)),
    backgroundColor:cColA(c), borderColor:cCol(c), borderWidth:1,
  }}));

  const unit = _tlGran==='yearly' ? 'år' : (_tlGran==='daily' ? 'dag' : 'månad');
  document.getElementById('timelineNote').textContent =
    labels.length ? `(${{labels.length}} ${{unit}})` : '(ingen data)';

  if (timelineChart) {{
    timelineChart.data = {{labels, datasets}};
    timelineChart.update('active');
  }} else {{
    timelineChart = new Chart(document.getElementById('timelineChart'), {{
      type:'bar',
      data:{{labels, datasets}},
      options:{{
        responsive:true,
        interaction:{{mode:'index',intersect:false}},
        plugins:{{legend:{{position:'bottom'}},
          tooltip:{{callbacks:{{label:ctx=>' '+ctx.parsed.y.toLocaleString('sv-SE',{{minimumFractionDigits:2}})+' SEK'}}}}}},
        scales:{{
          x:{{grid:{{display:false}}}},
          y:{{beginAtZero:true, ticks:{{callback:v=>v.toLocaleString('sv-SE')}}, grid:{{color:'#f5f5f5'}}}},
        }},
      }},
    }});
  }}
}}

// ── Surcharge chart ───────────────────────────────────────────────────────────
function renderSurchargeChart(scData) {{
  const agg={{}};
  scData.forEach(s=>{{ agg[s.surcharge_cat]=(agg[s.surcharge_cat]||0)+s.total; }});
  const sorted=Object.entries(agg).sort((a,b)=>b[1]-a[1]);
  const labels=sorted.map(x=>x[0]), data=sorted.map(x=>Math.round(x[1]*100)/100);

  document.getElementById('scCount').textContent = labels.length ? `(${{labels.length}} kategorier)` : '';

  if (surchargeChart) {{
    surchargeChart.data={{labels, datasets:[{{label:'SEK', data, backgroundColor:'#1565c0'}}]}};
    surchargeChart.update('active');
  }} else {{
    surchargeChart = new Chart(document.getElementById('surchargeChart'), {{
      type:'bar',
      data:{{labels, datasets:[{{label:'SEK', data, backgroundColor:'rgba(21,101,192,.75)'}}]}},
      options:{{
        indexAxis:'y', responsive:true,
        plugins:{{legend:{{display:false}}}},
        scales:{{x:{{beginAtZero:true, ticks:{{callback:v=>v.toLocaleString('sv-SE')}}}}}},
      }},
    }});
  }}
}}

// ── Avg cost trend per service type over time ─────────────────────────────────
function renderAvgCostTrendChart(filtData) {{
  const agg={{}};
  filtData.forEach(s=>{{
    if(!s.month||!s.count) return;
    const k=s.month+'||'+s.carrier+'||'+s.service_cat;
    if(!agg[k]) agg[k]={{month:s.month,carrier:s.carrier,cat:s.service_cat,count:0,grand:0}};
    agg[k].count+=s.count; agg[k].grand+=s.grand_total;
  }});
  const months=[...new Set(Object.values(agg).map(v=>v.month))].filter(Boolean).sort();
  // Unique (carrier,cat) combos sorted by total grand
  const seen=new Set(); const combos=[];
  Object.values(agg).forEach(v=>{{
    const k=v.carrier+'||'+v.cat;
    if(!seen.has(k)){{seen.add(k);combos.push({{carrier:v.carrier,cat:v.cat,total:0}});}}
    combos.find(c=>c.carrier===v.carrier&&c.cat===v.cat).total+=v.grand;
  }});
  combos.sort((a,b)=>b.total-a.total);
  // Line style per service index within carrier
  const DASHES=[[],[6,3],[2,3],[6,2,2,2]];
  const carIdx={{}};
  combos.forEach(c=>{{
    if(carIdx[c.carrier]===undefined) carIdx[c.carrier]=0;
    c.dashIdx=carIdx[c.carrier]++;
  }});
  const datasets=combos.map(c=>{{
    const color=cCol(c.carrier);
    const data=months.map(m=>{{
      const e=agg[m+'||'+c.carrier+'||'+c.cat];
      return e&&e.count>0?Math.round(e.grand/e.count*100)/100:null;
    }});
    return {{
      label:`${{c.carrier}} – ${{c.cat}}`,
      data, borderColor:color, backgroundColor:color+'18',
      borderWidth:2, borderDash:DASHES[c.dashIdx]||[],
      pointRadius:3, tension:.3, spanGaps:true,
    }};
  }});
  if(avgTrendChart){{
    avgTrendChart.data={{labels:months,datasets}};
    avgTrendChart.update('active');
  }}else{{
    avgTrendChart=new Chart(document.getElementById('avgTrendChart'),{{
      type:'line',
      data:{{labels:months,datasets}},
      options:{{
        responsive:true,
        interaction:{{mode:'index',intersect:false}},
        plugins:{{
          legend:{{position:'bottom',labels:{{boxWidth:14,font:{{size:11}}}}}},
          tooltip:{{callbacks:{{label:ctx=>' '+fmt(ctx.parsed.y)}}}},
        }},
        scales:{{
          x:{{grid:{{display:false}}}},
          y:{{beginAtZero:false,
             ticks:{{callback:v=>v.toLocaleString('sv-SE')}},
             grid:{{color:'#f5f5f5'}}}},
        }},
      }},
    }});
  }}
}}

// ── Chargeable weight trend per Parcel-family service ──────────────────────────
function renderWeightTrendChart(filtData) {{
  const agg={{}};
  filtData.forEach(w=>{{
    if(!w.month||!w.count) return;
    const k=w.month+'||'+w.carrier+'||'+w.service_cat;
    if(!agg[k]) agg[k]={{month:w.month,carrier:w.carrier,cat:w.service_cat,count:0,sum:0}};
    agg[k].count+=w.count; agg[k].sum+=w.weight_sum;
  }});
  const months=[...new Set(Object.values(agg).map(v=>v.month))].filter(Boolean).sort();

  // Unique (carrier,cat) combos, sorted by total kolli count
  const seen=new Set(); const combos=[];
  Object.values(agg).forEach(v=>{{
    const k=v.carrier+'||'+v.cat;
    if(!seen.has(k)){{seen.add(k);combos.push({{carrier:v.carrier,cat:v.cat,count:0,sum:0}});}}
    const c=combos.find(c=>c.carrier===v.carrier&&c.cat===v.cat);
    c.count+=v.count; c.sum+=v.sum;
  }});
  combos.sort((a,b)=>b.count-a.count);

  const DASHES=[[],[6,3],[2,3],[6,2,2,2]];
  const carIdx={{}};
  combos.forEach(c=>{{
    if(carIdx[c.carrier]===undefined) carIdx[c.carrier]=0;
    c.dashIdx=carIdx[c.carrier]++;
    c.totalAvg=c.count>0?c.sum/c.count:null;
  }});

  // Rolling 3-month average per combo per month (trailing window, not an
  // average-of-averages — months with more kolli weigh more)
  const datasets=combos.map(c=>{{
    const color=cCol(c.carrier);
    const data=months.map((m,i)=>{{
      const win=months.slice(Math.max(0,i-2),i+1);
      let wsum=0,wcnt=0;
      win.forEach(wm=>{{
        const e=agg[wm+'||'+c.carrier+'||'+c.cat];
        if(e){{wsum+=e.sum; wcnt+=e.count;}}
      }});
      return wcnt>0?Math.round(wsum/wcnt*100)/100:null;
    }});
    return {{
      label:`${{c.carrier}} – ${{c.cat}}`,
      data, borderColor:color, backgroundColor:color+'18',
      borderWidth:2, borderDash:DASHES[c.dashIdx]||[],
      pointRadius:3, tension:.3, spanGaps:true,
    }};
  }});

  document.getElementById('weightTrendNote').textContent =
    months.length ? `(${{months.length}} månad, 3-mån rullande snitt)` : '(ingen data)';

  const body = document.getElementById('weightTrendBody');
  body.innerHTML = combos.length
    ? combos.map(c => `<tr>
        <td>${{esc(c.carrier)}} – ${{esc(c.cat)}}</td>
        <td class="num">${{c.count.toLocaleString('sv-SE')}}</td>
        <td class="num">${{c.totalAvg!=null?fmt(c.totalAvg):'–'}} kg</td>
      </tr>`).join('')
    : '<tr><td colspan="3" class="no-data">Ingen viktdata.</td></tr>';

  if (weightTrendChart) {{
    weightTrendChart.data = {{labels:months, datasets}};
    weightTrendChart.update('active');
  }} else {{
    weightTrendChart = new Chart(document.getElementById('weightTrendChart'), {{
      type:'line',
      data:{{labels:months, datasets}},
      options:{{
        responsive:true,
        interaction:{{mode:'index',intersect:false}},
        plugins:{{
          legend:{{position:'bottom',labels:{{boxWidth:14,font:{{size:11}}}}}},
          tooltip:{{callbacks:{{label:ctx=>' '+fmt(ctx.parsed.y)+' kg'}}}},
        }},
        scales:{{
          x:{{grid:{{display:false}}}},
          y:{{beginAtZero:false,
             ticks:{{callback:v=>v.toLocaleString('sv-SE')+' kg'}},
             grid:{{color:'#f5f5f5'}}}},
        }},
      }},
    }});
  }}
}}

// ── Cost per service table (totals + averages + share) ────────────────────────
function renderServiceCostTable(filtData) {{
  const agg={{}};
  filtData.forEach(s=>{{
    const k=s.carrier+'||'+s.service_cat;
    if(!agg[k]) agg[k]={{carrier:s.carrier,cat:s.service_cat,count:0,base:0,sc:0,grand:0}};
    agg[k].count+=s.count; agg[k].base+=s.base_total;
    agg[k].sc+=s.sc_grand; agg[k].grand+=s.grand_total;
  }});
  const rows=Object.values(agg).sort((a,b)=>b.grand-a.grand);
  const grandTotal=rows.reduce((s,r)=>s+r.grand,0);
  const totalCount=rows.reduce((s,r)=>s+r.count,0);
  document.getElementById('svcCostCount').textContent=`(${{rows.length}} typ${{rows.length!==1?'er':''}})`;
  document.getElementById('svcCostHead').innerHTML=`<tr>
    <th>Tjänstetyp</th><th>Transportör</th><th class="num">Antal</th>
    <th class="num">Total</th><th class="num">Andel</th>
    <th class="num">Snitt bas</th><th class="num">Snitt tillägg</th>
    <th class="num">Snitt total</th></tr>`;
  if(!rows.length){{document.getElementById('svcCostBody').innerHTML='<tr><td colspan="8" class="no-data">Ingen data.</td></tr>';return;}}
  document.getElementById('svcCostBody').innerHTML=rows.map(r=>{{
    const n=r.count||1;
    const pct=grandTotal>0?r.grand/grandTotal*100:0;
    return `<tr>
      <td><strong>${{esc(r.cat)}}</strong></td><td>${{esc(r.carrier)}}</td>
      <td class="num">${{r.count.toLocaleString('sv-SE')}}</td>
      <td class="num">${{fmt(r.grand)}}</td>
      <td class="num">${{pct.toFixed(1)}}%</td>
      <td class="num">${{fmt(r.base/n)}}</td>
      <td class="num">${{fmt(r.sc/n)}}</td>
      <td class="num"><strong>${{fmt(r.grand/n)}}</strong></td></tr>`;
  }}).join('');
  document.getElementById('svcCostBody').innerHTML+=`<tr style="background:#f0f4ff;font-weight:700">
    <td>TOTAL</td><td>—</td>
    <td class="num">${{totalCount.toLocaleString('sv-SE')}}</td>
    <td class="num">${{fmt(grandTotal)}}</td>
    <td class="num">100.0%</td>
    <td colspan="3" class="num">—</td></tr>`;
}}

// ── Returns table (per carrier per month) ──────────────────────────────────────
let _returnDetailData = [];
function renderReturnsTable(filtData) {{
  _returnDetailData = filtData;
  const agg={{}};
  filtData.forEach(r=>{{
    const k=r.carrier+'||'+r.month;
    if(!agg[k]) agg[k]={{carrier:r.carrier,month:r.month,count:0,amount:0}};
    agg[k].count+=1; agg[k].amount+=r.amount;
  }});
  const rows=Object.values(agg).sort((a,b)=> (b.month||'').localeCompare(a.month||'') || b.amount-a.amount);
  const totalCount=filtData.length;
  const totalAmt=filtData.reduce((s,r)=>s+r.amount,0);
  document.getElementById('returnsCount').textContent=totalCount?`(${{totalCount}} st, ${{fmt(totalAmt)}} SEK)`:'';
  document.getElementById('returnsBody').innerHTML = rows.length
    ? rows.map(r=>`<tr class="inv-row" onclick="openModalView(showReturnDetail, ['${{esc(r.carrier)}}','${{esc(r.month)}}'])">
        <td>${{esc(r.carrier)}}</td><td>${{esc(r.month)||'–'}}</td>
        <td class="num">${{r.count.toLocaleString('sv-SE')}}</td>
        <td class="num">${{fmt(r.amount)}}</td></tr>`).join('')
    : '<tr><td colspan="4" class="no-data">Inga returer.</td></tr>';
}}

function showReturnDetail(carrier, month) {{
  const rows = _returnDetailData
    .filter(r => r.carrier===carrier && (r.month||'')===month)
    .sort((a,b) => (b.date||'').localeCompare(a.date||''));
  if (!rows.length) return;

  const total = rows.reduce((s,r)=>s+r.amount,0);
  const reason = carrier==='PostNord'
    ? 'Sändning till egen lageradress (43149) — retur.'
    : 'Avgift för misslyckat leveransförsök / retur.';
  const monthLabel = month || 'okänd månad';
  const repeatRows = rows.filter(r => r.repeat_count > 1);

  document.getElementById('modalTitle').textContent =
    `${{carrier}} · Returer ${{monthLabel}}`;

  const showOrigin = carrier === 'PostNord';
  const rowsHtml = rows.map(r => {{
    const repeatBadge = r.repeat_count > 1
      ? ` <span style="background:#fde8e8;color:#c62828;font-size:10px;font-weight:700;
          padding:1px 6px;border-radius:10px;margin-left:4px">returnerad ${{r.repeat_count}}x</span>`
      : '';
    const kolliTd = showOrigin
      ? `<td style="font-size:11px;white-space:nowrap">${{esc(r.kolli_id)||'—'}}${{
          r.kolli_id ? ` <button class="copy-btn" onclick="event.stopPropagation();copyToClipboard('${{esc(r.kolli_id)}}',this)">Kopiera</button>` : ''
        }}</td>` : '';
    const originTd = showOrigin
      ? `<td>${{esc(r.origin)||'<span style="color:#ccc">okänt</span>'}}</td>` : '';
    // PostNord rows drill into the kolli's full shipment history; Bring has
    // no kolli-level linkage (see reason text above) so falls back to the
    // invoice it was billed on.
    const rowClick = (showOrigin && r.kolli_id)
      ? `openModalView(showKolliDetail, ['${{esc(r.kolli_id)}}'])`
      : `openModalView(showInvoiceDetail, ['${{esc(r.invoice_no)}}'])`;
    return `<tr onclick="${{rowClick}}" style="cursor:pointer">
      <td>${{esc(r.date)||'—'}}</td>
      <td><strong style="font-size:11px">${{esc(r.invoice_no)}}</strong>${{repeatBadge}}</td>
      ${{kolliTd}}
      ${{originTd}}
      <td class="num">${{fmt(r.amount)}}</td>
    </tr>`;
  }}).join('');

  const kolliCol = showOrigin ? '<th>Kollinummer</th>' : '';
  const originCol = showOrigin ? '<th>Ursprunglig destination</th>' : '';
  const repeatNote = repeatRows.length
    ? `<p style="font-size:11px;color:#c62828;margin-top:8px">
        ${{repeatRows.length}} sändning${{repeatRows.length!==1?'ar':''}} denna månad har
        returnerats mer än en gång — se markerade rader.</p>`
    : '';

  document.getElementById('modalBody').innerHTML = `
    <div class="modal-section">
      <h4>Sammanfattning</h4>
      <div class="info-grid">
        <div class="info-item"><div class="ilabel">Transportör</div><div class="ivalue">${{esc(carrier)}}</div></div>
        <div class="info-item"><div class="ilabel">Månad</div><div class="ivalue">${{esc(monthLabel)}}</div></div>
        <div class="info-item"><div class="ilabel">Antal returer</div><div class="ivalue">${{rows.length}}</div></div>
        <div class="info-item"><div class="ilabel">Total kostnad</div><div class="ivalue">${{fmt(total)}} SEK</div></div>
      </div>
      <p style="font-size:11px;color:#999;margin-top:8px">${{esc(reason)}}</p>
      ${{repeatNote}}
    </div>
    <div class="modal-section">
      <h4>Returer denna månad</h4>
      <table>
        <thead><tr><th>Datum</th><th>Faktura #</th>${{kolliCol}}${{originCol}}<th class="num">Kostnad (SEK)</th></tr></thead>
        <tbody>${{rowsHtml}}</tbody>
      </table>
      <p style="font-size:11px;color:#999;margin-top:8px">${{
        showOrigin ? 'Klicka på en rad för att se hela kollits historik.'
                   : 'Klicka på en rad för att öppna fakturan.'
      }}</p>
    </div>`;

  document.getElementById('modalOverlay').classList.add('open');
  document.body.style.overflow='hidden';
}}

// ── Kolli detail (full shipment history for one PostNord kolli_id) ─────────────
function showKolliDetail(kolliId) {{
  const hist = (KOLLI_HISTORY[kolliId] || []).slice().sort((a,b) => (a.date||'').localeCompare(b.date||''));

  document.getElementById('modalTitle').innerHTML =
    `Kolli ${{esc(kolliId)}} <button class="copy-btn" onclick="copyToClipboard('${{esc(kolliId)}}',this)">Kopiera</button>`;

  if (!hist.length) {{
    document.getElementById('modalBody').innerHTML =
      '<div class="modal-section"><p class="no-data">Ingen historik hittad för detta kolli.</p></div>';
    document.getElementById('modalOverlay').classList.add('open');
    document.body.style.overflow='hidden';
    return;
  }}

  const returns = hist.filter(h => h.is_return);
  const total = hist.reduce((s,h) => s+h.amount, 0);

  const rowsHtml = hist.map(h => {{
    const dest = `${{esc(h.postal)}} ${{esc(h.city)}}`.trim() || '<span style="color:#ccc">okänt</span>';
    const typeBadge = h.is_return
      ? '<span style="background:#fde8e8;color:#c62828;font-size:10px;font-weight:700;padding:1px 7px;border-radius:10px">RETUR</span>'
      : '<span style="background:#e8f5e9;color:#2e7d32;font-size:10px;font-weight:700;padding:1px 7px;border-radius:10px">LEVERANS</span>';
    return `<tr onclick="openModalView(showInvoiceDetail, ['${{esc(h.invoice_no)}}'])" style="cursor:pointer">
      <td>${{esc(h.date)||'—'}}</td>
      <td>${{typeBadge}}</td>
      <td>${{dest}}</td>
      <td><strong style="font-size:11px">${{esc(h.invoice_no)}}</strong></td>
      <td class="num">${{fmt(h.amount)}}</td>
    </tr>`;
  }}).join('');

  const repeatNote = returns.length > 1
    ? `<p style="font-size:11px;color:#c62828;margin-top:8px">
        Detta kolli har returnerats ${{returns.length}} gånger — ovanligt, värt att undersöka.</p>`
    : '';

  document.getElementById('modalBody').innerHTML = `
    <div class="modal-section">
      <h4>Sammanfattning</h4>
      <div class="info-grid">
        <div class="info-item"><div class="ilabel">Kollinummer</div><div class="ivalue">${{esc(kolliId)}}</div></div>
        <div class="info-item"><div class="ilabel">Antal händelser</div><div class="ivalue">${{hist.length}}</div></div>
        <div class="info-item"><div class="ilabel">Antal returer</div><div class="ivalue">${{returns.length}}</div></div>
        <div class="info-item"><div class="ilabel">Total kostnad</div><div class="ivalue">${{fmt(total)}} SEK</div></div>
      </div>
      ${{repeatNote}}
    </div>
    <div class="modal-section">
      <h4>Historik för detta kolli</h4>
      <table>
        <thead><tr><th>Datum</th><th>Typ</th><th>Destination</th><th>Faktura #</th><th class="num">Kostnad (SEK)</th></tr></thead>
        <tbody>${{rowsHtml}}</tbody>
      </table>
      <p style="font-size:11px;color:#999;margin-top:8px">Klicka på en rad för att öppna fakturan.</p>
    </div>`;

  document.getElementById('modalOverlay').classList.add('open');
  document.body.style.overflow='hidden';
}}

// ── Anomaly table ─────────────────────────────────────────────────────────────
function renderAnomalyTable(data) {{
  document.getElementById('anomalyCount').textContent=data.length?`(${{data.length}})`:'';
  const body=document.getElementById('anomalyBody');
  if(!data.length){{body.innerHTML='<tr><td colspan="6" class="no-data">Inga avvikelser för valda filter.</td></tr>';return;}}
  const sS={{Warning:'background:#fff8e1;color:#f57f17',Error:'background:#ffebee;color:#c62828',Info:'background:#e3f2fd;color:#1565c0'}};
  body.innerHTML=data.map(a=>{{
    const sty=sS[a.severity]||'background:#f5f5f5;color:#333';
    const exp=a.explanation?`<span style="color:#555">${{esc(a.explanation)}}</span>`:`<span style="color:#ccc;font-style:italic">—</span>`;
    return `<tr><td>${{esc(a.carrier)}}</td><td>${{esc(a.invoice_no)}}</td><td>${{esc(a.type)}}</td>
      <td><span class="badge" style="${{sty}}">${{esc(a.severity)}}</span></td>
      <td>${{esc(a.description)}}</td><td>${{exp}}</td></tr>`;
  }}).join('');
}}

// ── Checks table ──────────────────────────────────────────────────────────────
function renderChecks() {{
  const body=document.getElementById('checkBody');
  if(!CHECK_DATA.length){{body.innerHTML='<tr><td colspan="5" class="no-data">Inga kontroller ännu.</td></tr>';return;}}
  body.innerHTML=CHECK_DATA.map(c=>`<tr>
    <td>${{esc(c.carrier)}}</td><td>${{esc(c.invoice_no)}}</td><td>${{esc(c.check)}}</td>
    <td>${{badge(c.status)}}</td><td>${{esc(c.message)}}</td></tr>`).join('');
}}

// ── Invoice detail modal ───────────────────────────────────────────────────────
function showInvoiceDetail(invNo) {{
  const inv=INV_DATA.find(i=>i.invoice_no===invNo);
  if(!inv) return;
  document.getElementById('modalTitle').textContent=`${{inv.carrier}} · Faktura ${{invNo}}`;
  const isPending=inv.status==='Pending';
  const isSpecOnly=inv.status==='SpecOnly';

  const fields=[
    ['Faktura #',invNo],['Transportör',inv.carrier],
    ['Fakturadatum',inv.date||'—'],['Förfallodatum',inv.due_date||'—'],
    ['Kundnummer',inv.customer_number||'—'],['Valuta',inv.currency||'SEK'],
    ['Total ex-moms', fmt(inv.total)+' SEK'+(isPending||isSpecOnly?' (spec)':'')],
    ['Moms', inv.vat_amount?fmt(inv.vat_amount)+' SEK':'—'],
    ['Total inkl moms', inv.total_inc_vat?fmt(inv.total_inc_vat)+' SEK':'—'],
    ['Status', badge(inv.status)],['Källfil', esc(inv.source_file||'—')],
  ];
  if(isPending&&inv.pending_note) fields.push(['Notering',esc(inv.pending_note)]);
  if(isSpecOnly) fields.push(['Notering','Historisk, endast specifikation — PDF kommer aldrig att levereras']);

  const infoHtml=`<div class="info-grid">${{fields.map(([l,v])=>
    `<div class="info-item"><div class="ilabel">${{esc(l)}}</div><div class="ivalue">${{v}}</div></div>`
  ).join('')}}</div>`;

  const svcRows=SVC_COST_DATA.filter(s=>s.invoice_no===invNo).sort((a,b)=>b.grand_total-a.grand_total);
  const svcHtml=svcRows.length?`<table>
    <thead><tr><th>Tjänstetyp</th><th class="num">Antal</th><th class="num">Bas (SEK)</th><th class="num">Tillägg (SEK)</th><th class="num">Total (SEK)</th></tr></thead>
    <tbody>${{svcRows.map(r=>`<tr><td><strong>${{esc(r.service_cat)}}</strong></td>
      <td class="num">${{r.count.toLocaleString('sv-SE')}}</td>
      <td class="num">${{fmt(r.base_total)}}</td><td class="num">${{fmt(r.sc_grand)}}</td>
      <td class="num"><strong>${{fmt(r.grand_total)}}</strong></td></tr>`).join('')}}</tbody>
  </table>`:'<p class="no-data">Ingen raddata.</p>';

  const scRows=SC_DATA.filter(s=>s.invoice_no===invNo).sort((a,b)=>b.total-a.total);
  const scHtml=scRows.length?`<table>
    <thead><tr><th>Tilläggstyp</th><th class="num">Belopp (SEK)</th></tr></thead>
    <tbody>${{scRows.map(r=>`<tr><td>${{esc(r.surcharge_cat)}}</td><td class="num">${{fmt(r.total)}}</td></tr>`).join('')}}</tbody>
  </table>`:'<p class="no-data">Inga tillägg.</p>';

  const anom=ANOMALY_DATA.filter(a=>a.invoice_no===invNo);
  const anomHtml=anom.length?`<table>
    <thead><tr><th>Typ</th><th>Allvarlighet</th><th>Beskrivning</th></tr></thead>
    <tbody>${{anom.map(a=>`<tr><td>${{esc(a.type)}}</td><td>${{badge(a.severity)}}</td><td>${{esc(a.description)}}</td></tr>`).join('')}}</tbody>
  </table>`:'<p class="no-data">Inga avvikelser.</p>';

  const chks=CHECK_DATA.filter(c=>c.invoice_no===invNo);
  const chkHtml=chks.length?`<table>
    <thead><tr><th>Kontroll</th><th>Status</th><th>Meddelande</th></tr></thead>
    <tbody>${{chks.map(c=>`<tr><td>${{esc(c.check)}}</td><td>${{badge(c.status)}}</td><td>${{esc(c.message)}}</td></tr>`).join('')}}</tbody>
  </table>`:'<p class="no-data">Inga kontrollresultat.</p>';

  document.getElementById('modalBody').innerHTML=`
    <div class="modal-section"><h4>Fakturainformation</h4>${{infoHtml}}</div>
    <div class="modal-section"><h4>Tjänstesammansättning</h4>${{svcHtml}}</div>
    <div class="modal-section"><h4>Tilläggsavgifter</h4>${{scHtml}}</div>
    <div class="modal-section"><h4>Avvikelser</h4>${{anomHtml}}</div>
    <div class="modal-section"><h4>Kontroller</h4>${{chkHtml}}</div>`;

  document.getElementById('modalOverlay').classList.add('open');
  document.body.style.overflow='hidden';
}}

function closeModal() {{
  document.getElementById('modalOverlay').classList.remove('open');
  document.body.style.overflow='';
  _modalStack = []; _currentView = null;
  document.getElementById('modalBackBtn').style.display = 'none';
}}
function closeModalOnBg(e) {{ if(e.target===document.getElementById('modalOverlay')) closeModal(); }}
document.addEventListener('keydown', e=>{{ if(e.key==='Escape') closeModal(); }});

// ── Modal navigation (back-stack) ──────────────────────────────────────────────
// Each show*Detail function only renders — it doesn't know or care whether it
// was reached via a fresh click or a "Tillbaka" press. openModalView() is the
// only thing that pushes history, so every entry point into the modal must go
// through it (never call show*Detail directly from onclick).
let _modalStack = [];
let _currentView = null;

function openModalView(fn, args) {{
  if (_currentView) _modalStack.push(_currentView);
  _currentView = {{fn, args}};
  fn(...args);
  document.getElementById('modalBackBtn').style.display = _modalStack.length ? 'inline-flex' : 'none';
}}

function modalGoBack() {{
  const prev = _modalStack.pop();
  if (!prev) return;
  _currentView = prev;
  prev.fn(...prev.args);
  document.getElementById('modalBackBtn').style.display = _modalStack.length ? 'inline-flex' : 'none';
}}

// ── Copy to clipboard ───────────────────────────────────────────────────────
function copyToClipboard(text, btnEl) {{
  const done = () => {{
    if (!btnEl) return;
    const orig = btnEl.textContent;
    btnEl.textContent = '✓ Kopierat';
    setTimeout(() => {{ btnEl.textContent = orig; }}, 1200);
  }};
  if (navigator.clipboard && navigator.clipboard.writeText) {{
    navigator.clipboard.writeText(text).then(done).catch(() => _fallbackCopy(text, done));
  }} else {{
    _fallbackCopy(text, done);
  }}
}}
function _fallbackCopy(text, cb) {{
  const ta = document.createElement('textarea');
  ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
  document.body.appendChild(ta); ta.select();
  try {{ document.execCommand('copy'); }} catch (e) {{}}
  document.body.removeChild(ta);
  if (cb) cb();
}}

// ── Init ──────────────────────────────────────────────────────────────────────
applyFilters();
renderChecks();
</script>
</body>
</html>"""

    out_path = config.DASHBOARD_HTML
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    logger.info("DashboardWriter", f"Dashboard updated: {out_path.name}")
