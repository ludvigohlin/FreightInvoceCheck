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


def write_html_dashboard(logger: ProcessingLogger) -> None:
    """Read all output CSVs and write 02_Output/dashboard.html."""
    invoices = _read_csv(config.INVOICE_HEADER_CSV)
    lines = _read_csv(config.INVOICE_LINES_CSV)
    surcharges = _read_csv(config.SURCHARGE_LINES_CSV)
    checks = _read_csv(config.INVOICE_CHECKS_CSV)
    anomalies_raw = _read_csv(config.ANOMALIES_CSV)
    pending_raw = _read_csv(config.PENDING_INVOICES_CSV)

    # Lookup: invoice_number → {invoice_date, carrier}
    inv_lookup = {inv["invoice_number"]: inv for inv in invoices}

    # ── Build JS data: one object per invoice ─────────────────────────────────
    # Fully reconciled invoices
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

    # Pending (incomplete document set) invoices — added with status="Pending"
    existing_inv_nos = {i["invoice_no"] for i in inv_js}
    for p in pending_raw:
        inv_no = p.get("invoice_number", "")
        if inv_no in existing_inv_nos:
            continue  # graduated to fully processed — skip pending record
        inv_js.append({
            "date": "",
            "year": "",
            "month": "",
            "carrier": p.get("carrier", "Bring"),
            "invoice_no": inv_no,
            "total": _safe_float(p.get("known_total_ex_vat")),
            "currency": "SEK",
            "status": "Pending",
            "customer_number": "",
            "due_date": "",
            "vat_amount": 0.0,
            "total_inc_vat": 0.0,
            "source_file": p.get("source_file", ""),
            "pending_note": p.get("note", ""),
        })

    # ── Build JS data: aggregated service breakdown per (invoice, service_cat) ─
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
        {
            "invoice_no": k[0],
            "service_cat": k[1],
            "carrier": v["carrier"],
            "date": v["date"],
            "year": v["year"],
            "month": v["month"],
            "count": v["count"],
            "total": round(v["total"], 2),
        }
        for k, v in svc_agg.items()
    ]

    # ── Build surcharge data: aggregated per (invoice, surcharge_cat) ─────────
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

    # ── Build service cost breakdown: base + surcharges per service type ─────
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
        inv_no = sc.get("invoice_number", "")
        svc_cat = sc.get("related_service_category", "") or "Unknown"
        sc_cat = sc.get("surcharge_category", "Unknown") or "Unknown"
        sc_svc_agg[(inv_no, svc_cat, sc_cat)] += _safe_float(sc.get("amount"))

    all_sc_cats = sorted({k[2] for k in sc_svc_agg})

    svc_cost_js = []
    for (inv_no, svc_cat), v in base_agg.items():
        sc_by_cat = {
            sc_cat: round(sc_svc_agg.get((inv_no, svc_cat, sc_cat), 0.0), 2)
            for sc_cat in all_sc_cats
        }
        sc_grand = sum(sc_by_cat.values())
        svc_cost_js.append({
            "invoice_no": inv_no,
            "service_cat": svc_cat,
            "carrier": v["carrier"],
            "year": v["year"],
            "month": v["month"],
            "count": v["count"],
            "base_total": round(v["total"], 2),
            "sc_by_cat": {k: val for k, val in sc_by_cat.items() if val > 0},
            "sc_grand": round(sc_grand, 2),
            "grand_total": round(v["total"] + sc_grand, 2),
        })

    # ── Build anomaly JS data ─────────────────────────────────────────────────
    anomaly_js = [
        {
            "carrier": a.get("carrier", ""),
            "invoice_no": a.get("invoice_number", ""),
            "type": a.get("anomaly_type", ""),
            "severity": a.get("severity", ""),
            "description": a.get("description", ""),
            "detail": a.get("detail", ""),
            "value": _safe_float(a.get("value")),
            "threshold": _safe_float(a.get("threshold")),
            "suggested_action": a.get("suggested_action", ""),
            "explanation": a.get("claude_explanation", ""),
        }
        for a in reversed(anomalies_raw[-200:] if len(anomalies_raw) > 200 else anomalies_raw)
    ]

    # ── Recent checks (last 60, newest first) ─────────────────────────────────
    check_js = [
        {"carrier": c.get("carrier", ""), "invoice_no": c.get("invoice_number", ""),
         "check": c.get("check_name", ""), "status": c.get("status", ""),
         "message": c.get("message", "")}
        for c in reversed(checks[-60:] if len(checks) > 60 else checks)
    ]

    # ── Filter option lists ───────────────────────────────────────────────────
    years = sorted({i["year"] for i in inv_js if i["year"]}, reverse=True)
    carriers = sorted({i["carrier"] for i in inv_js if i["carrier"]})
    service_cats = sorted({s["service_cat"] for s in svc_js if s["service_cat"]})
    months_all = sorted({i["month"] for i in inv_js if i["month"]}, reverse=True)

    year_opts = "".join(f'<option value="{_e(y)}">{_e(y)}</option>' for y in years)
    month_opts = "".join(f'<option value="{_e(m)}">{_e(m)}</option>' for m in months_all)
    carrier_opts = "".join(f'<option value="{_e(c)}">{_e(c)}</option>' for c in carriers)
    svc_opts = "".join(f'<option value="{_e(s)}">{_e(s)}</option>' for s in service_cats)

    last_updated = datetime.now().strftime("%Y-%m-%d %H:%M")

    # Embed data as JSON
    inv_json = json.dumps(inv_js)
    svc_json = json.dumps(svc_js)
    sc_json = json.dumps(sc_js)
    check_json = json.dumps(check_js)
    svc_cost_json = json.dumps(svc_cost_js)
    anomaly_json = json.dumps(anomaly_js)
    sc_cats_json = json.dumps(all_sc_cats)

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
            display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100}}
    header h1{{font-size:18px;font-weight:600}}
    header .meta{{font-size:12px;opacity:.8}}
    .filter-bar{{background:#fff;border-bottom:1px solid #ddd;padding:10px 24px;
                 display:flex;flex-wrap:wrap;gap:10px;align-items:center;
                 position:sticky;top:49px;z-index:99;box-shadow:0 1px 4px rgba(0,0,0,.06)}}
    .filter-bar label{{font-size:12px;color:#555;font-weight:600}}
    .filter-bar select,.filter-bar input{{
      padding:5px 8px;border:1px solid #ccc;border-radius:4px;
      font-size:13px;background:#fff;min-width:120px}}
    .filter-bar input{{min-width:160px}}
    .filter-bar button{{padding:5px 14px;background:#1565c0;color:#fff;
                        border:none;border-radius:4px;cursor:pointer;font-size:13px}}
    .filter-bar button:hover{{background:#0d47a1}}
    .filter-bar .reset{{background:#757575}}
    .filter-bar .reset:hover{{background:#424242}}
    .filter-sep{{width:1px;height:24px;background:#ddd;margin:0 4px}}
    main{{max-width:1300px;margin:0 auto;padding:20px 16px}}
    .kpi-row{{display:flex;gap:14px;margin-bottom:20px;flex-wrap:wrap}}
    .kpi{{background:#fff;border-radius:8px;padding:14px 18px;
           box-shadow:0 1px 3px rgba(0,0,0,.1);flex:1;min-width:160px}}
    .kpi .value{{font-size:26px;font-weight:700;color:#1565c0}}
    .kpi .label{{font-size:11px;color:#666;margin-top:3px;text-transform:uppercase;letter-spacing:.5px}}
    .kpi .sub{{font-size:12px;color:#999;margin-top:2px}}
    .kpi.pending-kpi .value{{color:#e65100}}
    .charts-row{{display:grid;grid-template-columns:2fr 1fr;gap:16px;margin-bottom:20px}}
    .card{{background:#fff;border-radius:8px;padding:18px;
           box-shadow:0 1px 3px rgba(0,0,0,.1);margin-bottom:20px}}
    .card h2{{font-size:14px;color:#1565c0;margin-bottom:14px;
              border-bottom:1px solid #e0e0e0;padding-bottom:8px;
              display:flex;align-items:center;justify-content:space-between}}
    .card h2 .count{{font-size:12px;color:#999;font-weight:400}}
    table{{width:100%;border-collapse:collapse;font-size:13px}}
    th{{background:#1565c0;color:#fff;text-align:left;padding:7px 10px;
        font-weight:600;white-space:nowrap;position:sticky;top:0;z-index:1}}
    th.sort{{cursor:pointer;user-select:none}}
    th.sort:hover{{background:#0d47a1}}
    td{{padding:6px 10px;border-bottom:1px solid #eee}}
    tr:hover td{{background:#f5f9ff}}
    td.num{{text-align:right;font-variant-numeric:tabular-nums}}
    .table-wrap{{max-height:420px;overflow-y:auto;border-radius:6px;
                 border:1px solid #e0e0e0}}
    .badge{{padding:2px 8px;border-radius:4px;font-weight:bold;font-size:12px;display:inline-block}}
    .badge-ok{{background:#e8f5e9;color:#2e7d32}}
    .badge-warn{{background:#fff8e1;color:#f57f17}}
    .badge-err{{background:#ffebee;color:#c62828}}
    .badge-pending{{background:#fff3e0;color:#e65100}}
    .badge-nc{{background:#f5f5f5;color:#757575}}
    .no-data{{color:#999;font-style:italic;padding:20px;text-align:center}}
    .toggle-grp{{display:flex;gap:4px}}
    .toggle-btn{{padding:3px 10px;border:1px solid #1565c0;border-radius:4px;
                 cursor:pointer;font-size:12px;background:#fff;color:#1565c0;font-weight:600}}
    .toggle-btn.active{{background:#1565c0;color:#fff}}
    .toggle-btn:hover:not(.active){{background:#e3f2fd}}
    .inv-row{{cursor:pointer}}
    .inv-row:hover td{{background:#e3f2fd !important}}
    /* Modal */
    .modal-overlay{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);
                    z-index:1000;overflow-y:auto;padding:40px 16px}}
    .modal-overlay.open{{display:flex;align-items:flex-start;justify-content:center}}
    .modal{{background:#fff;border-radius:10px;width:100%;max-width:900px;
            box-shadow:0 8px 32px rgba(0,0,0,.25);overflow:hidden}}
    .modal-header{{background:#1565c0;color:#fff;padding:16px 20px;
                   display:flex;align-items:center;justify-content:space-between}}
    .modal-header h3{{font-size:16px;font-weight:600}}
    .modal-close{{background:none;border:none;color:#fff;font-size:22px;
                  cursor:pointer;line-height:1;padding:0 4px}}
    .modal-close:hover{{opacity:.7}}
    .modal-body{{padding:20px;overflow-y:auto;max-height:75vh}}
    .modal-section{{margin-bottom:20px}}
    .modal-section h4{{font-size:13px;font-weight:700;color:#1565c0;
                       border-bottom:1px solid #e0e0e0;padding-bottom:6px;margin-bottom:10px}}
    .info-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:8px 16px}}
    .info-item .label{{font-size:11px;color:#888;text-transform:uppercase;letter-spacing:.4px}}
    .info-item .value{{font-size:13px;font-weight:600;color:#333}}
    @media(max-width:800px){{
      .charts-row{{grid-template-columns:1fr}}
      .kpi{{min-width:140px}}
      .filter-bar{{top:49px}}
    }}
  </style>
</head>
<body>
<header>
  <h1>Freight Invoice Control — Dashboard</h1>
  <div class="meta">Isicom AB &nbsp;|&nbsp; Updated {_e(last_updated)}</div>
</header>

<div class="filter-bar">
  <label>År</label>
  <select id="fYear" onchange="applyFilters()">
    <option value="">Alla år</option>{year_opts}
  </select>
  <label>Månad</label>
  <select id="fMonth" onchange="applyFilters()">
    <option value="">Alla månader</option>{month_opts}
  </select>
  <label>Transportör</label>
  <select id="fCarrier" onchange="applyFilters()">
    <option value="">Alla</option>{carrier_opts}
  </select>
  <label>Tjänstetyp</label>
  <select id="fService" onchange="applyFilters()">
    <option value="">Alla typer</option>{svc_opts}
  </select>
  <input id="fInvoice" type="text" placeholder="Sök fakturanr..." oninput="applyFilters()">
  <div class="filter-sep"></div>
  <label>Visa</label>
  <div class="toggle-grp">
    <button class="toggle-btn active" id="btnAll" onclick="setRecFilter('all')">Alla</button>
    <button class="toggle-btn" id="btnRecOnly" onclick="setRecFilter('reconciled')">Fullt avstämda</button>
  </div>
  <button class="reset" onclick="resetFilters()">Återställ</button>
</div>

<main>
  <!-- Average cost per service type — pinned to top -->
  <div class="card">
    <h2>Genomsnittskostnad per tjänstetyp (inkl. tillägg) <span class="count" id="svcCostCount"></span></h2>
    <div class="table-wrap">
      <table id="svcCostTable">
        <thead id="svcCostHead"></thead>
        <tbody id="svcCostBody"></tbody>
      </table>
    </div>
  </div>

  <!-- KPI row -->
  <div class="kpi-row" id="kpiRow"></div>

  <!-- Charts -->
  <div class="charts-row">
    <div class="card">
      <h2>Månadskostnad per transportör ex-moms (SEK) <span class="count" id="chartNote"></span></h2>
      <canvas id="monthlyChart" height="200"></canvas>
    </div>
    <div class="card">
      <h2>Fördelning per transportör</h2>
      <canvas id="carrierChart" height="200"></canvas>
    </div>
  </div>

  <!-- Cost timeline -->
  <div class="card">
    <h2>Kostnadsutveckling ex-moms (SEK)
      <span style="display:flex;align-items:center;gap:12px">
        <span class="count" id="timelineNote"></span>
        <span class="toggle-grp">
          <button class="toggle-btn active" id="btnMonthly" onclick="setGranularity('monthly')">Månadsvis</button>
          <button class="toggle-btn" id="btnYearly" onclick="setGranularity('yearly')">Årsvis</button>
        </span>
      </span>
    </h2>
    <canvas id="timelineChart" height="100"></canvas>
  </div>

  <!-- Service breakdown -->
  <div class="card">
    <h2>Försändelser per tjänstetyp <span class="count" id="svcCount"></span></h2>
    <div class="table-wrap">
      <table id="svcTable">
        <thead><tr>
          <th>Tjänstetyp</th><th>Transportör</th>
          <th class="num">Försändelser</th>
          <th class="num">Total ex-moms (SEK)</th>
          <th class="num">Snitt / Försändelse (SEK)</th>
          <th class="num">% av total</th>
        </tr></thead>
        <tbody id="svcBody"></tbody>
      </table>
    </div>
  </div>

  <!-- Surcharge breakdown -->
  <div class="card">
    <h2>Tilläggsavgifter <span class="count" id="scCount"></span></h2>
    <canvas id="surchargeChart" height="90"></canvas>
  </div>

  <!-- Anomalies -->
  <div class="card">
    <h2>Avvikelser <span class="count" id="anomalyCount"></span></h2>
    <div class="table-wrap">
      <table id="anomalyTable">
        <thead><tr>
          <th>Transportör</th><th>Faktura#</th><th>Typ</th><th>Allvarlighet</th>
          <th>Beskrivning</th><th>AI-förklaring</th>
        </tr></thead>
        <tbody id="anomalyBody"></tbody>
      </table>
    </div>
  </div>

  <!-- Invoice history -->
  <div class="card">
    <h2>Fakturahistorik <span class="count" id="invCount"></span></h2>
    <div class="table-wrap">
      <table id="invTable">
        <thead><tr>
          <th>Fakturadatum</th><th>Transportör</th><th>Faktura #</th>
          <th class="num">Totalt ex-moms (SEK)</th><th>Valuta</th><th>Status</th>
        </tr></thead>
        <tbody id="invBody"></tbody>
      </table>
    </div>
  </div>

  <!-- Recent checks -->
  <div class="card">
    <h2>Kontroller</h2>
    <div class="table-wrap">
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
      <h3 id="modalTitle">Fakturadetaljer</h3>
      <button class="modal-close" onclick="closeModal()">&#x2715;</button>
    </div>
    <div class="modal-body" id="modalBody"></div>
  </div>
</div>

<script>
const INV_DATA      = {inv_json};
const SVC_DATA      = {svc_json};
const SC_DATA       = {sc_json};
const CHECK_DATA    = {check_json};
const SVC_COST_DATA = {svc_cost_json};
const ANOMALY_DATA  = {anomaly_json};
const ALL_SC_CATS   = {sc_cats_json};

const CARRIER_COLORS = {{Bring:'#1565c0', PostNord:'#e65100'}};
const CARRIER_COLORS_ALPHA = {{Bring:'rgba(21,101,192,0.7)', PostNord:'rgba(230,81,0,0.7)'}};

function carrierColor(c) {{ return CARRIER_COLORS[c] || '#546e7a'; }}
function carrierColorA(c) {{ return CARRIER_COLORS_ALPHA[c] || 'rgba(84,110,122,0.7)'; }}

function badge(status) {{
  const map = {{
    OK:         ['badge-ok',      '✓'],
    Warning:    ['badge-warn',    '⚠'],
    Error:      ['badge-err',     '✗'],
    Pending:    ['badge-pending', '⏳'],
  }};
  const [cls, icon] = map[status] || ['badge-nc', '?'];
  return `<span class="badge ${{cls}}">${{icon}} ${{esc(status)}}</span>`;
}}

function esc(s) {{
  return String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}}

function fmt(n) {{ return n.toLocaleString('sv-SE', {{minimumFractionDigits:2, maximumFractionDigits:2}}); }}
function fmtInt(n) {{ return Math.round(n).toLocaleString('sv-SE'); }}

let monthlyChart, carrierChart, surchargeChart, timelineChart;
let _tlGran = 'monthly';
let _recFilter = 'all';

// ── Reconciliation filter ─────────────────────────────────────────────────────
function setRecFilter(f) {{
  _recFilter = f;
  document.getElementById('btnAll').classList.toggle('active', f === 'all');
  document.getElementById('btnRecOnly').classList.toggle('active', f === 'reconciled');
  applyFilters();
}}

// ── Filters ──────────────────────────────────────────────────────────────────
function applyFilters() {{
  const year    = document.getElementById('fYear').value;
  const month   = document.getElementById('fMonth').value;
  const carrier = document.getElementById('fCarrier').value;
  const svc     = document.getElementById('fService').value;
  const invQ    = document.getElementById('fInvoice').value.toLowerCase().trim();

  const filtInv = INV_DATA.filter(i =>
    (!year    || i.year    === year)    &&
    (!month   || i.month   === month)   &&
    (!carrier || i.carrier === carrier) &&
    (!invQ    || i.invoice_no.toLowerCase().includes(invQ)) &&
    (_recFilter !== 'reconciled' || i.status !== 'Pending')
  );
  const filtInvNos = new Set(filtInv.map(i => i.invoice_no));

  const filtSvc = SVC_DATA.filter(s =>
    filtInvNos.has(s.invoice_no) && (!svc || s.service_cat === svc)
  );
  const filtSc = SC_DATA.filter(s => filtInvNos.has(s.invoice_no));
  const filtSvcCost = SVC_COST_DATA.filter(s =>
    filtInvNos.has(s.invoice_no) && (!svc || s.service_cat === svc)
  );
  const filtAnomalies = ANOMALY_DATA.filter(a => filtInvNos.has(a.invoice_no));

  renderKPIs(filtInv, filtSvc);
  renderInvoiceTable(filtInv);
  renderTimelineChart(filtInv);
  renderServiceTable(filtSvc);
  renderServiceCostTable(filtSvcCost);
  renderAnomalyTable(filtAnomalies);
  renderSurchargeChart(filtSc);
  renderMonthlyChart(filtInv);
  renderCarrierChart(filtInv);
}}

function resetFilters() {{
  ['fYear','fMonth','fCarrier','fService'].forEach(id => document.getElementById(id).value = '');
  document.getElementById('fInvoice').value = '';
  _recFilter = 'all';
  document.getElementById('btnAll').classList.add('active');
  document.getElementById('btnRecOnly').classList.remove('active');
  applyFilters();
}}

// ── KPIs ─────────────────────────────────────────────────────────────────────
function renderKPIs(invData, svcData) {{
  const reconInv  = invData.filter(i => i.status !== 'Pending');
  const pendInv   = invData.filter(i => i.status === 'Pending');
  const totalCost = reconInv.reduce((s, i) => s + i.total, 0);
  const pendCost  = pendInv.reduce((s, i) => s + i.total, 0);
  const totalShipments = svcData.reduce((s, v) => s + v.count, 0);
  const totalSurcharge = SC_DATA
    .filter(s => new Set(reconInv.map(i => i.invoice_no)).has(s.invoice_no))
    .reduce((s, v) => s + v.total, 0);
  const scPct = totalCost > 0 ? (totalSurcharge / totalCost * 100) : 0;

  const svcAgg = {{}};
  svcData.forEach(s => {{ svcAgg[s.service_cat] = (svcAgg[s.service_cat] || 0) + s.count; }});
  const topSvc = Object.entries(svcAgg).sort((a,b) => b[1]-a[1])[0];

  const pendHtml = pendInv.length > 0
    ? `<div class="kpi pending-kpi">
        <div class="value">${{pendInv.length}}</div>
        <div class="label">Oavstämda fakturor</div>
        <div class="sub">${{fmt(pendCost)}} SEK (preliminärt)</div>
       </div>`
    : '';

  document.getElementById('kpiRow').innerHTML = `
    <div class="kpi">
      <div class="value">${{reconInv.length}}</div>
      <div class="label">Avstämda fakturor</div>
    </div>
    ${{pendHtml}}
    <div class="kpi">
      <div class="value">${{fmtInt(totalCost)}}</div>
      <div class="label">Total kostnad ex-moms (SEK)</div>
      <div class="sub">Bekräftade fakturor</div>
    </div>
    <div class="kpi">
      <div class="value">${{totalShipments.toLocaleString('sv-SE')}}</div>
      <div class="label">Försändelser</div>
      ${{topSvc ? `<div class="sub">Vanligast: ${{esc(topSvc[0])}} (${{topSvc[1]}})</div>` : ''}}
    </div>
    <div class="kpi">
      <div class="value">${{fmtInt(totalSurcharge)}}</div>
      <div class="label">Totala tillägg ex-moms (SEK)</div>
      <div class="sub">${{scPct.toFixed(1)}}% av total</div>
    </div>
  `;
}}

// ── Invoice table ─────────────────────────────────────────────────────────────
function renderInvoiceTable(invData) {{
  document.getElementById('invCount').textContent = `(${{invData.length}} fakturor)`;
  const body = document.getElementById('invBody');
  if (!invData.length) {{
    body.innerHTML = '<tr><td colspan="6" class="no-data">Inga fakturor matchar filtret.</td></tr>';
    return;
  }}
  body.innerHTML = invData.map(i => {{
    const pendStyle = i.status === 'Pending' ? ' style="opacity:.85"' : '';
    return `<tr class="inv-row"${{pendStyle}} onclick="showInvoiceDetail('${{esc(i.invoice_no)}}')">
      <td>${{i.status === 'Pending' ? '<em style="color:#aaa">—</em>' : esc(i.date)}}</td>
      <td>${{esc(i.carrier)}}</td>
      <td><strong>${{esc(i.invoice_no)}}</strong></td>
      <td class="num">${{fmt(i.total)}}${{i.status==='Pending'?' <small style="color:#aaa">(spec)</small>':''}}</td>
      <td>${{esc(i.currency)}}</td>
      <td>${{badge(i.status)}}</td>
    </tr>`;
  }}).join('');
}}

// ── Service breakdown table ───────────────────────────────────────────────────
function renderServiceTable(svcData) {{
  const agg = {{}};
  svcData.forEach(s => {{
    const key = s.carrier + '||' + s.service_cat;
    if (!agg[key]) agg[key] = {{carrier: s.carrier, service_cat: s.service_cat, count: 0, total: 0}};
    agg[key].count += s.count;
    agg[key].total += s.total;
  }});
  const rows = Object.values(agg).sort((a,b) => b.total - a.total);
  const grandTotal = rows.reduce((s, r) => s + r.total, 0);

  document.getElementById('svcCount').textContent = `(${{rows.length}} typer)`;
  const body = document.getElementById('svcBody');
  if (!rows.length) {{
    body.innerHTML = '<tr><td colspan="6" class="no-data">Ingen data.</td></tr>';
    return;
  }}
  body.innerHTML = rows.map(r => {{
    const avg = r.count > 0 ? r.total / r.count : 0;
    const pct = grandTotal > 0 ? r.total / grandTotal * 100 : 0;
    return `<tr>
      <td><strong>${{esc(r.service_cat)}}</strong></td>
      <td>${{esc(r.carrier)}}</td>
      <td class="num">${{r.count.toLocaleString('sv-SE')}}</td>
      <td class="num">${{fmt(r.total)}}</td>
      <td class="num">${{fmt(avg)}}</td>
      <td class="num">${{pct.toFixed(1)}}%</td>
    </tr>`;
  }}).join('');

  const totalCount = rows.reduce((s,r) => s + r.count, 0);
  body.innerHTML += `<tr style="background:#f0f4ff;font-weight:bold">
    <td>TOTAL</td><td>—</td>
    <td class="num">${{totalCount.toLocaleString('sv-SE')}}</td>
    <td class="num">${{fmt(grandTotal)}}</td>
    <td class="num">—</td>
    <td class="num">100.0%</td>
  </tr>`;
}}

// ── Cost timeline — stacked bar ───────────────────────────────────────────────
function setGranularity(g) {{
  _tlGran = g;
  document.getElementById('btnMonthly').classList.toggle('active', g === 'monthly');
  document.getElementById('btnYearly').classList.toggle('active',  g === 'yearly');
  applyFilters();
}}

function renderTimelineChart(invData) {{
  const reconInv = invData.filter(i => i.status !== 'Pending');
  const getKey = i => _tlGran === 'yearly' ? i.year : i.month;
  const labels  = [...new Set(reconInv.map(getKey))].filter(Boolean).sort();
  const cars    = [...new Set(reconInv.map(i => i.carrier))].sort();

  const datasets = cars.map(c => ({{
    label: c,
    data: labels.map(lbl =>
      reconInv.filter(i => getKey(i) === lbl && i.carrier === c)
              .reduce((s, i) => s + i.total, 0)
    ),
    backgroundColor: carrierColorA(c),
    borderColor: carrierColor(c),
    borderWidth: 1,
  }}));

  const unit = _tlGran === 'yearly' ? 'år' : 'månad';
  document.getElementById('timelineNote').textContent =
    labels.length ? `(${{labels.length}} ${{unit}})` : '(ingen data)';

  if (timelineChart) {{
    timelineChart.data = {{labels, datasets}};
    timelineChart.update('active');
  }} else {{
    timelineChart = new Chart(document.getElementById('timelineChart'), {{
      type: 'bar',
      data: {{labels, datasets}},
      options: {{
        responsive: true,
        interaction: {{mode: 'index', intersect: false}},
        plugins: {{
          legend: {{position: 'bottom'}},
          tooltip: {{
            callbacks: {{
              label: ctx => ' ' + ctx.parsed.y.toLocaleString('sv-SE', {{minimumFractionDigits:2}}) + ' SEK',
            }},
          }},
        }},
        scales: {{
          x: {{stacked: true, grid: {{display: false}}}},
          y: {{
            stacked: true,
            beginAtZero: true,
            ticks: {{callback: v => v.toLocaleString('sv-SE')}},
            grid: {{color: '#f0f0f0'}},
          }},
        }},
      }},
    }});
  }}
}}

// ── Monthly chart (stacked bar) ───────────────────────────────────────────────
function renderMonthlyChart(invData) {{
  const reconInv = invData.filter(i => i.status !== 'Pending');
  const months  = [...new Set(reconInv.map(i => i.month))].sort();
  const cars    = [...new Set(reconInv.map(i => i.carrier))].sort();
  const datasets = cars.map(c => ({{
    label: c,
    data: months.map(m => reconInv.filter(i => i.month===m && i.carrier===c).reduce((s,i) => s+i.total, 0)),
    backgroundColor: carrierColorA(c),
    borderColor: carrierColor(c),
    borderWidth: 1,
  }}));

  document.getElementById('chartNote').textContent = months.length ? '' : '(ingen data)';

  if (monthlyChart) {{
    monthlyChart.data = {{labels: months, datasets}};
    monthlyChart.update('active');
  }} else {{
    monthlyChart = new Chart(document.getElementById('monthlyChart'), {{
      type: 'bar',
      data: {{labels: months, datasets}},
      options: {{
        responsive: true,
        plugins: {{legend: {{position:'bottom'}}}},
        scales: {{
          x: {{stacked: true, grid:{{display:false}}}},
          y: {{stacked: true, beginAtZero:true, ticks:{{callback: v => v.toLocaleString('sv-SE')}}}},
        }},
      }},
    }});
  }}
}}

// ── Carrier chart — grouped bar ───────────────────────────────────────────────
function renderCarrierChart(invData) {{
  const reconInv = invData.filter(i => i.status !== 'Pending');
  const agg = {{}};
  reconInv.forEach(i => {{ agg[i.carrier] = (agg[i.carrier]||0) + i.total; }});
  const labels = Object.keys(agg).sort();
  const data   = labels.map(l => Math.round(agg[l]*100)/100);
  const colors = labels.map(carrierColorA);
  const borders = labels.map(carrierColor);

  if (carrierChart) {{
    carrierChart.data = {{
      labels,
      datasets:[{{data, backgroundColor:colors, borderColor:borders, borderWidth:1}}]
    }};
    carrierChart.update('active');
  }} else {{
    carrierChart = new Chart(document.getElementById('carrierChart'), {{
      type: 'bar',
      data: {{labels, datasets:[{{data, backgroundColor:colors, borderColor:borders, borderWidth:1}}]}},
      options: {{
        responsive: true,
        plugins: {{
          legend: {{display: false}},
          tooltip: {{callbacks:{{label: ctx => ' '+ctx.parsed.y.toLocaleString('sv-SE',{{minimumFractionDigits:2}})+' SEK'}}}},
        }},
        scales: {{
          y: {{beginAtZero: true, ticks:{{callback: v => v.toLocaleString('sv-SE')}}}},
          x: {{grid:{{display:false}}}},
        }},
      }},
    }});
  }}
}}

// ── Surcharge chart ───────────────────────────────────────────────────────────
function renderSurchargeChart(scData) {{
  const agg = {{}};
  scData.forEach(s => {{ agg[s.surcharge_cat] = (agg[s.surcharge_cat]||0) + s.total; }});
  const sorted = Object.entries(agg).sort((a,b) => b[1]-a[1]);
  const labels = sorted.map(x => x[0]);
  const data   = sorted.map(x => Math.round(x[1]*100)/100);

  document.getElementById('scCount').textContent = labels.length ? `(${{labels.length}} kategorier)` : '';

  if (surchargeChart) {{
    surchargeChart.data = {{labels, datasets:[{{label:'SEK', data, backgroundColor:'#1565c0'}}]}};
    surchargeChart.update('active');
  }} else {{
    surchargeChart = new Chart(document.getElementById('surchargeChart'), {{
      type: 'bar',
      data: {{labels, datasets:[{{label:'SEK', data, backgroundColor:'#1565c0'}}]}},
      options: {{
        indexAxis: 'y',
        responsive: true,
        plugins: {{legend:{{display:false}}}},
        scales: {{x: {{beginAtZero:true, ticks:{{callback: v => v.toLocaleString('sv-SE')}}}}}},
      }},
    }});
  }}
}}

// ── Service cost breakdown table ──────────────────────────────────────────────
function renderServiceCostTable(filtData) {{
  const scCats = [...new Set(filtData.flatMap(s => Object.keys(s.sc_by_cat)))].sort();
  const agg = {{}};
  filtData.forEach(s => {{
    const key = s.carrier + '||' + s.service_cat;
    if (!agg[key]) {{
      agg[key] = {{carrier: s.carrier, cat: s.service_cat, count: 0, base: 0, sc: {{}}, grand: 0}};
      scCats.forEach(c => agg[key].sc[c] = 0);
    }}
    agg[key].count += s.count;
    agg[key].base  += s.base_total;
    scCats.forEach(c => {{ agg[key].sc[c] += s.sc_by_cat[c] || 0; }});
    agg[key].grand += s.grand_total;
  }});

  const rows = Object.values(agg).sort((a, b) => b.grand - a.grand);
  document.getElementById('svcCostCount').textContent = `(${{rows.length}} typ${{rows.length !== 1 ? 'er' : ''}})`;

  const scHeaders = scCats.map(c => `<th class="num" style="white-space:nowrap">${{esc(c)}}<br><small>Snitt</small></th>`).join('');
  document.getElementById('svcCostHead').innerHTML = `<tr>
    <th>Tjänstetyp</th><th>Transportör</th>
    <th class="num">Försändelser</th>
    <th class="num">Snitt bas ex-moms (SEK)</th>
    ${{scHeaders}}
    <th class="num">Snitt total ex-moms (SEK)</th>
  </tr>`;

  if (!rows.length) {{
    document.getElementById('svcCostBody').innerHTML =
      '<tr><td colspan="99" class="no-data">Ingen data.</td></tr>';
    return;
  }}

  document.getElementById('svcCostBody').innerHTML = rows.map(r => {{
    const n = r.count || 1;
    const scCells = scCats.map(c =>
      `<td class="num">${{fmt(r.sc[c] / n)}}</td>`
    ).join('');
    return `<tr>
      <td><strong>${{esc(r.cat)}}</strong></td>
      <td>${{esc(r.carrier)}}</td>
      <td class="num">${{r.count.toLocaleString('sv-SE')}}</td>
      <td class="num">${{fmt(r.base / n)}}</td>
      ${{scCells}}
      <td class="num"><strong>${{fmt(r.grand / n)}}</strong></td>
    </tr>`;
  }}).join('');
}}

// ── Anomaly table ─────────────────────────────────────────────────────────────
function renderAnomalyTable(data) {{
  document.getElementById('anomalyCount').textContent = data.length ? `(${{data.length}})` : '';
  const body = document.getElementById('anomalyBody');
  if (!data.length) {{
    body.innerHTML = '<tr><td colspan="6" class="no-data">Inga avvikelser för valda filter.</td></tr>';
    return;
  }}
  const sevStyle = {{
    Warning: 'background:#fff8e1;color:#f57f17',
    Error:   'background:#ffebee;color:#c62828',
    Info:    'background:#e3f2fd;color:#1565c0',
  }};
  const sevIcon = {{Warning: '⚠', Error: '✗', Info: 'ℹ'}};
  body.innerHTML = data.map(a => {{
    const sty = sevStyle[a.severity] || 'background:#f5f5f5;color:#333';
    const ic  = sevIcon[a.severity] || '?';
    const expHtml = a.explanation
      ? `<span style="color:#555">${{esc(a.explanation)}}</span>`
      : `<span style="color:#bbb;font-style:italic">—</span>`;
    return `<tr>
      <td>${{esc(a.carrier)}}</td>
      <td>${{esc(a.invoice_no)}}</td>
      <td>${{esc(a.type)}}</td>
      <td><span class="badge" style="${{sty}}">${{ic}} ${{esc(a.severity)}}</span></td>
      <td>${{esc(a.description)}}</td>
      <td>${{expHtml}}</td>
    </tr>`;
  }}).join('');
}}

// ── Checks table ──────────────────────────────────────────────────────────────
function renderChecks() {{
  const body = document.getElementById('checkBody');
  if (!CHECK_DATA.length) {{
    body.innerHTML = '<tr><td colspan="5" class="no-data">Inga kontroller ännu.</td></tr>';
    return;
  }}
  body.innerHTML = CHECK_DATA.map(c => `
    <tr>
      <td>${{esc(c.carrier)}}</td>
      <td>${{esc(c.invoice_no)}}</td>
      <td>${{esc(c.check)}}</td>
      <td>${{badge(c.status)}}</td>
      <td>${{esc(c.message)}}</td>
    </tr>
  `).join('');
}}

// ── Invoice detail modal ───────────────────────────────────────────────────────
function showInvoiceDetail(invNo) {{
  const inv = INV_DATA.find(i => i.invoice_no === invNo);
  if (!inv) return;

  document.getElementById('modalTitle').textContent =
    `${{inv.carrier}} — Faktura ${{invNo}}`;

  // Header info
  const isPending = inv.status === 'Pending';
  const infoFields = [
    ['Faktura #',    invNo],
    ['Transportör',  inv.carrier],
    ['Fakturadatum', inv.date || '—'],
    ['Förfallodatum', inv.due_date || '—'],
    ['Kundnummer',   inv.customer_number || '—'],
    ['Valuta',       inv.currency || 'SEK'],
    ['Total ex-moms', fmt(inv.total) + ' SEK' + (isPending ? ' (preliminärt från spec)' : '')],
    ['Moms',         inv.vat_amount ? fmt(inv.vat_amount) + ' SEK' : '—'],
    ['Total inkl moms', inv.total_inc_vat ? fmt(inv.total_inc_vat) + ' SEK' : '—'],
    ['Status',       badge(inv.status)],
    ['Källfil',      inv.source_file || '—'],
  ];
  if (isPending && inv.pending_note) {{
    infoFields.push(['Notering', inv.pending_note]);
  }}

  const infoHtml = `<div class="info-grid">${{
    infoFields.map(([lbl, val]) =>
      `<div class="info-item"><div class="label">${{esc(lbl)}}</div><div class="value">${{val}}</div></div>`
    ).join('')
  }}</div>`;

  // Service breakdown
  const svcRows = SVC_COST_DATA.filter(s => s.invoice_no === invNo)
    .sort((a,b) => b.grand_total - a.grand_total);
  const svcHtml = svcRows.length ? `
    <table>
      <thead><tr>
        <th>Tjänstetyp</th>
        <th class="num">Försändelser</th>
        <th class="num">Bas ex-moms (SEK)</th>
        <th class="num">Tillägg (SEK)</th>
        <th class="num">Total ex-moms (SEK)</th>
      </tr></thead>
      <tbody>${{svcRows.map(r => `<tr>
        <td><strong>${{esc(r.service_cat)}}</strong></td>
        <td class="num">${{r.count.toLocaleString('sv-SE')}}</td>
        <td class="num">${{fmt(r.base_total)}}</td>
        <td class="num">${{fmt(r.sc_grand)}}</td>
        <td class="num"><strong>${{fmt(r.grand_total)}}</strong></td>
      </tr>`).join('')}}</tbody>
    </table>` : '<p class="no-data">Ingen raddata tillgänglig.</p>';

  // Surcharge breakdown
  const scRows = SC_DATA.filter(s => s.invoice_no === invNo)
    .sort((a,b) => b.total - a.total);
  const scHtml = scRows.length ? `
    <table>
      <thead><tr>
        <th>Tilläggstyp</th>
        <th class="num">Belopp (SEK)</th>
      </tr></thead>
      <tbody>${{scRows.map(r => `<tr>
        <td>${{esc(r.surcharge_cat)}}</td>
        <td class="num">${{fmt(r.total)}}</td>
      </tr>`).join('')}}</tbody>
    </table>` : '<p class="no-data">Inga tilläggsavgifter.</p>';

  // Anomalies
  const anom = ANOMALY_DATA.filter(a => a.invoice_no === invNo);
  const anomHtml = anom.length ? `
    <table>
      <thead><tr>
        <th>Typ</th><th>Allvarlighet</th><th>Beskrivning</th>
      </tr></thead>
      <tbody>${{anom.map(a => `<tr>
        <td>${{esc(a.type)}}</td>
        <td>${{badge(a.severity)}}</td>
        <td>${{esc(a.description)}}</td>
      </tr>`).join('')}}</tbody>
    </table>` : '<p class="no-data">Inga avvikelser.</p>';

  // Validation checks
  const chks = CHECK_DATA.filter(c => c.invoice_no === invNo);
  const chkHtml = chks.length ? `
    <table>
      <thead><tr>
        <th>Kontroll</th><th>Status</th><th>Meddelande</th>
      </tr></thead>
      <tbody>${{chks.map(c => `<tr>
        <td>${{esc(c.check)}}</td>
        <td>${{badge(c.status)}}</td>
        <td>${{esc(c.message)}}</td>
      </tr>`).join('')}}</tbody>
    </table>` : '<p class="no-data">Inga kontrollresultat tillgängliga.</p>';

  document.getElementById('modalBody').innerHTML = `
    <div class="modal-section"><h4>Fakturainformation</h4>${{infoHtml}}</div>
    <div class="modal-section"><h4>Tjänstesammansättning</h4>${{svcHtml}}</div>
    <div class="modal-section"><h4>Tilläggsavgifter</h4>${{scHtml}}</div>
    <div class="modal-section"><h4>Avvikelser</h4>${{anomHtml}}</div>
    <div class="modal-section"><h4>Kontroller</h4>${{chkHtml}}</div>
  `;

  document.getElementById('modalOverlay').classList.add('open');
  document.body.style.overflow = 'hidden';
}}

function closeModal() {{
  document.getElementById('modalOverlay').classList.remove('open');
  document.body.style.overflow = '';
}}

function closeModalOnBg(e) {{
  if (e.target === document.getElementById('modalOverlay')) closeModal();
}}

document.addEventListener('keydown', e => {{ if (e.key === 'Escape') closeModal(); }});

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
