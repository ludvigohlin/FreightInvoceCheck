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

    # Lookup: invoice_number → {invoice_date, carrier}
    inv_lookup = {inv["invoice_number"]: inv for inv in invoices}

    # ── Build JS data: one object per invoice ─────────────────────────────────
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
    # Aggregate base freight per (invoice_no, service_cat)
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

    # Aggregate surcharges per (invoice_no, related_service_cat, surcharge_cat)
    sc_svc_agg: dict = defaultdict(float)
    for sc in surcharges:
        inv_no = sc.get("invoice_number", "")
        svc_cat = sc.get("related_service_category", "") or "Unknown"
        sc_cat = sc.get("surcharge_category", "Unknown") or "Unknown"
        sc_svc_agg[(inv_no, svc_cat, sc_cat)] += _safe_float(sc.get("amount"))

    # Collect all surcharge categories that appear (for dynamic columns)
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
    main{{max-width:1300px;margin:0 auto;padding:20px 16px}}
    .kpi-row{{display:flex;gap:14px;margin-bottom:20px;flex-wrap:wrap}}
    .kpi{{background:#fff;border-radius:8px;padding:14px 18px;
           box-shadow:0 1px 3px rgba(0,0,0,.1);flex:1;min-width:160px}}
    .kpi .value{{font-size:26px;font-weight:700;color:#1565c0}}
    .kpi .label{{font-size:11px;color:#666;margin-top:3px;text-transform:uppercase;letter-spacing:.5px}}
    .kpi .sub{{font-size:12px;color:#999;margin-top:2px}}
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
    .badge-nc{{background:#f5f5f5;color:#757575}}
    .no-data{{color:#999;font-style:italic;padding:20px;text-align:center}}
    .toggle-grp{{display:flex;gap:4px}}
    .toggle-btn{{padding:3px 10px;border:1px solid #1565c0;border-radius:4px;
                 cursor:pointer;font-size:12px;background:#fff;color:#1565c0;font-weight:600}}
    .toggle-btn.active{{background:#1565c0;color:#fff}}
    .toggle-btn:hover:not(.active){{background:#e3f2fd}}
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
  <label>Year</label>
  <select id="fYear" onchange="applyFilters()">
    <option value="">All years</option>{year_opts}
  </select>
  <label>Month</label>
  <select id="fMonth" onchange="applyFilters()">
    <option value="">All months</option>{month_opts}
  </select>
  <label>Carrier</label>
  <select id="fCarrier" onchange="applyFilters()">
    <option value="">All carriers</option>{carrier_opts}
  </select>
  <label>Service Type</label>
  <select id="fService" onchange="applyFilters()">
    <option value="">All types</option>{svc_opts}
  </select>
  <input id="fInvoice" type="text" placeholder="Search invoice#..." oninput="applyFilters()">
  <button class="reset" onclick="resetFilters()">Reset</button>
</div>

<main>
  <!-- Average cost per service type — pinned to top -->
  <div class="card">
    <h2>Average Cost per Service Type (incl. surcharges) <span class="count" id="svcCostCount"></span></h2>
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
      <h2>Monthly Cost by Carrier (SEK) <span class="count" id="chartNote"></span></h2>
      <canvas id="monthlyChart" height="200"></canvas>
    </div>
    <div class="card">
      <h2>Carrier Split</h2>
      <canvas id="carrierChart" height="200"></canvas>
    </div>
  </div>

  <!-- Cost timeline -->
  <div class="card">
    <h2>Cost Timeline (SEK)
      <span style="display:flex;align-items:center;gap:12px">
        <span class="count" id="timelineNote"></span>
        <span class="toggle-grp">
          <button class="toggle-btn active" id="btnMonthly" onclick="setGranularity('monthly')">Monthly</button>
          <button class="toggle-btn" id="btnYearly" onclick="setGranularity('yearly')">Yearly</button>
        </span>
      </span>
    </h2>
    <canvas id="timelineChart" height="100"></canvas>
  </div>

  <!-- Service breakdown -->
  <div class="card">
    <h2>Shipments by Service Type <span class="count" id="svcCount"></span></h2>
    <div class="table-wrap">
      <table id="svcTable">
        <thead><tr>
          <th>Service Type</th><th>Carrier</th>
          <th class="num">Shipments</th>
          <th class="num">Total Cost (SEK)</th>
          <th class="num">Avg / Shipment</th>
          <th class="num">% of Total</th>
        </tr></thead>
        <tbody id="svcBody"></tbody>
      </table>
    </div>
  </div>

  <!-- Surcharge breakdown -->
  <div class="card">
    <h2>Surcharge Breakdown <span class="count" id="scCount"></span></h2>
    <canvas id="surchargeChart" height="90"></canvas>
  </div>

  <!-- Anomalies -->
  <div class="card">
    <h2>Anomalies <span class="count" id="anomalyCount"></span></h2>
    <div class="table-wrap">
      <table id="anomalyTable">
        <thead><tr>
          <th>Carrier</th><th>Invoice#</th><th>Type</th><th>Severity</th>
          <th>Description</th><th>AI Explanation</th>
        </tr></thead>
        <tbody id="anomalyBody"></tbody>
      </table>
    </div>
  </div>

  <!-- Invoice history -->
  <div class="card">
    <h2>Invoice History <span class="count" id="invCount"></span></h2>
    <div class="table-wrap">
      <table id="invTable">
        <thead><tr>
          <th>Date</th><th>Carrier</th><th>Invoice#</th>
          <th class="num">Total ex VAT</th><th>Currency</th><th>Status</th>
        </tr></thead>
        <tbody id="invBody"></tbody>
      </table>
    </div>
  </div>

  <!-- Recent checks -->
  <div class="card">
    <h2>Recent Validation Checks</h2>
    <div class="table-wrap">
      <table id="checkTable">
        <thead><tr>
          <th>Carrier</th><th>Invoice#</th><th>Check</th>
          <th>Status</th><th>Message</th>
        </tr></thead>
        <tbody id="checkBody"></tbody>
      </table>
    </div>
  </div>
</main>

<script>
const INV_DATA      = {inv_json};
const SVC_DATA      = {svc_json};
const SC_DATA       = {sc_json};
const CHECK_DATA    = {check_json};
const SVC_COST_DATA = {svc_cost_json};
const ANOMALY_DATA  = {anomaly_json};
const ALL_SC_CATS   = {sc_cats_json};

const CARRIER_COLORS = {{Bring:'#1565c0', PostNord:'#e65100'}};

function carrierColor(c) {{ return CARRIER_COLORS[c] || '#546e7a'; }}

function badge(status) {{
  const cls = {{OK:'badge-ok', Warning:'badge-warn', Error:'badge-err'}}[status] || 'badge-nc';
  const icon = {{OK:'✓', Warning:'⚠', Error:'✗'}}[status] || '?';
  return `<span class="badge ${{cls}}">${{icon}} ${{esc(status)}}</span>`;
}}

function esc(s) {{
  return String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}}

function fmt(n) {{ return n.toLocaleString('sv-SE', {{minimumFractionDigits:2, maximumFractionDigits:2}}); }}
function fmtInt(n) {{ return Math.round(n).toLocaleString('sv-SE'); }}

let monthlyChart, carrierChart, surchargeChart, timelineChart;
let _tlGran = 'monthly';

// ── Filters ──────────────────────────────────────────────────────────────────
function applyFilters() {{
  const year    = document.getElementById('fYear').value;
  const month   = document.getElementById('fMonth').value;
  const carrier = document.getElementById('fCarrier').value;
  const svc     = document.getElementById('fService').value;
  const invQ    = document.getElementById('fInvoice').value.toLowerCase().trim();

  // Filter invoices
  const filtInv = INV_DATA.filter(i =>
    (!year    || i.year    === year)    &&
    (!month   || i.month   === month)   &&
    (!carrier || i.carrier === carrier) &&
    (!invQ    || i.invoice_no.toLowerCase().includes(invQ))
  );
  const filtInvNos = new Set(filtInv.map(i => i.invoice_no));

  // Filter service data
  const filtSvc = SVC_DATA.filter(s =>
    filtInvNos.has(s.invoice_no) && (!svc || s.service_cat === svc)
  );

  // Filter surcharge data
  const filtSc = SC_DATA.filter(s => filtInvNos.has(s.invoice_no));

  // Filter svc_cost and anomaly data by the same filtered invoice set
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
  applyFilters();
}}

// ── KPIs ─────────────────────────────────────────────────────────────────────
function renderKPIs(invData, svcData) {{
  const totalCost     = invData.reduce((s, i) => s + i.total, 0);
  const totalShipments = svcData.reduce((s, v) => s + v.count, 0);
  const totalSurcharge = SC_DATA
    .filter(s => new Set(invData.map(i => i.invoice_no)).has(s.invoice_no))
    .reduce((s, v) => s + v.total, 0);
  const scPct = totalCost > 0 ? (totalSurcharge / totalCost * 100) : 0;

  // Top service type by shipment count
  const svcAgg = {{}};
  svcData.forEach(s => {{
    svcAgg[s.service_cat] = (svcAgg[s.service_cat] || 0) + s.count;
  }});
  const topSvc = Object.entries(svcAgg).sort((a,b) => b[1]-a[1])[0];

  document.getElementById('kpiRow').innerHTML = `
    <div class="kpi">
      <div class="value">${{invData.length}}</div>
      <div class="label">Invoices</div>
    </div>
    <div class="kpi">
      <div class="value">${{fmtInt(totalCost)}}</div>
      <div class="label">Total Cost (SEK)</div>
    </div>
    <div class="kpi">
      <div class="value">${{totalShipments.toLocaleString('sv-SE')}}</div>
      <div class="label">Shipments</div>
      ${{topSvc ? `<div class="sub">Top: ${{esc(topSvc[0])}} (${{topSvc[1]}})</div>` : ''}}
    </div>
    <div class="kpi">
      <div class="value">${{fmtInt(totalSurcharge)}}</div>
      <div class="label">Total Surcharges (SEK)</div>
      <div class="sub">${{scPct.toFixed(1)}}% of total</div>
    </div>
  `;
}}

// ── Invoice table ─────────────────────────────────────────────────────────────
function renderInvoiceTable(invData) {{
  document.getElementById('invCount').textContent = `(${{invData.length}} invoices)`;
  const body = document.getElementById('invBody');
  if (!invData.length) {{
    body.innerHTML = '<tr><td colspan="6" class="no-data">No invoices match the current filters.</td></tr>';
    return;
  }}
  body.innerHTML = invData.map(i => `
    <tr>
      <td>${{esc(i.date)}}</td>
      <td>${{esc(i.carrier)}}</td>
      <td><strong>${{esc(i.invoice_no)}}</strong></td>
      <td class="num">${{fmt(i.total)}}</td>
      <td>${{esc(i.currency)}}</td>
      <td>${{badge(i.status)}}</td>
    </tr>
  `).join('');
}}

// ── Service breakdown table ───────────────────────────────────────────────────
function renderServiceTable(svcData) {{
  // Aggregate by (carrier, service_cat)
  const agg = {{}};
  svcData.forEach(s => {{
    const key = s.carrier + '||' + s.service_cat;
    if (!agg[key]) agg[key] = {{carrier: s.carrier, service_cat: s.service_cat, count: 0, total: 0}};
    agg[key].count += s.count;
    agg[key].total += s.total;
  }});
  const rows = Object.values(agg).sort((a,b) => b.total - a.total);
  const grandTotal = rows.reduce((s, r) => s + r.total, 0);

  document.getElementById('svcCount').textContent = `(${{rows.length}} types)`;
  const body = document.getElementById('svcBody');
  if (!rows.length) {{
    body.innerHTML = '<tr><td colspan="6" class="no-data">No data.</td></tr>';
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

  // Total row
  const totalCount = rows.reduce((s,r) => s + r.count, 0);
  body.innerHTML += `<tr style="background:#f0f4ff;font-weight:bold">
    <td>TOTAL</td><td>—</td>
    <td class="num">${{totalCount.toLocaleString('sv-SE')}}</td>
    <td class="num">${{fmt(grandTotal)}}</td>
    <td class="num">—</td>
    <td class="num">100.0%</td>
  </tr>`;
}}

// ── Cost timeline ─────────────────────────────────────────────────────────────
function setGranularity(g) {{
  _tlGran = g;
  document.getElementById('btnMonthly').classList.toggle('active', g === 'monthly');
  document.getElementById('btnYearly').classList.toggle('active',  g === 'yearly');
  applyFilters();
}}

function renderTimelineChart(invData) {{
  const getKey = i => _tlGran === 'yearly' ? i.year : i.month;
  const labels  = [...new Set(invData.map(getKey))].filter(Boolean).sort();
  const cars    = [...new Set(invData.map(i => i.carrier))].sort();
  const pt      = labels.length <= 24 ? 4 : 2;

  const datasets = cars.map(c => ({{
    label: c,
    data: labels.map(lbl =>
      invData.filter(i => getKey(i) === lbl && i.carrier === c)
             .reduce((s, i) => s + i.total, 0)
    ),
    borderColor: carrierColor(c),
    backgroundColor: carrierColor(c) + '33',
    tension: 0.35,
    fill: true,
    pointRadius: pt,
    pointHoverRadius: pt + 3,
  }}));

  if (cars.length > 1) {{
    datasets.push({{
      label: 'Total',
      data: labels.map(lbl =>
        invData.filter(i => getKey(i) === lbl).reduce((s, i) => s + i.total, 0)
      ),
      borderColor: '#78909c',
      backgroundColor: 'transparent',
      borderDash: [6, 4],
      tension: 0.35,
      fill: false,
      pointRadius: pt - 1,
      pointHoverRadius: pt + 2,
    }});
  }}

  const unit = _tlGran === 'yearly' ? 'year' : 'month';
  document.getElementById('timelineNote').textContent =
    labels.length ? `(${{labels.length}} ${{unit}}${{labels.length !== 1 ? 's' : ''}})` : '(no data)';

  if (timelineChart) {{
    timelineChart.data = {{labels, datasets}};
    timelineChart.update('active');
  }} else {{
    timelineChart = new Chart(document.getElementById('timelineChart'), {{
      type: 'line',
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
          y: {{
            beginAtZero: true,
            ticks: {{callback: v => v.toLocaleString('sv-SE')}},
            grid: {{color: '#f0f0f0'}},
          }},
          x: {{grid: {{display: false}}}},
        }},
      }},
    }});
  }}
}}

// ── Monthly chart ─────────────────────────────────────────────────────────────
function renderMonthlyChart(invData) {{
  const months  = [...new Set(invData.map(i => i.month))].sort();
  const cars    = [...new Set(invData.map(i => i.carrier))].sort();
  const datasets = cars.map(c => ({{
    label: c,
    data: months.map(m => invData.filter(i => i.month===m && i.carrier===c).reduce((s,i) => s+i.total, 0)),
    backgroundColor: carrierColor(c),
  }}));

  document.getElementById('chartNote').textContent = months.length ? '' : '(no data)';

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
          y: {{beginAtZero:true, ticks:{{callback: v => v.toLocaleString('sv-SE')}}}},
        }},
      }},
    }});
  }}
}}

// ── Carrier chart ─────────────────────────────────────────────────────────────
function renderCarrierChart(invData) {{
  const agg = {{}};
  invData.forEach(i => {{ agg[i.carrier] = (agg[i.carrier]||0) + i.total; }});
  const labels = Object.keys(agg);
  const data   = labels.map(l => Math.round(agg[l]*100)/100);
  const colors = labels.map(carrierColor);

  if (carrierChart) {{
    carrierChart.data = {{labels, datasets:[{{data, backgroundColor:colors, borderWidth:2}}]}};
    carrierChart.update('active');
  }} else {{
    carrierChart = new Chart(document.getElementById('carrierChart'), {{
      type: 'doughnut',
      data: {{labels, datasets:[{{data, backgroundColor:colors, borderWidth:2}}]}},
      options: {{
        responsive:true,
        plugins: {{
          legend:{{position:'bottom'}},
          tooltip:{{callbacks:{{label: ctx => ' '+ctx.parsed.toLocaleString('sv-SE',{{minimumFractionDigits:2}})+' SEK'}}}},
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

  document.getElementById('scCount').textContent = labels.length ? `(${{labels.length}} categories)` : '';

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
  // Collect surcharge cats present in filtered data
  const scCats = [...new Set(filtData.flatMap(s => Object.keys(s.sc_by_cat)))].sort();

  // Aggregate by (carrier, service_cat)
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
  document.getElementById('svcCostCount').textContent = `(${{rows.length}} type${{rows.length !== 1 ? 's' : ''}})`;

  const scHeaders = scCats.map(c => `<th class="num" style="white-space:nowrap">${{esc(c)}}<br><small>Avg</small></th>`).join('');
  document.getElementById('svcCostHead').innerHTML = `<tr>
    <th>Service Type</th><th>Carrier</th>
    <th class="num">Shipments</th>
    <th class="num">Avg Base (SEK)</th>
    ${{scHeaders}}
    <th class="num">Avg Total (SEK)</th>
  </tr>`;

  if (!rows.length) {{
    document.getElementById('svcCostBody').innerHTML =
      '<tr><td colspan="99" class="no-data">No data.</td></tr>';
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
    body.innerHTML = '<tr><td colspan="6" class="no-data">No anomalies for the selected filters.</td></tr>';
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
    body.innerHTML = '<tr><td colspan="5" class="no-data">No checks yet.</td></tr>';
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
