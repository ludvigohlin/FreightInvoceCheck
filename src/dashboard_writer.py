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
            "invoice_no": inv_no, "service_cat": svc_cat, "carrier": v["carrier"],
            "year": v["year"], "month": v["month"], "count": v["count"],
            "base_total": round(v["total"], 2),
            "sc_by_cat": {k: val for k, val in sc_by_cat.items() if val > 0},
            "sc_grand": round(sc_grand, 2),
            "grand_total": round(v["total"] + sc_grand, 2),
        })

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
    sc_json        = json.dumps(sc_js)
    check_json     = json.dumps(check_js)
    svc_cost_json  = json.dumps(svc_cost_js)
    anomaly_json   = json.dumps(anomaly_js)
    sc_cats_json   = json.dumps(all_sc_cats)

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
    .table-wrap{{overflow-y:auto;border-radius:6px;border:1px solid #e8e8e8}}
    .table-wrap.fixed-h{{max-height:380px}}

    /* Badges */
    .badge{{padding:2px 8px;border-radius:4px;font-weight:700;font-size:11px;
            display:inline-block;letter-spacing:.2px}}
    .badge-ok{{background:#e8f5e9;color:#2e7d32}}
    .badge-warn{{background:#fff8e1;color:#f57f17}}
    .badge-err{{background:#ffebee;color:#c62828}}
    .badge-pending{{background:#fff3e0;color:#e65100}}
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
    .show-more-btn{{display:block;width:100%;padding:8px;border:none;
                    background:#f5f9ff;color:#1565c0;font-size:12px;font-weight:700;
                    cursor:pointer;border-top:1px solid #e0e0e0;border-radius:0 0 6px 6px;
                    text-align:center;letter-spacing:.3px}}
    .show-more-btn:hover{{background:#e3f2fd}}

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
      <h2>Antal försändelser per månad
        <span class="count" id="volumeNote"></span>
      </h2>
      <canvas id="volumeChart" height="195"></canvas>
    </div>
    <div class="card" style="display:flex;flex-direction:column">
      <h2>Senaste fakturor <span class="count" id="recentInvCount"></span></h2>
      <div style="flex:1;overflow:hidden;border-radius:6px;border:1px solid #e8e8e8">
        <table id="recentInvTable" style="font-size:12px">
          <thead><tr>
            <th>Datum</th><th>Transp.</th><th>Faktura #</th>
            <th class="num">Belopp</th><th>Status</th>
          </tr></thead>
          <tbody id="recentInvBody"></tbody>
        </table>
        <button class="show-more-btn" id="showMoreBtn" onclick="toggleAllInvoices()"></button>
      </div>
    </div>
  </div>

  <!-- Cost timeline -->
  <div class="card">
    <h2>Kostnadsutveckling ex-moms (SEK)
      <span style="display:flex;align-items:center;gap:10px">
        <span class="count" id="timelineNote"></span>
        <span class="toggle-grp">
          <button class="toggle-btn active" id="btnMonthly" onclick="setGranularity('monthly')">Månadsvis</button>
          <button class="toggle-btn" id="btnYearly" onclick="setGranularity('yearly')">Årsvis</button>
        </span>
      </span>
    </h2>
    <canvas id="timelineChart" height="90"></canvas>
  </div>

  <!-- Avg cost per service -->
  <div class="card">
    <h2>Genomsnittskostnad per tjänstetyp (inkl. tillägg) <span class="count" id="svcCostCount"></span></h2>
    <div class="table-wrap fixed-h">
      <table id="svcCostTable">
        <thead id="svcCostHead"></thead>
        <tbody id="svcCostBody"></tbody>
      </table>
    </div>
  </div>

  <!-- Service breakdown + Surcharges -->
  <div class="grid-3-2">
    <div class="card">
      <h2>Försändelser per tjänstetyp <span class="count" id="svcCount"></span></h2>
      <div class="table-wrap fixed-h">
        <table id="svcTable">
          <thead><tr>
            <th>Tjänstetyp</th><th>Transportör</th>
            <th class="num">Antal</th>
            <th class="num">Total (SEK)</th>
            <th class="num">Snitt (SEK)</th>
            <th class="num">Andel</th>
          </tr></thead>
          <tbody id="svcBody"></tbody>
        </table>
      </div>
    </div>
    <div class="card">
      <h2>Tilläggsavgifter <span class="count" id="scCount"></span></h2>
      <canvas id="surchargeChart" height="220"></canvas>
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

const CC  = {{Bring:'#1565c0', PostNord:'#e65100'}};
const CCA = {{Bring:'rgba(21,101,192,.72)', PostNord:'rgba(230,81,0,.72)'}};
function cCol(c)  {{ return CC[c]  || '#546e7a'; }}
function cColA(c) {{ return CCA[c] || 'rgba(84,110,122,.72)'; }}

function badge(status) {{
  const map = {{OK:['badge-ok','✓'], Warning:['badge-warn','⚠'],
                Error:['badge-err','✗'], Pending:['badge-pending','⏳']}};
  const [cls,icon] = map[status] || ['badge-nc','—'];
  return `<span class="badge ${{cls}}">${{icon}} ${{esc(status)}}</span>`;
}}
function esc(s) {{
  return String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}}
function fmt(n)    {{ return (+n).toLocaleString('sv-SE',{{minimumFractionDigits:2,maximumFractionDigits:2}}); }}
function fmtInt(n) {{ return Math.round(n).toLocaleString('sv-SE'); }}

let volumeChart, surchargeChart, timelineChart;
let _tlGran   = 'monthly';
let _recFilter = 'all';
let _showAllInv = false;
const RECENT_LIMIT = 5;

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
    (_recFilter !== 'reconciled' || i.status !== 'Pending')
  );
  const filtNos = new Set(filtInv.map(i => i.invoice_no));

  const filtSvc     = SVC_DATA.filter(s => filtNos.has(s.invoice_no) && (!svc||s.service_cat===svc));
  const filtSc      = SC_DATA.filter(s => filtNos.has(s.invoice_no));
  const filtSvcCost = SVC_COST_DATA.filter(s => filtNos.has(s.invoice_no) && (!svc||s.service_cat===svc));
  const filtAnom    = ANOMALY_DATA.filter(a => filtNos.has(a.invoice_no));

  renderKPIs(filtInv, filtSvc, filtAnom);
  renderRecentInvoices(filtInv);
  renderVolumeChart(filtSvc);
  renderTimelineChart(filtInv);
  renderServiceCostTable(filtSvcCost);
  renderServiceTable(filtSvc);
  renderSurchargeChart(filtSc);
  renderAnomalyTable(filtAnom);
}}

function resetFilters() {{
  ['fYear','fMonth','fCarrier','fService'].forEach(id => document.getElementById(id).value='');
  document.getElementById('fInvoice').value='';
  _recFilter='all'; _showAllInv=false;
  document.getElementById('btnAll').classList.add('active');
  document.getElementById('btnRecOnly').classList.remove('active');
  applyFilters();
}}

// ── KPIs ──────────────────────────────────────────────────────────────────────
function renderKPIs(invData, svcData, anomData) {{
  const reconInv  = invData.filter(i => i.status !== 'Pending');
  const pendInv   = invData.filter(i => i.status === 'Pending');
  const totalCost = reconInv.reduce((s,i) => s+i.total, 0);
  const pendCost  = pendInv.reduce((s,i)  => s+i.total, 0);
  const totalShip = svcData.reduce((s,v)  => s+v.count, 0);
  const avgPerShip = totalShip > 0 ? totalCost/totalShip : 0;

  const scNos = new Set(reconInv.map(i => i.invoice_no));
  const totalSc = SC_DATA.filter(s => scNos.has(s.invoice_no)).reduce((s,v) => s+v.total, 0);
  const scPct   = totalCost > 0 ? totalSc/totalCost*100 : 0;

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
      <div class="sub">${{fmt(pendCost)}} SEK (preliminärt)</div>
    </div>` : '';

  document.getElementById('kpiRow').innerHTML = `
    <div class="kpi blue">
      <div class="value">${{reconInv.length}}</div>
      <div class="label">Fakturor</div>
      <div class="sub">Avstämda</div>
    </div>
    ${{pendCard}}
    <div class="kpi blue">
      <div class="value">${{fmtInt(totalCost)}}</div>
      <div class="label">Total kostnad ex-moms</div>
      <div class="sub">SEK · bekräftade</div>
    </div>
    <div class="kpi grey">
      <div class="value">${{fmt(avgPerShip)}}</div>
      <div class="label">Snitt per försändelse</div>
      <div class="sub">SEK ex-moms</div>
    </div>
    <div class="kpi ${{anomCls}}">
      <div class="value">${{anomData.length}}</div>
      <div class="label">Avvikelser</div>
      <div class="sub">${{anomSub}}</div>
    </div>
    <div class="kpi grey">
      <div class="value">${{scPct.toFixed(1)}}%</div>
      <div class="label">Tillägg av total</div>
      <div class="sub">${{fmtInt(totalSc)}} SEK</div>
    </div>
  `;
}}

// ── Recent invoices (expandable) ──────────────────────────────────────────────
function toggleAllInvoices() {{
  _showAllInv = !_showAllInv;
  // Re-render with current filtered data — grab from last applyFilters call
  // Simplest: trigger full re-apply
  applyFilters();
}}

function renderRecentInvoices(invData) {{
  const total = invData.length;
  const shown = _showAllInv ? invData : invData.slice(0, RECENT_LIMIT);

  document.getElementById('recentInvCount').textContent =
    `(${{total}} faktura${{total!==1?'r':''}})`;

  const body = document.getElementById('recentInvBody');
  if (!total) {{
    body.innerHTML = '<tr><td colspan="5" class="no-data">Inga fakturor.</td></tr>';
    document.getElementById('showMoreBtn').style.display = 'none';
    return;
  }}

  body.innerHTML = shown.map(i => {{
    const pending = i.status === 'Pending';
    return `<tr class="inv-row${{pending?' pending-row':''}}" onclick="showInvoiceDetail('${{esc(i.invoice_no)}}')">
      <td style="white-space:nowrap">${{pending?'<em style="color:#ccc">—</em>':esc(i.date)}}</td>
      <td style="white-space:nowrap">${{esc(i.carrier)}}</td>
      <td><strong style="font-size:11px">${{esc(i.invoice_no)}}</strong></td>
      <td class="num" style="white-space:nowrap">${{fmtInt(i.total)}}${{pending?' <small style="color:#ccc">*</small>':''}}</td>
      <td>${{badge(i.status)}}</td>
    </tr>`;
  }}).join('');

  const btn = document.getElementById('showMoreBtn');
  if (total <= RECENT_LIMIT) {{
    btn.style.display = 'none';
  }} else {{
    btn.style.display = 'block';
    const hidden = total - RECENT_LIMIT;
    btn.textContent = _showAllInv
      ? '▲ Visa färre'
      : `▼ Visa ${{hidden}} äldre faktura${{hidden!==1?'r':''}}`;
  }}
}}

// ── Volume chart — shipments per month per carrier ────────────────────────────
function renderVolumeChart(svcData) {{
  // Aggregate shipment counts by (month, carrier) from SVC_DATA
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
  document.getElementById('btnMonthly').classList.toggle('active', g==='monthly');
  document.getElementById('btnYearly').classList.toggle('active',  g==='yearly');
  applyFilters();
}}

function renderTimelineChart(invData) {{
  const recon  = invData.filter(i => i.status !== 'Pending');
  const getKey = i => _tlGran==='yearly' ? i.year : i.month;
  const labels = [...new Set(recon.map(getKey))].filter(Boolean).sort();
  const cars   = [...new Set(recon.map(i => i.carrier))].sort();
  const datasets = cars.map(c => ({{
    label:c,
    data: labels.map(lbl=>recon.filter(i=>getKey(i)===lbl&&i.carrier===c).reduce((s,i)=>s+i.total,0)),
    backgroundColor:cColA(c), borderColor:cCol(c), borderWidth:1,
  }}));

  const unit = _tlGran==='yearly' ? 'år' : 'månad';
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

// ── Service breakdown ─────────────────────────────────────────────────────────
function renderServiceTable(svcData) {{
  const agg={{}};
  svcData.forEach(s=>{{
    const k=s.carrier+'||'+s.service_cat;
    if(!agg[k]) agg[k]={{carrier:s.carrier,service_cat:s.service_cat,count:0,total:0}};
    agg[k].count+=s.count; agg[k].total+=s.total;
  }});
  const rows=Object.values(agg).sort((a,b)=>b.total-a.total);
  const grand=rows.reduce((s,r)=>s+r.total,0);
  document.getElementById('svcCount').textContent=`(${{rows.length}} typer)`;
  const body=document.getElementById('svcBody');
  if(!rows.length){{body.innerHTML='<tr><td colspan="6" class="no-data">Ingen data.</td></tr>';return;}}
  body.innerHTML=rows.map(r=>{{
    const avg=r.count>0?r.total/r.count:0;
    const pct=grand>0?r.total/grand*100:0;
    return `<tr>
      <td><strong>${{esc(r.service_cat)}}</strong></td><td>${{esc(r.carrier)}}</td>
      <td class="num">${{r.count.toLocaleString('sv-SE')}}</td>
      <td class="num">${{fmt(r.total)}}</td><td class="num">${{fmt(avg)}}</td>
      <td class="num">${{pct.toFixed(1)}}%</td></tr>`;
  }}).join('');
  const tc=rows.reduce((s,r)=>s+r.count,0);
  body.innerHTML+=`<tr style="background:#f0f4ff;font-weight:700">
    <td>TOTAL</td><td>—</td><td class="num">${{tc.toLocaleString('sv-SE')}}</td>
    <td class="num">${{fmt(grand)}}</td><td class="num">—</td><td class="num">100.0%</td></tr>`;
}}

// ── Avg cost per service ──────────────────────────────────────────────────────
function renderServiceCostTable(filtData) {{
  const scCats=[...new Set(filtData.flatMap(s=>Object.keys(s.sc_by_cat)))].sort();
  const agg={{}};
  filtData.forEach(s=>{{
    const k=s.carrier+'||'+s.service_cat;
    if(!agg[k]){{agg[k]={{carrier:s.carrier,cat:s.service_cat,count:0,base:0,sc:{{}},grand:0}};
      scCats.forEach(c=>agg[k].sc[c]=0);}}
    agg[k].count+=s.count; agg[k].base+=s.base_total;
    scCats.forEach(c=>{{agg[k].sc[c]+=s.sc_by_cat[c]||0;}});
    agg[k].grand+=s.grand_total;
  }});
  const rows=Object.values(agg).sort((a,b)=>b.grand-a.grand);
  document.getElementById('svcCostCount').textContent=`(${{rows.length}} typ${{rows.length!==1?'er':''}})`;
  const scH=scCats.map(c=>`<th class="num">${{esc(c)}}<br><small style="font-weight:400">snitt</small></th>`).join('');
  document.getElementById('svcCostHead').innerHTML=`<tr>
    <th>Tjänstetyp</th><th>Transportör</th><th class="num">Antal</th>
    <th class="num">Snitt bas (SEK)</th>${{scH}}<th class="num">Snitt total (SEK)</th></tr>`;
  if(!rows.length){{document.getElementById('svcCostBody').innerHTML='<tr><td colspan="99" class="no-data">Ingen data.</td></tr>';return;}}
  document.getElementById('svcCostBody').innerHTML=rows.map(r=>{{
    const n=r.count||1;
    const sc=scCats.map(c=>`<td class="num">${{fmt(r.sc[c]/n)}}</td>`).join('');
    return `<tr><td><strong>${{esc(r.cat)}}</strong></td><td>${{esc(r.carrier)}}</td>
      <td class="num">${{r.count.toLocaleString('sv-SE')}}</td>
      <td class="num">${{fmt(r.base/n)}}</td>${{sc}}
      <td class="num"><strong>${{fmt(r.grand/n)}}</strong></td></tr>`;
  }}).join('');
}}

// ── Anomaly table ─────────────────────────────────────────────────────────────
function renderAnomalyTable(data) {{
  document.getElementById('anomalyCount').textContent=data.length?`(${{data.length}})`:'';
  const body=document.getElementById('anomalyBody');
  if(!data.length){{body.innerHTML='<tr><td colspan="6" class="no-data">Inga avvikelser för valda filter.</td></tr>';return;}}
  const sS={{Warning:'background:#fff8e1;color:#f57f17',Error:'background:#ffebee;color:#c62828',Info:'background:#e3f2fd;color:#1565c0'}};
  const sI={{Warning:'⚠',Error:'✗',Info:'ℹ'}};
  body.innerHTML=data.map(a=>{{
    const sty=sS[a.severity]||'background:#f5f5f5;color:#333';
    const ic=sI[a.severity]||'?';
    const exp=a.explanation?`<span style="color:#555">${{esc(a.explanation)}}</span>`:`<span style="color:#ccc;font-style:italic">—</span>`;
    return `<tr><td>${{esc(a.carrier)}}</td><td>${{esc(a.invoice_no)}}</td><td>${{esc(a.type)}}</td>
      <td><span class="badge" style="${{sty}}">${{ic}} ${{esc(a.severity)}}</span></td>
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

  const fields=[
    ['Faktura #',invNo],['Transportör',inv.carrier],
    ['Fakturadatum',inv.date||'—'],['Förfallodatum',inv.due_date||'—'],
    ['Kundnummer',inv.customer_number||'—'],['Valuta',inv.currency||'SEK'],
    ['Total ex-moms', fmt(inv.total)+' SEK'+(isPending?' (spec)':'')],
    ['Moms', inv.vat_amount?fmt(inv.vat_amount)+' SEK':'—'],
    ['Total inkl moms', inv.total_inc_vat?fmt(inv.total_inc_vat)+' SEK':'—'],
    ['Status', badge(inv.status)],['Källfil', esc(inv.source_file||'—')],
  ];
  if(isPending&&inv.pending_note) fields.push(['Notering',esc(inv.pending_note)]);

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
}}
function closeModalOnBg(e) {{ if(e.target===document.getElementById('modalOverlay')) closeModal(); }}
document.addEventListener('keydown', e=>{{ if(e.key==='Escape') closeModal(); }});

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
