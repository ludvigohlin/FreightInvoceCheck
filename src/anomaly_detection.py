"""Deterministic anomaly detection — code decides, Claude only explains."""

from __future__ import annotations

import csv
from collections import Counter
from dataclasses import dataclass
from typing import List, Optional

from src import config
from src.bring_parser import BringInvoiceHeader, BringInvoiceLine
from src.processing_logger import ProcessingLogger


@dataclass
class Anomaly:
    anomaly_type: str
    severity: str          # Info | Warning | Error
    carrier: str
    invoice_number: str
    description: str
    detail: str = ""
    line_no: Optional[int] = None
    value: Optional[float] = None
    threshold: Optional[float] = None
    suggested_action: str = ""
    claude_explanation: str = ""   # filled in later if Claude is enabled

    def to_dict(self) -> dict:
        return {
            "anomaly_type": self.anomaly_type,
            "severity": self.severity,
            "carrier": self.carrier,
            "invoice_number": self.invoice_number,
            "description": self.description,
            "detail": self.detail,
            "line_no": self.line_no,
            "value": self.value,
            "threshold": self.threshold,
            "suggested_action": self.suggested_action,
            "claude_explanation": self.claude_explanation,
        }


def detect_bring_anomalies(
    pdf_header: Optional[BringInvoiceHeader],
    excel_header: Optional[BringInvoiceHeader],
    lines: List[BringInvoiceLine],
    logger: ProcessingLogger,
) -> List[Anomaly]:
    """
    Detect anomalies in a Bring invoice deterministically.
    All decisions made here; Claude may add explanations later.
    """
    thresholds = config.load_anomaly_thresholds()
    anomalies: List[Anomaly] = []

    header = pdf_header or excel_header
    carrier = "Bring"
    inv_num = (header and header.invoice_number) or "Unknown"

    max_surcharge_pct = thresholds.get("max_surcharge_share_pct", 25)
    max_fuel_pct = thresholds.get("max_fuel_share_pct", 20)
    high_line_amount = thresholds.get("high_line_amount", 1000)
    max_parcel_cost = thresholds.get("max_parcel_cost", 300)
    max_pallet_cost = thresholds.get("max_pallet_cost", 2500)
    max_pickup_cost = thresholds.get("max_pickup_cost", 300)
    max_parcel_weight = thresholds.get("max_parcel_weight_kg", 35)
    max_pallet_weight = thresholds.get("max_pallet_weight_kg", 1800)
    cost_per_kg_multiplier = thresholds.get("max_cost_per_kg_multiplier", 5.0)
    min_cost_per_kg_absolute = thresholds.get("min_cost_per_kg_absolute", 150)

    svc_cost_limits = {
        "Parcel": max_parcel_cost,
        "Pickup": max_pickup_cost,
        "Pallet": max_pallet_cost,
    }
    svc_weight_limits = {
        "Parcel": max_parcel_weight,
        "Pallet": max_pallet_weight,
    }

    # ── Invoice-level anomalies ───────────────────────────────────────────────

    if not inv_num or inv_num == "Unknown":
        anomalies.append(Anomaly(
            anomaly_type="MissingInvoiceNumber",
            severity="Error",
            carrier=carrier,
            invoice_number=inv_num,
            description="Invoice number could not be detected.",
            suggested_action="Manually identify invoice number and re-process.",
        ))

    if pdf_header is None:
        anomalies.append(Anomaly(
            anomaly_type="MissingPDFInvoice",
            severity="Warning",
            carrier=carrier,
            invoice_number=inv_num,
            description="Specification received — PDF invoice not yet found in inbox.",
            suggested_action="Verify PDF invoice was received and placed in 00_Inbox.",
        ))

    if excel_header is None:
        anomalies.append(Anomaly(
            anomaly_type="MissingSpecification",
            severity="Warning",
            carrier=carrier,
            invoice_number=inv_num,
            description="No Bring Excel specification found for this invoice.",
            suggested_action="Request itemized specification from Bring.",
        ))

    if not lines:
        return anomalies  # No line-level checks possible

    # ── Line-level anomalies ──────────────────────────────────────────────────

    total_amount = sum(ln.amount for ln in lines if ln.amount is not None)
    base_lines = [ln for ln in lines if ln.line_type == "BaseFreight"]
    surcharge_lines = [ln for ln in lines if ln.line_type == "Surcharge"]

    surcharge_total = sum(ln.amount for ln in surcharge_lines if ln.amount is not None)
    fuel_total = sum(
        ln.amount for ln in surcharge_lines
        if ln.surcharge_category == "Fuel" and ln.amount is not None
    )

    # Surcharge share
    if total_amount > 0:
        surcharge_share = (surcharge_total / total_amount) * 100
        if surcharge_share > max_surcharge_pct:
            anomalies.append(Anomaly(
                anomaly_type="HighSurchargeShare",
                severity="Warning",
                carrier=carrier,
                invoice_number=inv_num,
                description=f"Surcharge share ({surcharge_share:.1f}%) exceeds threshold ({max_surcharge_pct}%).",
                value=round(surcharge_share, 1),
                threshold=max_surcharge_pct,
                suggested_action="Review surcharge composition with carrier.",
            ))

        fuel_share = (fuel_total / total_amount) * 100
        if fuel_share > max_fuel_pct:
            anomalies.append(Anomaly(
                anomaly_type="HighFuelSurchargeShare",
                severity="Warning",
                carrier=carrier,
                invoice_number=inv_num,
                description=f"Fuel surcharge share ({fuel_share:.1f}%) exceeds threshold ({max_fuel_pct}%).",
                value=round(fuel_share, 1),
                threshold=max_fuel_pct,
                suggested_action="Check current fuel surcharge rates vs contract.",
            ))

    # ── Service-type-aware high line amount ──────────────────────────────────
    for ln in base_lines:
        if ln.amount is None:
            continue
        limit = svc_cost_limits.get(ln.service_category, high_line_amount)
        if ln.amount > limit:
            ship_ref = f"shipment {ln.shipment_number}" if ln.shipment_number else f"line {ln.line_no}"
            anomalies.append(Anomaly(
                anomaly_type="HighLineAmount",
                severity="Warning",
                carrier=carrier,
                invoice_number=inv_num,
                description=(
                    f"{ship_ref.capitalize()} ({ln.service_category}) amount {ln.amount:,.2f} SEK "
                    f"exceeds {ln.service_category} threshold ({limit:,.0f} SEK)."
                ),
                line_no=ln.line_no,
                value=ln.amount,
                threshold=float(limit),
                detail=f"Shipment: {ln.shipment_number} | Service: {ln.service_name_raw}",
                suggested_action="Verify volume and rate match contract for this service type.",
            ))

    # ── Duplicate package numbers (kollinummer) ───────────────────────────────
    # A försändelse can contain multiple kolli — multiple lines per shipment_number
    # is normal. A duplicate kollinummer means the same physical package was billed twice.
    package_counts = Counter(
        ln.package_number for ln in base_lines if ln.package_number
    )
    for pkg_num, count in package_counts.items():
        if count > 1:
            ship_nums = list({
                ln.shipment_number for ln in base_lines
                if ln.package_number == pkg_num and ln.shipment_number
            })
            ship_ref = f" (försändelse {ship_nums[0]})" if ship_nums else ""
            anomalies.append(Anomaly(
                anomaly_type="DuplicatePackage",
                severity="Warning",
                carrier=carrier,
                invoice_number=inv_num,
                description=(
                    f"Kolli {pkg_num} appears {count} times on base freight lines{ship_ref}."
                ),
                detail=f"Kolli: {pkg_num}" + (f" | Försändelse: {ship_nums[0]}" if ship_nums else ""),
                value=float(count),
                suggested_action="Check for duplicate billing of the same kolli.",
            ))

    # ── Weight anomalies ──────────────────────────────────────────────────────
    weighted_lines = [ln for ln in base_lines if ln.weight_kg and ln.weight_kg > 0 and ln.amount]
    if weighted_lines:
        # Per-service-category cost/kg averages — pallets and parcels have very
        # different rates so comparing across categories produces false positives.
        from collections import defaultdict as _dd
        _svc_wt: dict = _dd(lambda: {"amt": 0.0, "kg": 0.0})
        for ln in weighted_lines:
            cat = ln.service_category or "Unknown"
            _svc_wt[cat]["amt"] += ln.amount
            _svc_wt[cat]["kg"] += ln.weight_kg
        svc_avg_cpkg = {
            cat: v["amt"] / v["kg"]
            for cat, v in _svc_wt.items() if v["kg"] > 0
        }

        # Weight-limit check
        for ln in base_lines:
            ship_ref = f"Shipment {ln.shipment_number}" if ln.shipment_number else f"Line {ln.line_no}"
            weight_limit = svc_weight_limits.get(ln.service_category)
            if weight_limit and ln.weight_kg and ln.weight_kg > weight_limit:
                anomalies.append(Anomaly(
                    anomaly_type="WeightExceedsLimit",
                    severity="Warning",
                    carrier=carrier,
                    invoice_number=inv_num,
                    description=(
                        f"{ship_ref} ({ln.service_category}) fraktberäknad vikt {ln.weight_kg:.1f} kg "
                        f"exceeds service limit ({weight_limit} kg)."
                    ),
                    line_no=ln.line_no,
                    value=ln.weight_kg,
                    threshold=float(weight_limit),
                    detail=f"Shipment: {ln.shipment_number} | Service: {ln.service_name_raw} | Weight: {ln.weight_kg} kg",
                    suggested_action="Verify fraktberäknad vikt — may indicate volume weight exceeds physical weight limit for this service type.",
                ))

        # Cost-per-kg outliers within same service category — require BOTH a relative
        # ratio AND an absolute floor to avoid noise. Cap at 10 to prevent token overflow.
        # Checked both directions: too high (overcharge risk) and too low (a bad
        # fraktberäknad vikt can also make a line implausibly cheap, which a
        # one-sided check would never surface).
        _high_cpkg: list[tuple[float, BringInvoiceLine]] = []
        _low_cpkg: list[tuple[float, BringInvoiceLine]] = []
        for ln in base_lines:
            if ln.weight_kg and ln.weight_kg > 0 and ln.amount:
                cat = ln.service_category or "Unknown"
                cat_avg = svc_avg_cpkg.get(cat, 0)
                if cat_avg > 0:
                    cost_per_kg = ln.amount / ln.weight_kg
                    ratio = cost_per_kg / cat_avg
                    if ratio > cost_per_kg_multiplier and cost_per_kg > min_cost_per_kg_absolute:
                        _high_cpkg.append((ratio, ln))
                    elif ratio < (1.0 / cost_per_kg_multiplier) and cost_per_kg > 0:
                        _low_cpkg.append((ratio, ln))
        for ratio, ln in sorted(_high_cpkg, key=lambda x: -x[0])[:10]:
            ship_ref = f"Shipment {ln.shipment_number}" if ln.shipment_number else f"Line {ln.line_no}"
            cost_per_kg = ln.amount / ln.weight_kg
            cat = ln.service_category or "Unknown"
            cat_avg = svc_avg_cpkg.get(cat, 0)
            anomalies.append(Anomaly(
                anomaly_type="HighCostPerKg",
                severity="Info",
                carrier=carrier,
                invoice_number=inv_num,
                description=(
                    f"{ship_ref} cost/fraktberäknad vikt is {cost_per_kg:.2f} SEK/kg — "
                    f"{ratio:.1f}x {cat} average ({cat_avg:.2f} SEK/kg)."
                ),
                line_no=ln.line_no,
                value=round(cost_per_kg, 2),
                threshold=round(cat_avg * cost_per_kg_multiplier, 2),
                detail=f"Shipment: {ln.shipment_number} | Service: {ln.service_name_raw} | Weight: {ln.weight_kg} kg",
                suggested_action="Check if fraktberäknad vikt is correct — volume weight may be inflating the chargeable weight.",
            ))
        for ratio, ln in sorted(_low_cpkg, key=lambda x: x[0])[:10]:
            ship_ref = f"Shipment {ln.shipment_number}" if ln.shipment_number else f"Line {ln.line_no}"
            cost_per_kg = ln.amount / ln.weight_kg
            cat = ln.service_category or "Unknown"
            cat_avg = svc_avg_cpkg.get(cat, 0)
            anomalies.append(Anomaly(
                anomaly_type="LowCostPerKg",
                severity="Info",
                carrier=carrier,
                invoice_number=inv_num,
                description=(
                    f"{ship_ref} cost/fraktberäknad vikt is only {cost_per_kg:.2f} SEK/kg — "
                    f"{1/ratio:.1f}x below {cat} average ({cat_avg:.2f} SEK/kg)."
                ),
                line_no=ln.line_no,
                value=round(cost_per_kg, 2),
                threshold=round(cat_avg / cost_per_kg_multiplier, 2),
                detail=f"Shipment: {ln.shipment_number} | Service: {ln.service_name_raw} | Weight: {ln.weight_kg} kg",
                suggested_action="Check if fraktberäknad vikt or weight is correct — an implausibly low rate can mean the wrong weight or service was billed.",
            ))

    # ── Negative amounts ──────────────────────────────────────────────────────
    for ln in lines:
        if ln.amount is not None and ln.amount < 0:
            anomalies.append(Anomaly(
                anomaly_type="NegativeAmount",
                severity="Warning",
                carrier=carrier,
                invoice_number=inv_num,
                description=f"Line {ln.line_no} has negative amount: {ln.amount:,.2f} SEK.",
                line_no=ln.line_no,
                value=ln.amount,
                suggested_action="Confirm this is a credit or correction, not a data error.",
            ))

    # ── Zero amounts ──────────────────────────────────────────────────────────
    zero_count = sum(1 for ln in lines if ln.amount == 0.0)
    if zero_count > 0:
        anomalies.append(Anomaly(
            anomaly_type="ZeroAmountLines",
            severity="Info",
            carrier=carrier,
            invoice_number=inv_num,
            description=f"{zero_count} line(s) have amount = 0.00.",
            value=float(zero_count),
            suggested_action="Verify zero-amount lines are expected (e.g., free services).",
        ))

    # ── Unknown classifications ───────────────────────────────────────────────
    unknown_service = [ln for ln in lines if ln.service_category == "Unknown"]
    if unknown_service:
        anomalies.append(Anomaly(
            anomaly_type="UnknownServiceCategory",
            severity="Warning",
            carrier=carrier,
            invoice_number=inv_num,
            description=f"{len(unknown_service)} line(s) could not be classified into a known service_category.",
            value=float(len(unknown_service)),
            suggested_action="Add service classification rules or enable Claude API for ambiguous classification.",
        ))

    unknown_surcharge = [ln for ln in surcharge_lines if ln.surcharge_category == "Unknown"]
    if unknown_surcharge:
        anomalies.append(Anomaly(
            anomaly_type="UnknownSurchargeCategory",
            severity="Warning",
            carrier=carrier,
            invoice_number=inv_num,
            description=f"{len(unknown_surcharge)} surcharge line(s) could not be classified.",
            value=float(len(unknown_surcharge)),
            suggested_action="Add surcharge classification rules or enable Claude API.",
        ))

    # ── Non-Nordic destinations ───────────────────────────────────────────────
    anomalies.extend(detect_non_nordic_destinations(carrier, inv_num, base_lines, logger))

    # ── Price increase vs. historical average ─────────────────────────────────
    anomalies.extend(detect_price_increase_vs_history(carrier, inv_num, lines, logger))

    for a in anomalies:
        log_fn = logger.warning if a.severity in ("Warning", "Error") else logger.info
        log_fn("AnomalyDetection", f"[{a.anomaly_type}] {a.description}")

    return anomalies


_NORDIC_COUNTRIES = {"SE", "NO", "DK", "FI"}


def detect_non_nordic_destinations(
    carrier: str,
    invoice_number: str,
    lines: list,
    logger: ProcessingLogger,
) -> List[Anomaly]:
    """
    Flag base-freight lines destined for countries outside the Nordic region.
    One anomaly per foreign country, listing shipment count and sample IDs.
    """
    from collections import defaultdict

    thresholds = config.load_anomaly_thresholds()
    if not thresholds.get("flag_non_nordic_destinations", True):
        return []
    high_amount = thresholds.get("high_non_nordic_amount", 5000)

    foreign: dict[str, list] = defaultdict(list)
    for ln in lines:
        if getattr(ln, "line_type", None) != "BaseFreight":
            continue
        country = getattr(ln, "to_country", None) or "Unknown"
        if country not in _NORDIC_COUNTRIES:
            foreign[country].append(ln)

    anomalies: List[Anomaly] = []
    for country, lns in sorted(foreign.items()):
        total_amt = sum(ln.amount or 0.0 for ln in lns)
        sample_ids = [
            getattr(ln, "shipment_number", None) or getattr(ln, "kolli_id", None) or ""
            for ln in lns
        ]
        sample_ids = [s for s in sample_ids if s][:5]
        sample_str = ", ".join(sample_ids) + (" …" if len(lns) > 5 else "")
        # Weight severity by amount — a 50 SEK shipment shouldn't compete for
        # attention with a 45,000 SEK one; only the latter should read as Warning.
        severity = "Warning" if total_amt >= high_amount else "Info"
        anomalies.append(Anomaly(
            anomaly_type="NonNordicDestination",
            severity=severity,
            carrier=carrier,
            invoice_number=invoice_number,
            description=(
                f"{len(lns)} shipment(s) to non-Nordic destination "
                f"'{country}' — total {total_amt:,.2f} SEK."
            ),
            detail=f"Country: {country} | Shipments: {sample_str}",
            value=float(len(lns)),
            threshold=high_amount,
            suggested_action=(
                "Verify these shipments are intentional — international rates "
                "may differ significantly from Nordic contract rates."
            ),
        ))
    return anomalies


def _historical_avg_price_by_category(carrier: str, exclude_invoice: str) -> dict:
    """Read invoice_lines.csv and return {service_category: {"amt": total, "qty": total}}
    for all previously-written BaseFreight lines of this carrier, excluding the
    invoice currently being checked. This runs before the current invoice's own
    lines are written (Step 5 in main.py), so the CSV only ever contains genuinely
    prior invoices — no need to filter by date."""
    path = config.INVOICE_LINES_CSV
    agg: dict = {}
    if not path.exists():
        return agg
    try:
        with open(path, encoding=config.CSV_ENCODING, newline="") as f:
            reader = csv.DictReader(f, delimiter=config.CSV_DELIMITER)
            for row in reader:
                if row.get("carrier") != carrier:
                    continue
                if row.get("invoice_number") == exclude_invoice:
                    continue
                if row.get("line_type") != "BaseFreight":
                    continue
                cat = row.get("service_category") or "Unknown"
                try:
                    amt = float(row.get("amount") or 0.0)
                    qty = max(float(row.get("quantity") or 1.0), 1.0)
                except ValueError:
                    continue
                d = agg.setdefault(cat, {"amt": 0.0, "qty": 0.0})
                d["amt"] += amt
                d["qty"] += qty
    except Exception:
        pass
    return agg


def detect_price_increase_vs_history(
    carrier: str,
    invoice_number: str,
    lines: list,
    logger: ProcessingLogger,
) -> List[Anomaly]:
    """
    Compare this invoice's average price per service category (amount / quantity)
    against the historical average across previously processed invoices for the
    same carrier. A gradual or one-off rate increase never trips a single-invoice
    reconciliation check (PDF and Excel can agree perfectly and still both be
    wrong relative to the contract), so this is the only place that would catch it.
    """
    thresholds = config.load_anomaly_thresholds()
    if not thresholds.get("flag_price_increase", True):
        return []
    max_increase_pct = thresholds.get("max_price_increase_pct", 10.0)
    min_history_units = thresholds.get("min_price_history_units", 3)

    base_lines = [
        ln for ln in lines
        if getattr(ln, "line_type", "") == "BaseFreight" and getattr(ln, "amount", None)
    ]
    if not base_lines:
        return []

    cur_agg: dict = {}
    for ln in base_lines:
        cat = getattr(ln, "service_category", None) or "Unknown"
        qty = max(float(getattr(ln, "quantity", 1) or 1), 1.0)
        d = cur_agg.setdefault(cat, {"amt": 0.0, "qty": 0.0})
        d["amt"] += ln.amount
        d["qty"] += qty

    hist_agg = _historical_avg_price_by_category(carrier, exclude_invoice=invoice_number)

    anomalies: List[Anomaly] = []
    for cat, d in cur_agg.items():
        if d["qty"] <= 0:
            continue
        cur_avg = d["amt"] / d["qty"]
        hist = hist_agg.get(cat)
        if not hist or hist["qty"] < min_history_units:
            continue  # not enough prior data to trust the baseline
        hist_avg = hist["amt"] / hist["qty"]
        if hist_avg <= 0:
            continue
        increase_pct = (cur_avg - hist_avg) / hist_avg * 100
        if increase_pct > max_increase_pct:
            anomalies.append(Anomaly(
                anomaly_type="PriceIncreaseVsPreviousInvoice",
                severity="Warning",
                carrier=carrier,
                invoice_number=invoice_number,
                description=(
                    f"{cat} average price {cur_avg:.2f} is {increase_pct:.1f}% higher than the "
                    f"historical average ({hist_avg:.2f}) across previous invoices."
                ),
                value=round(cur_avg, 2),
                threshold=round(hist_avg * (1 + max_increase_pct / 100), 2),
                detail=f"Historical avg based on {hist['qty']:.0f} unit(s) across prior invoices.",
                suggested_action="Verify the carrier hasn't changed rates outside the agreed contract.",
            ))
    return anomalies
