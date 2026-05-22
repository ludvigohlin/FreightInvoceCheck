"""Deterministic anomaly detection — code decides, Claude only explains."""

from __future__ import annotations

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
    cost_per_kg_multiplier = thresholds.get("max_cost_per_kg_multiplier", 2.5)

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
            description="No Bring PDF invoice found in inbox for this specification.",
            suggested_action="Verify PDF invoice was received and placed in 01_Inbox.",
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

    # ── Duplicate shipment numbers ────────────────────────────────────────────
    shipment_counts = Counter(
        ln.shipment_number for ln in base_lines if ln.shipment_number
    )
    for shipment_num, count in shipment_counts.items():
        if count > 1:
            anomalies.append(Anomaly(
                anomaly_type="DuplicateShipment",
                severity="Warning",
                carrier=carrier,
                invoice_number=inv_num,
                description=f"Shipment {shipment_num} appears {count} times on base freight lines.",
                detail=f"Shipment: {shipment_num}",
                value=float(count),
                suggested_action="Check for duplicate billing of the same shipment.",
            ))

    # ── Weight anomalies ──────────────────────────────────────────────────────
    weighted_lines = [ln for ln in base_lines if ln.weight_kg and ln.weight_kg > 0 and ln.amount]
    if weighted_lines:
        avg_cost_per_kg = (
            sum(ln.amount for ln in weighted_lines) /
            sum(ln.weight_kg for ln in weighted_lines)
        )
        for ln in base_lines:
            ship_ref = f"Shipment {ln.shipment_number}" if ln.shipment_number else f"Line {ln.line_no}"

            # Weight (fraktberäknad vikt) exceeds service type limit
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

            # Cost per kg outlier (based on fraktberäknad vikt)
            if (ln.weight_kg and ln.weight_kg > 0 and ln.amount and
                    avg_cost_per_kg > 0):
                cost_per_kg = ln.amount / ln.weight_kg
                if cost_per_kg > avg_cost_per_kg * cost_per_kg_multiplier:
                    anomalies.append(Anomaly(
                        anomaly_type="HighCostPerKg",
                        severity="Info",
                        carrier=carrier,
                        invoice_number=inv_num,
                        description=(
                            f"{ship_ref} cost/fraktberäknad vikt is {cost_per_kg:.2f} SEK/kg — "
                            f"{cost_per_kg/avg_cost_per_kg:.1f}x invoice average ({avg_cost_per_kg:.2f} SEK/kg)."
                        ),
                        line_no=ln.line_no,
                        value=round(cost_per_kg, 2),
                        threshold=round(avg_cost_per_kg * cost_per_kg_multiplier, 2),
                        detail=f"Shipment: {ln.shipment_number} | Service: {ln.service_name_raw} | Weight: {ln.weight_kg} kg",
                        suggested_action="Check if fraktberäknad vikt is correct — volume weight may be inflating the chargeable weight.",
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

    for a in anomalies:
        log_fn = logger.warning if a.severity in ("Warning", "Error") else logger.info
        log_fn("AnomalyDetection", f"[{a.anomaly_type}] {a.description}")

    return anomalies
