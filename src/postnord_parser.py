"""
PostNord parser — full per-shipment spec parsing from combined PDF.

PostNord invoices are a single PDF containing:
  Page 1   : Invoice header (totals, dates, invoice-level surcharges)
  Page 2   : Blank / second invoice page
  Pages 3+ : SPECIFIKATION — per-shipment lines grouped by pickup date

Shipment line formats:
  PostNord Parcel        {kolli_id}  {weight} {vol} {fraktdr}  {postal} [{city}]  {amount}
  PostNord Service Point {kolli_id}  {weight} {vol} {fraktdr}  {postal} [{city}]  {amount}
  PostNord Pallet ...    {kolli_id}  {postal} [{city}]  {amount}   (no weight/vol/fraktdr)

Finnish postal codes: FI-{digits}  (no city text on shipment line)
Swedish postal codes: 5 digits     (city sometimes wraps to next line)
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Tuple

from src.processing_logger import ProcessingLogger
from src.utils import parse_swedish_number, parse_date, infer_country_from_postal_code

# ── Regex patterns ────────────────────────────────────────────────────────────

# Parcel / Service Point / Parcel Locker — weight/vol/fraktdr columns present
_PARCEL_RE = re.compile(
    r"^(PostNord\s+(?:Parcel(?:\s+Locker)?|Service\s+Point))\s+"
    r"(\d{14,22})\s+"
    r"(\d+,\d+)\s+(\d+,\d+)\s+(\d+,\d+)\s+"  # weight_kg  vol_m3  fraktdr_vikt (all have commas)
    r"(FI-\d+|DK-\d+|NO-\d+|\d{5})\s*"        # postal code (SE/FI/DK/NO)
    r"([A-ZÅÄÖÉÜ][A-ZÅÄÖÉÜÀÂÆŒÇÄÖÜ\s\-]*?)?"  # city (optional, may wrap to next line)
    r"\s*([\d][\d\s\xa0]*,\d{2,})\s*$",        # amount — allow 2+ decimals for PDF artifacts
    re.UNICODE,
)

# Fallback: Parcel with only weight+vol columns (no fraktdr_vikt)
_PARCEL_SHORT_RE = re.compile(
    r"^(PostNord\s+(?:Parcel(?:\s+Locker)?|Service\s+Point))\s+"
    r"(\d{14,22})\s+"
    r"(\d+,\d+)\s+(\d+,\d+)\s+"              # weight_kg  vol_m3
    r"(FI-\d+|DK-\d+|NO-\d+|\d{5})\s*"
    r"([A-ZÅÄÖÉÜ][A-ZÅÄÖÉÜÀÂÆŒÇÄÖÜ\s\-]*?)?"
    r"\s*([\d][\d\s\xa0]*,\d{2,})\s*$",
    re.UNICODE,
)

# Pallet — no weight/vol/fraktdr columns
_PALLET_RE = re.compile(
    r"^(PostNord\s+Pallet\s+\S+)\s+"
    r"(\d{14,22})\s+"
    r"(FI-\d+|DK-\d+|NO-\d+|\d{5})\s*"
    r"([A-ZÅÄÖÉÜ][A-ZÅÄÖÉÜÀÂÆŒÇÄÖÜ\s\-]*?)?"
    r"\s*([\d][\d\s\xa0]*,\d{2,})\s*$",
    re.UNICODE,
)

# Surcharge/add-on line: text followed by a decimal amount
_SURCHARGE_RE = re.compile(
    r"^(.+?)\s+([\d][\d\s\xa0]*,\d{2})\s*$",
    re.UNICODE,
)

# Section pickup-date header
_PICKUP_DATE_RE = re.compile(r"Inhämtningsdatum:\s*(\d{4}-\d{2}-\d{2})", re.IGNORECASE)

# Section summary line
_SUMMA_RE = re.compile(r"^Summa:\s*\d+\s*kolli\(n\)", re.IGNORECASE)

# Invoice-level surcharges on page 1 (these are billed at invoice level, not per shipment)
_INV_SURCHARGES = [
    (re.compile(r"Drivmedelstillägg\S*,?\s*Paket.*?([\d\s\xa0]+,\d{2})", re.IGNORECASE),
     "Drivmedelstillägg Paket", "Fuel"),
    (re.compile(r"Drivmedelstillägg\S*,?\s*Pall.*?([\d\s\xa0]+,\d{2})", re.IGNORECASE),
     "Drivmedelstillägg Pall", "Fuel"),
    (re.compile(r"Svaveltillägg\S*\s*\(.*?([\d\s\xa0]+,\d{2})", re.IGNORECASE),
     "Svaveltillägg", "Sulphur"),
    (re.compile(r"Valutatillägg\S*\s*\(.*?([\d\s\xa0]+,\d{2})", re.IGNORECASE),
     "Valutatillägg", "Currency"),
]

# Lines to skip unconditionally (boilerplate on every spec page + summary markers)
_SKIP_EXACT = frozenset({
    "SPECIFIKATION",
    "KOLLITAXERAT",
})

_SKIP_CONTAINS = [
    "PostNord Sverige AB",
    "Kundservice Faktura",
    "200 05 Malm",
    "(cid:",
    "117005400",
    "IsiCom AB",
    "ckstensgatan",     # Bäckstensgatan (Swedish ä may vary)
    "431 49",
    "Vår referens:",
    "Var referens:",
    "Faktura-/OCR-nummer",
    "Fakturadatum:",
    "Epost:",
    "Internet:",
    "Kundnummer:",
    "Er referens:",
    "Kund:",
    "Specifikationsnummer",
    "Utförd Tjänst",
    "Kundreferens Vikt",
    "SAMMANFATTNING",
    "Tjänst Totalt",
    "Totalt:",
    "Summa exkl. moms för",
    "Orgnr 556711",
    "Momsregnr",
    "Sidan:",
    "www.postnord.se",
    "Uppge faktura",
    "Från färfallodagen",
    "ndring av priser",    # Ändring av priser
    "ndrar vi priser",
    "postnord.se/priser",
    "postnord.se/villkor",
    "Momssatsen",
    "detaljerad information",
    "Bankgiro",
    "IBAN",
    "Swift/BIC",
    "Momsregnr",
    "Styrelsens",
    "godkänt för F-skatt",
    "Betalning oss",
    "Fakturaavgift",
    "Totalmoms",
    "Summa att betala",
    "Summa exklusive moms",
    "Summa inkl",
    "PostNord Kundservice",
]


def _should_skip(line: str) -> bool:
    s = line.strip()
    if not s:
        return True
    if s in _SKIP_EXACT:
        return True
    sl = s.lower()
    for kw in _SKIP_CONTAINS:
        if kw.lower() in sl:
            return True
    return False


# ── Service and surcharge classification ─────────────────────────────────────

def _classify_service(service_name_raw: str, service_mapping: dict) -> str:
    name_lower = service_name_raw.lower()
    for category, keywords in service_mapping.get("service_categories", {}).items():
        if any(k.lower() in name_lower for k in keywords):
            return category
    return "Unknown"


def _classify_surcharge(surcharge_raw: str, surcharge_mapping: dict) -> str:
    s_lower = surcharge_raw.lower()
    for category, keywords in surcharge_mapping.get("surcharge_categories", {}).items():
        if any(k.lower() in s_lower for k in keywords):
            return category
    return "Unknown"


# ── Data classes ──────────────────────────────────────────────────────────────

@dataclass
class PostNordInvoiceHeader:
    run_id: str
    processed_timestamp: str
    carrier: str = "PostNord"
    invoice_number: str = ""
    invoice_date: str = ""
    due_date: str = ""
    customer_number: str = ""
    customer_reference: str = ""
    period_from: str = ""
    period_to: str = ""
    currency: str = ""
    total_ex_vat: Optional[float] = None
    vat_amount: Optional[float] = None
    total_inc_vat: Optional[float] = None
    source_file: str = ""
    document_type: str = "InvoiceAndSpecification"
    reconciliation_status: str = "NotChecked"
    error_message: str = ""

    def to_dict(self) -> dict:
        return {
            "run_id": self.run_id,
            "processed_timestamp": self.processed_timestamp,
            "carrier": self.carrier,
            "invoice_number": self.invoice_number,
            "invoice_date": self.invoice_date,
            "due_date": self.due_date,
            "customer_number": self.customer_number,
            "customer_reference": self.customer_reference,
            "period_from": self.period_from,
            "period_to": self.period_to,
            "currency": self.currency,
            "total_ex_vat": self.total_ex_vat,
            "vat_amount": self.vat_amount,
            "total_inc_vat": self.total_inc_vat,
            "source_file": self.source_file,
            "document_type": self.document_type,
            "reconciliation_status": self.reconciliation_status,
            "error_message": self.error_message,
        }


@dataclass
class PostNordInvoiceLine:
    run_id: str
    processed_timestamp: str
    invoice_number: str
    carrier: str = "PostNord"
    source_file: str = ""
    line_no: int = 0
    # shipment identifiers
    kolli_id: str = ""
    customer_ref: str = ""
    pickup_date: str = ""
    # address
    to_postal: str = ""
    to_city: str = ""
    from_country: str = "SE"
    to_country: str = "SE"
    from_postal: str = "43149"   # IsiCom warehouse
    # service
    service_name_raw: str = ""
    service_category: str = "Unknown"
    line_type: str = "BaseFreight"   # BaseFreight | Surcharge
    # surcharge fields (populated for line_type == Surcharge)
    surcharge_name_raw: str = ""
    surcharge_category: str = ""
    # measurements
    weight_kg: Optional[float] = None
    vol_m3: Optional[float] = None
    fraktdr_vikt: Optional[float] = None
    # pricing
    quantity: float = 1.0
    unit: str = "st"
    unit_price: Optional[float] = None
    amount: Optional[float] = None
    discount_percent: Optional[float] = 0.0
    vat_type: str = "25%"
    # classification metadata
    classified_by: str = "PDF"
    classification_confidence: float = 1.0
    manual_review_required: bool = False

    def to_dict(self) -> dict:
        return {
            "run_id": self.run_id,
            "processed_timestamp": self.processed_timestamp,
            "carrier": self.carrier,
            "invoice_number": self.invoice_number,
            "source_file": self.source_file,
            "line_no": self.line_no,
            "article_number": "",
            "service_code": self.kolli_id,
            "service_name_raw": self.service_name_raw,
            "service_category": self.service_category,
            "from_country": self.from_country,
            "to_country": self.to_country,
            "quantity": self.quantity,
            "unit": self.unit,
            "unit_price": self.unit_price,
            "discount_percent": self.discount_percent,
            "vat_type": self.vat_type,
            "amount": self.amount,
            "line_type": self.line_type,
            "classified_by": self.classified_by,
            "classification_confidence": self.classification_confidence,
            "manual_review_required": self.manual_review_required,
        }

    def to_surcharge_dict(self) -> dict:
        return {
            "run_id": self.run_id,
            "processed_timestamp": self.processed_timestamp,
            "carrier": self.carrier,
            "invoice_number": self.invoice_number,
            "source_file": self.source_file,
            "line_no": self.line_no,
            "surcharge_raw": self.surcharge_name_raw,
            "surcharge_category": self.surcharge_category,
            "service_name_raw": self.service_name_raw,
            "quantity": self.quantity,
            "unit_price": self.unit_price,
            "amount": self.amount,
            "related_service_category": self.service_category,
            "classified_by": self.classified_by,
            "classification_confidence": self.classification_confidence,
            "manual_review_required": self.manual_review_required,
        }


# ── Page 1 header extraction ──────────────────────────────────────────────────

def _parse_header_page(text: str, run_id: str, source_file: str) -> PostNordInvoiceHeader:
    ts = datetime.now().isoformat(timespec="seconds")
    h = PostNordInvoiceHeader(run_id=run_id, processed_timestamp=ts, source_file=source_file)

    # Invoice / OCR number
    m = re.search(r"Faktura-?/OCR-?nummer\s+(\d+)", text, re.IGNORECASE)
    if m:
        h.invoice_number = m.group(1)

    # Invoice date
    m = re.search(r"Fakturadatum:\s*(\d{4}-\d{2}-\d{2})", text, re.IGNORECASE)
    if m:
        h.invoice_date = parse_date(m.group(1)) or m.group(1)

    # Due date
    m = re.search(r"senast\s+(\d{4}-\d{2}-\d{2})", text, re.IGNORECASE)
    if m:
        h.due_date = parse_date(m.group(1)) or m.group(1)

    # Customer number
    m = re.search(r"Kundnummer:\s*(\d+)", text, re.IGNORECASE)
    if m:
        h.customer_number = m.group(1)

    # Currency (from "Summa att betala, SEK")
    m = re.search(r"Summa att betala,\s*([A-Z]{3})", text, re.IGNORECASE)
    if m:
        h.currency = m.group(1).upper()

    # Total incl VAT (first occurrence of "Summa att betala, SEK AMOUNT")
    m = re.search(r"Summa att betala,\s*[A-Z]{3}\s+([\d\s\xa0]+,\d{2})", text, re.IGNORECASE)
    if m:
        h.total_inc_vat = parse_swedish_number(m.group(1))

    # Total excl VAT
    m = re.search(r"Summa exklusive moms\s+([\d\s\xa0]+,\d{2})", text, re.IGNORECASE)
    if m:
        h.total_ex_vat = parse_swedish_number(m.group(1))

    # VAT
    m = re.search(r"Totalmoms\*?\s+([\d\s\xa0]+,\d{2})", text, re.IGNORECASE)
    if m:
        h.vat_amount = parse_swedish_number(m.group(1))

    # Derived VAT if not found explicitly
    if h.vat_amount is None and h.total_inc_vat is not None and h.total_ex_vat is not None:
        h.vat_amount = round(h.total_inc_vat - h.total_ex_vat, 2)

    return h


def _parse_invoice_surcharges(
    page1_text: str,
    run_id: str,
    invoice_number: str,
    source_file: str,
    service_mapping: dict,
    surcharge_mapping: dict,
    line_counter: list,     # mutable [int] counter
    logger: ProcessingLogger,
) -> List[PostNordInvoiceLine]:
    """Extract invoice-level surcharges from page 1 and return as PostNordInvoiceLine list.

    Uses findall so multiple occurrences of the same charge type (e.g. two
    Drivmedelstillägg Paket lines — one VAT-exempt, one at 25%) are all captured.
    """
    ts = datetime.now().isoformat(timespec="seconds")
    lines = []
    for pattern, clean_name, default_category in _INV_SURCHARGES:
        for m in pattern.finditer(page1_text):
            amount = parse_swedish_number(m.group(1))
            if amount is None:
                logger.warning(
                    "PostNordParser",
                    f"Invoice-level surcharge '{clean_name}' matched but amount "
                    f"could not be parsed from {m.group(1)!r} — surcharge dropped.",
                    file_name=source_file,
                )
                continue
            if amount != 0.0:
                category = _classify_surcharge(clean_name, surcharge_mapping) or default_category
                line_counter[0] += 1
                ln = PostNordInvoiceLine(
                    run_id=run_id,
                    processed_timestamp=ts,
                    invoice_number=invoice_number,
                    source_file=source_file,
                    line_no=line_counter[0],
                    line_type="Surcharge",
                    surcharge_name_raw=clean_name,
                    surcharge_category=category,
                    service_name_raw="",
                    service_category="",
                    amount=amount,
                    unit_price=amount,
                )
                lines.append(ln)
    return lines


# ── Spec pages parsing ────────────────────────────────────────────────────────

def _is_swedish_postal(postal: str) -> bool:
    return bool(re.match(r"^\d{5}$", postal))


def _needs_city(postal: str) -> bool:
    """True if this postal code belongs to a country that has a city on the shipment line."""
    p = postal.upper()
    return not p.startswith("FI-") and not p.startswith("DK-") and not p.startswith("NO-")


def _match_shipment(line: str) -> Optional[Tuple[str, str, Optional[float], Optional[float],
                                                  Optional[float], str, str, float]]:
    """
    Try to match a shipment line.
    Returns (service_name, kolli_id, weight_kg, vol_m3, fraktdr_vikt, postal, city, amount)
    or None if not a shipment line.
    city may be "" if it should come from the next line.
    """
    # Parcel / Service Point (3-column: weight, vol, fraktdr)
    m = _PARCEL_RE.match(line)
    if m:
        service = m.group(1).strip()
        kolli_id = m.group(2)
        weight = parse_swedish_number(m.group(3))
        vol = parse_swedish_number(m.group(4))
        fraktdr = parse_swedish_number(m.group(5))
        postal = m.group(6).strip()
        city = (m.group(7) or "").strip()
        amount = parse_swedish_number(m.group(8))
        return service, kolli_id, weight, vol, fraktdr, postal, city, amount

    # Parcel / Service Point fallback (2-column: weight, vol only — rare malformed line)
    m = _PARCEL_SHORT_RE.match(line)
    if m:
        service = m.group(1).strip()
        kolli_id = m.group(2)
        weight = parse_swedish_number(m.group(3))
        vol = parse_swedish_number(m.group(4))
        postal = m.group(5).strip()
        city = (m.group(6) or "").strip()
        amount = parse_swedish_number(m.group(7))
        return service, kolli_id, weight, vol, None, postal, city, amount

    # Pallet
    m = _PALLET_RE.match(line)
    if m:
        service = m.group(1).strip()
        kolli_id = m.group(2)
        postal = m.group(3).strip()
        city = (m.group(4) or "").strip()
        amount = parse_swedish_number(m.group(5))
        return service, kolli_id, None, None, None, postal, city, amount

    return None


def _parse_spec_pages(
    pages: List[str],
    run_id: str,
    invoice_number: str,
    source_file: str,
    service_mapping: dict,
    surcharge_mapping: dict,
    line_counter: list,
    logger: ProcessingLogger,
) -> List[PostNordInvoiceLine]:
    """
    Parse all spec pages (pages[2:]) line-by-line via state machine.
    Returns list of PostNordInvoiceLine (BaseFreight + non-zero per-shipment Surcharges).
    """
    ts = datetime.now().isoformat(timespec="seconds")
    result: List[PostNordInvoiceLine] = []

    # State
    current: Optional[PostNordInvoiceLine] = None
    pending_city: bool = False
    in_surcharges: bool = False
    pickup_date: str = ""

    def emit_current():
        nonlocal current, pending_city, in_surcharges
        if current is not None:
            result.append(current)
        current = None
        pending_city = False
        in_surcharges = False

    for page_idx, page_text in enumerate(pages[2:], start=3):
        for raw_line in page_text.split("\n"):
            line = raw_line.strip()
            if not line:
                continue
            if _should_skip(line):
                continue

            # Pickup date header
            pd_m = _PICKUP_DATE_RE.match(line)
            if pd_m:
                pickup_date = pd_m.group(1)
                continue

            # Section summary — skip
            if _SUMMA_RE.match(line):
                continue

            # Try to match a shipment line
            ship = _match_shipment(line)
            if ship:
                emit_current()
                service_name, kolli_id, weight, vol, fraktdr, postal, city, amount = ship
                # PDF artifacts can append extra digits to amounts (e.g. 375,002 instead of
                # 375,00). Round to 2 decimal places to recover the correct value.
                if amount is not None:
                    amount = round(amount, 2)
                service_category = _classify_service(service_name, service_mapping)
                to_country = infer_country_from_postal_code(postal)
                line_counter[0] += 1
                current = PostNordInvoiceLine(
                    run_id=run_id,
                    processed_timestamp=ts,
                    invoice_number=invoice_number,
                    source_file=source_file,
                    line_no=line_counter[0],
                    kolli_id=kolli_id,
                    pickup_date=pickup_date,
                    service_name_raw=service_name,
                    service_category=service_category,
                    line_type="BaseFreight",
                    to_postal=postal,
                    to_city=city,
                    to_country=to_country,
                    weight_kg=weight,
                    vol_m3=vol,
                    fraktdr_vikt=fraktdr,
                    amount=amount,
                    unit_price=amount,
                )
                # Determine if city is pending (Swedish/Norwegian postal, city absent)
                if city == "" and _needs_city(postal):
                    pending_city = True
                else:
                    pending_city = False
                    in_surcharges = False
                continue

            # City continuation line (uppercase city name wrapping from previous line)
            if current is not None and pending_city:
                if re.match(r"^[A-ZÅÄÖÉÜÀÂÆŒÇÀÈÙ\s\-]+$", line, re.UNICODE) and not line.isdigit():
                    current.to_city = line
                    pending_city = False
                    continue
                # Not a city — fall through (will be treated as customer ref or surcharge)
                pending_city = False

            # Customer reference line — the first non-skippable, non-shipment line after
            # a shipment (before surcharges). May be pure digits OR alphanumeric (e.g. SR1134).
            # Exception: if the line already looks like a surcharge (some shipments have no
            # customer reference on the invoice), skip straight to surcharge handling.
            if current is not None and not in_surcharges:
                if _SURCHARGE_RE.match(line):
                    in_surcharges = True
                    # fall through to surcharge block below
                else:
                    current.customer_ref = line
                    in_surcharges = True
                    continue

            # Surcharge line
            if current is not None and in_surcharges:
                m = _SURCHARGE_RE.match(line)
                if m:
                    surcharge_name = m.group(1).strip()
                    surcharge_amount = parse_swedish_number(m.group(2))
                    if surcharge_amount is not None:
                        surcharge_amount = round(surcharge_amount, 2)
                    if surcharge_amount is not None and surcharge_amount != 0.0:
                        surcharge_category = _classify_surcharge(surcharge_name, surcharge_mapping)
                        line_counter[0] += 1
                        surcharge_ln = PostNordInvoiceLine(
                            run_id=run_id,
                            processed_timestamp=ts,
                            invoice_number=invoice_number,
                            source_file=source_file,
                            line_no=line_counter[0],
                            line_type="Surcharge",
                            surcharge_name_raw=surcharge_name,
                            surcharge_category=surcharge_category,
                            service_name_raw=current.service_name_raw,
                            service_category=current.service_category,
                            kolli_id=current.kolli_id,
                            pickup_date=pickup_date,
                            to_postal=current.to_postal,
                            to_city=current.to_city,
                            to_country=current.to_country,
                            amount=surcharge_amount,
                            unit_price=surcharge_amount,
                        )
                        result.append(surcharge_ln)
                    continue

    # Emit final pending shipment
    emit_current()

    return result


# ── Public entry point ────────────────────────────────────────────────────────

def parse_postnord_pdf(
    file_path: Path,
    run_id: str,
    logger: ProcessingLogger,
    service_mapping: dict = None,
    surcharge_mapping: dict = None,
) -> Tuple[Optional[PostNordInvoiceHeader], List[PostNordInvoiceLine]]:
    """
    Parse a PostNord combined PDF invoice.
    Returns (header, lines) where lines includes BaseFreight + Surcharge entries.
    """
    from src import config
    from src.pdf_utils import extract_text_by_page

    if service_mapping is None:
        service_mapping = config.load_service_mapping()
    if surcharge_mapping is None:
        surcharge_mapping = config.load_surcharge_mapping()

    try:
        pages = extract_text_by_page(file_path)
    except Exception as e:
        logger.error("PostNordParser", f"Cannot read PDF: {e}",
                     file_name=file_path.name, error=e)
        return None, []

    if not pages:
        logger.error("PostNordParser", "PDF has no pages.", file_name=file_path.name)
        return None, []

    # ── Header from page 1 ────────────────────────────────────────────────────
    header = _parse_header_page(pages[0], run_id, file_path.name)

    if not header.invoice_number:
        logger.warning("PostNordParser",
                       "Could not extract invoice number from page 1.",
                       file_name=file_path.name)

    logger.info(
        "PostNordParser",
        f"PDF header extracted: invoice={header.invoice_number}, "
        f"total_ex_vat={header.total_ex_vat}, date={header.invoice_date}",
        file_name=file_path.name,
    )

    if len(pages) < 3:
        logger.warning("PostNordParser",
                       "PDF has fewer than 3 pages — no spec pages to parse.",
                       file_name=file_path.name)
        return header, []

    # ── Spec parsing ──────────────────────────────────────────────────────────
    line_counter = [0]  # mutable counter shared across helpers

    all_lines: List[PostNordInvoiceLine] = []

    # Invoice-level surcharges from page 1
    inv_surcharges = _parse_invoice_surcharges(
        pages[0], run_id, header.invoice_number, file_path.name,
        service_mapping, surcharge_mapping, line_counter, logger,
    )
    all_lines.extend(inv_surcharges)

    # Per-shipment lines from spec pages
    spec_lines = _parse_spec_pages(
        pages, run_id, header.invoice_number, file_path.name,
        service_mapping, surcharge_mapping, line_counter, logger,
    )
    all_lines.extend(spec_lines)

    base_count = sum(1 for ln in spec_lines if ln.line_type == "BaseFreight")
    surcharge_count = sum(1 for ln in spec_lines if ln.line_type == "Surcharge")
    inv_surcharge_count = len(inv_surcharges)

    logger.info(
        "PostNordParser",
        f"Spec parsed: {base_count} shipments, {surcharge_count} per-shipment surcharges, "
        f"{inv_surcharge_count} invoice-level surcharges",
        file_name=file_path.name,
    )

    return header, all_lines
