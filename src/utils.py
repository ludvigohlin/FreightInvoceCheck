"""Shared utility functions."""

from __future__ import annotations

import hashlib
import re
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional


def generate_run_id() -> str:
    """Generate a unique run identifier based on timestamp + short UUID."""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    short = str(uuid.uuid4()).split("-")[0]
    return f"{ts}_{short}"


def parse_swedish_number(value: str) -> Optional[float]:
    """
    Parse Swedish-formatted numbers to float.
    Handles thousand separators (space, non-breaking space) and comma decimal.
    Examples: '15 556,25' → 15556.25, '1\xa0955,84' → 1955.84
    """
    if value is None:
        return None
    s = str(value).strip()
    if not s or s in ("-", ""):
        return None
    # Remove non-breaking spaces and regular spaces used as thousand separators
    s = s.replace("\xa0", "").replace(" ", "").replace(" ", "")
    # Replace comma decimal separator
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def parse_date(value: str) -> Optional[str]:
    """
    Parse various date formats and return ISO format (YYYY-MM-DD).
    Handles: 2026-05-16, 2026-05-16 00:00:00, 2026.05.16, 18.05.2026
    """
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    s = str(value).strip()
    if not s:
        return None
    # Try to import dateutil for flexible parsing
    try:
        from dateutil import parser as du_parser
        dt = du_parser.parse(s, dayfirst=True)
        return dt.strftime("%Y-%m-%d")
    except Exception:
        pass
    # Fallback: try common patterns
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%Y.%m.%d", "%d/%m/%Y"):
        try:
            dt = datetime.strptime(s[:10], fmt)
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def safe_float(value) -> Optional[float]:
    """Convert value to float, returning None on failure."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip())
    except (ValueError, TypeError):
        return None


def hash_string(s: str) -> str:
    """Return SHA-256 hex digest of a string, for audit logging."""
    return hashlib.sha256(s.encode("utf-8", errors="replace")).hexdigest()[:16]


def normalize_column_name(name: str) -> str:
    """Lowercase, strip, replace spaces/special chars with underscores."""
    if name is None:
        return ""
    s = str(name).strip().lower()
    s = re.sub(r"[^a-z0-9äöå]+", "_", s)
    s = s.strip("_")
    return s


def infer_country_from_postal_code(postal_code: str) -> str:
    """
    Infer country from postal code format.
    FI-xxxxx → Finland, DK-xxxx → Denmark,
    5-digit → Sweden, 4-digit → Norway.
    """
    if postal_code is None:
        return "Unknown"
    s = str(postal_code).strip().replace(" ", "").upper()
    if s.startswith("FI-"):
        return "FI"
    if s.startswith("DK-"):
        return "DK"
    if len(s) == 5 and s.isdigit():
        return "SE"
    if len(s) == 4 and s.isdigit():
        return "NO"
    return "Unknown"


def ensure_directories(paths: list[Path]) -> None:
    """Create directories if they don't exist."""
    for p in paths:
        p.mkdir(parents=True, exist_ok=True)
