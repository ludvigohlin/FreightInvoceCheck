"""Configuration loader — reads .env and JSON config files."""

from __future__ import annotations

import json
import os
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # python-dotenv optional; env vars may be set externally

# ── Project root (one level above this file) ─────────────────────────────────
ROOT = Path(__file__).parent.parent

# ── Folder paths ──────────────────────────────────────────────────────────────
INBOX_DIR = ROOT / "00_Inbox"
OUTPUT_DIR = ROOT / "02_Output"
LOGS_DIR = ROOT / "03_Logs"
CONFIG_DIR = ROOT / "04_Config"

DASHBOARD_DATA_DIR = OUTPUT_DIR / "Dashboard_Data"
SUMMARIES_DIR = OUTPUT_DIR / "Summaries"
FOR_EMAIL_DIR = OUTPUT_DIR / "For_Email"
DASHBOARD_HTML = OUTPUT_DIR / "dashboard.html"
CLAUDE_LOGS_DIR = LOGS_DIR / "Claude_API"

# ── Output file paths ─────────────────────────────────────────────────────────
FILE_INVENTORY_CSV = DASHBOARD_DATA_DIR / "file_inventory.csv"
INVOICE_HEADER_CSV = DASHBOARD_DATA_DIR / "invoice_header.csv"
INVOICE_LINES_CSV = DASHBOARD_DATA_DIR / "invoice_lines.csv"
SURCHARGE_LINES_CSV = DASHBOARD_DATA_DIR / "surcharge_lines.csv"
INVOICE_CHECKS_CSV = DASHBOARD_DATA_DIR / "invoice_checks.csv"
ANOMALIES_CSV = DASHBOARD_DATA_DIR / "anomalies.csv"
PROCESSING_LOG_CSV = LOGS_DIR / "processing_log.csv"

# ── Supported inbox file extensions ──────────────────────────────────────────
SUPPORTED_EXTENSIONS = {".pdf", ".xlsx", ".xls", ".csv"}

# ── Runtime flags from environment ───────────────────────────────────────────
USE_CLAUDE_API: bool = os.getenv("USE_CLAUDE_API", "false").lower() == "true"
ANTHROPIC_API_KEY: str = os.getenv("ANTHROPIC_API_KEY", "")
CLAUDE_MODEL: str = os.getenv("CLAUDE_MODEL", "claude-sonnet-4-6")
MOVE_FILES_AFTER_PROCESSING: bool = (
    os.getenv("MOVE_FILES_AFTER_PROCESSING", "false").lower() == "true"
)

# ── CSV output settings ───────────────────────────────────────────────────────
CSV_DELIMITER = ";"
CSV_ENCODING = "utf-8-sig"  # BOM for Excel compatibility


# ── JSON config loaders ───────────────────────────────────────────────────────
def _load_json(filename: str) -> dict:
    path = CONFIG_DIR / filename
    if not path.exists():
        return {}
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def load_carrier_rules() -> dict:
    return _load_json("carrier_rules.json")


def load_service_mapping() -> dict:
    return _load_json("service_mapping.json")


def load_surcharge_mapping() -> dict:
    return _load_json("surcharge_mapping.json")


def load_validation_rules() -> dict:
    return _load_json("validation_rules.json")


def load_anomaly_thresholds() -> dict:
    return _load_json("anomaly_thresholds.json")


def load_report_settings() -> dict:
    return _load_json("report_settings.json")


def ensure_all_directories() -> None:
    """Create all required output and processing directories."""
    dirs = [
        INBOX_DIR,
        DASHBOARD_DATA_DIR, SUMMARIES_DIR, FOR_EMAIL_DIR, CLAUDE_LOGS_DIR,
    ]
    for d in dirs:
        d.mkdir(parents=True, exist_ok=True)
