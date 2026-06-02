"""
Claude API client — optional, auditable, never authoritative for financial figures.

Financial calculations, invoice totals, VAT, reconciliation status, and
dashboard data are always computed deterministically in Python code.

Claude may only:
- Classify ambiguous service or surcharge names
- Explain anomalies already detected by code
- Generate a written management summary from pre-calculated numbers
- Help interpret unclear text when deterministic parsing has already failed

The script runs fully without Claude. If API is disabled, unavailable, or
returns invalid output, processing continues with deterministic fallback.
"""

from __future__ import annotations

import json
import traceback
from datetime import datetime
from pathlib import Path
from typing import Optional

from src import config
from src.processing_logger import ProcessingLogger
from src.prompts import (
    PROMPT_VERSION,
    SYSTEM_PROMPT,
    CLASSIFICATION_USER_PROMPT,
    MANAGEMENT_SUMMARY_PROMPT,
    ANOMALY_EXPLANATION_PROMPT,
    VALIDATION_EXPLANATION_PROMPT,
    UNKNOWN_CARRIER_EXTRACTION_PROMPT,
)
from src.utils import hash_string

# Valid allowed categories — Claude output is validated against these
VALID_SERVICE_CATEGORIES = {
    "Parcel", "Pallet", "Service Point", "Parcel Locker", "Pickup", "Return", "Other", "Unknown"
}
VALID_SURCHARGE_CATEGORIES = {
    "Fuel", "Remote Area", "City", "Notification", "Special Handling",
    "Heavy", "Box Address", "Currency", "Sulphur", "Delivery Attempt", "Return", "Other", "Unknown"
}
VALID_LINE_TYPES = {"BaseFreight", "Surcharge", "Tax", "Fee", "Other", "Unknown"}


def is_claude_enabled() -> bool:
    """Return True only if USE_CLAUDE_API=true and ANTHROPIC_API_KEY is set."""
    return config.USE_CLAUDE_API and bool(config.ANTHROPIC_API_KEY)


def _get_client():
    """Lazy-initialize the Anthropic client."""
    from anthropic import Anthropic
    return Anthropic(api_key=config.ANTHROPIC_API_KEY)


def _strip_code_fence(text: str) -> str:
    """Remove markdown ```json ... ``` wrappers Claude sometimes adds."""
    text = text.strip()
    if text.startswith("```"):
        first_nl = text.find("\n")
        if first_nl != -1:
            text = text[first_nl + 1:]
        last_fence = text.rfind("```")
        if last_fence != -1:
            text = text[:last_fence]
    return text.strip()


def _log_request(run_id: str, task_type: str, payload: dict, response_text: str,
                 success: bool, error_msg: str = "") -> None:
    """Write request/response to audit JSONL files in 03_Logs/Claude_API/."""
    ts = datetime.now().isoformat(timespec="seconds")
    payload_str = json.dumps(payload, ensure_ascii=False)

    req_entry = {
        "run_id": run_id,
        "timestamp": ts,
        "task_type": task_type,
        "input_hash": hash_string(payload_str),
        "prompt_version": PROMPT_VERSION,
        "request_payload": payload,
    }
    resp_entry = {
        "run_id": run_id,
        "timestamp": ts,
        "task_type": task_type,
        "response": response_text,
        "success": success,
        "error_message": error_msg,
    }

    req_path = config.CLAUDE_LOGS_DIR / f"claude_requests_{run_id}.jsonl"
    resp_path = config.CLAUDE_LOGS_DIR / f"claude_responses_{run_id}.jsonl"

    with open(req_path, "a", encoding="utf-8") as f:
        f.write(json.dumps(req_entry, ensure_ascii=False) + "\n")
    with open(resp_path, "a", encoding="utf-8") as f:
        f.write(json.dumps(resp_entry, ensure_ascii=False) + "\n")


def classify_ambiguous_line(
    run_id: str,
    line_payload: dict,
    logger: ProcessingLogger,
) -> dict:
    """
    Ask Claude to classify an ambiguous invoice line.
    Returns classification dict. Falls back to Unknown on any failure.
    """
    fallback = {
        "service_category": "Unknown",
        "surcharge_category": "Unknown",
        "line_type": "Unknown",
        "confidence": 0.0,
        "reasoning_short": "Claude not called",
        "should_review_manually": True,
        "classified_by": "Rules",
    }

    if not is_claude_enabled():
        return fallback

    try:
        client = _get_client()
        user_msg = CLASSIFICATION_USER_PROMPT.format(
            input_json=json.dumps(line_payload, ensure_ascii=False, indent=2)
        )
        response = client.messages.create(
            model=config.CLAUDE_MODEL,
            max_tokens=512,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_msg}],
        )
        response_text = response.content[0].text.strip()
        _log_request(run_id, "classify_line", line_payload, response_text, True)

        # Parse and validate JSON (Claude sometimes wraps in ```json fences)
        result = json.loads(_strip_code_fence(response_text))

        # Validate categories — revert invalid values to Unknown
        if result.get("service_category") not in VALID_SERVICE_CATEGORIES:
            logger.warning("ClaudeClient", f"Invalid service_category from Claude: {result.get('service_category')}")
            result["service_category"] = "Unknown"
        if result.get("surcharge_category") not in VALID_SURCHARGE_CATEGORIES | {""}:
            result["surcharge_category"] = "Unknown"
        if result.get("line_type") not in VALID_LINE_TYPES:
            result["line_type"] = "Unknown"
        if not isinstance(result.get("confidence"), (int, float)):
            result["confidence"] = 0.0

        result["classified_by"] = "Claude"
        result["classification_confidence"] = float(result.get("confidence", 0.0))
        result["manual_review_required"] = bool(result.get("should_review_manually", True))
        return result

    except Exception as e:
        err = traceback.format_exc()
        _log_request(run_id, "classify_line", line_payload, "", False, str(e))
        logger.warning("ClaudeClient", f"Claude classification failed: {e}", error=e)
        fallback["classified_by"] = "Rules"
        return fallback


def generate_management_summary(
    run_id: str,
    summary_payload: dict,
    logger: ProcessingLogger,
) -> Optional[str]:
    """
    Ask Claude to write a management summary from pre-calculated data.
    Returns markdown string or None on failure.
    """
    if not is_claude_enabled():
        return None

    try:
        client = _get_client()
        user_msg = MANAGEMENT_SUMMARY_PROMPT.format(
            summary_json=json.dumps(summary_payload, ensure_ascii=False, indent=2)
        )
        response = client.messages.create(
            model=config.CLAUDE_MODEL,
            max_tokens=1024,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_msg}],
        )
        response_text = response.content[0].text.strip()
        _log_request(run_id, "management_summary", summary_payload, response_text, True)
        return response_text

    except Exception as e:
        _log_request(run_id, "management_summary", summary_payload, "", False, str(e))
        logger.warning("ClaudeClient", f"Claude summary generation failed: {e}", error=e)
        return None


def extract_unknown_carrier_invoice(
    run_id: str,
    raw_text: str,
    source_file: str,
    logger: ProcessingLogger,
) -> Optional[dict]:
    """
    Ask Claude to extract invoice data from an unknown carrier's PDF text.
    Returns extraction dict or None on failure.
    All returned amounts must be treated as unverified — manual review required.
    """
    if not is_claude_enabled():
        return None

    try:
        client = _get_client()
        user_msg = UNKNOWN_CARRIER_EXTRACTION_PROMPT.format(
            raw_text=raw_text[:6000],  # cap to avoid token overflow on large PDFs
            source_file=source_file,
        )
        response = client.messages.create(
            model=config.CLAUDE_MODEL,
            max_tokens=1536,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_msg}],
        )
        response_text = response.content[0].text.strip()
        payload = {"source_file": source_file, "text_length": len(raw_text)}
        _log_request(run_id, "unknown_carrier_extraction", payload, response_text, True)

        result = json.loads(_strip_code_fence(response_text))
        if not isinstance(result, dict):
            return None
        return result

    except Exception as e:
        payload = {"source_file": source_file}
        _log_request(run_id, "unknown_carrier_extraction", payload, "", False, str(e))
        logger.warning("ClaudeClient", f"Unknown carrier extraction failed: {e}", error=e)
        return None


def explain_validation_issues(
    run_id: str,
    issues_payload: list[dict],
    logger: ProcessingLogger,
) -> list[dict]:
    """
    Ask Claude to explain Warning/Error validation checks in plain language.
    Returns list of explanation dicts. Falls back to empty list on failure.
    Status values are never changed — Claude only adds human-readable context.
    """
    if not is_claude_enabled() or not issues_payload:
        return []

    try:
        client = _get_client()
        user_msg = VALIDATION_EXPLANATION_PROMPT.format(
            issues_json=json.dumps(issues_payload, ensure_ascii=False, indent=2)
        )
        response = client.messages.create(
            model=config.CLAUDE_MODEL,
            max_tokens=1024,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_msg}],
        )
        response_text = response.content[0].text.strip()
        _log_request(run_id, "explain_validation", {"issues": issues_payload}, response_text, True)

        explanations = json.loads(_strip_code_fence(response_text))
        if not isinstance(explanations, list):
            return []
        return explanations

    except Exception as e:
        _log_request(run_id, "explain_validation", {"issues": issues_payload}, "", False, str(e))
        logger.warning("ClaudeClient", f"Claude validation explanation failed: {e}", error=e)
        return []


def explain_anomalies(
    run_id: str,
    anomaly_payload: list[dict],
    logger: ProcessingLogger,
) -> list[dict]:
    """
    Ask Claude to explain anomalies already detected by code.
    Returns list of explanation dicts. Falls back to empty list on failure.
    """
    if not is_claude_enabled() or not anomaly_payload:
        return []

    try:
        client = _get_client()
        user_msg = ANOMALY_EXPLANATION_PROMPT.format(
            anomaly_json=json.dumps(anomaly_payload, ensure_ascii=False, indent=2)
        )
        response = client.messages.create(
            model=config.CLAUDE_MODEL,
            max_tokens=4096,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_msg}],
        )
        response_text = response.content[0].text.strip()
        _log_request(run_id, "explain_anomalies", {"anomalies": anomaly_payload}, response_text, True)

        explanations = json.loads(_strip_code_fence(response_text))
        if not isinstance(explanations, list):
            return []
        return explanations

    except Exception as e:
        _log_request(run_id, "explain_anomalies", {"anomalies": anomaly_payload}, "", False, str(e))
        logger.warning("ClaudeClient", f"Claude anomaly explanation failed: {e}", error=e)
        return []
