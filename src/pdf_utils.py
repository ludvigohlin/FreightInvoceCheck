"""PDF utility functions — extract text and tables from PDF files."""

from __future__ import annotations

from pathlib import Path
from typing import List, Optional


def extract_text_from_pdf(file_path: Path) -> str:
    """
    Extract all text from a PDF using pdfplumber.
    Returns concatenated text from all pages.
    Raises on failure so caller can handle gracefully.
    """
    import pdfplumber
    with pdfplumber.open(str(file_path)) as pdf:
        parts = []
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                parts.append(t)
        return "\n".join(parts)


def extract_text_by_page(file_path: Path) -> List[str]:
    """Return list of text strings, one per page."""
    import pdfplumber
    pages = []
    with pdfplumber.open(str(file_path)) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            pages.append(t)
    return pages


def safe_extract_text(file_path: Path) -> Optional[str]:
    """Extract text, returning None if extraction fails."""
    try:
        return extract_text_from_pdf(file_path)
    except Exception:
        return None
