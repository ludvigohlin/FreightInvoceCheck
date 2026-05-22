"""Processing logger — writes to processing_log.csv and prints to console."""

from __future__ import annotations

import csv
import traceback
from datetime import datetime
from pathlib import Path
from typing import Optional

from src import config


class ProcessingLogger:
    def __init__(self, run_id: str):
        self.run_id = run_id
        self._log_path = config.PROCESSING_LOG_CSV
        self._rows: list[dict] = []
        self._fieldnames = [
            "run_id", "timestamp", "level", "step", "file_name", "message", "error_detail"
        ]
        self._ensure_header()

    def _ensure_header(self) -> None:
        if not self._log_path.exists():
            with open(self._log_path, "w", newline="", encoding=config.CSV_ENCODING) as f:
                w = csv.DictWriter(f, fieldnames=self._fieldnames, delimiter=config.CSV_DELIMITER)
                w.writeheader()

    def _write(self, level: str, step: str, file_name: str, message: str,
               error_detail: str = "") -> None:
        row = {
            "run_id": self.run_id,
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "level": level,
            "step": step,
            "file_name": file_name or "",
            "message": message,
            "error_detail": error_detail or "",
        }
        self._rows.append(row)
        with open(self._log_path, "a", newline="", encoding=config.CSV_ENCODING) as f:
            w = csv.DictWriter(f, fieldnames=self._fieldnames, delimiter=config.CSV_DELIMITER)
            w.writerow(row)
        try:
            print(f"[{level}] {step}: {message}")
        except UnicodeEncodeError:
            print(f"[{level}] {step}: {message}".encode("ascii", errors="replace").decode("ascii"))

    def info(self, step: str, message: str, file_name: str = "") -> None:
        self._write("INFO", step, file_name, message)

    def warning(self, step: str, message: str, file_name: str = "",
                error: Optional[Exception] = None) -> None:
        detail = traceback.format_exc() if error else ""
        self._write("WARNING", step, file_name, message, detail)

    def error(self, step: str, message: str, file_name: str = "",
              error: Optional[Exception] = None) -> None:
        detail = traceback.format_exc() if error else ""
        self._write("ERROR", step, file_name, message, detail)

    def get_counts(self) -> dict:
        counts = {"INFO": 0, "WARNING": 0, "ERROR": 0}
        for r in self._rows:
            counts[r["level"]] = counts.get(r["level"], 0) + 1
        return counts
