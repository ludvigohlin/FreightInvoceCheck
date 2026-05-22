"""Scans 01_Inbox for supported files and returns file metadata records."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import List

from src import config
from src.processing_logger import ProcessingLogger


@dataclass
class FileRecord:
    run_id: str
    scan_timestamp: str
    file_name: str
    file_path: str
    file_extension: str
    file_size_bytes: int
    file_modified_timestamp: str
    detected_carrier: str = "Unknown"
    detected_document_type: str = "Unknown"
    detected_invoice_number: str = ""
    processing_status: str = "Found"
    error_message: str = ""

    def to_dict(self) -> dict:
        return {
            "run_id": self.run_id,
            "scan_timestamp": self.scan_timestamp,
            "file_name": self.file_name,
            "file_path": self.file_path,
            "file_extension": self.file_extension,
            "file_size_bytes": self.file_size_bytes,
            "file_modified_timestamp": self.file_modified_timestamp,
            "detected_carrier": self.detected_carrier,
            "detected_document_type": self.detected_document_type,
            "detected_invoice_number": self.detected_invoice_number,
            "processing_status": self.processing_status,
            "error_message": self.error_message,
        }


def scan_inbox(run_id: str, logger: ProcessingLogger,
               inbox_dir: Path = None) -> List[FileRecord]:
    """
    Scan the inbox folder for processable files.
    Returns a list of FileRecord objects.
    Only considers files directly in inbox_dir (not in subfolders).
    """
    if inbox_dir is None:
        inbox_dir = config.INBOX_DIR

    records: List[FileRecord] = []
    scan_ts = datetime.now().isoformat(timespec="seconds")

    if not inbox_dir.exists():
        logger.warning("FileScanner", f"Inbox folder does not exist: {inbox_dir}")
        return records

    # Only scan files at the root of inbox (not subfolders, which are classification folders)
    files = [p for p in inbox_dir.iterdir() if p.is_file()]

    if not files:
        logger.info("FileScanner", "01_Inbox is empty — nothing to process.")
        return records

    logger.info("FileScanner", f"Found {len(files)} file(s) in {inbox_dir}")

    for fp in sorted(files):
        ext = fp.suffix.lower()
        stat = fp.stat()
        modified_ts = datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds")

        if ext not in config.SUPPORTED_EXTENSIONS:
            logger.info(
                "FileScanner",
                f"Skipping unsupported file type '{ext}': {fp.name}",
                file_name=fp.name,
            )
            rec = FileRecord(
                run_id=run_id,
                scan_timestamp=scan_ts,
                file_name=fp.name,
                file_path=str(fp),
                file_extension=ext,
                file_size_bytes=stat.st_size,
                file_modified_timestamp=modified_ts,
                processing_status="SkippedUnsupportedType",
            )
            records.append(rec)
            continue

        rec = FileRecord(
            run_id=run_id,
            scan_timestamp=scan_ts,
            file_name=fp.name,
            file_path=str(fp),
            file_extension=ext,
            file_size_bytes=stat.st_size,
            file_modified_timestamp=modified_ts,
            processing_status="Found",
        )
        records.append(rec)
        logger.info("FileScanner", f"Found: {fp.name} ({stat.st_size:,} bytes)", file_name=fp.name)

    return records
