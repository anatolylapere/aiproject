"""Source file access: filelist parsing, existence checks, password handling, workbook opening.

Workbooks are always opened read-only into memory and never saved back to their
source path, so the original files are never modified.
"""

import csv
import io
import json
from dataclasses import dataclass
from pathlib import Path

import msoffcrypto
import openpyxl

from src.logging_utils import NULL_LOGGER


class WorkbookOpenError(Exception):
    """Raised when a workbook cannot be inspected/opened (corrupt, wrong password, unsupported)."""


@dataclass
class FileListEntry:
    source_file: Path
    ingestion_type: str  # "risk" | "claims"


def read_filelist(csv_path):
    """Parse input/filelist.csv into a list of FileListEntry. Validates the Type column."""
    entries = []
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            raw_type = (row.get("Type") or "").strip().lower()
            raw_path = (row.get("File path") or "").strip()
            if raw_type not in ("risk", "claims"):
                raise ValueError(f"Unrecognized Type '{row.get('Type')}' in filelist for: {raw_path}")
            entries.append(FileListEntry(source_file=Path(raw_path), ingestion_type=raw_type))
    return entries


def check_file_accessible(path):
    """Return True if the file exists and can be opened for reading."""
    path = Path(path)
    if not path.is_file():
        return False
    try:
        with open(path, "rb"):
            pass
    except OSError:
        return False
    return True


def load_password(password_config_path):
    """Return the configured password string. Never log this value."""
    with open(password_config_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data["password"]


def is_password_protected(path):
    """Return True if the workbook at path is encrypted."""
    try:
        with open(path, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            return office_file.is_encrypted()
    except Exception as e:
        raise WorkbookOpenError(f"Failed to inspect workbook {path}: {type(e).__name__}") from e


def open_workbook(path, password=None):
    """Open a workbook read-only into memory. Raises WorkbookOpenError on failure."""
    try:
        if password is not None:
            decrypted = io.BytesIO()
            with open(path, "rb") as f:
                office_file = msoffcrypto.OfficeFile(f)
                office_file.load_key(password=password)
                office_file.decrypt(decrypted)
            decrypted.seek(0)
            return openpyxl.load_workbook(decrypted, data_only=True)
        return openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        raise WorkbookOpenError(f"Failed to open workbook {path}: {type(e).__name__}") from e


def open_source_file(entry, password_config_path, logger=NULL_LOGGER):
    """Check access, check/handle password protection, and open the workbook -
    logging each step as it happens. Raises WorkbookOpenError (already logged) on
    any failure. Returns (workbook, password_required).
    """
    accessible = check_file_accessible(entry.source_file)
    logger.log_step(
        entry.source_file, entry.ingestion_type, None, "file_access_check",
        "ok" if accessible else "fail",
        "file exists and is readable" if accessible else "file not found or not readable",
    )
    if not accessible:
        raise WorkbookOpenError(f"File not accessible: {entry.source_file}")

    try:
        protected = is_password_protected(entry.source_file)
    except WorkbookOpenError as e:
        logger.log_step(entry.source_file, entry.ingestion_type, None, "password_check", "fail", str(e))
        raise
    logger.log_step(
        entry.source_file, entry.ingestion_type, None, "password_check", "ok",
        f"password_protected={protected}",
    )

    password = load_password(password_config_path) if protected else None

    try:
        workbook = open_workbook(entry.source_file, password)
    except WorkbookOpenError as e:
        logger.log_step(entry.source_file, entry.ingestion_type, None, "workbook_open", "fail", str(e))
        raise
    logger.log_step(entry.source_file, entry.ingestion_type, None, "workbook_open", "ok", "opened successfully")

    return workbook, protected
