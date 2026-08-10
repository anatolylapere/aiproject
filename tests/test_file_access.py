from pathlib import Path

import pytest

from src.file_access import (
    WorkbookOpenError,
    check_file_accessible,
    is_password_protected,
    load_password,
    open_workbook,
    read_filelist,
)

TEST_FILES_DIR = Path(__file__).resolve().parent / "test files"
CONFIG_DIR = Path(__file__).resolve().parent.parent / "config"


def test_read_filelist_parses_type_and_path(tmp_path):
    csv_path = tmp_path / "filelist.csv"
    csv_path.write_text(
        "Type,File path\nClaims,C:\\data\\claims.xlsx\nRisk,C:\\data\\risk.xlsx\n",
        encoding="utf-8",
    )

    entries = read_filelist(csv_path)

    assert [e.ingestion_type for e in entries] == ["claims", "risk"]
    assert entries[0].source_file == Path("C:\\data\\claims.xlsx")


def test_read_filelist_rejects_unrecognized_type(tmp_path):
    csv_path = tmp_path / "filelist.csv"
    csv_path.write_text("Type,File path\nBogus,C:\\data\\file.xlsx\n", encoding="utf-8")

    with pytest.raises(ValueError):
        read_filelist(csv_path)


def test_check_file_accessible_true_for_existing_file():
    assert check_file_accessible(TEST_FILES_DIR / "file1_Risk_test.xlsx") is True


def test_check_file_accessible_false_for_missing_file(tmp_path):
    assert check_file_accessible(tmp_path / "does_not_exist.xlsx") is False


def test_load_password_returns_configured_string():
    password = load_password(CONFIG_DIR / "password.json")
    assert isinstance(password, str) and password != ""


def test_sample_workbooks_are_not_password_protected():
    assert is_password_protected(TEST_FILES_DIR / "file1_Risk_test.xlsx") is False


def test_open_workbook_reads_sample_file():
    workbook = open_workbook(TEST_FILES_DIR / "file1_Risk_test.xlsx")
    assert "Risk_Main" in workbook.sheetnames


def test_open_workbook_raises_on_corrupt_file(tmp_path):
    bad_file = tmp_path / "corrupt.xlsx"
    bad_file.write_bytes(b"not a real workbook")

    with pytest.raises(WorkbookOpenError):
        open_workbook(bad_file)
