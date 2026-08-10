from pathlib import Path

import openpyxl
import pytest

TEST_FILES_DIR = Path(__file__).resolve().parent / "test files"


def _load_sheet(filename, sheet_name):
    workbook = openpyxl.load_workbook(TEST_FILES_DIR / filename, data_only=True)
    return workbook[sheet_name]


@pytest.fixture
def risk_test_workbook():
    return openpyxl.load_workbook(TEST_FILES_DIR / "file1_Risk_test.xlsx", data_only=True)


@pytest.fixture
def risk_main1_workbook():
    return openpyxl.load_workbook(TEST_FILES_DIR / "file1_Risk_test1.xlsx", data_only=True)


@pytest.fixture
def claims_test_workbook():
    return openpyxl.load_workbook(TEST_FILES_DIR / "file2_Claims_test.xlsx", data_only=True)


@pytest.fixture
def risk_anchor_pairs():
    return [
        ("Policy Number", "Inception Date"),
        ("Policy Number External", "Sum Insured"),
        ("Broker", "Effective Date"),
    ]


@pytest.fixture
def claims_anchor_pairs():
    return [
        ("Claim Number", "Claim Amount"),
        ("Insured Name", "Paid Amount"),
        ("Policy Number", "Date Reported"),
        ("Policy Number Source", "Effective Date Value"),
    ]
