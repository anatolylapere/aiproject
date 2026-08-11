import openpyxl

from src.metadata_extraction import extract_metadata, join_metadata


def test_risk_main_metadata_flattens_key_value_rows(risk_test_workbook):
    ws = risk_test_workbook["Risk_Main"]
    metadata = extract_metadata(ws, header_row_index=5)

    assert metadata.values == ["Broker", "Northstar Brokers", "Reporting Period", "2026-07", "Currency", "GBP"]


def test_risk_altanchor_metadata_keeps_all_cells_in_a_row(risk_test_workbook):
    # Row 2 has 3 non-empty cells (Batch ID, RISK-2026-008, anatoly) - a rigid 2-column
    # key/value rule would silently drop the third cell.
    ws = risk_test_workbook["Risk_AltAnchor"]
    metadata = extract_metadata(ws, header_row_index=4)

    assert metadata.values == ["Source System", "Legacy Platform", "Batch ID", "RISK-2026-008", "anatoly"]


def test_join_metadata_uses_pipe_delimiter(risk_test_workbook):
    ws = risk_test_workbook["Risk_AltAnchor"]
    metadata = extract_metadata(ws, header_row_index=4)

    assert join_metadata(metadata) == "Source System|Legacy Platform|Batch ID|RISK-2026-008|anatoly"


def test_no_metadata_above_header_yields_empty(risk_test_workbook):
    ws = risk_test_workbook["Risk_NoData"]
    metadata = extract_metadata(ws, header_row_index=1)

    assert metadata.values == []
    assert join_metadata(metadata) == ""


# ---------------------------------------------------------------------------
# internal "CR0..." code rows (e.g. Bw0197222 claims.xlsx) - dropped like a blank row
# ---------------------------------------------------------------------------

def _worksheet_from_rows(rows):
    workbook = openpyxl.Workbook()
    ws = workbook.active
    for row in rows:
        ws.append(row)
    return ws


def test_code_row_with_three_or_more_cr0_cells_is_dropped():
    ws = _worksheet_from_rows([
        ["CR001", "CR002", "CR003", "CR004"],
        ["Client", "Example Insurance Ltd", None, None],
        ["Claim Number", "Claim Amount", None, None],
    ])
    metadata = extract_metadata(ws, header_row_index=3)

    assert metadata.values == ["Client", "Example Insurance Ltd"]


def test_code_row_detection_is_case_insensitive():
    ws = _worksheet_from_rows([
        ["cr001", "Cr002", "CR003", None],
        ["Claim Number", "Claim Amount", None, None],
    ])
    metadata = extract_metadata(ws, header_row_index=2)

    assert metadata.values == []


def test_row_with_only_two_cr0_cells_is_kept_as_real_metadata():
    # below the 3-cell threshold - not mistaken for a full code row
    ws = _worksheet_from_rows([
        ["CR001", "CR002", "Some real value", None],
        ["Claim Number", "Claim Amount", None, None],
    ])
    metadata = extract_metadata(ws, header_row_index=2)

    assert metadata.values == ["CR001", "CR002", "Some real value"]


def test_non_code_rows_around_a_code_row_are_still_kept():
    ws = _worksheet_from_rows([
        ["Client", "Example Insurance Ltd", None, None],
        ["CR001", "CR002", "CR003", "CR004"],
        ["Extract Date", "2026-08-01", None, None],
        ["Claim Number", "Claim Amount", None, None],
    ])
    metadata = extract_metadata(ws, header_row_index=4)

    assert metadata.values == ["Client", "Example Insurance Ltd", "Extract Date", "2026-08-01"]
