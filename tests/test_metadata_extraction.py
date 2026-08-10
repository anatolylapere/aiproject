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
