import openpyxl

from src.anchor_detection import find_header_row, load_anchor_pairs


def test_risk_main_header_detected(risk_test_workbook, risk_anchor_pairs):
    ws = risk_test_workbook["Risk_Main"]
    match = find_header_row(ws, risk_anchor_pairs)

    assert match.row_index == 5
    assert match.start_col == 1
    assert match.end_col == 6
    assert match.matched_pair == ("Policy Number", "Inception Date")
    assert match.header_values == ["Policy Number", "Inception Date", "Insured Name", "Premium", "Premium", "Region"]
    assert match.occurrence_rows == [5]


def test_risk_alt_anchor_uses_non_adjacent_columns(risk_test_workbook, risk_anchor_pairs):
    ws = risk_test_workbook["Risk_AltAnchor"]
    match = find_header_row(ws, risk_anchor_pairs)

    assert match.row_index == 4
    assert match.start_col == 1
    assert match.end_col == 5
    assert match.matched_pair == ("Policy Number External", "Sum Insured")
    # non-adjacent anchor columns: "Policy Number External" in col A(1), "Sum Insured" in col D(4)
    assert match.matched_pair_cols == {"Policy Number External": 1, "Sum Insured": 4}


def test_risk_nodata_header_found_no_rows_below(risk_test_workbook, risk_anchor_pairs):
    ws = risk_test_workbook["Risk_NoData"]
    match = find_header_row(ws, risk_anchor_pairs)

    assert match is not None
    assert match.row_index == 3


def test_notes_sheet_has_no_header_match(risk_test_workbook, risk_anchor_pairs):
    ws = risk_test_workbook["Notes"]
    match = find_header_row(ws, risk_anchor_pairs)

    assert match is None


def test_risk_main1_collects_both_header_occurrences(risk_main1_workbook, risk_anchor_pairs):
    ws = risk_main1_workbook["Risk_Main1"]
    match = find_header_row(ws, risk_anchor_pairs)

    assert match.row_index == 5
    assert match.occurrence_rows == [5, 15]


def test_claims_main_header_detected(claims_test_workbook, claims_anchor_pairs):
    ws = claims_test_workbook["Claims_Main"]
    match = find_header_row(ws, claims_anchor_pairs)

    assert match.row_index == 5
    assert match.matched_pair == ("Claim Number", "Claim Amount")


def test_claims_secondary_header_detected_via_non_adjacent_pair(claims_test_workbook, claims_anchor_pairs):
    ws = claims_test_workbook["Claims_Secondary"]
    match = find_header_row(ws, claims_anchor_pairs)

    assert match.row_index == 3
    assert match.matched_pair == ("Policy Number Source", "Effective Date Value")


def test_claims_nodata_header_found(claims_test_workbook, claims_anchor_pairs):
    ws = claims_test_workbook["Claims_NoData"]
    match = find_header_row(ws, claims_anchor_pairs)

    assert match is not None
    assert match.row_index == 3


def test_readme_sheet_has_no_header_match(claims_test_workbook, claims_anchor_pairs):
    ws = claims_test_workbook["ReadMe"]
    match = find_header_row(ws, claims_anchor_pairs)

    assert match is None


def test_broker_alone_does_not_false_positive_match(risk_test_workbook, risk_anchor_pairs):
    # Risk_Main's metadata row 1 contains "Broker" (part of the unused pair
    # ["Broker", "Effective Date"]) but not "Effective Date" - must not match.
    ws = risk_test_workbook["Risk_Main"]
    match = find_header_row(ws, risk_anchor_pairs)

    assert match.matched_pair != ("Broker", "Effective Date")
    assert match.row_index == 5  # real header, not the metadata row


def test_pair_priority_wins_over_row_position():
    # Earlier-priority pair matches a LOWER row than a later-priority pair's match on
    # a HIGHER row. Pair-priority (not row-priority) must govern: the earlier pair's
    # (lower) row wins, even though a naive top-to-bottom-any-pair scan would stop
    # earlier at the later pair's row.
    workbook = openpyxl.Workbook()
    ws = workbook.active
    ws.append(["Foo", "Bar"])          # row 1: matches the later-priority pair
    ws.append([None, None])            # row 2
    ws.append(["Alpha", "Beta"])       # row 3: matches the earlier-priority pair

    pairs = [("Alpha", "Beta"), ("Foo", "Bar")]
    match = find_header_row(ws, pairs)

    assert match.matched_pair == ("Alpha", "Beta")
    assert match.row_index == 3


def test_load_anchor_pairs_reads_real_config():
    from pathlib import Path

    config_path = Path(__file__).resolve().parent.parent / "config" / "anchor_pairs.json"
    anchors = load_anchor_pairs(config_path)

    assert "risk" in anchors and "claims" in anchors
    assert ("Policy Number", "Inception Date") in anchors["risk"]
