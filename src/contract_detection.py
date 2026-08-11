"""ContractCodeYear detection: identifies the base contract (and, where applicable,
its section) that an extracted table belongs to, using config/contract_codes.json.

Detection searches four evidence tiers in priority order - metadata, table column
names/values, worksheet name, source filename - stopping at the first tier that
yields a match. See CLAUDE.md-approved plan for the full rationale.

A base contract match can come from either an exact configured alias (word-boundary
isolated) or the code's digit-core followed by exactly two more digits (e.g. BW01972's
'1972' matches the '197225' in a numeric broker-ref field), isolated from any longer
digit run so it can't fire inside an unrelated bigger number.

Section matching falls back, in order, to two row-by-row column sources - tried only
after all four phrase tiers come up empty:
  1. A table column literally named 'Property Sec'/'Property Section': each row's own
     value IS the section code.
  2. A column named 'Province': each row's value maps to a section (BC/B.C./British
     Columbia -> A, AB/A.B./Alberta -> B).
Both are read per row, not just the first row. If every row with a resolvable value
agrees, the whole sheet resolves to that section as usual. If rows disagree, the sheet
genuinely covers more than one section - ContractDetectionResult.row_sections is set to
a {worksheet_row_number: section} mapping, and the caller (sheet_processing.py) splits
the extracted rows into separate per-section results instead of using one section for
the whole sheet.

Every match (base contract and section) is tracked back to the specific cell it came
from ('A1'-style, or '(worksheet name)'/'(filename)'/'(metadata)' where there's no
single cell) - surfaced in ContractDetectionResult.detail, including the ambiguous-
match warning, e.g. "multiple contracts matched in table_columns_and_values:
{'BW01972': 'G6', 'BW05407': 'C12'}".

Free-text narrative columns ('Loss Details', 'Loss Description') are excluded entirely
from table evidence (header and data cells alike) - unlike structured reference fields,
they hold arbitrary human-written text prone to coincidentally matching an unrelated
contract's alias and producing false ambiguity. This does not affect what gets
extracted/written to output - only what counts as detection evidence.

BW01972-specific special case, checked before the row-by-row fallback above: if
'Property Sec'/'Property Section'/'Section No'/'Sec No'/'Section'/'Property' reads 'C'
(or 'Section C'/'Sec C') on every row, the sheet is reclassified entirely to BW05407
(Legal Expenses) - irrespective of Province, which is never consulted once this
triggers. Only whole-sheet agreement triggers it; a mix of 'C' and other values falls
through to normal section resolution untouched. Any result resolving to BW05407 (via
this override or otherwise) is also grouped into its own output folder independently
of the rest of its source file - see main.py's _write_sectioned_groups.
"""

import json
import re
from dataclasses import dataclass

from openpyxl.utils import get_column_letter

from src.logging_utils import NULL_LOGGER


@dataclass
class ContractCodeConfig:
    contracts: dict  # canonical_code -> {"aliases": [...], "sections": {letter: [...aliases]} | None}


@dataclass
class ContractDetectionResult:
    contract_code_year: str  # e.g. "BW0197221A", "BW01973", or "" if unresolved
    status: str  # "resolved" | "not_found" | "ambiguous"
    tier: str  # evidence tier(s) that resolved it, "" if unresolved
    detail: str  # human-readable trace string, for logging + WorksheetResult.warnings
    section: str = None  # resolved section letter alone (e.g. "A"), or None if not applicable
    base_code: str = None  # the resolved base contract, set whenever status == "resolved"
    row_sections: dict = None  # {worksheet_row_number: section_or_None} when rows disagree on
    # section (Property Sec/Province, read row by row) and the sheet must be split - the
    # caller partitions the extracted rows by this mapping instead of using a single
    # contract_code_year/section for the whole sheet


def load_contract_codes(config_path):
    """Load and validate config/contract_codes.json into a ContractCodeConfig."""
    with open(config_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, dict):
        raise ValueError(f"Contract codes config must be a JSON object: {config_path}")

    contracts = data.get("contracts")
    if not isinstance(contracts, dict):
        raise ValueError(f"Contract codes config missing 'contracts' object: {config_path}")

    for code, entry in contracts.items():
        if not isinstance(entry, dict) or not isinstance(entry.get("aliases"), list):
            raise ValueError(f"Contract '{code}' must define an 'aliases' list: {config_path}")
        sections = entry.get("sections")
        if sections is not None:
            if not isinstance(sections, dict):
                raise ValueError(f"Contract '{code}' 'sections' must be an object: {config_path}")
            for section_letter, section_entry in sections.items():
                if not isinstance(section_entry, dict) or not isinstance(section_entry.get("aliases"), list):
                    raise ValueError(
                        f"Contract '{code}' section '{section_letter}' must define an 'aliases' list: {config_path}"
                    )

    return ContractCodeConfig(contracts=contracts)


def _alias_pattern(alias):
    """Word aliases (no digits, e.g. 'Hardy') require word-boundary isolation so they
    don't match inside unrelated text. Code-style aliases (contain a digit, e.g. the
    canonical 'BW01972') are specific enough on their own that a plain substring search
    is safe - and is what lets a broker-ref style value like 'B1262BW0197219' resolve.
    """
    alias = str(alias)
    if any(c.isdigit() for c in alias):
        return re.compile(re.escape(alias), re.IGNORECASE)
    return re.compile(r"(?<![A-Za-z0-9])" + re.escape(alias) + r"(?![A-Za-z0-9])", re.IGNORECASE)


def _digit_core(code):
    """The code's digits with any leading zeros dropped, e.g. 'BW01972' -> '1972'."""
    return re.sub(r"\D", "", str(code)).lstrip("0")


def _numeric_year_pattern(code):
    """Match the code's digit-core immediately followed by exactly two more digits
    (e.g. 1972 -> matches the '197225' in a 'Broker Ref: 197225' cell), isolated from
    any longer digit run so it can't match inside an unrelated bigger number. None if
    the code has too few digits to be a specific enough fragment (avoids over-broad
    matches like any 2-digit number in the workbook).
    """
    digit_core = _digit_core(code)
    if len(digit_core) < 3:
        return None
    return re.compile(r"(?<!\d)" + re.escape(digit_core) + r"\d{2}(?!\d)")


def _build_alias_index(aliases_by_code):
    """aliases_by_code: {code: [alias, ...]} -> [(compiled_pattern, code), ...]"""
    return [(_alias_pattern(alias), code) for code, aliases in aliases_by_code.items() for alias in aliases]


def _build_contract_index(contracts):
    """Base-contract matching: exact alias match OR the code's digit-core+year pattern,
    e.g. BW01972's aliases plus the '1972NN' numeric pattern.
    """
    index = []
    for code, entry in contracts.items():
        for alias in entry["aliases"]:
            index.append((_alias_pattern(alias), code))
        numeric_pattern = _numeric_year_pattern(code)
        if numeric_pattern is not None:
            index.append((numeric_pattern, code))
    return index


def _cell_ref(col, row):
    """'A1'-style reference, or a generic placeholder if position info isn't known."""
    if col is None or row is None:
        return "(unknown cell)"
    return f"{get_column_letter(col)}{row}"


def _find_alias_matches(alias_index, evidence):
    """evidence: [(text, cell_ref), ...]. Returns {code: cell_ref} - the first cell
    where each code was matched, for traceability (e.g. in the ambiguous-match warning).
    """
    matched = {}
    for text, cell_ref in evidence:
        for pattern, code in alias_index:
            if code not in matched and pattern.search(text):
                matched[code] = cell_ref
    return matched


_SECTION_COLUMN_NAMES = {"property sec", "property section"}
_PROVINCE_COLUMN_NAMES = {"province"}
_PROVINCE_TO_SECTION = {
    "bc": "A", "b.c": "A", "b.c.": "A", "british columbia": "A",
    "ab": "B", "a.b": "B", "a.b.": "B", "alberta": "B",
}


def _province_value(raw):
    return _PROVINCE_TO_SECTION.get(raw.strip().lower())


_LEGAL_EXPENSES_TRIGGER_CONTRACT = "BW01972"
LEGAL_EXPENSES_CODE = "BW05407"  # public: main.py groups any result resolving to this
# code into its own output folder, independently of what else shares its source file -
# same treatment as a resolved section.
_LEGAL_EXPENSES_COLUMN_NAMES = {
    "property sec", "property section", "section no", "sec no", "section", "property",
}


def _is_legal_expenses_value(raw):
    """'C', 'Section C', 'Sec C' (case-insensitive, trimmed) all count."""
    normalized = raw.strip().lower()
    for prefix in ("section ", "sec "):
        if normalized.startswith(prefix):
            normalized = normalized[len(prefix):]
            break
    return normalized == "c"


def _legal_expenses_override(header_values, data_rows, data_row_numbers, start_col):
    """BW01972-specific special case: if a 'Property Sec'/'Property Section'/'Section
    No'/'Sec No'/'Section'/'Property' column reads 'C' (or 'Section C'/'Sec C') on
    every row, the sheet doesn't belong to a section of BW01972 at all - it's actually
    a different contract, BW05407 (Legal Expenses), irrespective of Province (never
    consulted once this triggers). Only whole-sheet agreement triggers it; if the
    column exists but isn't uniformly 'C', this returns None and normal section
    resolution proceeds untouched.
    """
    idx = _column_index(header_values, _LEGAL_EXPENSES_COLUMN_NAMES)
    if idx is None:
        return None

    row_numbers = data_row_numbers if len(data_row_numbers) == len(data_rows) else [None] * len(data_rows)
    first_cell = None
    for row, row_num in zip(data_rows, row_numbers):
        value = row[idx] if idx < len(row) else None
        if value is None or not _is_legal_expenses_value(str(value)):
            return None
        if first_cell is None:
            col = (start_col + idx) if start_col is not None else None
            first_cell = _cell_ref(col, row_num)
    return first_cell


def _column_index(header_values, names):
    """Index of the first header matching one of the given names (case-insensitive,
    trimmed), or None if no such column exists.
    """
    for idx, header in enumerate(header_values):
        if header is not None and str(header).strip().lower() in names:
            return idx
    return None


def _resolve_row_based_section(
    header_values, data_rows, data_row_numbers, sections, column_names, start_col, value_map=None,
):
    """Read a section from a column identified by name (e.g. 'Property Sec', 'Province')
    - the cell value itself IS the section code, or maps to one via value_map (e.g.
    Province's 'BC'/'Alberta' text) - evaluated row by row, not just the first row.

    Returns None if no such column exists, or if the column exists but no row has a
    value that resolves to a configured section (falls through to the next fallback
    either way). Otherwise a tuple:
      ("single", section, cell_ref) - every row with a resolvable value agrees
      ("split", {row_number: section_or_None}) - rows disagree; caller must split
    """
    idx = _column_index(header_values, column_names)
    if idx is None:
        return None

    row_numbers = data_row_numbers if len(data_row_numbers) == len(data_rows) else [None] * len(data_rows)
    per_row = {}
    for row, row_num in zip(data_rows, row_numbers):
        value = row[idx] if idx < len(row) else None
        if value is None:
            per_row[row_num] = None
            continue
        raw = str(value).strip()
        mapped = value_map(raw) if value_map else raw
        per_row[row_num] = mapped if mapped in sections else None

    distinct = {section for section in per_row.values() if section is not None}
    if not distinct:
        return None
    if len(distinct) == 1:
        section = next(iter(distinct))
        first_row = next(row_num for row_num, s in per_row.items() if s == section)
        col = (start_col + idx) if start_col is not None else None
        return "single", section, _cell_ref(col, first_row)
    return "split", per_row


def _metadata_evidence(metadata):
    cells = metadata.cells if len(metadata.cells) == len(metadata.values) else [None] * len(metadata.values)
    return [(str(v), cell or "(metadata)") for v, cell in zip(metadata.values, cells)]


_EXCLUDED_EVIDENCE_COLUMNS = {"loss details", "loss description"}


def _table_evidence(header_values, data_rows, header_row, start_col, data_row_numbers):
    """Free-text narrative columns (Loss Details/Loss Description) are excluded
    entirely, header and data cells alike - unlike structured reference fields
    (Agreement Number, Broker Ref), they hold arbitrary human-written text that's
    prone to coincidentally matching an unrelated contract's alias, producing false
    'multiple contracts matched' ambiguity instead of genuine evidence.
    """
    excluded_cols = {
        idx for idx, v in enumerate(header_values)
        if v is not None and str(v).strip().lower() in _EXCLUDED_EVIDENCE_COLUMNS
    }

    evidence = []
    for idx, v in enumerate(header_values):
        if v is not None and idx not in excluded_cols:
            col = (start_col + idx) if start_col is not None else None
            evidence.append((str(v), _cell_ref(col, header_row)))

    row_numbers = data_row_numbers if len(data_row_numbers) == len(data_rows) else [None] * len(data_rows)
    for row, row_num in zip(data_rows, row_numbers):
        for idx, v in enumerate(row):
            if v is not None and idx not in excluded_cols:
                col = (start_col + idx) if start_col is not None else None
                evidence.append((str(v), _cell_ref(col, row_num)))
    return evidence


def _evidence_tiers(source_file, worksheet_name, metadata, header_values, data_rows, header_row, start_col, data_row_numbers):
    return [
        ("metadata", _metadata_evidence(metadata)),
        ("table_columns_and_values", _table_evidence(header_values, data_rows, header_row, start_col, data_row_numbers)),
        ("worksheet_name", [(worksheet_name, "(worksheet name)")]),
        ("filename", [(source_file.stem, "(filename)")]),
    ]


def detect_contract_code_year(
    *, source_file, worksheet_name, metadata, header_values, data_rows,
    header_row, start_col, data_row_numbers,
    contract_config, logger=NULL_LOGGER, ingestion_type=None,
):
    tiers = _evidence_tiers(
        source_file, worksheet_name, metadata, header_values, data_rows, header_row, start_col, data_row_numbers,
    )

    contract_alias_index = _build_contract_index(contract_config.contracts)

    code = None
    base_tier = ""
    base_cell = ""
    for tier_name, evidence in tiers:
        matched = _find_alias_matches(contract_alias_index, evidence)
        if len(matched) == 1:
            code = next(iter(matched))
            base_tier = tier_name
            base_cell = matched[code]
            break
        if len(matched) > 1:
            return ContractDetectionResult(
                contract_code_year="", status="ambiguous", tier=tier_name,
                detail=f"multiple contracts matched in {tier_name}: {matched}",
            )

    if code is None:
        return ContractDetectionResult(
            contract_code_year="", status="not_found", tier="",
            detail="no contract evidence found in metadata, table, worksheet name, or filename",
        )

    sections = contract_config.contracts[code].get("sections")
    if not sections:
        return ContractDetectionResult(
            contract_code_year=code, status="resolved", tier=base_tier,
            detail=f"matched {code} via {base_tier} ({base_cell})", base_code=code,
        )

    # 1) phrase-based tiers: metadata, table text, worksheet name, filename (unchanged
    # priority order, sequential stop-at-first-match, same as base-contract resolution)
    section_alias_index = _build_alias_index(
        {letter: entry["aliases"] for letter, entry in sections.items()}
    )
    for tier_name, evidence in tiers:
        matched_sections = _find_alias_matches(section_alias_index, evidence)
        if len(matched_sections) == 1:
            section = next(iter(matched_sections))
            return ContractDetectionResult(
                contract_code_year=f"{code}{section}", status="resolved", tier=f"{base_tier}+{tier_name}",
                detail=(
                    f"matched {code} via {base_tier} ({base_cell}); "
                    f"section {section} via {tier_name} ({matched_sections[section]})"
                ),
                section=section, base_code=code,
            )
        if len(matched_sections) > 1:
            return ContractDetectionResult(
                contract_code_year=code, status="resolved", tier=base_tier,
                detail=(
                    f"matched {code} via {base_tier} ({base_cell}); section ambiguous in {tier_name}: "
                    f"{matched_sections} - section dropped"
                ),
                base_code=code,
            )

    # 1.5) BW01972-specific special case: a Property Sec/Property Section/Section No/
    # Sec No/Section/Property column reading 'C' (or 'Section C'/'Sec C') on every row
    # means the sheet is really BW05407 (Legal Expenses), not a section of BW01972 -
    # checked before the normal row-based fallback below since it reads the very same
    # columns.
    if code == _LEGAL_EXPENSES_TRIGGER_CONTRACT:
        override_cell = _legal_expenses_override(header_values, data_rows, data_row_numbers, start_col)
        if override_cell is not None:
            return ContractDetectionResult(
                contract_code_year=LEGAL_EXPENSES_CODE, status="resolved",
                tier=f"{base_tier}+legal_expenses_override",
                detail=(
                    f"matched {code} via {base_tier} ({base_cell}); reclassified to "
                    f"{LEGAL_EXPENSES_CODE} (Legal Expenses) via legal_expenses_override ({override_cell})"
                ),
                base_code=LEGAL_EXPENSES_CODE,
            )

    # 2) last resort, only reached if every phrase tier came up empty: 'Property Sec'/
    # 'Property Section' column, read row by row (not just the first row) - rows that
    # disagree mean the sheet itself covers more than one section and must be split.
    # 3) if that column doesn't exist at all, 'Province' column, same row-by-row logic,
    # mapping BC/British Columbia -> A and AB/Alberta -> B.
    for source_name, column_names, value_map in (
        ("property_column", _SECTION_COLUMN_NAMES, None),
        ("province_column", _PROVINCE_COLUMN_NAMES, _province_value),
    ):
        row_result = _resolve_row_based_section(
            header_values, data_rows, data_row_numbers, sections, column_names, start_col, value_map,
        )
        if row_result is None:
            continue
        kind = row_result[0]
        if kind == "single":
            _, section, cell = row_result
            return ContractDetectionResult(
                contract_code_year=f"{code}{section}", status="resolved", tier=f"{base_tier}+{source_name}",
                detail=f"matched {code} via {base_tier} ({base_cell}); section {section} via {source_name} ({cell})",
                section=section, base_code=code,
            )
        _, row_sections = row_result
        return ContractDetectionResult(
            contract_code_year=code, status="resolved", tier=base_tier,
            detail=(
                f"matched {code} via {base_tier} ({base_cell}); section split by {source_name} per row: "
                f"{row_sections}"
            ),
            base_code=code, row_sections=row_sections,
        )

    return ContractDetectionResult(
        contract_code_year=code, status="resolved", tier=base_tier,
        detail=f"matched {code} via {base_tier} ({base_cell}); no section evidence found", base_code=code,
    )
