"""Extraction of metadata content above a detected header row."""

from dataclasses import dataclass


@dataclass
class MetadataBlock:
    values: list


def extract_metadata(worksheet, header_row_index):
    """Flatten all non-empty cell values from rows 1..header_row_index-1 into reading
    order (row-major, top-to-bottom, then left-to-right). Blank cells and fully-blank
    rows are dropped. Not a fixed key/value pairing - rows may carry any number of
    non-empty cells.
    """
    values = []
    max_col = worksheet.max_column
    for row_index in range(1, header_row_index):
        for col in range(1, max_col + 1):
            value = worksheet.cell(row=row_index, column=col).value
            if value is not None:
                values.append(value)
    return MetadataBlock(values=values)


def join_metadata(metadata_block):
    """Join a MetadataBlock's flattened values into a single '|'-delimited string."""
    return "|".join(str(v) for v in metadata_block.values)
