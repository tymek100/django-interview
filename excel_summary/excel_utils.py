# app_name/excel_utils.py (or services.py, etc.)

from typing import Any, Dict, List, Optional, Tuple



def detect_header_row(ws, requested_columns: List[str], max_header_search: int = 5):
    """
    Try to detect which row contains the headers by matching requested column
    names against the first `max_header_search` rows.

    Returns: (row_index, row_values) or (None, None) if nothing found.
    """
    requested_norm = [normalize_header(c) for c in requested_columns]

    best_row_idx: Optional[int] = None
    best_match_count = 0
    best_row_values = None

    # 1) Try to find the row where most requested column names appear
    for idx, row in enumerate(
        ws.iter_rows(min_row=1, max_row=max_header_search, values_only=True),
        start=1,
    ):
        if not any(cell not in (None, "") for cell in row):
            # skip completely empty rows
            continue

        header_map = {
            normalize_header(cell): col_idx
            for col_idx, cell in enumerate(row)
            if normalize_header(cell)
        }

        match_count = sum(1 for c in requested_norm if c in header_map)

        if match_count > best_match_count:
            best_match_count = match_count
            best_row_idx = idx
            best_row_values = row

    if best_row_idx is not None and best_match_count > 0:
        # Found a row that matches at least one requested column
        return best_row_idx, best_row_values

    # 2) Fallback: first non-empty row
    for idx, row in enumerate(
        ws.iter_rows(min_row=1, max_row=max_header_search, values_only=True),
        start=1,
    ):
        if any(cell not in (None, "") for cell in row):
            return idx, row

    return None, None


def normalize_header(header: Any) -> str:
    """Normalize header cell to allow case-insensitive matching."""
    if header is None:
        return ""
    return str(header).strip().lower()


def coerce_to_number(value: Any) -> Optional[float]:
    """
    Try to convert various cell values to float.

    Handles:
    - numeric types
    - strings like "$90,00", "90,00", "90.00", "1,234.56"
    """
    if value is None:
        return None

    if isinstance(value, (int, float)):
        return float(value)

    if isinstance(value, str):
        s = value.strip()
        if not s:
            return None

        # Remove common currency symbols and spaces
        for ch in ["$", "€", "£"]:
            s = s.replace(ch, "")
        s = s.replace(" ", "")

        # Handle decimal/thousands separators
        if "," in s and "." in s:
            # Assume '.' is decimal, ',' is thousands separator: "1,234.56"
            s = s.replace(",", "")
        elif "," in s and "." not in s:
            # European style: "90,00" -> "90.00"
            s = s.replace(",", ".")

        try:
            return float(s)
        except ValueError:
            return None

    return None

class HeaderNotFoundError(Exception):
    pass


def summarize_excel_columns(ws, columns: list) -> Tuple[list, list, list]:
    """
    Given an openpyxl worksheet and a list of column names,
    return (summaries, missing_columns, available_columns).

    summaries: [
        {"column": "CURRENT USD", "sum": 1234.5, "avg": 56.7},
        ...
    ]
    """
    # Find header row
    header_row_index, header_row_values = detect_header_row(ws, columns)
    if header_row_index is None or header_row_values is None:
        raise HeaderNotFoundError("Could not detect header row in the Excel sheet.")

    # Map normalized header -> column index
    header_map: Dict[str, int] = {}
    for col_idx, cell in enumerate(header_row_values):
        key = normalize_header(cell)
        if key:
            header_map[key] = col_idx

    requested_summaries: List[dict] = []
    missing_columns: List[str] = []

    # Iterate over requested columns and compute sum + avg
    for col_name in columns:
        norm = normalize_header(col_name)
        if norm not in header_map:
            missing_columns.append(col_name)
            continue

        col_idx = header_map[norm]
        total = 0.0
        count = 0

        # Iterate data rows, starting after header
        for row in ws.iter_rows(
            min_row=header_row_index + 1,
            max_row=ws.max_row,
            values_only=True,
        ):
            if col_idx >= len(row):
                continue
            cell_value = row[col_idx]
            number = coerce_to_number(cell_value)
            if number is None:
                continue
            total += number
            count += 1

        if count > 0:
            avg = total / count
            requested_summaries.append(
                {
                    "column": col_name,
                    "sum": round(total, 2),
                    "avg": round(avg, 2),
                }
            )
        else:
            requested_summaries.append(
                {
                    "column": col_name,
                    "sum": 0.0,
                    "avg": 0.0,
                }
            )

    available_columns = list(header_map.keys())
    return requested_summaries, missing_columns, available_columns
