"""Excel formula resolution engine.

Handles: simple cell references, RECHERCHEV (VLOOKUP), ARRONDI (ROUND), cell ranges.
"""

import logging
import re
from datetime import datetime
from typing import Any, Optional, Union

logger = logging.getLogger(__name__)


def _format_cell_value(value: Any) -> Optional[str]:
    """Format a cell value to string."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, str):
        result = value.strip()
        return result if result else None
    return str(value)


def _col_letter_to_num(col: str) -> int:
    """Convert column letter(s) to number. A=1, B=2, ..., Z=26, AA=27."""
    result = 0
    for char in col.upper():
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result


def _parse_cell_ref(ref: str) -> tuple[str, str]:
    """Parse 'A1' into ('A', '1'). Returns (col_letters, row_number_str)."""
    match = re.match(r"([A-Za-z]+)(\d+)", ref.strip())
    if not match:
        raise ValueError(f"Invalid cell reference: {ref}")
    return match.group(1).upper(), match.group(2)


def _find_sheet(workbook, sheet_name: str):
    """Find a worksheet by name, with fuzzy matching for numeric names.

    Handles cases like sheet_name='3' matching '3. Hypothèses'.
    """
    # Exact match
    if sheet_name in workbook.sheetnames:
        return workbook[sheet_name]
    # Fuzzy: sheet_name is a prefix (e.g. "3" matches "3. Hypothèses")
    for ws_name in workbook.sheetnames:
        if ws_name.startswith(sheet_name + ".") or ws_name.startswith(sheet_name + " "):
            return workbook[ws_name]
    # Case-insensitive
    for ws_name in workbook.sheetnames:
        if ws_name.lower() == sheet_name.lower():
            return workbook[ws_name]
    raise KeyError(f"Worksheet {sheet_name} does not exist.")


def resolve_cell_reference(sheet_name: str, cell_ref: str, workbook) -> Optional[str]:
    """Resolve a single cell reference to a formatted string value."""
    try:
        ws = _find_sheet(workbook, sheet_name)
        cell = ws[cell_ref]
        return _format_cell_value(cell.value)
    except Exception as e:
        logger.warning(f"Error resolving {sheet_name}!{cell_ref}: {e}")
        return None


def resolve_cell_range(sheet_name: str, range_ref: str, workbook) -> list[str]:
    """Resolve a cell range (e.g., 'E38:E41') to a list of non-empty string values."""
    try:
        start_ref, end_ref = range_ref.split(":")
        start_col, start_row = _parse_cell_ref(start_ref)
        end_col, end_row = _parse_cell_ref(end_ref)

        ws = _find_sheet(workbook, sheet_name)
        values = []
        for row_num in range(int(start_row), int(end_row) + 1):
            cell_ref = f"{start_col}{row_num}"
            cell = ws[cell_ref]
            formatted = _format_cell_value(cell.value)
            if formatted:
                values.append(formatted)
        return values
    except Exception as e:
        logger.warning(f"Error resolving range {sheet_name}!{range_ref}: {e}")
        return []


def _parse_sheet_and_cell(reference: str) -> tuple[str, str]:
    """Parse "'SheetName'!B5" or "SheetName!B5" into (sheet_name, cell_ref)."""
    ref = reference.strip()
    # Remove leading = if present
    if ref.startswith("="):
        ref = ref[1:]
    # Remove workbook references like [1] or [Classeur1]
    ref = re.sub(r"\[[^\]]*\]", "", ref)

    if "!" not in ref:
        raise ValueError(f"No sheet reference in: {reference}")

    parts = ref.split("!", 1)
    sheet_name = parts[0].strip().strip("'")
    cell_ref = parts[1].strip()
    return sheet_name, cell_ref


def _split_formula_args(args_str: str) -> list[str]:
    """Split formula arguments respecting quotes and parentheses nesting."""
    args = []
    current = ""
    depth = 0
    in_quote = False

    for char in args_str:
        if char == "'" and depth == 0:
            in_quote = not in_quote
            current += char
        elif char == "(" and not in_quote:
            depth += 1
            current += char
        elif char == ")" and not in_quote:
            depth -= 1
            current += char
        elif char == ";" and depth == 0 and not in_quote:
            args.append(current.strip())
            current = ""
        else:
            current += char

    if current.strip():
        args.append(current.strip())
    return args


def _resolve_recherchev(formula: str, workbook) -> Optional[str]:
    """Resolve RECHERCHEV (VLOOKUP) formula."""
    # Extract arguments from RECHERCHEV(...)
    match = re.match(r"RECHERCHEV\((.+)\)", formula, re.IGNORECASE)
    if not match:
        return None

    args = _split_formula_args(match.group(1))
    if len(args) < 3:
        logger.warning(f"RECHERCHEV needs at least 3 args, got {len(args)}: {formula}")
        return None

    lookup_value_ref = args[0]
    table_range_ref = args[1]
    col_index = int(args[2])

    # Resolve lookup value (may be a cell reference)
    lookup_value = resolve_formula(f"={lookup_value_ref}", workbook) if "!" in lookup_value_ref else lookup_value_ref.strip("'\"")

    if lookup_value is None:
        return None

    # Parse table range: 'SheetName'!B9:P14
    sheet_name, range_ref = _parse_sheet_and_cell(table_range_ref)
    start_ref, end_ref = range_ref.split(":")
    start_col, start_row = _parse_cell_ref(start_ref)
    end_col, end_row = _parse_cell_ref(end_ref)

    ws = workbook[sheet_name]
    start_col_num = _col_letter_to_num(start_col)

    # Search for lookup_value in the first column of the range
    for row_num in range(int(start_row), int(end_row) + 1):
        cell = ws[f"{start_col}{row_num}"]
        cell_val = cell.value

        if cell_val is None:
            continue

        # Try numeric comparison first, then string
        match_found = False
        try:
            if float(str(cell_val)) == float(str(lookup_value)):
                match_found = True
        except (ValueError, TypeError):
            if str(cell_val).strip().lower() == str(lookup_value).strip().lower():
                match_found = True

        if match_found:
            # Get value from the target column
            target_col_num = start_col_num + col_index - 1
            # Convert column number back to letter
            target_col = ""
            num = target_col_num
            while num > 0:
                num -= 1
                target_col = chr(num % 26 + ord("A")) + target_col
                num //= 26

            result_cell = ws[f"{target_col}{row_num}"]
            return _format_cell_value(result_cell.value)

    logger.warning(f"RECHERCHEV: lookup value '{lookup_value}' not found in range")
    return None


def _resolve_arrondi(formula: str, workbook) -> Optional[str]:
    """Resolve ARRONDI (ROUND) formula."""
    match = re.match(r"ARRONDI\((.+)\)", formula, re.IGNORECASE)
    if not match:
        return None

    args = _split_formula_args(match.group(1))
    if len(args) < 2:
        logger.warning(f"ARRONDI needs 2 args, got {len(args)}: {formula}")
        return None

    inner_value = resolve_formula(f"={args[0]}", workbook)
    if inner_value is None:
        return None

    try:
        num_decimals = int(args[1])
        result = round(float(inner_value), num_decimals)
        return str(result)
    except (ValueError, TypeError) as e:
        logger.warning(f"ARRONDI error: {e}")
        return None


def resolve_formula(formula: str, workbook) -> Optional[str]:
    """
    Resolve an Excel formula to its value.

    Supports:
    - Simple cell references: ='Sheet'!B5
    - RECHERCHEV (VLOOKUP)
    - ARRONDI (ROUND)
    - Cell ranges: Sheet!E38:E41 (returns first value; use resolve_cell_range for list)
    """
    if not formula or not isinstance(formula, str):
        return None

    f = formula.strip()
    if f.startswith("="):
        f = f[1:]

    # Remove workbook references like [1] or [Classeur1]
    f = re.sub(r"\[[^\]]*\]", "", f)

    # ARRONDI
    if f.upper().startswith("ARRONDI("):
        return _resolve_arrondi(f, workbook)

    # RECHERCHEV
    if f.upper().startswith("RECHERCHEV("):
        return _resolve_recherchev(f, workbook)

    # Simple cell reference: 'Sheet'!B5 or Sheet!B5
    if "!" in f:
        try:
            sheet_name, cell_ref = _parse_sheet_and_cell(f)
            if ":" in cell_ref:
                # It's a range — return first non-empty value
                values = resolve_cell_range(sheet_name, cell_ref, workbook)
                return values[0] if values else None
            return resolve_cell_reference(sheet_name, cell_ref, workbook)
        except Exception as e:
            logger.warning(f"Error parsing formula '{formula}': {e}")
            return None

    logger.warning(f"Unrecognized formula: {formula}")
    return None
