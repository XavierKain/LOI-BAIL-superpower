# tests/test_formula_engine.py
import pytest
from datetime import datetime
from unittest.mock import MagicMock
from core.formula_engine import resolve_formula, resolve_cell_reference, resolve_cell_range


def _make_workbook(sheets: dict) -> MagicMock:
    """Helper: create a mock workbook with sheet data.
    sheets = {"SheetName": {"A1": value, "B2": value, ...}}
    """
    wb = MagicMock()
    wb.sheetnames = list(sheets.keys())

    for sheet_name, cells in sheets.items():
        ws = MagicMock()
        ws.title = sheet_name

        def make_cell(val):
            cell = MagicMock()
            cell.value = val
            return cell

        def getitem(key, _cells=cells):
            if key in _cells:
                return make_cell(_cells[key])
            return make_cell(None)

        ws.__getitem__ = getitem
        wb.__getitem__ = lambda key, _sheets=sheets: (
            _make_sheet_mock(key, _sheets)
        )

    # Simpler approach: build real sheet mocks
    sheet_mocks = {}
    for sheet_name, cells in sheets.items():
        ws = MagicMock()
        ws.title = sheet_name
        ws.__getitem__ = MagicMock(side_effect=lambda key, _cells=cells: _make_cell_mock(_cells.get(key)))
        sheet_mocks[sheet_name] = ws

    wb.__getitem__ = MagicMock(side_effect=lambda key: sheet_mocks[key])
    wb.sheetnames = list(sheets.keys())
    return wb


def _make_cell_mock(value):
    cell = MagicMock()
    cell.value = value
    return cell


def test_simple_cell_reference():
    wb = _make_workbook({"Validation": {"B5": "Hello"}})
    result = resolve_cell_reference("Validation", "B5", wb)
    assert result == "Hello"


def test_cell_reference_date():
    wb = _make_workbook({"Validation": {"B5": datetime(2025, 12, 1)}})
    result = resolve_cell_reference("Validation", "B5", wb)
    assert result == "01/12/2025"


def test_cell_reference_number():
    wb = _make_workbook({"Validation": {"B5": 42}})
    result = resolve_cell_reference("Validation", "B5", wb)
    assert result == "42"


def test_cell_reference_float():
    wb = _make_workbook({"Validation": {"B5": 3.14}})
    result = resolve_cell_reference("Validation", "B5", wb)
    assert result == "3.14"


def test_cell_reference_none():
    wb = _make_workbook({"Validation": {}})
    result = resolve_cell_reference("Validation", "B5", wb)
    assert result is None


def test_resolve_formula_simple_ref():
    wb = _make_workbook({"Hypotheses": {"E5": "test value"}})
    result = resolve_formula("='Hypotheses'!E5", wb)
    assert result == "test value"


def test_resolve_formula_strips_workbook_ref():
    wb = _make_workbook({"Hypotheses": {"E5": "test"}})
    result = resolve_formula("=[1]'Hypotheses'!E5", wb)
    assert result == "test"


def test_resolve_formula_none_on_error():
    wb = _make_workbook({})
    result = resolve_formula("=INVALID", wb)
    assert result is None
