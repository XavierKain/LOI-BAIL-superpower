"""Tests for core.excel_parser module."""

import pytest
from unittest.mock import patch, MagicMock
from datetime import datetime
from core.excel_parser import ExcelParser


def test_format_cell_value_string():
    parser = ExcelParser.__new__(ExcelParser)
    assert parser._format_cell_value("  hello  ") == "hello"


def test_format_cell_value_date():
    parser = ExcelParser.__new__(ExcelParser)
    assert parser._format_cell_value(datetime(2026, 4, 1)) == "01/04/2026"


def test_format_cell_value_int():
    parser = ExcelParser.__new__(ExcelParser)
    assert parser._format_cell_value(42) == "42"


def test_format_cell_value_float():
    parser = ExcelParser.__new__(ExcelParser)
    assert parser._format_cell_value(3.14) == "3.14"


def test_format_cell_value_none():
    parser = ExcelParser.__new__(ExcelParser)
    assert parser._format_cell_value(None) is None


def test_format_cell_value_empty_string():
    parser = ExcelParser.__new__(ExcelParser)
    assert parser._format_cell_value("   ") is None


def test_format_cell_value_other_type():
    parser = ExcelParser.__new__(ExcelParser)
    assert parser._format_cell_value(True) == "True"


def test_adds_date_aujourdhui():
    """The parser should add Date d'aujourd'hui = today's date."""
    today = datetime.now().strftime("%d/%m/%Y")
    parser = ExcelParser.__new__(ExcelParser)
    variables = {}
    parser._add_system_variables(variables)
    assert variables["Date d'aujourd'hui"] == today
