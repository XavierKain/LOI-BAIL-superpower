# LOI-BAIL Generator v2 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Complete rewrite of the LOI/BAIL commercial real estate document generator with clean architecture, unified Word engine, shared business logic, and full test coverage.

**Architecture:** 4-layer architecture (core / generators / renderers / UI). Single Word engine for all placeholder replacement. Shared calculations module eliminates all LOI/BAIL duplication. Typed dataclasses replace raw dicts.

**Tech Stack:** Python 3.11+, Streamlit, python-docx, openpyxl, pandas, cloudscraper, beautifulsoup4, python-dateutil, ratelimit, pytest

**Spec:** `docs/superpowers/specs/2026-04-02-loi-bail-rewrite-design.md`

**Reference codebase (read-only):** `/Users/xavier/VSCode3/FA_Baux_LOI_V2a/`

---

## File Structure

| File | Responsibility |
|------|---------------|
| `app.py` | Streamlit UI only — upload, tabs, generate buttons, download |
| `core/__init__.py` | Package init |
| `core/models.py` | Dataclasses: DossierData, InpiData, SocieteInfo, ArticleResult |
| `core/number_to_french.py` | Number-to-French-words conversion |
| `core/formula_engine.py` | Excel formula resolution (RECHERCHEV, ARRONDI, cell refs) |
| `core/excel_parser.py` | Excel file extraction + INPI enrichment trigger |
| `core/inpi_client.py` | INPI API + scraping fallbacks |
| `generators/__init__.py` | Package init |
| `generators/shared.py` | Variable normalization, derived calculations, societe detection, number formatting |
| `generators/loi_generator.py` | LOI-specific logic (optional sections, clear list) |
| `generators/bail_generator.py` | BAIL condition evaluation, article generation, conditions suspensives |
| `renderers/__init__.py` | Package init |
| `renderers/word_engine.py` | Shared Word engine: placeholder replacement, char-map reconstruction, format copy |
| `renderers/loi_renderer.py` | LOI Word rendering: blue sections, headers/footers, special sections |
| `renderers/bail_renderer.py` | BAIL Word rendering: article placeholders, HTML tags, headings, TOC |
| `config/settings.py` | Centralized config, secrets loading |
| `tests/conftest.py` | Shared pytest fixtures |
| `tests/test_number_to_french.py` | Number conversion tests |
| `tests/test_formula_engine.py` | Formula resolution tests |
| `tests/test_shared.py` | Shared calculations tests |
| `tests/test_word_engine.py` | Word engine tests |
| `tests/test_bail_generator.py` | BAIL condition evaluation tests |
| `tests/test_inpi_client.py` | INPI client tests (mocked) |
| `tests/test_integration.py` | End-to-end tests |

---

### Task 1: Project Scaffolding

**Files:**
- Create: `requirements.txt`, `.gitignore`, `.env.example`, `.streamlit/config.toml`, `config/settings.py`, `core/__init__.py`, `generators/__init__.py`, `renderers/__init__.py`, `output/.gitkeep`

- [ ] **Step 1: Create requirements.txt**

```
streamlit>=1.31.0
python-docx>=1.1.0
openpyxl>=3.1.2
pandas>=2.0.0
python-dateutil>=2.8.2
python-dotenv>=1.0.0
requests>=2.31.0
ratelimit>=2.2.1
cloudscraper>=1.2.71
beautifulsoup4>=4.12.0
pytest>=7.4.0
```

- [ ] **Step 2: Create .gitignore**

```
__pycache__/
*.py[cod]
*$py.class
*.so
.Python
env/
venv/
.venv/
build/
dist/
*.egg-info/
*.egg
.vscode/
.idea/
*.swp
*.swo
*~
.DS_Store
Thumbs.db
output/*.docx
!output/.gitkeep
*.log
.env
.streamlit/secrets.toml
```

- [ ] **Step 3: Create .env.example**

```
# INPI API credentials
# Create account at https://data.inpi.fr/login
INPI_USERNAME=votre_email@example.com
INPI_PASSWORD=votre_mot_de_passe
```

- [ ] **Step 4: Create .streamlit/config.toml**

```toml
[client]
showErrorDetails = true
showWarningOnDirectExecution = false

[logger]
level = "info"

[server]
headless = true
port = 8501
enableXsrfProtection = true
maxUploadSize = 200

[theme]
primaryColor = "#1f77b4"
backgroundColor = "#ffffff"
secondaryBackgroundColor = "#f0f2f6"
textColor = "#31333F"
font = "sans serif"
```

- [ ] **Step 5: Create config/settings.py**

```python
"""Centralized configuration. Supports .env (local) and Streamlit Secrets (production)."""

import os
from pathlib import Path
from dotenv import load_dotenv

env_path = Path(__file__).parent.parent / ".env"
load_dotenv(dotenv_path=env_path)


def _get_secret(key: str, default: str = "") -> str:
    """Get secret from Streamlit Secrets (production) or .env (local)."""
    try:
        import streamlit as st
        if hasattr(st, "secrets") and key in st.secrets:
            return st.secrets[key]
    except (ImportError, FileNotFoundError, KeyError):
        pass
    return os.getenv(key, default)


# File paths (relative to project root)
PROJECT_ROOT = Path(__file__).parent.parent
TEMPLATE_LOI = PROJECT_ROOT / "templates" / "Template LOI avec placeholder.docx"
TEMPLATE_BAIL = PROJECT_ROOT / "templates" / "Template BAIL avec placeholder.docx"
CONFIG_LOI = PROJECT_ROOT / "config" / "redaction_loi.xlsx"
CONFIG_BAIL = PROJECT_ROOT / "config" / "redaction_bail.xlsx"
OUTPUT_DIR = PROJECT_ROOT / "output"

# INPI
INPI_BASE_URL = "https://registre-national-entreprises.inpi.fr/api"
INPI_MAX_CALLS = 5
INPI_PERIOD = 60  # seconds
INPI_CACHE_DURATION = 3600  # 1 hour


def get_inpi_credentials() -> dict:
    """Return INPI credentials dict."""
    return {
        "username": _get_secret("INPI_USERNAME", ""),
        "password": _get_secret("INPI_PASSWORD", ""),
    }


def validate_inpi_credentials() -> bool:
    """Check that INPI credentials are configured."""
    creds = get_inpi_credentials()
    return bool(creds["username"] and creds["password"])
```

- [ ] **Step 6: Create package __init__.py files and output/.gitkeep**

Create empty `core/__init__.py`, `generators/__init__.py`, `renderers/__init__.py`, and empty `output/.gitkeep`.

- [ ] **Step 7: Create directories for templates and config data**

```bash
mkdir -p templates config tests
```

- [ ] **Step 8: Copy data files from the old project**

```bash
cp "/Users/xavier/VSCode3/FA_Baux_LOI_V2a/Rédaction LOI.xlsx" config/redaction_loi.xlsx
cp "/Users/xavier/VSCode3/FA_Baux_LOI_V2a/Redaction BAIL.xlsx" config/redaction_bail.xlsx
cp "/Users/xavier/VSCode3/FA_Baux_LOI_V2a/Template LOI avec placeholder.docx" templates/
cp "/Users/xavier/VSCode3/FA_Baux_LOI_V2a/2025 - Template BAIL.docx" "templates/Template BAIL avec placeholder.docx"
```

Note: Check the exact BAIL template filename — old project uses `2025 - Template BAIL.docx` or `Template BAIL avec placeholder.docx`. Use the one with placeholders.

- [ ] **Step 9: Install dependencies**

```bash
pip install -r requirements.txt
```

- [ ] **Step 10: Commit**

```bash
git add requirements.txt .gitignore .env.example .streamlit/config.toml config/settings.py core/__init__.py generators/__init__.py renderers/__init__.py output/.gitkeep templates/ config/redaction_loi.xlsx config/redaction_bail.xlsx
git commit -m "chore: project scaffolding with dependencies, config, and data files"
```

---

### Task 2: Data Models (`core/models.py`)

**Files:**
- Create: `core/models.py`
- Test: `tests/test_models.py`

- [ ] **Step 1: Write the test**

```python
# tests/test_models.py
from pathlib import Path
from core.models import DossierData, InpiData, SocieteInfo, ArticleResult


def test_dossier_data_creation():
    data = DossierData(
        variables={"Nom Preneur": "Test SAS"},
        variables_derivees={"Adresse Locaux Loues": "10 rue de Paris, Paris"},
        inpi_data=None,
        source_file=Path("/tmp/test.xlsx"),
    )
    assert data.variables["Nom Preneur"] == "Test SAS"
    assert data.inpi_data is None


def test_inpi_data_success():
    inpi = InpiData(
        nom_societe="Ma Societe",
        type_societe="SAS",
        capital_social="10 000 €",
        localite_rcs="Paris",
        adresse_domiciliation="10 rue de la Paix, 75002 Paris",
        president="Jean Dupont",
        fonction="Président",
        status="success",
        error_message=None,
    )
    assert inpi.status == "success"
    assert inpi.president == "Jean Dupont"


def test_inpi_data_failed():
    inpi = InpiData(
        nom_societe="",
        type_societe="",
        capital_social="",
        localite_rcs="",
        adresse_domiciliation="",
        president="",
        fonction="",
        status="failed",
        error_message="SIRET invalide",
    )
    assert inpi.status == "failed"
    assert inpi.error_message == "SIRET invalide"


def test_societe_info():
    info = SocieteInfo(nom="Bailleur SCI", header_text="BAILLEUR SCI", footer_text="Ligne 1\nLigne 2")
    assert info.header_text == "BAILLEUR SCI"
    assert "\n" in info.footer_text


def test_article_result():
    article = ArticleResult(
        designation="Article 1",
        contenu="Le bail est conclu pour <b>9 ans</b>.",
        placeholders_manquants=["Montant du loyer"],
    )
    assert article.designation == "Article 1"
    assert len(article.placeholders_manquants) == 1
```

- [ ] **Step 2: Run test to verify it fails**

```bash
pytest tests/test_models.py -v
```

Expected: FAIL — `ModuleNotFoundError: No module named 'core.models'`

- [ ] **Step 3: Write the implementation**

```python
# core/models.py
"""Typed data structures for the LOI-BAIL generator."""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


@dataclass
class InpiData:
    """Company data retrieved from INPI."""
    nom_societe: str
    type_societe: str
    capital_social: str
    localite_rcs: str
    adresse_domiciliation: str
    president: str
    fonction: str
    status: str  # "success" | "failed"
    error_message: Optional[str] = None


@dataclass
class SocieteInfo:
    """Landlord company info for document headers/footers."""
    nom: str
    header_text: str
    footer_text: str


@dataclass
class ArticleResult:
    """Generated article content for BAIL documents."""
    designation: str
    contenu: str
    placeholders_manquants: list[str] = field(default_factory=list)


@dataclass
class DossierData:
    """All extracted variables for a dossier (Excel + INPI + derived)."""
    variables: dict[str, str]
    variables_derivees: dict[str, str]
    inpi_data: Optional[InpiData]
    source_file: Path
```

- [ ] **Step 4: Run test to verify it passes**

```bash
pytest tests/test_models.py -v
```

Expected: 5 passed

- [ ] **Step 5: Commit**

```bash
git add core/models.py tests/test_models.py
git commit -m "feat: add typed data models (DossierData, InpiData, SocieteInfo, ArticleResult)"
```

---

### Task 3: Number to French (`core/number_to_french.py`)

**Files:**
- Create: `core/number_to_french.py`
- Test: `tests/test_number_to_french.py`

- [ ] **Step 1: Write the tests**

```python
# tests/test_number_to_french.py
from core.number_to_french import number_to_french_words


def test_zero():
    assert number_to_french_words(0) == "ZÉRO"


def test_ones():
    assert number_to_french_words(1) == "UN"
    assert number_to_french_words(5) == "CINQ"
    assert number_to_french_words(9) == "NEUF"


def test_teens():
    assert number_to_french_words(10) == "DIX"
    assert number_to_french_words(11) == "ONZE"
    assert number_to_french_words(16) == "SEIZE"
    assert number_to_french_words(19) == "DIX-NEUF"


def test_french_specific_tens():
    # 70 = soixante-dix
    assert number_to_french_words(70) == "SOIXANTE-DIX"
    assert number_to_french_words(71) == "SOIXANTE-ONZE"
    assert number_to_french_words(79) == "SOIXANTE-DIX-NEUF"
    # 80 = quatre-vingts (with S)
    assert number_to_french_words(80) == "QUATRE-VINGTS"
    # 81 = quatre-vingt-un (no S)
    assert number_to_french_words(81) == "QUATRE-VINGT-UN"
    # 90 = quatre-vingt-dix
    assert number_to_french_words(90) == "QUATRE-VINGT-DIX"
    assert number_to_french_words(99) == "QUATRE-VINGT-DIX-NEUF"


def test_et_linking():
    assert number_to_french_words(21) == "VINGT ET UN"
    assert number_to_french_words(31) == "TRENTE ET UN"
    assert number_to_french_words(41) == "QUARANTE ET UN"
    assert number_to_french_words(51) == "CINQUANTE ET UN"
    assert number_to_french_words(61) == "SOIXANTE ET UN"


def test_hundreds():
    assert number_to_french_words(100) == "CENT"
    assert number_to_french_words(200) == "DEUX CENTS"  # with S
    assert number_to_french_words(201) == "DEUX CENT UN"  # no S
    assert number_to_french_words(500) == "CINQ CENTS"


def test_thousands():
    assert number_to_french_words(1000) == "MILLE"
    assert number_to_french_words(2000) == "DEUX MILLE"
    assert number_to_french_words(5000) == "CINQ MILLE"
    assert number_to_french_words(40000) == "QUARANTE MILLE"


def test_real_amounts():
    assert number_to_french_words(160000) == "CENT SOIXANTE MILLE"
    assert number_to_french_words(1234) == "MILLE DEUX CENT TRENTE-QUATRE"


def test_millions():
    assert number_to_french_words(1000000) == "UN MILLION"
    assert number_to_french_words(2000000) == "DEUX MILLIONS"
    assert number_to_french_words(1500000) == "UN MILLION CINQ CENT MILLE"


def test_float_rounds():
    assert number_to_french_words(99.7) == "CENT"
    assert number_to_french_words(99.4) == "QUATRE-VINGT-DIX-NEUF"


def test_negative():
    assert number_to_french_words(-5) == "MOINS CINQ"
```

- [ ] **Step 2: Run test to verify it fails**

```bash
pytest tests/test_number_to_french.py -v
```

Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: Write the implementation**

```python
# core/number_to_french.py
"""Convert numbers to French words in uppercase."""


def number_to_french_words(number: float) -> str:
    """
    Convert a number to French words in uppercase.

    Examples:
        5000 -> "CINQ MILLE"
        160000 -> "CENT SOIXANTE MILLE"
    """
    n = int(round(number))

    if n == 0:
        return "ZÉRO"

    if n < 0:
        return "MOINS " + number_to_french_words(abs(n))

    ones = ["", "UN", "DEUX", "TROIS", "QUATRE", "CINQ", "SIX", "SEPT", "HUIT", "NEUF"]
    teens = [
        "DIX", "ONZE", "DOUZE", "TREIZE", "QUATORZE", "QUINZE", "SEIZE",
        "DIX-SEPT", "DIX-HUIT", "DIX-NEUF",
    ]
    tens = [
        "", "DIX", "VINGT", "TRENTE", "QUARANTE", "CINQUANTE",
        "SOIXANTE", "SOIXANTE", "QUATRE-VINGT", "QUATRE-VINGT",
    ]

    def _below_thousand(n: int) -> str:
        if n == 0:
            return ""
        if n < 10:
            return ones[n]
        if n < 20:
            return teens[n - 10]
        if n < 100:
            t, o = n // 10, n % 10
            if t == 7:
                return "SOIXANTE-" + teens[o]
            if t == 9:
                return "QUATRE-VINGT-" + teens[o]
            if t == 8 and o == 0:
                return "QUATRE-VINGTS"
            if t == 8:
                return "QUATRE-VINGT-" + ones[o]
            if o == 0:
                return tens[t]
            if o == 1:
                return tens[t] + " ET " + ones[o]
            return tens[t] + "-" + ones[o]

        h, rest = n // 100, n % 100
        if h == 1:
            result = "CENT"
        else:
            result = ones[h] + " CENT"
            if rest == 0:
                result += "S"
        if rest > 0:
            result += " " + _below_thousand(rest)
        return result

    parts = []

    if n >= 1_000_000:
        millions = n // 1_000_000
        parts.append("UN MILLION" if millions == 1 else _below_thousand(millions) + " MILLIONS")
        n %= 1_000_000

    if n >= 1000:
        thousands = n // 1000
        parts.append("MILLE" if thousands == 1 else _below_thousand(thousands) + " MILLE")
        n %= 1000

    if n > 0:
        parts.append(_below_thousand(n))

    return " ".join(parts)
```

- [ ] **Step 4: Run test to verify it passes**

```bash
pytest tests/test_number_to_french.py -v
```

Expected: All passed

- [ ] **Step 5: Commit**

```bash
git add core/number_to_french.py tests/test_number_to_french.py
git commit -m "feat: add number-to-French-words conversion"
```

---

### Task 4: Formula Engine (`core/formula_engine.py`)

**Files:**
- Create: `core/formula_engine.py`
- Test: `tests/test_formula_engine.py`

- [ ] **Step 1: Write the tests**

```python
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
        for cell_ref, val in cells.items():
            cell = MagicMock()
            cell.value = val
            ws.__getitem__ = lambda key, _cells=cells: _make_cell_mock(_cells.get(key))
        sheet_mocks[sheet_name] = ws

    wb.__getitem__ = lambda key: sheet_mocks[key]
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
```

- [ ] **Step 2: Run test to verify it fails**

```bash
pytest tests/test_formula_engine.py -v
```

Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: Write the implementation**

```python
# core/formula_engine.py
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


def resolve_cell_reference(sheet_name: str, cell_ref: str, workbook) -> Optional[str]:
    """Resolve a single cell reference to a formatted string value."""
    try:
        ws = workbook[sheet_name]
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

        ws = workbook[sheet_name]
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
```

- [ ] **Step 4: Run test to verify it passes**

```bash
pytest tests/test_formula_engine.py -v
```

Expected: All passed

- [ ] **Step 5: Commit**

```bash
git add core/formula_engine.py tests/test_formula_engine.py
git commit -m "feat: add formula engine (RECHERCHEV, ARRONDI, cell references)"
```

---

### Task 5: Shared Calculations (`generators/shared.py`)

**Files:**
- Create: `generators/shared.py`
- Test: `tests/test_shared.py`

- [ ] **Step 1: Write the tests**

```python
# tests/test_shared.py
import pytest
from generators.shared import (
    normaliser_noms_variables,
    calculer_variables_derivees,
    est_societe,
    formater_nombre,
)
from core.models import InpiData


def test_normaliser_duree_bail():
    variables = {"Duree du Bail": "9"}
    result = normaliser_noms_variables(variables)
    assert result["Duree Bail"] == "9"


def test_normaliser_montant_palier():
    variables = {"Montant Palier 1": "5000"}
    result = normaliser_noms_variables(variables)
    assert result["Montant du palier 1"] == "5000"


def test_normaliser_date_prise_effet():
    variables = {"Date prise d'effet": "01/01/2026"}
    result = normaliser_noms_variables(variables)
    assert result["Date de prise d'effet"] == "01/01/2026"


def test_normaliser_typo_gapd():
    variables = {"Dure GAPD": "6"}
    result = normaliser_noms_variables(variables)
    assert result["Duree GAPD"] == "6"


def test_normaliser_case_insensitive_dedup():
    variables = {"Nom Preneur": "Test", "nom preneur": ""}
    result = normaliser_noms_variables(variables)
    assert result["nom preneur"] == "Test"


def test_adresse_rue_ville():
    variables = {"Numero et rue": "10 rue de la Paix", "Ville ou arrondissement": "Paris 2ème"}
    result = calculer_variables_derivees(variables, None)
    assert result["Adresse Locaux Loues"] == "10 rue de la Paix, Paris 2ème"


def test_adresse_only_rue():
    variables = {"Numero et rue": "10 rue de la Paix"}
    result = calculer_variables_derivees(variables, None)
    assert result["Adresse Locaux Loues"] == "10 rue de la Paix"


def test_paliers():
    variables = {"Montant du loyer": "160 000", "Loyer annee 1": "140000"}
    result = calculer_variables_derivees(variables, None)
    assert result["Montant du palier 1"] == "20 000"


def test_type_bail_9():
    variables = {"Duree Bail": "9"}
    result = calculer_variables_derivees(variables, None)
    assert result["Type Bail"] == "3/6/9"


def test_type_bail_10():
    variables = {"Duree Bail": "10"}
    result = calculer_variables_derivees(variables, None)
    assert result["Type Bail"] == "6/9/10"


def test_type_bail_other():
    variables = {"Duree Bail": "12"}
    result = calculer_variables_derivees(variables, None)
    assert result["Type Bail"] == "12 ans"


def test_date_signature():
    variables = {"Date d'aujourd'hui": "01/04/2026"}
    result = calculer_variables_derivees(variables, None)
    assert result["Date de signature"] == "22/04/2026"


def test_date_offre_valable():
    variables = {"Date d'aujourd'hui": "01/04/2026"}
    result = calculer_variables_derivees(variables, None)
    assert result["Date offre valable"] == "08/04/2026"


def test_date_prise_effet_plus_9_ans():
    variables = {"Date de prise d'effet": "01/06/2026"}
    result = calculer_variables_derivees(variables, None)
    # relativedelta(years=9) -> 01/06/2035
    assert result["Date de prise d'effet + 9 ans"] == "01/06/2035"


def test_surface_r_moins_1():
    variables = {"Surface totale": "200", "Surface RDC": "150"}
    result = calculer_variables_derivees(variables, None)
    assert result["Surface R-1"] == "50"


def test_montant_dg():
    variables = {"Montant du loyer": "120000", "Duree DG": "3"}
    result = calculer_variables_derivees(variables, None)
    assert result["Montant du DG"] == "30 000"


def test_periode_dg():
    variables = {"Duree DG": "3"}
    result = calculer_variables_derivees(variables, None)
    assert result["Periode DG"] == "quart"


def test_periode_dg_6():
    variables = {"Duree DG": "6"}
    result = calculer_variables_derivees(variables, None)
    assert result["Periode DG"] == "moitié"


def test_est_societe_sas():
    assert est_societe("SAS") is True


def test_est_societe_sarl():
    assert est_societe("SARL") is True


def test_est_societe_sasu():
    assert est_societe("SASU") is True


def test_est_societe_individual():
    assert est_societe("Personne physique") is False


def test_est_societe_passage_not_matched():
    assert est_societe("PASSAGE") is False


def test_est_societe_case_insensitive():
    assert est_societe("sas") is True
    assert est_societe("Sarl") is True


def test_formater_nombre_entier():
    assert formater_nombre(160000) == "160 000"


def test_formater_nombre_decimal():
    assert formater_nombre(1234.56) == "1 234,56"


def test_formater_nombre_string():
    assert formater_nombre("160000") == "160 000"


def test_formater_nombre_with_spaces():
    assert formater_nombre("160 000") == "160 000"


def test_inpi_data_merged():
    inpi = InpiData(
        nom_societe="Test SAS",
        type_societe="SAS",
        capital_social="10 000 €",
        localite_rcs="Paris",
        adresse_domiciliation="10 rue de la Paix",
        president="Jean Dupont",
        fonction="Président",
        status="success",
        error_message=None,
    )
    result = calculer_variables_derivees({}, inpi)
    assert result["NOM DE LA SOCIETE"] == "Test SAS"
    assert result["PRESIDENT DE LA SOCIETE"] == "Jean Dupont"
    assert result["FONCTION INPI"] == "Président"
```

- [ ] **Step 2: Run test to verify it fails**

```bash
pytest tests/test_shared.py -v
```

Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: Write the implementation**

```python
# generators/shared.py
"""Shared calculations for LOI and BAIL generators.

Single source of truth for: variable normalization, derived calculations,
societe detection, number formatting.
"""

import logging
import re
from datetime import datetime
from typing import Optional

from dateutil.relativedelta import relativedelta

from core.models import InpiData
from core.number_to_french import number_to_french_words

logger = logging.getLogger(__name__)

# Centralized variable name normalization mapping
_VARIABLE_ALIASES: dict[str, str] = {
    "Duree du Bail": "Duree Bail",
    "Durée du Bail": "Duree Bail",
    "Duree du DG": "Duree DG",
    "Durée du DG": "Duree DG",
    "Durée DG": "Duree DG",
    "Date prise d'effet": "Date de prise d'effet",
    "Date de prise d'effet du bail": "Date de prise d'effet",
    "Date debut bail": "Date de prise d'effet",
    "Date de Prise d'effet + 9 ans": "Date de prise d'effet + 9 ans",
    "Statut Locaux loués": "Statut Locaux Loués",
    "Statut Locaux loues": "Statut Locaux Loués",
    "Durée ferme bail": "Durée ferme Bail",
    "Duree ferme bail": "Durée ferme Bail",
    "Dure GAPD": "Duree GAPD",
    "Duré GAPD": "Duree GAPD",
}

# Add palier aliases for 1-6
for _i in range(1, 7):
    _VARIABLE_ALIASES[f"Montant Palier {_i}"] = f"Montant du palier {_i}"
    _VARIABLE_ALIASES[f"Montant du Palier {_i}"] = f"Montant du palier {_i}"
    _VARIABLE_ALIASES[f"Montant palier {_i}"] = f"Montant du palier {_i}"

# Legal forms for societe detection (exact match, case-insensitive)
_FORMES_JURIDIQUES = {
    "SAS", "SARL", "EURL", "SA", "SCI", "SNC", "SASU",
    "SOCIÉTÉ", "SOCIETE", "SOCIÉTÉ EN FORMATION", "SOCIETE EN FORMATION",
}


def normaliser_noms_variables(variables: dict[str, str]) -> dict[str, str]:
    """Normalize variable names using the centralized alias mapping.

    Also performs case-insensitive deduplication: if two keys match
    case-insensitively and one has a value while the other is empty,
    the value is copied to the empty key.
    """
    result = dict(variables)

    # Apply alias mapping
    for old_name, new_name in _VARIABLE_ALIASES.items():
        if old_name in result and new_name not in result:
            result[new_name] = result[old_name]
        elif old_name in result and new_name in result and not result[new_name]:
            result[new_name] = result[old_name]

    # Case-insensitive deduplication
    lower_map: dict[str, list[str]] = {}
    for key in result:
        lower_map.setdefault(key.lower(), []).append(key)

    for keys in lower_map.values():
        if len(keys) > 1:
            # Find the one with a value
            value = ""
            for k in keys:
                if result.get(k):
                    value = result[k]
                    break
            # Copy to all empty ones
            if value:
                for k in keys:
                    if not result.get(k):
                        result[k] = value

    return result


def _clean_number(value: str) -> Optional[float]:
    """Parse a number from string, handling French formatting (spaces, commas)."""
    if not value:
        return None
    try:
        cleaned = str(value).strip().replace(" ", "").replace("\u00a0", "").replace(",", ".")
        return float(cleaned)
    except (ValueError, TypeError):
        return None


def formater_nombre(value) -> str:
    """Format a number with French conventions: space thousands, comma decimals.

    160000 -> "160 000"
    1234.56 -> "1 234,56"
    Strips trailing ,00.
    """
    num = _clean_number(str(value)) if not isinstance(value, (int, float)) else float(value)
    if num is None:
        return str(value)

    formatted = f"{num:,.2f}".replace(",", " ").replace(".", ",")
    # Strip ,00 suffix
    if formatted.endswith(",00"):
        formatted = formatted[:-3]
    return formatted


def est_societe(type_preneur: str) -> bool:
    """Check if the tenant type is a company (exact match on legal forms)."""
    if not type_preneur:
        return False
    upper = type_preneur.strip().upper()
    return upper in {f.upper() for f in _FORMES_JURIDIQUES}


def calculer_variables_derivees(
    variables: dict[str, str],
    inpi_data: Optional[InpiData],
) -> dict[str, str]:
    """Calculate all derived variables from raw data. Called once, shared by LOI and BAIL."""
    result: dict[str, str] = {}

    # --- INPI data ---
    if inpi_data and inpi_data.status == "success":
        result["NOM DE LA SOCIETE"] = inpi_data.nom_societe
        result["TYPE DE SOCIETE"] = inpi_data.type_societe
        result["CAPITAL SOCIAL"] = inpi_data.capital_social
        result["LOCALITE RCS"] = inpi_data.localite_rcs
        result["ADRESSE DE DOMICILIATION"] = inpi_data.adresse_domiciliation
        result["PRESIDENT DE LA SOCIETE"] = inpi_data.president
        result["FONCTION INPI"] = inpi_data.fonction

    # --- Address: "rue, ville" ---
    rue = variables.get("Numero et rue", "").strip()
    ville = variables.get("Ville ou arrondissement", "").strip()
    if rue and ville:
        result["Adresse Locaux Loues"] = f"{rue}, {ville}"
    elif rue:
        result["Adresse Locaux Loues"] = rue
    elif ville:
        result["Adresse Locaux Loues"] = ville

    # --- Paliers ---
    loyer_base = _clean_number(variables.get("Montant du loyer", ""))
    if loyer_base:
        for i in range(1, 7):
            loyer_annee = _clean_number(variables.get(f"Loyer annee {i}", ""))
            if loyer_annee is not None:
                remise = loyer_base - loyer_annee
                if remise > 0:
                    result[f"Montant du palier {i}"] = formater_nombre(int(remise))

    # --- Surfaces ---
    surface_totale = _clean_number(variables.get("Surface totale", ""))
    surface_rdc = _clean_number(variables.get("Surface RDC", ""))
    if surface_totale is not None and surface_rdc is not None:
        result["Surface R-1"] = str(int(surface_totale - surface_rdc))

    # --- Type Bail ---
    duree_bail = _clean_number(variables.get("Duree Bail", ""))
    if duree_bail is not None:
        d = int(duree_bail)
        if d == 9:
            result["Type Bail"] = "3/6/9"
        elif d == 10:
            result["Type Bail"] = "6/9/10"
        else:
            result["Type Bail"] = f"{d} ans"

    # --- Dates ---
    date_aujourdhui_str = variables.get("Date d'aujourd'hui", "")
    if date_aujourdhui_str:
        try:
            date_aujourdhui = datetime.strptime(date_aujourdhui_str.strip(), "%d/%m/%Y")
            result["Date de signature"] = (date_aujourdhui + relativedelta(days=21)).strftime("%d/%m/%Y")
            result["Date offre valable"] = (date_aujourdhui + relativedelta(days=7)).strftime("%d/%m/%Y")
        except ValueError:
            logger.warning(f"Cannot parse Date d'aujourd'hui: {date_aujourdhui_str}")

    # Date de prise d'effet + 9 ans (using relativedelta for calendar accuracy)
    date_prise_str = variables.get("Date de prise d'effet", "")
    if date_prise_str:
        for fmt in ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d.%m.%Y"]:
            try:
                date_prise = datetime.strptime(str(date_prise_str).strip(), fmt)
                date_fin = date_prise + relativedelta(years=9)
                result["Date de prise d'effet + 9 ans"] = date_fin.strftime("%d/%m/%Y")
                result["Date de Prise d'effet + 9 ans"] = date_fin.strftime("%d/%m/%Y")
                break
            except ValueError:
                continue

    # --- Montant du DG ---
    duree_dg = _clean_number(variables.get("Duree DG", ""))
    if loyer_base and duree_dg:
        montant_dg = (loyer_base / 12) * duree_dg
        result["Montant du DG"] = formater_nombre(int(montant_dg))

    # --- Periode DG ---
    if duree_dg:
        periode_map = {3: "quart", 4: "tiers", 6: "moitié"}
        result["Periode DG"] = periode_map.get(int(duree_dg), "")

    return result
```

- [ ] **Step 4: Run test to verify it passes**

```bash
pytest tests/test_shared.py -v
```

Expected: All passed

- [ ] **Step 5: Commit**

```bash
git add generators/shared.py tests/test_shared.py
git commit -m "feat: add shared calculations (normalization, derived values, societe detection)"
```

---

### Task 6: Word Engine (`renderers/word_engine.py`)

**Files:**
- Create: `renderers/word_engine.py`
- Test: `tests/test_word_engine.py`

- [ ] **Step 1: Write the tests**

```python
# tests/test_word_engine.py
import pytest
from docx import Document
from docx.shared import RGBColor, Pt
from renderers.word_engine import WordEngine


@pytest.fixture
def engine():
    return WordEngine()


def _make_paragraph(doc, text, color=None, bold=None, font_name=None, font_size=None):
    """Helper: create a paragraph with a single run."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    if color:
        run.font.color.rgb = color
    if bold is not None:
        run.font.bold = bold
    if font_name:
        run.font.name = font_name
    if font_size:
        run.font.size = font_size
    return p


def _make_fragmented_paragraph(doc, fragments):
    """Helper: create a paragraph with multiple runs (fragmented placeholder).
    fragments = [("text", {optional font attrs}), ...]
    """
    p = doc.add_paragraph()
    for text, attrs in fragments:
        run = p.add_run(text)
        for k, v in attrs.items():
            if k == "bold":
                run.font.bold = v
            elif k == "name":
                run.font.name = v
            elif k == "size":
                run.font.size = v
    return p


def test_simple_replacement(engine):
    doc = Document()
    p = _make_paragraph(doc, "Le preneur [Nom Preneur] signe.")
    variables = {"Nom Preneur": "Test SAS"}
    result = engine.replace_placeholders(p, variables)
    assert result is None  # no deletion
    assert "Test SAS" in p.text
    assert "[Nom Preneur]" not in p.text


def test_missing_placeholder_turns_red(engine):
    doc = Document()
    p = _make_paragraph(doc, "Montant: [Montant du loyer]")
    variables = {}
    engine.replace_placeholders(p, variables)
    # The placeholder should still be there
    assert "[Montant du loyer]" in p.text
    # At least one run should be red
    has_red = any(
        run.font.color.rgb == RGBColor(255, 0, 0)
        for run in p.runs
        if run.font.color.rgb is not None
    )
    assert has_red


def test_fragmented_placeholder_replacement(engine):
    doc = Document()
    p = _make_fragmented_paragraph(doc, [
        ("Le preneur [Nom", {}),
        (" Preneur] signe.", {}),
    ])
    variables = {"Nom Preneur": "Test SAS"}
    engine.replace_placeholders(p, variables)
    assert "Test SAS" in p.text
    assert "[" not in p.text


def test_optional_blue_paragraph_deleted_when_no_data(engine):
    doc = Document()
    p = _make_paragraph(doc, "Optionnel: [Variable rare]", color=RGBColor(0, 0, 255))
    variables = {}
    result = engine.replace_placeholders(p, variables)
    assert result == "delete"


def test_optional_blue_paragraph_kept_when_data(engine):
    doc = Document()
    p = _make_paragraph(doc, "Optionnel: [Variable rare]", color=RGBColor(0, 0, 255))
    variables = {"Variable rare": "valeur"}
    result = engine.replace_placeholders(p, variables)
    assert result is None
    assert "valeur" in p.text
    # Runs should now be black (blue removed)
    for run in p.runs:
        if run.font.color.rgb is not None:
            assert run.font.color.rgb == RGBColor(0, 0, 0)


def test_clear_list_placeholder_replaced_with_empty(engine):
    doc = Document()
    p = _make_paragraph(doc, "President: [PRESIDENT DE LA SOCIETE]")
    variables = {"PRESIDENT DE LA SOCIETE": "Jean Dupont"}
    clear_list = ["PRESIDENT DE LA SOCIETE"]
    engine.replace_placeholders(p, variables, clear_list=clear_list)
    assert "Jean Dupont" not in p.text
    assert "[PRESIDENT DE LA SOCIETE]" not in p.text


def test_is_paragraph_optional_blue(engine):
    doc = Document()
    p = _make_paragraph(doc, "Blue text", color=RGBColor(0, 0, 200))
    assert engine.is_paragraph_optional(p) is True


def test_is_paragraph_optional_black(engine):
    doc = Document()
    p = _make_paragraph(doc, "Black text", color=RGBColor(0, 0, 0))
    assert engine.is_paragraph_optional(p) is False


def test_is_paragraph_optional_no_color(engine):
    doc = Document()
    p = _make_paragraph(doc, "No color")
    assert engine.is_paragraph_optional(p) is False


def test_format_preservation_simple(engine):
    doc = Document()
    p = _make_paragraph(doc, "Bold: [Nom]", bold=True, font_name="Arial", font_size=Pt(14))
    variables = {"Nom": "Test"}
    engine.replace_placeholders(p, variables)
    assert p.runs[0].font.bold is True
    assert p.runs[0].font.name == "Arial"


def test_en_lettres_placeholder(engine):
    doc = Document()
    p = _make_paragraph(doc, "Montant: [Montant du loyer en lettres]")
    variables = {"Montant du loyer": "160000"}
    engine.replace_placeholders(p, variables)
    assert "CENT SOIXANTE MILLE" in p.text


def test_multiple_placeholders_in_one_paragraph(engine):
    doc = Document()
    p = _make_paragraph(doc, "[Nom Preneur] signe le [Date LOI].")
    variables = {"Nom Preneur": "Test SAS", "Date LOI": "01/04/2026"}
    engine.replace_placeholders(p, variables)
    assert "Test SAS" in p.text
    assert "01/04/2026" in p.text


def test_process_document_body(engine):
    doc = Document()
    doc.add_paragraph("Hello [Nom]")
    doc.add_paragraph("World [Date]")
    variables = {"Nom": "Test", "Date": "2026"}
    engine.process_document_body(doc, variables)
    assert "Test" in doc.paragraphs[0].text
    assert "2026" in doc.paragraphs[1].text
```

- [ ] **Step 2: Run test to verify it fails**

```bash
pytest tests/test_word_engine.py -v
```

Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: Write the implementation**

```python
# renderers/word_engine.py
"""Shared Word document engine for placeholder replacement and formatting.

Handles both [Variable] and {{ARTICLE}} patterns. Single implementation
for char-map reconstruction, format copying, and optional section detection.
"""

import logging
import re
from typing import Optional

from docx.shared import RGBColor
from docx.oxml.ns import qn

from core.number_to_french import number_to_french_words

logger = logging.getLogger(__name__)

_PLACEHOLDER_RE = re.compile(r"\[([^\]]+)\]")


class WordEngine:
    """Unified Word engine for placeholder replacement with format preservation."""

    def copy_run_format(self, source, target, override_color=None):
        """Copy font formatting from source run to target run.

        Args:
            source: Source run to copy formatting from.
            target: Target run to apply formatting to.
            override_color: If set, use this RGBColor instead of source color.
        """
        for attr in ("name", "size", "bold", "italic", "underline", "strike",
                      "subscript", "superscript", "all_caps", "small_caps"):
            val = getattr(source.font, attr, None)
            if val is not None:
                setattr(target.font, attr, val)

        if override_color is not None:
            target.font.color.rgb = override_color
        else:
            # Only copy explicit RGB colors (type == 1), not theme colors
            try:
                if source.font.color.type is not None and source.font.color.type == 1:
                    target.font.color.rgb = source.font.color.rgb
            except (AttributeError, TypeError):
                pass

    def is_paragraph_optional(self, paragraph) -> bool:
        """Check if a paragraph is optional (has blue-colored text).

        A paragraph is optional if ANY run has an explicit RGB color
        where blue > red AND blue > green.
        """
        for run in paragraph.runs:
            try:
                if run.font.color.type is not None and run.font.color.type == 1:
                    rgb = run.font.color.rgb
                    if rgb and rgb[2] > rgb[0] and rgb[2] > rgb[1]:
                        return True
            except (AttributeError, TypeError):
                continue
        return False

    def _resolve_placeholder_value(self, name: str, variables: dict, clear_list: Optional[list] = None) -> tuple[Optional[str], bool]:
        """Resolve a placeholder name to its value.

        Returns:
            (value, should_clear): value is the replacement string, should_clear
            means the placeholder should be replaced with empty string.
        """
        # Check clear list first
        if clear_list and name in clear_list:
            return "", False

        # Handle "en lettres" suffix
        if name.endswith(" en lettres"):
            base_name = name[:-len(" en lettres")]
            base_value = self._lookup_variable(base_name, variables)
            if base_value:
                try:
                    num = float(str(base_value).replace(" ", "").replace(",", ".").replace("\u00a0", ""))
                    return number_to_french_words(num) + " ", False
                except (ValueError, TypeError):
                    pass
            return None, False

        value = self._lookup_variable(name, variables)
        if value is not None and value != "":
            return str(value), False
        return None, False

    def _lookup_variable(self, name: str, variables: dict) -> Optional[str]:
        """Look up a variable by name with case-insensitive fallback."""
        # Exact match
        if name in variables:
            val = variables[name]
            return str(val) if val is not None and str(val).strip() else None

        # Case-insensitive fallback
        name_lower = name.lower()
        for key, val in variables.items():
            if key.lower() == name_lower:
                return str(val) if val is not None and str(val).strip() else None
        return None

    def replace_placeholders(self, paragraph, variables: dict, clear_list: Optional[list] = None) -> Optional[str]:
        """Replace [Variable] placeholders in a paragraph.

        Returns:
            "delete" if the paragraph is optional and has no data.
            None otherwise.
        """
        full_text = paragraph.text
        matches = _PLACEHOLDER_RE.findall(full_text)
        if not matches:
            return None

        is_optional = self.is_paragraph_optional(paragraph)

        # Check which placeholders have values
        resolved = {}
        all_empty = True
        for name in matches:
            value, _ = self._resolve_placeholder_value(name, variables, clear_list)
            resolved[name] = value
            if value is not None and value != "":
                all_empty = False

        # Optional paragraph with no data -> delete
        if is_optional and all_empty:
            return "delete"

        # Optional paragraph with data -> turn all runs black
        if is_optional:
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(0, 0, 0)

        # Check if any placeholder is fragmented across runs
        has_fragmented = False
        for name in matches:
            placeholder = f"[{name}]"
            found_in_run = any(placeholder in run.text for run in paragraph.runs)
            if not found_in_run:
                has_fragmented = True
                break

        # Also check if we have missing placeholders (need char-map for red coloring)
        has_missing = any(v is None for v in resolved.values())

        if not has_fragmented and not has_missing:
            # Simple path: all placeholders are in single runs and all have values
            for run in paragraph.runs:
                for name, value in resolved.items():
                    placeholder = f"[{name}]"
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, value if value is not None else "")
            return None

        # Complex path: char-map reconstruction
        self._reconstruct_paragraph(paragraph, matches, resolved, is_optional)
        return None

    def _reconstruct_paragraph(self, paragraph, placeholder_names: list, resolved: dict, was_optional: bool):
        """Reconstruct paragraph using char-map algorithm for fragmented/missing placeholders."""
        full_text = paragraph.text
        runs = list(paragraph.runs)

        if not runs:
            return

        # Build char_to_run_map
        char_to_run_map = []
        for run in runs:
            for _ in run.text:
                char_to_run_map.append(run)

        # Build segments
        segments = []  # (text, source_run, is_missing)
        i = 0
        while i < len(full_text):
            # Check if a placeholder starts here
            placeholder_match = None
            for name in placeholder_names:
                placeholder = f"[{name}]"
                if full_text[i:i + len(placeholder)] == placeholder:
                    placeholder_match = (name, placeholder)
                    break

            if placeholder_match:
                name, placeholder = placeholder_match
                value = resolved.get(name)
                source_run = char_to_run_map[i] if i < len(char_to_run_map) else runs[-1]

                if value is not None:
                    segments.append((value, source_run, False))
                else:
                    # Missing: keep placeholder text, mark as missing
                    segments.append((placeholder, source_run, True))

                i += len(placeholder)
            else:
                # Regular text: accumulate until run boundary or next placeholder
                start = i
                current_run = char_to_run_map[i] if i < len(char_to_run_map) else runs[-1]
                while i < len(full_text):
                    # Check for placeholder
                    is_placeholder = False
                    for name in placeholder_names:
                        if full_text[i:].startswith(f"[{name}]"):
                            is_placeholder = True
                            break
                    if is_placeholder:
                        break
                    # Check run boundary
                    next_run = char_to_run_map[i] if i < len(char_to_run_map) else runs[-1]
                    if next_run != current_run:
                        break
                    i += 1
                text = full_text[start:i]
                segments.append((text, current_run, False))

        # Clear all existing runs
        for run in runs:
            run.text = ""

        # Rebuild with new runs
        for text, source_run, is_missing in segments:
            if not text:
                continue
            new_run = paragraph.add_run(text)
            if is_missing:
                self.copy_run_format(source_run, new_run, override_color=RGBColor(255, 0, 0))
            elif was_optional:
                # Optional paragraph kept: force black to remove blue
                self.copy_run_format(source_run, new_run, override_color=RGBColor(0, 0, 0))
            else:
                self.copy_run_format(source_run, new_run)

    def process_document_body(self, document, variables: dict, clear_list: Optional[list] = None) -> list:
        """Process all paragraphs and table cells. Returns list of paragraphs to delete."""
        to_delete = []

        for paragraph in document.paragraphs:
            result = self.replace_placeholders(paragraph, variables, clear_list)
            if result == "delete":
                to_delete.append(paragraph)

        # Process tables
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self.replace_placeholders(paragraph, variables, clear_list)

        return to_delete

    @staticmethod
    def delete_paragraphs(paragraphs: list):
        """Remove paragraphs from the document by deleting their XML elements."""
        for p in paragraphs:
            element = p._element
            parent = element.getparent()
            if parent is not None:
                parent.remove(element)
```

- [ ] **Step 4: Run test to verify it passes**

```bash
pytest tests/test_word_engine.py -v
```

Expected: All passed

- [ ] **Step 5: Commit**

```bash
git add renderers/word_engine.py tests/test_word_engine.py
git commit -m "feat: add shared Word engine (placeholder replacement, char-map, format copy)"
```

---

### Task 7: INPI Client (`core/inpi_client.py`)

**Files:**
- Create: `core/inpi_client.py`
- Test: `tests/test_inpi_client.py`

- [ ] **Step 1: Write the tests**

```python
# tests/test_inpi_client.py
import pytest
from unittest.mock import patch, MagicMock
from core.inpi_client import INPIClient, _format_dirigeant_name, _extract_siren
from core.models import InpiData


def test_extract_siren_from_siret():
    assert _extract_siren("12345678901234") == "123456789"


def test_extract_siren_from_siren():
    assert _extract_siren("123456789") == "123456789"


def test_extract_siren_with_spaces():
    assert _extract_siren("123 456 789 01234") == "123456789"


def test_extract_siren_invalid():
    assert _extract_siren("12345") is None


def test_extract_siren_empty():
    assert _extract_siren("") is None


def test_format_name_all_uppercase():
    assert _format_dirigeant_name("DUPONT", "JEAN") == "Jean Dupont"


def test_format_name_mixed_case():
    assert _format_dirigeant_name("Dupont", "Jean") == "Jean Dupont"


def test_format_name_no_prenom():
    assert _format_dirigeant_name("DUPONT", "") == "Dupont"


def test_get_company_info_empty_siret():
    client = INPIClient(username="test", password="test")
    result = client.get_company_info("")
    assert result.status == "failed"
    assert "manquant" in result.error_message.lower() or "invalide" in result.error_message.lower()


def test_get_company_info_invalid_siret():
    client = INPIClient(username="test", password="test")
    result = client.get_company_info("123")
    assert result.status == "failed"


@patch("core.inpi_client.requests.post")
@patch("core.inpi_client.requests.get")
def test_api_success(mock_get, mock_post):
    # Mock auth
    mock_post.return_value = MagicMock(
        status_code=200,
        json=lambda: {"token": "fake-token"},
    )
    # Mock search
    mock_get.return_value = MagicMock(
        status_code=200,
        json=lambda: {
            "results": [{
                "formality": {
                    "content": {
                        "personneMorale": {
                            "identite": {
                                "entreprise": {
                                    "denomination": "TEST SAS",
                                }
                            },
                            "etablissementPrincipal": {
                                "descriptionEtablissement": {},
                                "adresse": {
                                    "commune": "PARIS",
                                    "codePostal": "75002",
                                    "voie": "10 RUE DE LA PAIX",
                                }
                            },
                            "composition": {
                                "pouvoirs": [{
                                    "roleEntreprise": "30",
                                    "typeDePersonne": "INDIVIDU",
                                    "actif": True,
                                    "individu": {
                                        "descriptionPersonne": {
                                            "nom": "DUPONT",
                                            "prenoms": ["JEAN"],
                                        }
                                    }
                                }]
                            }
                        },
                        "natureCreation": {
                            "formeJuridique": "SAS",
                        }
                    }
                }
            }]
        },
    )

    client = INPIClient(username="test", password="test")
    result = client.get_company_info("123456789")
    assert result.status == "success"
    assert result.nom_societe == "TEST SAS"
    assert result.president == "Jean Dupont"
    assert result.fonction == "Président"
```

- [ ] **Step 2: Run test to verify it fails**

```bash
pytest tests/test_inpi_client.py -v
```

Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: Write the implementation**

This is a large file. The implementation closely follows the old `modules/inpi_client.py` from `/Users/xavier/VSCode3/FA_Baux_LOI_V2a/modules/inpi_client.py` but with these changes:
- Returns `InpiData` instead of raw dict
- Unified `"Prenom Nom"` format everywhere (the old API path returned `"Nom Prenom"`)
- No `lru_cache` (caching handled by Streamlit)
- Helper functions `_format_dirigeant_name` and `_extract_siren` extracted as module-level for testability

Reference the old file at `/Users/xavier/VSCode3/FA_Baux_LOI_V2a/modules/inpi_client.py` for the complete scraping logic (BeautifulSoup HTML parsing, Playwright fallback). Adapt the `get_company_info` return type to `InpiData`.

Key structure:
```python
# core/inpi_client.py
"""INPI client: API + scraping fallbacks for French company data."""

import logging
import re
import time
import requests
from typing import Optional
from ratelimit import limits, sleep_and_retry

from core.models import InpiData
from config.settings import INPI_BASE_URL, INPI_MAX_CALLS, INPI_PERIOD

# Conditional imports for scraping
try:
    import cloudscraper
    from bs4 import BeautifulSoup
    SCRAPING_AVAILABLE = True
except ImportError:
    SCRAPING_AVAILABLE = False

try:
    from playwright.sync_api import sync_playwright
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False

logger = logging.getLogger(__name__)

ROLES_LIBELLES = {
    "30": "Président",
    "71": "Président",
    "50": "Gérant",
    "10": "Directeur général",
}

ROLES_DIRIGEANTS = ["30", "71", "50"]


def _extract_siren(siret: str) -> Optional[str]:
    """Extract 9-digit SIREN from SIRET or SIREN string."""
    if not siret:
        return None
    cleaned = str(siret).replace(" ", "").strip()
    if len(cleaned) == 14:
        return cleaned[:9]
    if len(cleaned) == 9:
        return cleaned
    return None


def _format_dirigeant_name(nom: str, prenom: str = "") -> str:
    """Format dirigeant name as 'Prenom Nom'."""
    nom_fmt = nom.capitalize() if nom.isupper() else nom
    if not prenom:
        return nom_fmt
    prenom_fmt = prenom.capitalize() if prenom.isupper() else prenom
    return f"{prenom_fmt} {nom_fmt}"


class INPIClient:
    """Client for INPI company registry (API + scraping fallbacks)."""

    def __init__(self, username: str = "", password: str = ""):
        self.base_url = INPI_BASE_URL
        self.username = username
        self.password = password
        self.token = None
        self._token_expiry = 0

    def _authenticate(self) -> bool:
        """Authenticate with INPI API, cache token for 1 hour."""
        if not self.username or not self.password:
            return False
        if self.token and time.time() < self._token_expiry:
            return True
        try:
            response = requests.post(
                f"{self.base_url}/sso/login",
                json={"username": self.username, "password": self.password},
                headers={"Content-Type": "application/json"},
                timeout=10,
            )
            if response.status_code == 200:
                self.token = response.json().get("token")
                self._token_expiry = time.time() + 3600
                return True
            logger.error(f"INPI auth failed: {response.status_code}")
            return False
        except Exception as e:
            logger.error(f"INPI auth error: {e}")
            return False

    @sleep_and_retry
    @limits(calls=INPI_MAX_CALLS, period=INPI_PERIOD)
    def _api_search(self, siren: str) -> Optional[dict]:
        """Search company by SIREN via INPI API."""
        if not self._authenticate():
            return None
        try:
            response = requests.get(
                f"{self.base_url}/companies",
                params={"siren[]": siren},
                headers={"Authorization": f"Bearer {self.token}"},
                timeout=15,
            )
            if response.status_code == 200:
                results = response.json().get("results", [])
                return results[0] if results else None
            return None
        except Exception as e:
            logger.error(f"INPI API search error: {e}")
            return None

    def _extract_from_api(self, company_data: dict) -> InpiData:
        """Extract InpiData from API response."""
        formality = company_data.get("formality", {})
        content = formality.get("content", {})
        pm = content.get("personneMorale", {})
        nature = content.get("natureCreation", {})

        identite = pm.get("identite", {})
        entreprise = identite.get("entreprise", {})
        etab = pm.get("etablissementPrincipal", {})
        desc_etab = etab.get("descriptionEtablissement", {})
        adresse = etab.get("adresse", {})

        nom = (
            desc_etab.get("nomCommercial")
            or desc_etab.get("enseigne")
            or pm.get("denomination")
            or entreprise.get("denomination")
            or ""
        )

        forme = nature.get("formeJuridique", "")
        # Map code to name if needed
        # (old code had FORMES_JURIDIQUES_MAP, simplified here)

        # Capital
        capital_str = ""
        capital = entreprise.get("capital", {})
        if isinstance(capital, dict):
            montant = capital.get("montant")
            if montant:
                capital_str = f"{int(float(montant)):,}".replace(",", " ") + " €"

        # Address
        rue = adresse.get("voie", "")
        cp = adresse.get("codePostal", "")
        commune = adresse.get("commune", "")
        adresse_str = f"{rue}, {cp} {commune}".strip(", ")

        # RCS locality
        localite = commune
        # Strip arrondissement for Paris/Lyon/Marseille
        for city in ["PARIS", "LYON", "MARSEILLE"]:
            if city in localite.upper():
                localite = city
                break

        # Dirigeant
        president = ""
        fonction = ""
        composition = pm.get("composition", {})
        for pouvoir in composition.get("pouvoirs", []):
            role = pouvoir.get("roleEntreprise")
            if (pouvoir.get("actif") and role in ROLES_DIRIGEANTS
                    and pouvoir.get("typeDePersonne") == "INDIVIDU"):
                individu = pouvoir.get("individu", {})
                desc = individu.get("descriptionPersonne", {})
                nom_dir = desc.get("nom", "")
                prenoms = desc.get("prenoms", [])
                prenom_dir = prenoms[0] if prenoms else ""
                if nom_dir:
                    president = _format_dirigeant_name(nom_dir, prenom_dir)
                    fonction = ROLES_LIBELLES.get(role, "Dirigeant")
                    break

        return InpiData(
            nom_societe=nom,
            type_societe=forme,
            capital_social=capital_str,
            localite_rcs=localite,
            adresse_domiciliation=adresse_str,
            president=president,
            fonction=fonction,
            status="success",
            error_message=None,
        )

    def _scrape_beautifulsoup(self, siren: str) -> Optional[InpiData]:
        """Scrape company data from data.inpi.fr using BeautifulSoup."""
        if not SCRAPING_AVAILABLE:
            return None
        # ... (port from old inpi_client.py _scrape_inpi_beautifulsoup,
        #  adapting to return InpiData instead of dict,
        #  using _format_dirigeant_name for "Prenom Nom" format)
        # Reference: /Users/xavier/VSCode3/FA_Baux_LOI_V2a/modules/inpi_client.py
        try:
            scraper = cloudscraper.create_scraper()
            url = f"https://data.inpi.fr/entreprises/{siren}"
            response = scraper.get(url, timeout=15)
            if response.status_code != 200:
                return None

            soup = BeautifulSoup(response.text, "html.parser")

            # Extract fields from HTML structure
            # ... (same parsing logic as old code)
            # Return InpiData with status="success"
            # Fallback return None on any error
            return None  # Placeholder — port full BS parsing from old code
        except Exception as e:
            logger.error(f"BeautifulSoup scraping error: {e}")
            return None

    def get_company_info(self, siret: str) -> InpiData:
        """Get company info. Always returns InpiData (never None)."""
        siren = _extract_siren(siret)
        if not siren:
            return InpiData(
                nom_societe="", type_societe="", capital_social="",
                localite_rcs="", adresse_domiciliation="",
                president="", fonction="",
                status="failed",
                error_message=f"SIRET invalide: {siret}",
            )

        # Try API first
        company_data = self._api_search(siren)
        if company_data:
            result = self._extract_from_api(company_data)
            # If no dirigeant from API, try scraping
            if not result.president and SCRAPING_AVAILABLE:
                scraped = self._scrape_beautifulsoup(siren)
                if scraped and scraped.president:
                    result = InpiData(
                        nom_societe=result.nom_societe or scraped.nom_societe,
                        type_societe=result.type_societe or scraped.type_societe,
                        capital_social=result.capital_social or scraped.capital_social,
                        localite_rcs=result.localite_rcs or scraped.localite_rcs,
                        adresse_domiciliation=result.adresse_domiciliation or scraped.adresse_domiciliation,
                        president=scraped.president,
                        fonction=scraped.fonction,
                        status="success",
                        error_message=None,
                    )
            return result

        # Fallback: BeautifulSoup scraping
        if SCRAPING_AVAILABLE:
            scraped = self._scrape_beautifulsoup(siren)
            if scraped:
                return scraped

        return InpiData(
            nom_societe="", type_societe="", capital_social="",
            localite_rcs="", adresse_domiciliation="",
            president="", fonction="",
            status="failed",
            error_message="Entreprise non trouvée (API et scraping échoués)",
        )
```

**IMPORTANT**: The `_scrape_beautifulsoup` method must be fully ported from the old code at `/Users/xavier/VSCode3/FA_Baux_LOI_V2a/modules/inpi_client.py` (the `_scrape_inpi_beautifulsoup` method, ~100 lines). The placeholder `return None` above must be replaced with the actual parsing logic, adapted to return `InpiData` and use `_format_dirigeant_name`.

- [ ] **Step 4: Run test to verify it passes**

```bash
pytest tests/test_inpi_client.py -v
```

Expected: All passed

- [ ] **Step 5: Commit**

```bash
git add core/inpi_client.py tests/test_inpi_client.py
git commit -m "feat: add INPI client (API + BeautifulSoup fallback, unified name format)"
```

---

### Task 8: Excel Parser (`core/excel_parser.py`)

**Files:**
- Create: `core/excel_parser.py`
- Test: `tests/test_excel_parser.py`

- [ ] **Step 1: Write the tests**

```python
# tests/test_excel_parser.py
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


def test_adds_date_aujourdhui():
    """The parser should add Date d'aujourd'hui = today's date."""
    # This tests the behavior without actual Excel files
    today = datetime.now().strftime("%d/%m/%Y")
    parser = ExcelParser.__new__(ExcelParser)
    variables = {}
    parser._add_system_variables(variables)
    assert variables["Date d'aujourd'hui"] == today
```

- [ ] **Step 2: Run test to verify it fails**

```bash
pytest tests/test_excel_parser.py -v
```

Expected: FAIL

- [ ] **Step 3: Write the implementation**

```python
# core/excel_parser.py
"""Unified Excel parser for LOI and BAIL data extraction."""

import logging
from datetime import datetime
from pathlib import Path
from typing import Optional

import openpyxl

from core.formula_engine import resolve_formula
from core.inpi_client import INPIClient
from core.models import DossierData, InpiData, SocieteInfo
from config.settings import get_inpi_credentials, validate_inpi_credentials

logger = logging.getLogger(__name__)


class ExcelParser:
    """Parse Excel decision files and extract variables."""

    def __init__(self, source_path: str, config_path: str):
        """
        Args:
            source_path: Path to uploaded Excel file (Fiche de decision).
            config_path: Path to config Excel (Redaction LOI.xlsx or Redaction BAIL.xlsx).
        """
        self.source_path = Path(source_path)
        self.config_path = Path(config_path)

    def _format_cell_value(self, value) -> Optional[str]:
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

    def _add_system_variables(self, variables: dict):
        """Add system-generated variables."""
        variables["Date d'aujourd'hui"] = datetime.now().strftime("%d/%m/%Y")

    def extract_variables(self) -> dict[str, str]:
        """Extract all variables from the source Excel using config mapping."""
        # Load source workbook (data_only=True to get cached formula values)
        source_wb = openpyxl.load_workbook(str(self.source_path), data_only=True)

        # Load config workbook in dual mode
        config_wb_values = openpyxl.load_workbook(str(self.config_path), data_only=True)
        config_wb_formulas = openpyxl.load_workbook(str(self.config_path), data_only=False)

        # Read config mapping from first sheet
        config_sheet_name = config_wb_values.sheetnames[0]
        ws_values = config_wb_values[config_sheet_name]
        ws_formulas = config_wb_formulas[config_sheet_name]

        variables: dict[str, str] = {}

        for row in range(2, ws_values.max_row + 1):
            nom = self._format_cell_value(ws_values.cell(row=row, column=1).value)
            if not nom:
                continue

            # Try data_only value first
            source_val = ws_values.cell(row=row, column=2).value
            # Fallback to formula if value is None or #REF!
            if source_val is None or (isinstance(source_val, str) and "#REF" in source_val):
                source_val = ws_formulas.cell(row=row, column=2).value

            source = self._format_cell_value(source_val) if source_val is not None else None

            if not source:
                continue

            if isinstance(source, str) and source.startswith("=") and "!" in source:
                # Cell reference formula -> resolve against source workbook
                value = resolve_formula(source, source_wb)
                if value:
                    variables[nom] = value
            elif isinstance(source, str) and ("[" in source and "]" in source):
                # Formula for later processing
                variables[f"_formula_{nom}"] = source
            else:
                # Direct value or description
                variables[nom] = source

        self._add_system_variables(variables)

        source_wb.close()
        config_wb_values.close()
        config_wb_formulas.close()

        return variables

    def extract_societe_info(self) -> dict[str, SocieteInfo]:
        """Extract Societe Bailleur info from config workbook."""
        config_wb = openpyxl.load_workbook(str(self.config_path), data_only=True)

        result: dict[str, SocieteInfo] = {}
        if "Societe Bailleur" not in config_wb.sheetnames:
            config_wb.close()
            return result

        ws = config_wb["Societe Bailleur"]
        for row in range(2, ws.max_row + 1):
            nom = self._format_cell_value(ws.cell(row=row, column=1).value)
            if not nom:
                continue
            header = self._format_cell_value(ws.cell(row=row, column=2).value) or nom
            footer = self._format_cell_value(ws.cell(row=row, column=3).value) or ""
            result[nom] = SocieteInfo(nom=nom, header_text=header, footer_text=footer)

        config_wb.close()
        return result

    def enrich_from_inpi(self, variables: dict) -> Optional[InpiData]:
        """Enrich variables with INPI data if SIRET is available."""
        siret = variables.get("SIRET", "")
        if not siret or not validate_inpi_credentials():
            return None

        try:
            creds = get_inpi_credentials()
            client = INPIClient(username=creds["username"], password=creds["password"])
            return client.get_company_info(siret)
        except Exception as e:
            logger.error(f"INPI enrichment error: {e}")
            return None

    def get_output_filename_loi(self, variables: dict) -> str:
        """Generate LOI output filename: 'YYYY MM DD - LOI NomPreneur.docx'."""
        nom_preneur = variables.get("Nom Preneur", "Preneur")
        date_loi = variables.get("Date LOI", "")
        try:
            dt = datetime.strptime(date_loi, "%d/%m/%Y")
            date_str = dt.strftime("%Y %m %d")
        except (ValueError, TypeError):
            date_str = datetime.now().strftime("%Y %m %d")
        return f"{date_str} - LOI {nom_preneur}.docx"

    def get_output_filename_bail(self, variables: dict) -> str:
        """Generate BAIL output filename: 'BAIL - NomPreneur - DateLOI.docx'."""
        nom_preneur = variables.get("Nom Preneur", "Preneur").replace("/", "-").replace("\\", "-")
        date_loi = variables.get("Date LOI", datetime.now().strftime("%d/%m/%Y"))
        return f"BAIL - {nom_preneur} - {date_loi}.docx"
```

- [ ] **Step 4: Run test to verify it passes**

```bash
pytest tests/test_excel_parser.py -v
```

Expected: All passed

- [ ] **Step 5: Commit**

```bash
git add core/excel_parser.py tests/test_excel_parser.py
git commit -m "feat: add unified Excel parser with config mapping and INPI enrichment"
```

---

### Task 9: LOI Generator + Renderer

**Files:**
- Create: `generators/loi_generator.py`, `renderers/loi_renderer.py`

- [ ] **Step 1: Write LOI generator**

```python
# generators/loi_generator.py
"""LOI-specific generation logic: optional sections, clear list."""

import logging
from typing import Optional

from core.models import DossierData
from generators.shared import est_societe

logger = logging.getLogger(__name__)


class LOIGenerator:
    """LOI document generation logic."""

    def __init__(self, dossier: DossierData):
        self.dossier = dossier
        self.clear_list: list[str] = []
        self._build_clear_list()

    def _build_clear_list(self):
        """Build list of placeholders to replace with empty string."""
        all_vars = {**self.dossier.variables, **self.dossier.variables_derivees}
        type_preneur = all_vars.get("Type Preneur", "")
        if not est_societe(type_preneur):
            self.clear_list.extend(["PRESIDENT DE LA SOCIETE", "FONCTION INPI"])

    def get_all_variables(self) -> dict[str, str]:
        """Merge all variable sources into a single dict for rendering."""
        result = dict(self.dossier.variables)
        result.update(self.dossier.variables_derivees)
        return result

    def has_palier_data(self) -> bool:
        """Check if any palier data exists."""
        all_vars = self.get_all_variables()
        return any(all_vars.get(f"Montant du palier {i}") for i in range(1, 7))

    def has_conditions_suspensives(self) -> bool:
        """Check if any condition suspensive exists."""
        all_vars = self.get_all_variables()
        return any(
            all_vars.get(f"Condition suspensive {i}")
            for i in range(1, 5)
        )

    def has_honoraires_preneurs(self) -> bool:
        """Check if Honoraires Preneurs data exists."""
        all_vars = self.get_all_variables()
        return bool(all_vars.get("Honoraires Preneurs"))
```

- [ ] **Step 2: Write LOI renderer**

```python
# renderers/loi_renderer.py
"""LOI Word document renderer: blue sections, headers/footers, special sections."""

import logging
from pathlib import Path

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from core.models import SocieteInfo
from generators.loi_generator import LOIGenerator
from renderers.word_engine import WordEngine

logger = logging.getLogger(__name__)


class LOIRenderer:
    """Render a LOI Word document from template + variables."""

    def __init__(self, template_path: str):
        self.template_path = Path(template_path)
        self.engine = WordEngine()

    def render(self, generator: LOIGenerator, societe_info: dict[str, SocieteInfo], output_path: str):
        """Generate the LOI document.

        Args:
            generator: LOIGenerator with all variables and clear list.
            societe_info: Dict of SocieteInfo keyed by societe name.
            output_path: Path to save the generated document.
        """
        doc = Document(str(self.template_path))
        variables = generator.get_all_variables()
        clear_list = generator.clear_list

        has_paliers = generator.has_palier_data()
        has_conditions = generator.has_conditions_suspensives()
        has_honoraires = generator.has_honoraires_preneurs()

        # Phase 1: Handle special sections
        to_delete = []
        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()

            # Special: "Honoraires de commercialisation" (content-based deletion)
            if "Honoraires de commercialisation" in text and not has_honoraires:
                to_delete.append(paragraph)
                continue

            # Special: Blue "Remises" section
            if self.engine.is_paragraph_optional(paragraph) and "Remises" in text and "loyer" in text.lower():
                if not has_paliers:
                    to_delete.append(paragraph)
                    continue
                else:
                    # Keep and turn black
                    for run in paragraph.runs:
                        run.font.color.rgb = RGBColor(0, 0, 0)

            # Special: Blue "Condition(s) suspensive(s)" section
            if self.engine.is_paragraph_optional(paragraph) and "Condition" in text and "suspensive" in text.lower():
                if not has_conditions:
                    to_delete.append(paragraph)
                    continue
                else:
                    for run in paragraph.runs:
                        run.font.color.rgb = RGBColor(0, 0, 0)
                    # Fall through to normal processing for [.] replacement

        # Phase 2: Standard placeholder replacement
        standard_deletes = self.engine.process_document_body(doc, variables, clear_list)
        to_delete.extend(standard_deletes)

        # Phase 3: Delete marked paragraphs
        self.engine.delete_paragraphs(to_delete)

        # Phase 4: Update headers/footers
        bailleur_name = variables.get("Societe Bailleur", "")
        if bailleur_name and bailleur_name in societe_info:
            self._update_headers_footers(doc, societe_info[bailleur_name])

        # Save
        doc.save(output_path)
        logger.info(f"LOI document saved to {output_path}")

    def _update_headers_footers(self, document, societe: SocieteInfo):
        """Update document headers and footers for the societe bailleur."""
        for section in document.sections:
            section.top_margin = Inches(0.5)
            section.header_distance = Inches(0.3)
            section.bottom_margin = Inches(0.5)
            section.footer_distance = Inches(0.3)

            # Header
            header = section.header
            for p in list(header.paragraphs):
                p._element.getparent().remove(p._element)
            hp = header.add_paragraph()
            hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            hr = hp.add_run(societe.header_text)
            hr.font.bold = True
            hr.font.size = Pt(22)
            hp.paragraph_format.space_after = Pt(12)

            # Footer
            footer = section.footer
            for p in list(footer.paragraphs):
                p._element.getparent().remove(p._element)
            lines = societe.footer_text.split("\n")
            for idx, line in enumerate(lines):
                fp = footer.add_paragraph()
                fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                fr = fp.add_run(line)
                fr.font.size = Pt(9)
                if idx == 0:
                    fp.paragraph_format.space_before = Pt(12)
```

- [ ] **Step 3: Run existing tests to verify no regressions**

```bash
pytest -v
```

Expected: All previously passing tests still pass

- [ ] **Step 4: Commit**

```bash
git add generators/loi_generator.py renderers/loi_renderer.py
git commit -m "feat: add LOI generator and renderer (blue sections, headers/footers)"
```

---

### Task 10: BAIL Generator (`generators/bail_generator.py`)

**Files:**
- Create: `generators/bail_generator.py`
- Test: `tests/test_bail_generator.py`

- [ ] **Step 1: Write the tests**

```python
# tests/test_bail_generator.py
import pytest
from generators.bail_generator import evaluer_condition, generer_conditions_suspensives


def test_condition_empty():
    assert evaluer_condition("", {}) is True
    assert evaluer_condition(None, {}) is True


def test_condition_equals_oui():
    assert evaluer_condition('Si [Actualisation] = \'Oui\'', {"Actualisation": "Oui"}) is True
    assert evaluer_condition('Si [Actualisation] = \'Oui\'', {"Actualisation": "Non"}) is False


def test_condition_equals_case_insensitive():
    assert evaluer_condition('Si [Actualisation] = \'oui\'', {"Actualisation": "Oui"}) is True


def test_condition_greater_than():
    assert evaluer_condition('Si [Duree Bail] > 9', {"Duree Bail": "10"}) is True
    assert evaluer_condition('Si [Duree Bail] > 9', {"Duree Bail": "9"}) is False


def test_condition_superieur_a():
    assert evaluer_condition('Si [Duree Bail] superieur a 9', {"Duree Bail": "10"}) is True


def test_condition_not_equal():
    assert evaluer_condition('Si [Type] != \'SAS\'', {"Type": "SARL"}) is True
    assert evaluer_condition('Si [Type] != \'SAS\'', {"Type": "SAS"}) is False


def test_condition_non_vide():
    assert evaluer_condition('Si [Loyer année 1] non vide', {"Loyer année 1": "5000"}) is True
    assert evaluer_condition('Si [Loyer année 1] non vide', {"Loyer année 1": ""}) is False
    assert evaluer_condition('Si [Loyer année 1] non vide', {}) is False


def test_condition_non_vide_zero_is_empty():
    assert evaluer_condition('Si [Val] non vide', {"Val": "0"}) is False
    assert evaluer_condition('Si [Val] non vide', {"Val": 0}) is False


def test_condition_plusieurs_suspensives():
    donnees = {
        "Condition suspensive 1": "Financement",
        "Condition suspensive 2": "Extraction",
        "Condition suspensive 3": "",
        "Condition suspensive 4": "",
    }
    assert evaluer_condition("Si plusieurs conditions suspensives", donnees) is True


def test_condition_plusieurs_suspensives_only_one():
    donnees = {
        "Condition suspensive 1": "Financement",
        "Condition suspensive 2": "",
    }
    assert evaluer_condition("Si plusieurs conditions suspensives", donnees) is False


def test_condition_typographic_quotes():
    # Smart quotes should be normalized
    assert evaluer_condition("Si [Type] = \u2018Oui\u2019", {"Type": "Oui"}) is True


def test_condition_unrecognized():
    assert evaluer_condition("Something weird", {}) is False


def test_conditions_suspensives_single():
    donnees = {"Condition suspensive 1": "Financement"}
    result = generer_conditions_suspensives(donnees, "Option 1 text", "Option 2 text")
    assert result == "Option 1 text"


def test_conditions_suspensives_multiple():
    donnees = {
        "Condition suspensive 1": "Financement",
        "Condition suspensive 2": "Extraction",
    }
    result = generer_conditions_suspensives(donnees, "Option 1", "Texte avec suivantes :\n\nblabla\n\nCi-après")
    assert "a." in result
    assert "b." in result
    assert "Financement" not in result  # The actual legal text replaces the condition name
    assert "prêt bancaire" in result.lower() or "pret bancaire" in result.lower() or "Obtention" in result


def test_conditions_suspensives_none():
    donnees = {}
    result = generer_conditions_suspensives(donnees, "Option 1", "Option 2")
    assert result == ""
```

- [ ] **Step 2: Run test to verify it fails**

```bash
pytest tests/test_bail_generator.py -v
```

Expected: FAIL

- [ ] **Step 3: Write the implementation**

```python
# generators/bail_generator.py
"""BAIL business logic: condition evaluation, article generation, conditions suspensives."""

import logging
import re
from typing import Any, Optional

import openpyxl
import pandas as pd

from core.formula_engine import resolve_formula, resolve_cell_range
from core.models import ArticleResult

logger = logging.getLogger(__name__)

# Hard-coded legal texts for conditions suspensives
_TEXTES_CONDITIONS = {
    "Financement": (
        "Obtention d'un prêt bancaire par le Preneur d'un montant de [X] € "
        "([EN LETTRES] EUROS) auprès d'un organisme prêteur ;"
    ),
    "Autorisations administratives": (
        "Obtention par le Preneur de l'ensemble des autorisations administratives "
        "pour la réalisation des travaux nécessaires à la bonne exploitation de son "
        "activité et à la mise en place dans les Locaux Loués de son concept, à savoir "
        "une Déclaration Préalable pour les modifications de façade, une demande "
        "d'Autorisation d'aménager un établissement recevant du public (ERP) ou toute "
        "autre autorisation administrative qui s'avérerait indispensable à l'ouverture "
        "et à l'exploitation des Locaux Loués au vu de la localisation spécifique de ces "
        "derniers. Le Preneur s'engage à justifier au Bailleur du dépôt de ces "
        "autorisations administratives au plus tard 1 (UN) mois après la signature des "
        "présentes. ;"
    ),
    "Extraction": (
        "Obtention par le bailleur d'une autorisation de copropriété pour la "
        "réalisation de travaux d'installation d'un conduit d'extraction extérieur "
        "permettant notamment l'exercice d'une activité de restauration au sein des "
        "Locaux Loués ;"
    ),
    "Libération des locaux": (
        "Libération effective des Locaux Loués par l'actuel occupant, étant précisé "
        "que l'occupant a donné congé pour le [X] ;"
    ),
    "Autre": "[.]",
}

# Article generation order
_ARTICLES_ORDER = [
    ("Comparution Bailleur", None),
    ("Comparution Preneur", None),
    ("Article preliminaire", None),
    ("Article 1", None),
    ("Article 2", None),
    ("Article 3", None),
    ("Article 5.3", None),
    ("Article 7.1", None),
    ("Article 7.2", None),
    ("Article  7.3", None),  # Note: double space preserved from template
    ("Article 7.6", None),
    ("Article 8", None),
    ("Article 19", None),
    ("Article 22.2", None),
    ("Article 26", None),
    ("Article 26.1", None),
    ("Article 26.2", None),
]


def _normalize_quotes(text: str) -> str:
    """Replace typographic quotes with straight quotes."""
    return text.replace("\u2018", "'").replace("\u2019", "'").replace("\u201c", '"').replace("\u201d", '"')


def evaluer_condition(condition_str: Optional[str], donnees: dict[str, Any]) -> bool:
    """Evaluate a textual condition against data.

    Supports:
    - Empty/None -> True
    - "Si [X] = 'value'" -> string comparison
    - "Si [X] > N" -> numeric comparison
    - "Si [X] non vide" -> non-empty check
    - "Si plusieurs conditions suspensives" -> count check
    """
    if not condition_str or not str(condition_str).strip():
        return True

    condition = _normalize_quotes(str(condition_str).strip())

    # Special: multiple conditions suspensives
    if "plusieurs conditions suspensives" in condition.lower():
        count = sum(
            1 for i in range(1, 5)
            if donnees.get(f"Condition suspensive {i}")
            and str(donnees[f"Condition suspensive {i}"]).strip()
            and str(donnees[f"Condition suspensive {i}"]).lower() != "none"
        )
        return count > 1

    # "Si [Variable] non vide/nul"
    non_vide_match = re.search(r"Si\s+\[([^\]]+)\]\s+non\s+(vide|nul)", condition, re.IGNORECASE)
    if non_vide_match:
        var_name = non_vide_match.group(1)
        value = donnees.get(var_name)
        return bool(value) and value != 0 and str(value).strip() != "0" and str(value).strip() != ""

    # Comparison: "Si [Variable] operator value"
    comp_match = re.search(
        r"Si\s+\"?([^\"\[\]]+|\[[^\]]+\])\"?\s*(=|>|<|>=|<=|!=|superieur a|superieure a)\s*[\"']?([^\"']+)[\"']?",
        condition,
        re.IGNORECASE,
    )
    if comp_match:
        var_ref = comp_match.group(1).strip().strip("[]")
        operator = comp_match.group(2).strip().lower()
        expected = comp_match.group(3).strip()

        # Normalize variable name
        actual = donnees.get(var_ref, "")

        if operator in ("=",):
            return str(actual).strip().lower() == expected.lower()
        if operator in ("!=",):
            return str(actual).strip().lower() != expected.lower()

        # Numeric comparisons
        try:
            actual_num = float(str(actual).replace(" ", "").replace(",", "."))
            expected_num = float(expected.replace(" ", "").replace(",", "."))
        except (ValueError, TypeError):
            return False

        if operator in (">", "superieur a", "superieure a"):
            return actual_num > expected_num
        if operator == ">=":
            return actual_num >= expected_num
        if operator == "<":
            return actual_num < expected_num
        if operator == "<=":
            return actual_num <= expected_num

    logger.warning(f"Unrecognized condition: {condition_str}")
    return False


def generer_conditions_suspensives(
    donnees: dict[str, Any],
    option1_text: str,
    option2_text: str,
) -> str:
    """Generate conditions suspensives text.

    1 condition -> option1_text directly.
    Multiple -> option2_text with a./b./c./d. lettered items.
    0 conditions -> empty string.
    """
    conditions = []
    for i in range(1, 5):
        key = f"Condition suspensive {i}"
        value = donnees.get(key)
        if value and str(value).strip() and str(value).lower() != "none":
            conditions.append((key, str(value).strip()))

    if not conditions:
        return ""

    if len(conditions) == 1:
        return option1_text

    # Multiple conditions: use option2_text with letter replacement
    lettres = ["a", "b", "c", "d"]
    conditions_text = []
    for idx, (key, value) in enumerate(conditions):
        if idx < len(lettres):
            lettre = lettres[idx]
            texte = _TEXTES_CONDITIONS.get(value, f"[Condition: {value}]")
            conditions_text.append(f"{lettre}. {texte}")

    replacement_lines = "\n\n".join(conditions_text)
    replacement = f"suivantes :\n\n{replacement_lines}\n\nCi-après"

    result = re.sub(
        r"suivantes\s*:\s*\n\s*\n(.+?)\n\s*\nCi-après",
        replacement,
        option2_text,
        flags=re.DOTALL,
    )
    return result


class BailGenerator:
    """BAIL document generator with conditional article selection."""

    def __init__(self, config_path: str, source_workbook_path: Optional[str] = None):
        self.config_path = config_path
        self.source_workbook = None
        self.regles_df = None

        if source_workbook_path:
            self.source_workbook = openpyxl.load_workbook(source_workbook_path, data_only=True)

        self._load_rules()

    def _load_rules(self):
        """Load rules from Redaction BAIL.xlsx."""
        wb = openpyxl.load_workbook(self.config_path, data_only=True)
        ws = None
        for name in wb.sheetnames:
            if "bail" in name.lower() and "redaction" in name.lower():
                ws = wb[name]
                break
            if "rédaction" in name.lower() and "bail" in name.lower():
                ws = wb[name]
                break
        if ws is None:
            # Fallback: first sheet
            ws = wb[wb.sheetnames[0]]

        headers = [cell.value for cell in ws[1]]
        data = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            data.append(dict(zip(headers, row)))
        self.regles_df = pd.DataFrame(data)
        wb.close()

    def _resolve_source(self, source_ref: str) -> Any:
        """Resolve a data source reference (formula or literal)."""
        if not source_ref or not self.source_workbook:
            return source_ref
        source_str = str(source_ref).strip()
        if source_str.startswith("=") and "!" in source_str:
            if ":" in source_str.split("!")[-1]:
                # Range reference
                sheet, rng = source_str.lstrip("=").split("!", 1)
                sheet = sheet.strip("'")
                return resolve_cell_range(sheet, rng, self.source_workbook)
            return resolve_formula(source_str, self.source_workbook)
        return source_ref

    def _get_article_rows(self, article_name: str) -> list[dict]:
        """Get all rows for an article (including continuation rows)."""
        rows = []
        found = False
        for _, row in self.regles_df.iterrows():
            art = row.get("Article")
            if pd.notna(art) and str(art).strip() == article_name:
                found = True
                rows.append(row)
            elif found and (pd.isna(art) or str(art).strip() == ""):
                rows.append(row)
            elif found and pd.notna(art) and str(art).strip() != article_name:
                break
        return rows

    def generer_bail(self, variables: dict[str, Any]) -> list[ArticleResult]:
        """Generate all BAIL articles from rules and variables."""
        articles = []

        for article_name, designation in _ARTICLES_ORDER:
            rows = self._get_article_rows(article_name)
            if not rows:
                logger.warning(f"No rules found for {article_name}")
                continue

            textes = []
            manquants = []

            for row in rows:
                condition = row.get("Condition")
                option1 = row.get("Entrée correspondante - Option 1", "")
                condition2 = row.get("Condition Option 2")
                option2 = row.get("Entrée correspondante - Option 2", "")

                # Check for conditions suspensives special case
                nom_source = str(row.get("Nom Source", "")) if pd.notna(row.get("Nom Source")) else ""
                if (article_name == "Article preliminaire"
                        and "Condition" in nom_source and "suspensive" in nom_source.lower()):
                    opt1 = str(option1) if pd.notna(option1) else ""
                    opt2 = str(option2) if pd.notna(option2) else ""
                    texte = generer_conditions_suspensives(variables, opt1, opt2)
                    if texte:
                        textes.append(texte)
                    continue

                # Standard condition evaluation
                if evaluer_condition(condition, variables):
                    if pd.notna(option1) and str(option1).strip():
                        textes.append(str(option1))
                elif evaluer_condition(condition2, variables):
                    if pd.notna(option2) and str(option2).strip():
                        textes.append(str(option2))

            contenu = "\n\n".join(textes)

            # Replace [Variable] placeholders in the generated text
            from generators.shared import formater_nombre
            for match in re.findall(r"\[([^\]]+)\]", contenu):
                value = variables.get(match)
                if value is not None and str(value).strip():
                    # Format numbers
                    try:
                        num = float(str(value).replace(" ", "").replace(",", "."))
                        formatted = formater_nombre(num)
                        contenu = contenu.replace(f"[{match}]", formatted)
                    except (ValueError, TypeError):
                        contenu = contenu.replace(f"[{match}]", str(value))
                else:
                    manquants.append(match)

            # Use designation from first row as key
            desig = str(rows[0].get("Designation", article_name)) if rows else article_name
            articles.append(ArticleResult(
                designation=desig,
                contenu=contenu,
                placeholders_manquants=manquants,
            ))

        return articles
```

- [ ] **Step 4: Run test to verify it passes**

```bash
pytest tests/test_bail_generator.py -v
```

Expected: All passed

- [ ] **Step 5: Commit**

```bash
git add generators/bail_generator.py tests/test_bail_generator.py
git commit -m "feat: add BAIL generator (condition evaluation, article generation, conditions suspensives)"
```

---

### Task 11: BAIL Renderer (`renderers/bail_renderer.py`)

**Files:**
- Create: `renderers/bail_renderer.py`

- [ ] **Step 1: Write the implementation**

```python
# renderers/bail_renderer.py
"""BAIL Word document renderer: article placeholders, HTML tags, headings, TOC."""

import logging
import re
from pathlib import Path
from typing import Optional

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from lxml import etree

from core.models import ArticleResult
from renderers.word_engine import WordEngine

logger = logging.getLogger(__name__)

# Mapping from article designation to template placeholder
_ARTICLE_PLACEHOLDER_MAP = {
    "Comparution Bailleur": "{{COMPARUTION_BAILLEUR}}",
    "Comparution Preneur": "{{COMPARUTION_PRENEUR}}",
    "Article preliminaire": "{{ARTICLE_PRELIMINAIRE}}",
    "Article 1": "{{ARTICLE_1}}",
    "Article 2": "{{ARTICLE_2}}",
    "Article 3": "{{ARTICLE_3}}",
    "Article 5.3": "{{ARTICLE_5_3}}",
    "Article 7.1": "{{ARTICLE_7_1}}",
    "Article 7.2": "{{ARTICLE_7_2}}",
    "Article  7.3": "{{ARTICLE_7_3}}",  # Double space in designation
    "Article 7.6": "{{ARTICLE_7_6}}",
    "Article 8": "{{ARTICLE_8}}",
    "Article 19": "{{ARTICLE_19}}",
    "Article 22.2": "{{ARTICLE_22_2}}",
    "Article 26": "{{ARTICLE_26}}",
    "Article 26.1": "{{ARTICLE_26_1}}",
    "Article 26.2": "{{ARTICLE_26_2}}",
}

_FORMATTING_TAG_RE = re.compile(r"<(/?)([biu])>", re.IGNORECASE)


def parse_formatting_tags(text: str) -> list[tuple[str, dict]]:
    """Parse HTML-like formatting tags (<b>, <i>, <u>) from text.

    Returns list of (text_segment, {bold: bool, italic: bool, underline: bool}).
    Stack-based: supports nesting.
    """
    if not _FORMATTING_TAG_RE.search(text):
        return [(text, {})]

    segments = []
    stack = {"b": 0, "i": 0, "u": 0}
    current_text = ""
    current_format = {}

    i = 0
    while i < len(text):
        match = _FORMATTING_TAG_RE.match(text, i)
        if match:
            # Save current segment
            if current_text:
                segments.append((current_text, dict(current_format)))
                current_text = ""

            is_closing = match.group(1) == "/"
            tag = match.group(2).lower()

            if is_closing:
                stack[tag] = max(0, stack[tag] - 1)
            else:
                stack[tag] += 1

            # Update current format
            current_format = {}
            if stack["b"] > 0:
                current_format["bold"] = True
            if stack["i"] > 0:
                current_format["italic"] = True
            if stack["u"] > 0:
                current_format["underline"] = True

            i = match.end()
        else:
            current_text += text[i]
            i += 1

    if current_text:
        segments.append((current_text, dict(current_format)))

    return segments


class BailRenderer:
    """Render a BAIL Word document from template + articles + variables."""

    def __init__(self, template_path: str):
        self.template_path = Path(template_path)
        self.engine = WordEngine()

    def render(self, articles: list[ArticleResult], variables: dict[str, str], output_path: str):
        """Generate the BAIL document.

        Args:
            articles: Generated articles from BailGenerator.
            variables: All merged variables for [Variable] replacement.
            output_path: Path to save the generated document.
        """
        doc = Document(str(self.template_path))

        # Phase 1: Replace {{ARTICLE}} placeholders
        self._replace_article_placeholders(doc, articles, variables)

        # Phase 2: Replace remaining [Variable] placeholders
        self.engine.process_document_body(doc, variables)

        # Phase 3: Clean unreplaced {{}} placeholders
        self._clean_unreplaced_placeholders(doc)

        # Phase 4: Fix heading indentation
        self._fix_heading_indentation(doc)

        # Phase 5: Mark TOC dirty
        self._update_toc(doc)

        doc.save(output_path)
        logger.info(f"BAIL document saved to {output_path}")

    def _replace_article_placeholders(self, doc, articles: list[ArticleResult], variables: dict):
        """Replace {{ARTICLE_xxx}} placeholders with generated content."""
        # Also handle {{VILLE}} and {{DATE_SIGNATURE}}
        ville = variables.get("Ville ou arrondissement", "")
        if "(" in ville:
            ville = ville.split("(")[0].strip()

        for paragraph in list(doc.paragraphs):
            text = paragraph.text.strip()
            if not text.startswith("{{"):
                continue

            # Check for {{VILLE}}
            if "{{VILLE}}" in text:
                for run in paragraph.runs:
                    run.text = run.text.replace("{{VILLE}}", ville)
                continue

            # Check for {{DATE_SIGNATURE}}
            if "{{DATE_SIGNATURE}}" in text:
                date_sig = variables.get("Date de signature", "")
                for run in paragraph.runs:
                    run.text = run.text.replace("{{DATE_SIGNATURE}}", date_sig)
                continue

            # Check for article placeholders
            for article in articles:
                placeholder = _ARTICLE_PLACEHOLDER_MAP.get(article.designation)
                if not placeholder or placeholder not in text:
                    continue

                if not article.contenu.strip():
                    continue

                self._insert_article_content(paragraph, article.contenu)
                break

    def _insert_article_content(self, paragraph, content: str):
        """Replace a paragraph with multi-paragraph article content."""
        # Split on double newlines for paragraph breaks
        parts = content.split("\n\n")

        # Further split on heading markers
        all_parts = []
        for part in parts:
            sub_parts = re.split(r"(\n(?=\*{2,4}))", part)
            for sp in sub_parts:
                if sp.strip():
                    all_parts.append(sp.strip())

        if not all_parts:
            return

        # Process first part: reuse existing paragraph
        first_part = all_parts[0]
        self._render_paragraph_content(paragraph, first_part)

        # Process remaining parts: insert new paragraphs after
        prev_element = paragraph._element
        for part in all_parts[1:]:
            if not part.strip():
                # Empty paragraph for spacing
                new_p = etree.SubElement(prev_element.getparent(), qn("w:p"))
                prev_element.addnext(new_p)
                prev_element = new_p
                continue

            # Create new paragraph element
            new_p_element = etree.SubElement(prev_element.getparent(), qn("w:p"))
            prev_element.addnext(new_p_element)
            prev_element = new_p_element

            # Wrap in a docx Paragraph object for easier manipulation
            from docx.text.paragraph import Paragraph
            new_p = Paragraph(new_p_element, paragraph._parent)

            self._render_paragraph_content(new_p, part)

    def _render_paragraph_content(self, paragraph, text: str):
        """Render text with heading markers and formatting tags into a paragraph."""
        # Detect heading level
        heading_level = None
        clean_text = text
        if text.startswith("****"):
            heading_level = 4
            clean_text = text[4:].lstrip()
        elif text.startswith("***"):
            heading_level = 3
            clean_text = text[3:].lstrip()
        elif text.startswith("**"):
            heading_level = 2
            clean_text = text[2:].lstrip()

        if heading_level:
            paragraph.style = paragraph.part.document.styles[f"Heading {heading_level}"]
            paragraph.paragraph_format.left_indent = None
            paragraph.paragraph_format.first_line_indent = None
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        else:
            try:
                paragraph.style = paragraph.part.document.styles["Normal"]
            except Exception:
                pass

        # Clear existing runs
        for run in paragraph.runs:
            run.text = ""

        # Parse formatting tags and create runs
        segments = parse_formatting_tags(clean_text)
        for seg_text, formatting in segments:
            if not seg_text:
                continue
            run = paragraph.add_run(seg_text)
            # Apply formatting from tags
            if formatting.get("bold"):
                run.font.bold = True
            if formatting.get("italic"):
                run.font.italic = True
            if formatting.get("underline"):
                run.font.underline = True

    def _clean_unreplaced_placeholders(self, doc):
        """Remove paragraphs that contain only {{...}} placeholders."""
        to_remove = []
        for p in doc.paragraphs:
            text = p.text.strip()
            if text and re.match(r"^(\{\{[^}]*\}\}\s*)+$", text):
                to_remove.append(p)
        self.engine.delete_paragraphs(to_remove)

    def _fix_heading_indentation(self, doc):
        """Reset indentation for all heading paragraphs."""
        for p in doc.paragraphs:
            if p.style and p.style.name and p.style.name.startswith("Heading"):
                p.paragraph_format.left_indent = None
                p.paragraph_format.first_line_indent = None
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                # Remove empty runs
                for run in list(p.runs):
                    if not run.text:
                        run._element.getparent().remove(run._element)

    def _update_toc(self, doc):
        """Mark TOC as dirty so Word regenerates it on open."""
        for p in doc.paragraphs:
            for run in p.runs:
                for fld_char in run._element.findall(qn("w:fldChar")):
                    if fld_char.get(qn("w:fldCharType")) == "begin":
                        fld_char.set(qn("w:dirty"), "1")

        # Update document settings
        settings = doc.settings.element
        update_fields = settings.find(qn("w:updateFields"))
        if update_fields is None:
            update_fields = etree.SubElement(settings, qn("w:updateFields"))
        update_fields.set(qn("w:val"), "true")
```

- [ ] **Step 2: Run all tests**

```bash
pytest -v
```

Expected: All passed

- [ ] **Step 3: Commit**

```bash
git add renderers/bail_renderer.py
git commit -m "feat: add BAIL renderer (articles, HTML tags, headings, TOC)"
```

---

### Task 12: Streamlit App (`app.py`)

**Files:**
- Create: `app.py`

- [ ] **Step 1: Write the app**

```python
# app.py
"""LOI-BAIL Document Generator — Streamlit application."""

import hashlib
import logging
import tempfile
from pathlib import Path

import streamlit as st

from config.settings import (
    TEMPLATE_LOI, TEMPLATE_BAIL, CONFIG_LOI, CONFIG_BAIL, OUTPUT_DIR,
)
from core.excel_parser import ExcelParser
from core.models import DossierData
from generators.shared import normaliser_noms_variables, calculer_variables_derivees
from generators.loi_generator import LOIGenerator
from generators.bail_generator import BailGenerator
from renderers.loi_renderer import LOIRenderer
from renderers.bail_renderer import BailRenderer

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

st.set_page_config(page_title="Générateur LOI & BAIL", page_icon="📄", layout="wide")


def _check_required_files():
    """Check that all required files exist."""
    required = {
        "Template LOI": TEMPLATE_LOI,
        "Template BAIL": TEMPLATE_BAIL,
        "Config LOI": CONFIG_LOI,
        "Config BAIL": CONFIG_BAIL,
    }
    missing = [name for name, path in required.items() if not Path(path).exists()]
    if missing:
        st.error(f"Fichiers manquants: {', '.join(missing)}")
        st.stop()


@st.cache_data(show_spinner=False)
def _parse_excel(file_content: bytes, file_name: str, config_path: str, _cache_key: str):
    """Parse Excel file with daily cache invalidation."""
    # Write to temp file
    file_hash = hashlib.sha256(file_content).hexdigest()[:12]
    tmp_path = Path(tempfile.gettempdir()) / f"temp_{file_hash}.xlsx"
    try:
        tmp_path.write_bytes(file_content)
        parser = ExcelParser(str(tmp_path), config_path)
        variables = parser.extract_variables()
        societes = parser.extract_societe_info()
        inpi_data = parser.enrich_from_inpi(variables)
        return variables, societes, inpi_data, str(tmp_path)
    finally:
        if tmp_path.exists():
            tmp_path.unlink()


def main():
    st.title("📄 Générateur LOI & BAIL")
    _check_required_files()

    uploaded_file = st.file_uploader(
        "Charger la Fiche de décision (Excel)",
        type=["xlsx", "xls"],
    )

    if not uploaded_file:
        st.info("Veuillez charger un fichier Excel pour commencer.")
        return

    # Parse with daily cache
    from datetime import datetime
    cache_key = f"{uploaded_file.name}_{datetime.now().strftime('%Y-%m-%d')}"
    file_content = uploaded_file.read()

    with st.spinner("Extraction des données..."):
        variables, societes, inpi_data, source_path = _parse_excel(
            file_content, uploaded_file.name, str(CONFIG_LOI), cache_key,
        )

    # Normalize + derive
    variables = normaliser_noms_variables(variables)
    variables_derivees = calculer_variables_derivees(variables, inpi_data)

    dossier = DossierData(
        variables=variables,
        variables_derivees=variables_derivees,
        inpi_data=inpi_data,
        source_file=Path(source_path),
    )

    # Show extraction summary
    nom_preneur = variables.get("Nom Preneur", "—")
    st.success(f"Données extraites pour: **{nom_preneur}**")

    if inpi_data and inpi_data.status == "success":
        st.info(f"INPI: {inpi_data.nom_societe} — {inpi_data.president} ({inpi_data.fonction})")

    # Tabs
    tab_loi, tab_bail = st.tabs(["📝 LOI", "📋 BAIL"])

    with tab_loi:
        if st.button("Générer la LOI", key="btn_gen_loi"):
            with st.spinner("Génération de la LOI..."):
                generator = LOIGenerator(dossier)
                renderer = LOIRenderer(str(TEMPLATE_LOI))
                parser = ExcelParser(source_path, str(CONFIG_LOI))
                output_name = parser.get_output_filename_loi(variables)
                output_path = OUTPUT_DIR / output_name
                OUTPUT_DIR.mkdir(exist_ok=True)
                renderer.render(generator, societes, str(output_path))

            with open(output_path, "rb") as f:
                st.download_button(
                    "⬇️ Télécharger la LOI",
                    f.read(),
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="download_loi",
                )

    with tab_bail:
        if st.button("Générer le BAIL", key="btn_gen_bail"):
            with st.spinner("Génération du BAIL..."):
                all_vars = {**variables, **variables_derivees}
                bail_gen = BailGenerator(str(CONFIG_BAIL), source_workbook_path=source_path)
                articles = bail_gen.generer_bail(all_vars)
                renderer = BailRenderer(str(TEMPLATE_BAIL))
                parser = ExcelParser(source_path, str(CONFIG_LOI))
                output_name = parser.get_output_filename_bail(variables)
                output_path = OUTPUT_DIR / output_name
                OUTPUT_DIR.mkdir(exist_ok=True)
                renderer.render(articles, all_vars, str(output_path))

            with open(output_path, "rb") as f:
                st.download_button(
                    "⬇️ Télécharger le BAIL",
                    f.read(),
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="download_bail",
                )


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Verify it starts**

```bash
streamlit run app.py --server.headless true
```

Expected: App starts without import errors. Ctrl+C to stop.

- [ ] **Step 3: Commit**

```bash
git add app.py
git commit -m "feat: add Streamlit app (upload, LOI/BAIL generation, download)"
```

---

### Task 13: Integration Test

**Files:**
- Create: `tests/test_integration.py`, `tests/conftest.py`

- [ ] **Step 1: Write conftest with shared fixtures**

```python
# tests/conftest.py
"""Shared pytest fixtures."""
import sys
from pathlib import Path

# Ensure project root is in path
sys.path.insert(0, str(Path(__file__).parent.parent))
```

- [ ] **Step 2: Write integration test**

```python
# tests/test_integration.py
"""End-to-end integration tests using real templates and data files."""
import pytest
from pathlib import Path
from docx import Document

from config.settings import TEMPLATE_LOI, TEMPLATE_BAIL, CONFIG_LOI, CONFIG_BAIL


@pytest.fixture
def real_templates_available():
    """Skip if real template files are not available."""
    if not Path(TEMPLATE_LOI).exists():
        pytest.skip("Template LOI not found")
    if not Path(TEMPLATE_BAIL).exists():
        pytest.skip("Template BAIL not found")


def test_loi_template_has_placeholders(real_templates_available):
    """Verify the LOI template contains expected placeholders."""
    doc = Document(str(TEMPLATE_LOI))
    full_text = "\n".join(p.text for p in doc.paragraphs)
    assert "[Nom Preneur]" in full_text or "[" in full_text


def test_bail_template_has_article_placeholders(real_templates_available):
    """Verify the BAIL template contains {{ARTICLE}} placeholders."""
    doc = Document(str(TEMPLATE_BAIL))
    full_text = "\n".join(p.text for p in doc.paragraphs)
    assert "{{" in full_text


def test_word_engine_on_real_loi_template(real_templates_available):
    """Test WordEngine on the real LOI template with sample data."""
    from renderers.word_engine import WordEngine

    doc = Document(str(TEMPLATE_LOI))
    engine = WordEngine()

    variables = {
        "Nom Preneur": "TEST SAS",
        "Montant du loyer": "160 000",
        "Duree Bail": "9",
    }

    # Process should not raise
    to_delete = engine.process_document_body(doc, variables)

    # Check replacement worked
    full_text = "\n".join(p.text for p in doc.paragraphs)
    assert "TEST SAS" in full_text


def test_full_loi_generation(real_templates_available):
    """Test complete LOI generation pipeline with mock data."""
    from core.models import DossierData
    from generators.shared import normaliser_noms_variables, calculer_variables_derivees
    from generators.loi_generator import LOIGenerator
    from renderers.loi_renderer import LOIRenderer
    import tempfile

    variables = normaliser_noms_variables({
        "Nom Preneur": "INTEGRATION TEST SAS",
        "Montant du loyer": "160000",
        "Duree Bail": "9",
        "Numero et rue": "10 rue de la Paix",
        "Ville ou arrondissement": "Paris 2ème",
        "Date d'aujourd'hui": "01/04/2026",
        "Type Preneur": "SAS",
        "Societe Bailleur": "",
    })
    derivees = calculer_variables_derivees(variables, None)

    dossier = DossierData(
        variables=variables,
        variables_derivees=derivees,
        inpi_data=None,
        source_file=Path("/tmp/fake.xlsx"),
    )

    generator = LOIGenerator(dossier)
    renderer = LOIRenderer(str(TEMPLATE_LOI))

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        output_path = f.name

    renderer.render(generator, {}, output_path)

    # Verify output exists and is a valid DOCX
    assert Path(output_path).exists()
    doc = Document(output_path)
    full_text = "\n".join(p.text for p in doc.paragraphs)
    assert "INTEGRATION TEST SAS" in full_text
    assert "3/6/9" in full_text  # Type Bail derived value

    Path(output_path).unlink()
```

- [ ] **Step 3: Run all tests**

```bash
pytest -v
```

Expected: All tests pass (integration tests may skip if templates not present)

- [ ] **Step 4: Commit**

```bash
git add tests/conftest.py tests/test_integration.py
git commit -m "feat: add integration tests for end-to-end LOI generation"
```

---

### Task 14: GitHub Repository Setup

- [ ] **Step 1: Create the GitHub repo**

```bash
gh repo create XavierKain/LOI-BAIL-superpower --public --source=. --remote=origin
```

- [ ] **Step 2: Push all code**

```bash
git branch -M main
git push -u origin main
```

- [ ] **Step 3: Copy .env file for local development**

```bash
cp "/Users/xavier/VSCode3/FA_Baux_LOI_V2a/.env" .env
```

- [ ] **Step 4: Run the app locally to verify**

```bash
streamlit run app.py
```

Expected: App loads, file upload works, LOI generation produces a DOCX.

- [ ] **Step 5: Final commit with any fixes**

```bash
git add -A
git commit -m "chore: final setup and configuration"
git push
```

---

## Self-Review Checklist

**Spec coverage:**
- [x] Section 1: Models — Task 2
- [x] Section 2: Excel parser + formula engine — Tasks 4, 8
- [x] Section 3: INPI client — Task 7
- [x] Section 4: Shared calculations — Task 5
- [x] Section 5: Word engine — Task 6
- [x] Section 6: LOI renderer — Task 9
- [x] Section 7: BAIL generator + renderer — Tasks 10, 11
- [x] Section 8: Streamlit app — Task 12
- [x] Section 9: Config — Task 1
- [x] Section 10: Tests — Tasks 2-8, 13
- [x] Bug fixes table — All addressed in Tasks 5 (shared), 6 (word engine), 7 (INPI)

**Placeholder scan:** No TBD/TODO found. The INPI BeautifulSoup scraping method has a note to port from old code — this is intentional as the full HTML parsing logic (~100 lines) should be ported verbatim.

**Type consistency:** `DossierData`, `InpiData`, `SocieteInfo`, `ArticleResult` used consistently. `evaluer_condition` and `generer_conditions_suspensives` are module-level functions in `bail_generator.py`, referenced correctly in tests and by `BailGenerator.generer_bail`.
