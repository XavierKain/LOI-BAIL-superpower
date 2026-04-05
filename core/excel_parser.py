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
        self.source_path = Path(source_path)
        self.config_path = Path(config_path)
        self._effective_config_path = self._resolve_config_path()

    def _resolve_config_path(self) -> Path:
        """If the source file contains a 'Rédaction LOI' sheet, use it as config."""
        try:
            wb = openpyxl.load_workbook(str(self.source_path), read_only=True)
            has_redaction = any(
                s.lower() in ("rédaction loi", "redaction loi")
                for s in wb.sheetnames
            )
            wb.close()
            if has_redaction:
                logger.info("Source file contains 'Rédaction LOI' — using as config.")
                return self.source_path
        except Exception as e:
            logger.debug(f"Could not inspect source for config sheet: {e}")
        return self.config_path

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

    def _find_config_sheet(self, workbook) -> Optional[str]:
        """Find the Rédaction LOI sheet in a workbook."""
        for name in workbook.sheetnames:
            if name.lower() in ("rédaction loi", "redaction loi"):
                return name
        return None

    def _is_section_header(self, value_a: str, value_b: str = None) -> bool:
        """Check if a row is a section header (not a variable row).

        A section header is "Société Bailleur" with "Entête" in column B,
        or "Condition suspensive Fiche de décision".
        Row 4 has "Société Bailleur" as a VARIABLE name (with a formula in B),
        while row 35 has it as a SECTION header (with "Entête" in B).
        """
        if not value_a:
            return False
        lower_a = value_a.lower()

        # "Société Bailleur" is only a section header if col B = "Entête"
        if "société bailleur" in lower_a or "societe bailleur" in lower_a:
            if value_b and "entête" in value_b.lower():
                return True
            if value_b and "entete" in value_b.lower():
                return True
            return False

        if "condition suspensive" in lower_a and "fiche" in lower_a:
            return True

        return False

    def extract_variables(self) -> dict[str, str]:
        """Extract all variables from the source Excel using config mapping.

        Reads primary variables from columns A+B and secondary from C+D.
        Uses the FORMULA version of the config to get formulas, resolves them
        against the source workbook.
        """
        source_wb = openpyxl.load_workbook(str(self.source_path), data_only=True)

        # Always read config in formula mode to get the formulas
        config_wb_formulas = openpyxl.load_workbook(
            str(self._effective_config_path), data_only=False
        )

        config_sheet_name = self._find_config_sheet(config_wb_formulas)
        if not config_sheet_name:
            config_sheet_name = config_wb_formulas.sheetnames[0]

        ws = config_wb_formulas[config_sheet_name]

        variables: dict[str, str] = {}
        empty_row_count = 0

        for row in range(2, ws.max_row + 1):
            nom_a = self._format_cell_value(ws.cell(row=row, column=1).value)
            nom_c = self._format_cell_value(ws.cell(row=row, column=3).value)

            # Stop at section headers
            val_b = self._format_cell_value(ws.cell(row=row, column=2).value)
            if nom_a and self._is_section_header(nom_a, val_b):
                break

            # Skip header row (row 2 with "Nom", "Source" etc)
            if nom_a and nom_a.lower() in ("nom", "source"):
                continue

            # Count empty rows to detect end of data
            if not nom_a and not nom_c:
                empty_row_count += 1
                if empty_row_count >= 2:
                    break
                continue
            empty_row_count = 0

            # Primary variable (columns A + B)
            if nom_a:
                self._extract_variable(nom_a, ws, row, 2, source_wb, variables)

            # Secondary variable (columns C + D)
            if nom_c:
                self._extract_variable(nom_c, ws, row, 4, source_wb, variables)

        self._add_system_variables(variables)

        source_wb.close()
        config_wb_formulas.close()

        return variables

    def _extract_variable(self, nom, ws, row, col, source_wb, variables):
        """Extract a single variable from the config sheet (formula mode)."""
        source_val = ws.cell(row=row, column=col).value
        source = self._format_cell_value(source_val) if source_val is not None else None
        if not source:
            return

        if isinstance(source, str) and source.startswith("=") and "!" in source:
            value = resolve_formula(source, source_wb)
            if value:
                variables[nom] = value
        elif isinstance(source, str) and "[" in source and "]" in source:
            variables[f"_formula_{nom}"] = source
        else:
            variables[nom] = source

    def extract_societe_info(self) -> dict[str, SocieteInfo]:
        """Extract Societe Bailleur info.

        Looks for data in:
        1. A dedicated 'Société Bailleur' sheet (legacy: A=name, B=header, C=footer).
        2. Embedded in 'Rédaction LOI' sheet (new: A=name, B=header, D=footer).
        """
        result: dict[str, SocieteInfo] = {}

        # Strategy 1: dedicated sheet (read with data_only=True for values)
        config_wb_val = openpyxl.load_workbook(
            str(self._effective_config_path), data_only=True
        )
        for name in config_wb_val.sheetnames:
            if (
                "bailleur" in name.lower()
                and ("societ" in name.lower() or "société" in name.lower())
            ):
                ws = config_wb_val[name]
                for row in range(2, ws.max_row + 1):
                    nom = self._format_cell_value(ws.cell(row=row, column=1).value)
                    if not nom:
                        continue
                    header = (
                        self._format_cell_value(ws.cell(row=row, column=2).value) or nom
                    )
                    footer = (
                        self._format_cell_value(ws.cell(row=row, column=3).value) or ""
                    )
                    result[nom] = SocieteInfo(
                        nom=nom, header_text=header, footer_text=footer
                    )
                if result:
                    config_wb_val.close()
                    return result
        config_wb_val.close()

        # Strategy 2: embedded in Rédaction LOI — read in formula mode
        config_wb = openpyxl.load_workbook(
            str(self._effective_config_path), data_only=False
        )
        redaction_name = self._find_config_sheet(config_wb)
        if redaction_name:
            ws = config_wb[redaction_name]
            start_row = None
            for row in range(1, ws.max_row + 1):
                val_a = self._format_cell_value(ws.cell(row=row, column=1).value)
                val_b = self._format_cell_value(ws.cell(row=row, column=2).value)
                # The section header has "Société Bailleur" in A and "Entête" in B
                if val_a and "bailleur" in val_a.lower() and (
                    "societ" in val_a.lower() or "société" in val_a.lower()
                ) and val_b and "ent" in val_b.lower():
                    start_row = row + 1
                    break
            if start_row:
                for row in range(start_row, ws.max_row + 1):
                    nom = self._format_cell_value(ws.cell(row=row, column=1).value)
                    if not nom:
                        break
                    header = (
                        self._format_cell_value(ws.cell(row=row, column=2).value)
                        or nom
                    )
                    footer = (
                        self._format_cell_value(ws.cell(row=row, column=4).value)
                        or ""
                    )
                    result[nom] = SocieteInfo(
                        nom=nom, header_text=header, footer_text=footer
                    )

        config_wb.close()
        return result

    def extract_conditions_suspensives_mapping(self) -> dict[str, str]:
        """Extract the condition suspensive name -> LOI text mapping.

        In the new format, this is in the Rédaction LOI sheet after the
        Société Bailleur section: col A = fiche name, col C = LOI text.
        """
        config_wb = openpyxl.load_workbook(
            str(self._effective_config_path), data_only=True
        )
        mapping: dict[str, str] = {}

        redaction_name = self._find_config_sheet(config_wb)
        if not redaction_name:
            config_wb.close()
            return mapping

        ws = config_wb[redaction_name]
        in_conditions = False
        for row in range(1, ws.max_row + 1):
            val = self._format_cell_value(ws.cell(row=row, column=1).value)
            if val and "condition suspensive" in val.lower() and "fiche" in val.lower():
                in_conditions = True
                continue
            if in_conditions:
                if not val:
                    break
                text = self._format_cell_value(ws.cell(row=row, column=3).value) or ""
                mapping[val] = text

        config_wb.close()
        return mapping

    def enrich_from_inpi(self, variables: dict) -> Optional[InpiData]:
        """Enrich variables with INPI data if SIRET is available."""
        siret = (
            variables.get("SIRET", "")
            or variables.get("N° DE SIRET", "")
            or variables.get("N° de SIRET", "")
            or variables.get("Numéro SIRET", "")
            or variables.get("Siret Preneur", "")
        )
        if not siret or not validate_inpi_credentials():
            return None

        try:
            creds = get_inpi_credentials()
            client = INPIClient(
                username=creds["username"], password=creds["password"]
            )
            return client.get_company_info(siret)
        except Exception as e:
            logger.error(f"INPI enrichment error: {e}")
            return None

    def get_output_filename_loi(self, variables: dict) -> str:
        """Generate LOI output filename."""
        nom_preneur = variables.get("Nom Preneur", "Preneur")
        date_loi = variables.get("Date LOI", "")
        try:
            dt = datetime.strptime(date_loi, "%d/%m/%Y")
            date_str = dt.strftime("%Y %m %d")
        except (ValueError, TypeError):
            date_str = datetime.now().strftime("%Y %m %d")
        return f"{date_str} - LOI {nom_preneur}.docx"

    def get_output_filename_bail(self, variables: dict) -> str:
        """Generate BAIL output filename."""
        nom_preneur = (
            variables.get("Nom Preneur", "Preneur")
            .replace("/", "-")
            .replace("\\", "-")
        )
        date_loi = variables.get(
            "Date LOI", datetime.now().strftime("%d-%m-%Y")
        )
        date_loi = date_loi.replace("/", "-").replace("\\", "-")
        return f"BAIL - {nom_preneur} - {date_loi}.docx"
