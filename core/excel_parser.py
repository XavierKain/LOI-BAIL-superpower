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
