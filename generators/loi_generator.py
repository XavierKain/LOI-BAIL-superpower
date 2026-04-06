"""LOI-specific generation logic: optional sections, clear list."""

import logging
from typing import Optional

from core.models import DossierData
from generators.shared import est_societe

logger = logging.getLogger(__name__)


class LOIGenerator:
    """LOI document generation logic."""

    def __init__(self, dossier: DossierData, conditions_mapping: dict[str, str] = None):
        self.dossier = dossier
        self.conditions_mapping = conditions_mapping or {}
        self.clear_list: list[str] = []
        self._build_clear_list()

    def _build_clear_list(self):
        """Build list of placeholders to replace with empty string."""
        all_vars = {**self.dossier.variables, **self.dossier.variables_derivees}
        type_preneur = all_vars.get("Type Preneur", "")

        # Only clear PRESIDENT/FONCTION if preneur is NOT a societe
        # AND we don't have INPI data for these fields
        has_president = bool(all_vars.get("PRESIDENT DE LA SOCIETE", "").strip())
        if not est_societe(type_preneur) and not has_president:
            self.clear_list.extend(["PRESIDENT DE LA SOCIETE", "FONCTION INPI"])

        # Clear empty conditions suspensives
        for i in range(1, 5):
            key = f"Condition suspensive {i}"
            if not all_vars.get(key):
                self.clear_list.append(key)

    def get_all_variables(self) -> dict[str, str]:
        """Merge all variable sources into a single dict for rendering.

        Also maps condition suspensive values to their LOI text using
        the conditions mapping from the Excel config.
        """
        result = dict(self.dossier.variables)
        result.update(self.dossier.variables_derivees)

        # Replace condition suspensive raw names with mapped LOI texts
        if self.conditions_mapping:
            for i in range(1, 5):
                key = f"Condition suspensive {i}"
                raw_value = result.get(key, "")
                if raw_value and raw_value in self.conditions_mapping:
                    mapped = self.conditions_mapping[raw_value]
                    # Strip leading "- " to avoid double dash (template already has "- ")
                    result[key] = mapped.lstrip("- ").lstrip("-")

        return result

    def has_palier_data(self) -> bool:
        """Check if any palier data exists."""
        all_vars = {**self.dossier.variables, **self.dossier.variables_derivees}
        return any(all_vars.get(f"Montant du palier {i}") for i in range(1, 7))

    def has_conditions_suspensives(self) -> bool:
        """Check if any condition suspensive exists."""
        all_vars = {**self.dossier.variables, **self.dossier.variables_derivees}
        return any(
            all_vars.get(f"Condition suspensive {i}")
            for i in range(1, 5)
        )

    def has_honoraires_preneurs(self) -> bool:
        """Check if Honoraires Preneurs data exists."""
        all_vars = {**self.dossier.variables, **self.dossier.variables_derivees}
        return bool(
            all_vars.get("Honoraires Preneurs")
            or all_vars.get("Honoraires Preneur")
        )
