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
    # Aliases for BAIL conditions that use "du" variant
    "Durée du Bail": "Durée Bail",
    "Duree du Bail": "Duree Bail",
    "Option Accession": "Accession",
    "Honoraires Preneurs": "Honoraires Preneur",
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

    # Apply alias mapping (bidirectional: copy values both ways)
    for old_name, new_name in _VARIABLE_ALIASES.items():
        if old_name in result and new_name not in result:
            result[new_name] = result[old_name]
        elif old_name in result and new_name in result and not result[new_name]:
            result[new_name] = result[old_name]
        elif new_name in result and old_name not in result:
            result[old_name] = result[new_name]
        elif new_name in result and old_name in result and not result[old_name]:
            result[old_name] = result[new_name]

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


def _get_var(variables: dict, *keys: str) -> str:
    """Get a variable value trying multiple key names (accent variants)."""
    for key in keys:
        val = variables.get(key, "")
        if val and str(val).strip():
            return str(val).strip()
    return ""


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
    rue = _get_var(variables, "Numero et rue", "Numéro et rue")
    ville = _get_var(variables, "Ville ou arrondissement")
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
    duree_bail = _clean_number(_get_var(variables, "Duree Bail", "Durée Bail"))
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
            result["Date de signature"] = (date_aujourdhui + relativedelta(days=15)).strftime("%d/%m/%Y")
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
    duree_dg = _clean_number(_get_var(variables, "Duree DG", "Durée DG"))
    if loyer_base and duree_dg:
        montant_dg = (loyer_base / 12) * duree_dg
        result["Montant du DG"] = formater_nombre(int(montant_dg))

    # --- Periode DG ---
    if duree_dg:
        periode_map = {3: "quart", 4: "tiers", 6: "moitié"}
        result["Periode DG"] = periode_map.get(int(duree_dg), "")

    return result
