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
    "Taxe Fonciere": "Taxe foncière",
    "Participation travaux": "Participation Travaux",
    "Charges copro": "Charges Copro",
    "Droit d'entree": "Droit d'entrée",
    "Droit d'entrée": "Droit d'entrée",
    "Honoraires preneurs": "Honoraires Preneur",
    "Honoraires preneur": "Honoraires Preneur",
    "Honoraires bailleur": "Honoraires Bailleur",
    "Durée franchise": "Durée Franchise",
    "Duree Franchise": "Durée Franchise",
    "Duree ferme Bail": "Durée ferme Bail",
    "Durée ferme": "Durée ferme Bail",
    "Duree ferme": "Durée ferme Bail",
    "Periode paliers": "Periode paliers",
    "Durée du DG": "Durée DG",
    "Duree du DG": "Duree DG",
    "Montant du DG en lettres": "Montant du DG en lettres",
    # SIRET aliases
    "Siret Preneur": "N° DE SIRET",
    "SIRET": "N° DE SIRET",
    "N° de SIRET": "N° DE SIRET",
    "Numéro SIRET": "N° DE SIRET",
}

# Add palier aliases for 1-6 (all variants: with/without "du", capitalization)
for _i in range(1, 7):
    _VARIABLE_ALIASES[f"Montant Palier {_i}"] = f"Montant du palier {_i}"
    _VARIABLE_ALIASES[f"Montant du Palier {_i}"] = f"Montant du palier {_i}"
    _VARIABLE_ALIASES[f"Montant palier {_i}"] = f"Montant du palier {_i}"
    # "en lettres" variants will be resolved by the word engine
    _VARIABLE_ALIASES[f"Montant Palier {_i} en lettres"] = f"Montant du palier {_i} en lettres"
    _VARIABLE_ALIASES[f"Montant du Palier {_i} en lettres"] = f"Montant du palier {_i} en lettres"
    # Loyer année variants
    _VARIABLE_ALIASES[f"Loyer année {_i} en lettres"] = f"Loyer année {_i} en lettres"

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

    # Apply alias mapping (bidirectional: ensure all variants have the value)
    # Run twice to propagate through chains
    for _ in range(2):
        for old_name, new_name in _VARIABLE_ALIASES.items():
            old_val = result.get(old_name, "")
            new_val = result.get(new_name, "")
            if old_val and not new_val:
                result[new_name] = old_val
            elif new_val and not old_val:
                result[old_name] = new_val

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
    """Get a variable value trying multiple key names (case/accent insensitive)."""
    import unicodedata as _ud
    def _norm(s):
        return "".join(c for c in _ud.normalize("NFKD", str(s).lower()) if not _ud.combining(c))
    # Build a normalized lookup map once
    norm_map = {_norm(k): k for k in variables}
    for key in keys:
        # Exact match first
        val = variables.get(key, "")
        if val and str(val).strip():
            return str(val).strip()
        # Case/accent-insensitive fallback
        actual_key = norm_map.get(_norm(key))
        if actual_key:
            val = variables.get(actual_key, "")
            if val and str(val).strip():
                return str(val).strip()
    return ""


def _resolve_internal_formulas(variables: dict[str, str]):
    """Resolve internal formulas like =[Montant du loyer] - [Loyer année 1].

    These are stored by the parser as _formula_Name = formula_string.
    Resolves them and stores the result as the variable name.
    """
    import re as _re

    formula_keys = [k for k in variables if k.startswith("_formula_")]
    for fk in formula_keys:
        var_name = fk[len("_formula_"):]
        formula = variables[fk]

        # Skip if already calculated
        if var_name in variables and variables[var_name]:
            continue

        # Try to resolve: =[A] - [B], =[A] + [B], =[A] / N * [B], etc
        # Simple case: =[X] - [Y]
        match = _re.match(r"=\[([^\]]+)\]\s*([+\-*/])\s*\[([^\]]+)\]", formula)
        if match:
            var_a, op, var_b = match.group(1), match.group(2), match.group(3)
            val_a = _clean_number(_get_var(variables, var_a))
            val_b = _clean_number(_get_var(variables, var_b))
            if val_a is not None and val_b is not None:
                try:
                    if op == "-":
                        result_val = val_a - val_b
                    elif op == "+":
                        result_val = val_a + val_b
                    elif op == "*":
                        result_val = val_a * val_b
                    elif op == "/":
                        result_val = val_a / val_b if val_b != 0 else 0
                    else:
                        continue
                    variables[var_name] = formater_nombre(int(result_val)) if result_val == int(result_val) else formater_nombre(result_val)
                except (ValueError, TypeError, ZeroDivisionError):
                    pass
            continue

        # Complex: =[A]/N*[B]
        match2 = _re.match(r"=\[([^\]]+)\]\s*/\s*(\d+)\s*\*\s*\[([^\]]+)\]", formula)
        if match2:
            var_a, divisor, var_b = match2.group(1), int(match2.group(2)), match2.group(3)
            val_a = _clean_number(_get_var(variables, var_a))
            val_b = _clean_number(_get_var(variables, var_b))
            if val_a is not None and val_b is not None:
                result_val = (val_a / divisor) * val_b
                variables[var_name] = formater_nombre(int(result_val)) if result_val == int(result_val) else formater_nombre(result_val)


def calculer_variables_derivees(
    variables: dict[str, str],
    inpi_data: Optional[InpiData],
) -> dict[str, str]:
    """Calculate all derived variables from raw data. Called once, shared by LOI and BAIL."""
    result: dict[str, str] = {}

    # --- Resolve internal formulas (stored as _formula_Name) ---
    _resolve_internal_formulas(variables)

    # --- INPI data (only set non-empty values) ---
    if inpi_data and inpi_data.status == "success":
        for key, val in [
            ("NOM DE LA SOCIETE", inpi_data.nom_societe),
            ("TYPE DE SOCIETE", inpi_data.type_societe),
            ("CAPITAL SOCIAL", inpi_data.capital_social),
            ("LOCALITE RCS", inpi_data.localite_rcs),
            ("ADRESSE DE DOMICILIATION", inpi_data.adresse_domiciliation),
            ("PRESIDENT DE LA SOCIETE", inpi_data.president),
            ("FONCTION INPI", inpi_data.fonction),
        ]:
            if val and str(val).strip():
                result[key] = val

    # --- Type Preneur normalization ---
    # If Type Preneur is not one of the recognized values (SAS/SARL/EURL/
    # Société en formation/Personne Physique), default to "SAS".
    # Reason: a numeric/unknown value (e.g. legal form code) used to silently
    # match the first row in BAIL rules (Personne Physique) which is wrong for
    # most companies.
    type_preneur_raw = _get_var(variables, "Type Preneur")
    if type_preneur_raw:
        recognized = {"sas", "sarl", "eurl", "société en formation", "societe en formation",
                      "personne physique"}
        if type_preneur_raw.strip().lower() not in recognized:
            result["Type Preneur"] = "SAS"

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
    loyer_base = _clean_number(_get_var(variables, "Montant du loyer", "Montant du loyer "))
    palier_count = 0
    if loyer_base:
        for i in range(1, 7):
            loyer_annee = _clean_number(_get_var(variables, f"Loyer année {i}", f"Loyer annee {i}"))
            if loyer_annee is not None:
                remise = loyer_base - loyer_annee
                if remise > 0:
                    result[f"Montant du palier {i}"] = formater_nombre(int(remise))
                    palier_count = i
    if palier_count > 0:
        # Periode paliers = ordinal of the last palier year + 1
        ordinal_map = {1: "deuxième", 2: "troisième", 3: "quatrième", 4: "cinquième", 5: "sixième", 6: "septième"}
        result["Periode paliers"] = ordinal_map.get(palier_count, f"{palier_count + 1}ème")

    # --- Surfaces ---
    surface_totale = _clean_number(_get_var(variables, "Surface totale"))
    surface_rdc = _clean_number(_get_var(variables, "Surface RDC"))
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
