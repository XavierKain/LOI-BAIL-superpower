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

# Article generation order: (article_name_in_excel, designation_filter, output_key)
# article_name_in_excel: matches the "Article" column in the Excel rules
# designation_filter: if set, only rows with this Designation are used (for "Comparution" which has 2)
# output_key: the key used in ArticleResult.designation (maps to template placeholders)
_ARTICLES_ORDER = [
    ("Comparution", "Comparution Bailleur", "Comparution Bailleur"),
    ("Comparution", "Comparution Preneur", "Comparution Preneur"),
    ("Article préliminaire", None, "Article preliminaire"),
    ("Article 1", None, "Article 1"),
    ("Article 2", None, "Article 2"),
    ("Article 3", None, "Article 3"),
    ("Article 5.3", None, "Article 5.3"),
    ("Article 7.1", None, "Article 7.1"),
    ("Article 7.2", None, "Article 7.2"),
    ("Article  7.3", None, "Article  7.3"),  # Note: double space preserved from template
    ("Article 7.6", None, "Article 7.6"),
    ("Article 8", None, "Article 8"),
    ("Article 19", None, "Article 19"),
    ("Article 22.2", None, "Article 22.2"),
    ("Article 26", None, "Article 26"),
    ("Article 26.1", None, "Article 26.1"),
    ("Article 26.2", None, "Article 26.2"),
]


def _normalize_quotes(text: str) -> str:
    """Replace typographic quotes with straight quotes."""
    return (
        text.replace("\u2018", "'").replace("\u2019", "'")
        .replace("\u201c", '"').replace("\u201d", '"')
    )


def _strip_accents(text: str) -> str:
    """Remove accents for fuzzy matching."""
    import unicodedata
    nfkd = unicodedata.normalize("NFKD", text)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def _lookup_variable(name: str, donnees: dict) -> Any:
    """Look up a variable by name with case/accent-insensitive fallback."""
    # Exact match
    if name in donnees:
        return donnees[name]
    # Case-insensitive
    name_lower = name.lower()
    for key, val in donnees.items():
        if key.lower() == name_lower:
            return val
    # Accent-insensitive
    name_stripped = _strip_accents(name_lower)
    for key, val in donnees.items():
        if _strip_accents(key.lower()) == name_stripped:
            return val
    return None


def evaluer_condition(condition_str: Optional[str], donnees: dict[str, Any]) -> bool:
    """Evaluate a textual condition against data.

    Supports:
    - Empty/None -> True
    - "Si [X] = 'value'" or 'Si "X" = value' -> string comparison
    - "Si [X] > N" or "Si X supérieur à N" -> numeric comparison
    - "Si [X] non vide" -> non-empty check
    - "Si plusieurs conditions suspensives" -> count check
    """
    if not condition_str:
        return True
    # Handle pandas NaN and empty strings
    cond_s = str(condition_str).strip()
    if not cond_s or cond_s == "nan" or cond_s == "None":
        return True

    condition = _normalize_quotes(cond_s)

    # Special: multiple conditions suspensives
    if "plusieurs conditions suspensives" in condition.lower():
        count = sum(
            1 for i in range(1, 5)
            if donnees.get(f"Condition suspensive {i}")
            and str(donnees[f"Condition suspensive {i}"]).strip()
            and str(donnees[f"Condition suspensive {i}"]).lower() != "none"
        )
        return count > 1

    # "Si [Variable] non vide/nul" or "Si [Variable] non nul"
    # Must handle apostrophes in variable names like "Droit d'entrée"
    non_vide_match = re.search(
        r"Si\s+\[([^\]]+)\]\s+non\s+(vide|nul)", condition, re.IGNORECASE
    )
    if non_vide_match:
        var_name = non_vide_match.group(1).strip()
        value = _lookup_variable(var_name, donnees)
        if value is None:
            return False
        s = str(value).strip()
        return bool(s) and s != "0"

    # Normalize accented operators for matching
    # "supérieur à" -> "superieur a", "supérieure à" -> "superieure a"
    condition_normalized = _strip_accents(condition)

    # Comparison pattern: supports [Variable], "Variable", or bare Variable
    # Operators: =, !=, >, <, >=, <=, superieur a, superieure a
    comp_match = re.search(
        r"Si\s+[\[\"']?([^\]\"'=<>!]+?)[\]\"']?\s*(>=|<=|!=|=|>|<|superieur a|superieure a)\s*[\"']?([^\"']+?)[\"']?\s*$",
        condition_normalized,
        re.IGNORECASE,
    )
    if comp_match:
        var_ref = comp_match.group(1).strip()
        operator = comp_match.group(2).strip().lower()
        expected = comp_match.group(3).strip()

        actual = _lookup_variable(var_ref, donnees)
        if actual is None:
            actual = ""

        if operator == "=":
            return _strip_accents(str(actual).strip().lower()) == _strip_accents(expected.lower())
        if operator == "!=":
            return _strip_accents(str(actual).strip().lower()) != _strip_accents(expected.lower())

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
            self.source_workbook = openpyxl.load_workbook(
                source_workbook_path, data_only=True
            )

        self._load_rules()

    def _load_rules(self):
        """Load rules from Redaction BAIL.xlsx."""
        wb = openpyxl.load_workbook(self.config_path, data_only=True)
        ws = None
        for name in wb.sheetnames:
            if "bail" in name.lower() and ("redaction" in name.lower() or "rédaction" in name.lower()):
                ws = wb[name]
                break
        if ws is None:
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
                sheet, rng = source_str.lstrip("=").split("!", 1)
                sheet = sheet.strip("'")
                return resolve_cell_range(sheet, rng, self.source_workbook)
            return resolve_formula(source_str, self.source_workbook)
        return source_ref

    def _get_article_rows(self, article_name: str, designation: str = None) -> list[dict]:
        """Get all rows for an article (including continuation rows).

        If designation is set, only returns rows where the Designation column
        matches (used for "Comparution" which has both Bailleur and Preneur).
        """
        rows = []
        found = False
        for _, row in self.regles_df.iterrows():
            art = row.get("Article")
            desig = row.get("Désignation") or row.get("Designation")

            if pd.notna(art) and str(art).strip() == article_name:
                if designation:
                    desig_str = str(desig).strip() if pd.notna(desig) else ""
                    if desig_str != designation:
                        if found:
                            # We were in our section, hit a different designation -> stop
                            break
                        continue
                found = True
                rows.append(row)
            elif found and (pd.isna(art) or str(art).strip() == ""):
                rows.append(row)
            elif found and pd.notna(art) and str(art).strip() != article_name:
                break
        return rows

    @staticmethod
    def _strip_empty_palier_sections(contenu: str, variables: dict) -> str:
        """Remove palier paragraph blocks for palier years with no data.

        Each palier year is a block separated by \\n\\n. A block belongs to
        palier N if it references [Montant du palier N] or [Montant du Palier N]
        or mentions the Nth year ordinal and has no resolved palier value.
        """
        blocks = contenu.split("\n\n")
        filtered = []
        for block in blocks:
            # Check if this block references a palier placeholder
            palier_refs = re.findall(
                r"\[Montant du [Pp]alier (\d+)(?:\s+en lettres)?\]", block
            )
            if palier_refs:
                # Keep block only if at least one referenced palier has data
                has_data = False
                for p_num in palier_refs:
                    val = _lookup_variable(f"Montant du palier {p_num}", variables)
                    if val is not None and str(val).strip():
                        has_data = True
                        break
                if not has_data:
                    continue
            filtered.append(block)
        return "\n\n".join(filtered)

    def _parse_nom_source_variables(self, nom_source: str, donnees: dict) -> list[str]:
        """Parse Nom Source into individual variable names.

        Handles patterns like "Conditions suspensives 1, 2, 3, 4." ->
        ["Condition suspensive 1", "Condition suspensive 2", ...]
        """
        # Check for pattern "BASE 1, 2, 3, 4"
        pattern_match = re.match(r"^(.+?)\s+(\d+)(?:,\s*(\d+))*", str(nom_source))
        if pattern_match and "," in str(nom_source):
            base = pattern_match.group(1).strip().rstrip(".")
            # Singularize: "Conditions suspensives" -> "Condition suspensive"
            base_words = base.split()
            base_singular = " ".join(
                w.rstrip("s") if w.endswith("s") and len(w) > 1 else w
                for w in base_words
            )
            # Check if singular form exists in data
            if f"{base_singular} 1" in donnees:
                base = base_singular
            numbers = re.findall(r"\d+", str(nom_source))
            return [f"{base} {num}" for num in numbers]
        else:
            # Split by newline
            return [n.strip().rstrip(".") for n in str(nom_source).split("\n") if n.strip()]

    def _check_donnee_source_match(
        self, nom_source: str, donnee_source, article_name: str, donnees: dict
    ) -> bool:
        """Check if the data source value matches the expected value (lookup logic)."""
        noms = self._parse_nom_source_variables(nom_source, donnees)

        # Resolve formula if needed
        valeur_attendue = donnee_source
        if str(donnee_source).startswith("="):
            resolved = self._resolve_source(donnee_source)
            if resolved is not None:
                valeur_attendue = resolved
            else:
                return False

        # For ranges (list of values)
        if isinstance(valeur_attendue, list):
            for nom in noms:
                val = _lookup_variable(nom, donnees)
                if val and str(val) in valeur_attendue:
                    return True
            return False

        # Special case: conditions suspensives - check if at least one is non-empty
        if ("préliminaire" in article_name.lower()
                and "Condition" in nom_source
                and "suspensive" in nom_source.lower()):
            for nom in noms:
                val = _lookup_variable(nom, donnees)
                if val and str(val).strip():
                    return True
            return False

        # Simple value comparison
        for nom in noms:
            val = _lookup_variable(nom, donnees)
            if val is not None and str(val).strip() == str(valeur_attendue).strip():
                return True
        return False

    def generer_bail(self, variables: dict[str, Any]) -> list[ArticleResult]:
        """Generate all BAIL articles from rules and variables."""
        from generators.shared import formater_nombre

        articles = []

        for article_name, designation_filter, output_key in _ARTICLES_ORDER:
            rows = self._get_article_rows(article_name, designation_filter)
            if not rows:
                logger.warning(f"No rules found for {article_name}")
                continue

            textes = []
            manquants = []

            # Special handling: Article préliminaire with conditions suspensives
            # Process ONCE using the first row's options, not per-row
            is_cond_suspensive = False
            if "préliminaire" in article_name.lower() and rows:
                first_ns = str(rows[0].get("Nom Source", "")).strip() if pd.notna(rows[0].get("Nom Source")) else ""
                if "Condition" in first_ns and "suspensive" in first_ns.lower():
                    is_cond_suspensive = True
                    opt1 = str(rows[0].get("Entrée correspondante - Option 1", "")) if pd.notna(rows[0].get("Entrée correspondante - Option 1")) else ""
                    opt2 = str(rows[0].get("Entrée correspondante - Option 2", "")) if pd.notna(rows[0].get("Entrée correspondante - Option 2")) else ""
                    texte = generer_conditions_suspensives(variables, opt1, opt2)
                    if texte:
                        textes.append(texte)

            if not is_cond_suspensive:
                for row in rows:
                    donnee_source = row.get("Donnée source")
                    nom_source_raw = row.get("Nom Source")
                    nom_source = str(nom_source_raw).strip() if pd.notna(nom_source_raw) else ""
                    condition = row.get("Condition")
                    option1 = row.get("Entrée correspondante - Option 1", "")
                    condition2 = row.get("Condition Option 2")
                    option2 = row.get("Entrée correspondante - Option 2", "")

                    # Step 1: Check Donnee source / Nom Source lookup (if both present)
                    if pd.notna(donnee_source) and nom_source:
                        if not self._check_donnee_source_match(
                            nom_source, donnee_source, article_name, variables
                        ):
                            continue

                    # Step 2: Evaluate Condition -> Option 1
                    if evaluer_condition(condition, variables):
                        if pd.notna(option1) and str(option1).strip():
                            textes.append(str(option1))
                            continue

                    # Step 3: Evaluate Condition Option 2 -> Option 2
                    if evaluer_condition(condition2, variables):
                        if pd.notna(option2) and str(option2).strip():
                            textes.append(str(option2))

            contenu = "\n\n".join(textes)

            # Remove instruction/meta lines (Excel notes not meant for the document)
            contenu = re.sub(
                r"En fonction du n° de SIRET[^\n]*(?:Pappers|Validation)[^\n]*:?\s*\n?",
                "",
                contenu,
            ).strip()

            # For Article 26.1 (Paliers): remove paragraph blocks for paliers
            # that have no data (e.g. palier 3 when only 2 paliers exist)
            if output_key == "Article 26.1":
                contenu = self._strip_empty_palier_sections(contenu, variables)

            # Replace [Variable] placeholders in the generated text
            for match in re.findall(r"\[([^\]]+)\]", contenu):
                value = _lookup_variable(match, variables)
                if value is not None and str(value).strip():
                    try:
                        num = float(
                            str(value).replace(" ", "").replace(",", ".")
                            .replace("\u00a0", "")
                        )
                        formatted = formater_nombre(num)
                        contenu = contenu.replace(f"[{match}]", formatted)
                    except (ValueError, TypeError):
                        contenu = contenu.replace(f"[{match}]", str(value))
                else:
                    manquants.append(match)

            articles.append(
                ArticleResult(
                    designation=output_key,
                    contenu=contenu,
                    placeholders_manquants=manquants,
                )
            )

        return articles
