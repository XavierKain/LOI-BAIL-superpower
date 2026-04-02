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
    return (
        text.replace("\u2018", "'").replace("\u2019", "'")
        .replace("\u201c", '"').replace("\u201d", '"')
    )


def evaluer_condition(condition_str: Optional[str], donnees: dict[str, Any]) -> bool:
    """Evaluate a textual condition against data.

    Supports:
    - Empty/None -> True
    - "Si [X] = 'value'" -> string comparison (case-insensitive)
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
    non_vide_match = re.search(
        r"Si\s+\[([^\]]+)\]\s+non\s+(vide|nul)", condition, re.IGNORECASE
    )
    if non_vide_match:
        var_name = non_vide_match.group(1)
        value = donnees.get(var_name)
        if value is None:
            return False
        s = str(value).strip()
        return bool(s) and s != "0"

    # Comparison: "Si [Variable] operator value"
    comp_match = re.search(
        r"Si\s+\"?([^\"\[\]]+|\[[^\]]+\])\"?\s*(>=|<=|!=|=|>|<|superieur a|superieure a)\s*[\"']?([^\"']+)[\"']?",
        condition,
        re.IGNORECASE,
    )
    if comp_match:
        var_ref = comp_match.group(1).strip().strip("[]")
        operator = comp_match.group(2).strip().lower()
        expected = comp_match.group(3).strip()

        actual = donnees.get(var_ref, "")

        if operator == "=":
            return str(actual).strip().lower() == expected.lower()
        if operator == "!=":
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
        from generators.shared import formater_nombre

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
                nom_source = (
                    str(row.get("Nom Source", ""))
                    if pd.notna(row.get("Nom Source"))
                    else ""
                )
                if (
                    article_name == "Article preliminaire"
                    and "Condition" in nom_source
                    and "suspensive" in nom_source.lower()
                ):
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
            for match in re.findall(r"\[([^\]]+)\]", contenu):
                value = variables.get(match)
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

            desig = (
                str(rows[0].get("Designation", article_name))
                if rows
                else article_name
            )
            if pd.isna(desig):
                desig = article_name

            articles.append(
                ArticleResult(
                    designation=desig,
                    contenu=contenu,
                    placeholders_manquants=manquants,
                )
            )

        return articles
