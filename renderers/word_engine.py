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

    @staticmethod
    def _strip_accents(text: str) -> str:
        """Remove French accents for fuzzy matching."""
        import unicodedata
        nfkd = unicodedata.normalize("NFKD", text)
        return "".join(c for c in nfkd if not unicodedata.combining(c))

    def _lookup_variable(self, name: str, variables: dict) -> Optional[str]:
        """Look up a variable by name with case-insensitive and accent-insensitive fallback."""
        # Exact match
        if name in variables:
            val = variables[name]
            return str(val) if val is not None and str(val).strip() else None

        # Case-insensitive fallback
        name_lower = name.lower()
        for key, val in variables.items():
            if key.lower() == name_lower:
                return str(val) if val is not None and str(val).strip() else None

        # Accent-insensitive fallback
        name_stripped = self._strip_accents(name_lower)
        for key, val in variables.items():
            if self._strip_accents(key.lower()) == name_stripped:
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
