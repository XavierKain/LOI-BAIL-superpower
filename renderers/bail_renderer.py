"""BAIL Word document renderer: article placeholders, HTML tags, headings, TOC."""

import logging
import re
from pathlib import Path

from docx import Document
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
            if current_text:
                segments.append((current_text, dict(current_format)))
                current_text = ""

            is_closing = match.group(1) == "/"
            tag = match.group(2).lower()

            if is_closing:
                stack[tag] = max(0, stack[tag] - 1)
            else:
                stack[tag] += 1

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

    def render(
        self,
        articles: list[ArticleResult],
        variables: dict[str, str],
        output_path: str,
    ):
        """Generate the BAIL document."""
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

    def _replace_article_placeholders(
        self, doc, articles: list[ArticleResult], variables: dict
    ):
        """Replace {{ARTICLE_xxx}} placeholders with generated content."""
        ville = variables.get("Ville ou arrondissement", "")
        if "(" in ville:
            ville = ville.split("(")[0].strip()

        for paragraph in list(doc.paragraphs):
            text = paragraph.text.strip()

            # Handle {{VILLE}} and {{DATE_SIGNATURE}} anywhere in paragraph
            if "{{VILLE}}" in text or "{{DATE_SIGNATURE}}" in text:
                date_sig = variables.get("Date de signature", "")
                for run in paragraph.runs:
                    if "{{VILLE}}" in run.text:
                        run.text = run.text.replace("{{VILLE}}", ville)
                    if "{{DATE_SIGNATURE}}" in run.text:
                        run.text = run.text.replace("{{DATE_SIGNATURE}}", date_sig)
                continue

            if "{{" not in text:
                continue

            # Check for article placeholders
            for article in articles:
                placeholder = _ARTICLE_PLACEHOLDER_MAP.get(article.designation)
                if not placeholder or placeholder not in text:
                    continue

                if not article.contenu.strip():
                    continue

                self._insert_article_content(paragraph, article.contenu, doc)
                break

    def _insert_article_content(self, paragraph, content: str, doc):
        """Replace a paragraph with multi-paragraph article content."""
        # Split on double newlines for paragraph breaks
        parts = content.split("\n\n")

        # Further split on heading markers
        all_parts = []
        for part in parts:
            sub_parts = re.split(r"(\n(?=\*{2,4}))", part)
            for sp in sub_parts:
                stripped = sp.strip()
                if stripped:
                    all_parts.append(stripped)

        if not all_parts:
            return

        # Process first part: reuse existing paragraph
        self._render_paragraph_content(paragraph, all_parts[0], doc)

        # Process remaining parts: insert new paragraphs after
        prev_element = paragraph._element
        for idx, part in enumerate(all_parts[1:]):
            # Create new paragraph XML element
            new_p_element = etree.SubElement(
                prev_element.getparent(), qn("w:p")
            )
            prev_element.addnext(new_p_element)
            prev_element = new_p_element

            if not part.strip():
                # Empty paragraph for spacing
                continue

            # Wrap in a docx Paragraph object
            from docx.text.paragraph import Paragraph

            new_p = Paragraph(new_p_element, paragraph._parent)

            self._render_paragraph_content(new_p, part, doc)

    def _render_paragraph_content(self, paragraph, text: str, doc):
        """Render text with heading markers and formatting tags into a paragraph."""
        # Capture the template font from existing runs before clearing
        template_font_name = None
        template_font_size = None
        for run in paragraph.runs:
            if run.font.name:
                template_font_name = run.font.name
                template_font_size = run.font.size
                break

        # Fallback: get font from the Normal style
        if not template_font_name:
            try:
                normal_font = doc.styles["Normal"].font
                template_font_name = normal_font.name or "Calibri"
                template_font_size = normal_font.size
            except (KeyError, AttributeError):
                template_font_name = "Calibri"
        # Default to 11pt if no size found (template inherits from theme)
        if not template_font_size:
            from docx.shared import Pt as _Pt
            template_font_size = _Pt(11)

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
            try:
                paragraph.style = doc.styles[f"Heading {heading_level}"]
            except KeyError:
                pass
            paragraph.paragraph_format.left_indent = None
            paragraph.paragraph_format.first_line_indent = None
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        else:
            try:
                paragraph.style = doc.styles["Normal"]
            except (KeyError, AttributeError):
                pass

        # Clear existing runs
        for run in paragraph.runs:
            run.text = ""

        # Parse formatting tags and create runs with template font
        segments = parse_formatting_tags(clean_text)
        for seg_text, formatting in segments:
            if not seg_text:
                continue
            run = paragraph.add_run(seg_text)
            # Apply template font (preserve Calibri from template)
            if not heading_level:
                run.font.name = template_font_name
                if template_font_size:
                    run.font.size = template_font_size
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
