"""LOI-BAIL Document Generator — Streamlit application."""

import hashlib
import logging
import tempfile
from datetime import datetime
from pathlib import Path

import streamlit as st

from config.settings import (
    TEMPLATE_LOI,
    TEMPLATE_BAIL,
    CONFIG_LOI,
    CONFIG_BAIL,
    OUTPUT_DIR,
)
from core.excel_parser import ExcelParser
from core.models import DossierData
from generators.shared import normaliser_noms_variables, calculer_variables_derivees
from generators.loi_generator import LOIGenerator
from generators.bail_generator import BailGenerator
from renderers.loi_renderer import LOIRenderer
from renderers.bail_renderer import BailRenderer

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
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
    file_hash = hashlib.sha256(file_content).hexdigest()[:12]
    tmp_path = Path(tempfile.gettempdir()) / f"temp_{file_hash}.xlsx"
    try:
        tmp_path.write_bytes(file_content)
        parser = ExcelParser(str(tmp_path), config_path)
        variables = parser.extract_variables()
        societes = parser.extract_societe_info()
        inpi_data = parser.enrich_from_inpi(variables)
        output_name_loi = parser.get_output_filename_loi(variables)
        output_name_bail = parser.get_output_filename_bail(variables)
        return variables, societes, inpi_data, str(tmp_path), output_name_loi, output_name_bail
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
    cache_key = f"{uploaded_file.name}_{datetime.now().strftime('%Y-%m-%d')}"
    file_content = uploaded_file.read()

    with st.spinner("Extraction des données..."):
        variables, societes, inpi_data, source_path, output_name_loi, output_name_bail = _parse_excel(
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
                output_path = OUTPUT_DIR / output_name_loi
                OUTPUT_DIR.mkdir(exist_ok=True)
                renderer.render(generator, societes, str(output_path))

            with open(output_path, "rb") as f:
                st.download_button(
                    "⬇️ Télécharger la LOI",
                    f.read(),
                    file_name=output_name_loi,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="download_loi",
                )

    with tab_bail:
        if st.button("Générer le BAIL", key="btn_gen_bail"):
            with st.spinner("Génération du BAIL..."):
                all_vars = {**variables, **variables_derivees}
                bail_gen = BailGenerator(
                    str(CONFIG_BAIL), source_workbook_path=source_path
                )
                articles = bail_gen.generer_bail(all_vars)
                renderer = BailRenderer(str(TEMPLATE_BAIL))
                output_path = OUTPUT_DIR / output_name_bail
                OUTPUT_DIR.mkdir(exist_ok=True)
                renderer.render(articles, all_vars, str(output_path))

            with open(output_path, "rb") as f:
                st.download_button(
                    "⬇️ Télécharger le BAIL",
                    f.read(),
                    file_name=output_name_bail,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="download_bail",
                )


if __name__ == "__main__":
    main()
