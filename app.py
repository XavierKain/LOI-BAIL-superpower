"""LOI-BAIL Document Generator — Streamlit application."""

import hashlib
import logging
import tempfile
import traceback
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


def _ensure_source_file(file_content: bytes, file_name: str) -> str:
    """Ensure the uploaded Excel file exists on disk. Returns path."""
    file_hash = hashlib.sha256(file_content).hexdigest()[:12]
    tmp_path = Path(tempfile.gettempdir()) / f"loi_bail_{file_hash}.xlsx"
    if not tmp_path.exists():
        tmp_path.write_bytes(file_content)
    return str(tmp_path)


@st.cache_data(show_spinner=False)
def _parse_excel(file_content: bytes, file_name: str, config_path: str, _cache_key: str):
    """Parse Excel file with daily cache invalidation."""
    source_path = _ensure_source_file(file_content, file_name)
    parser = ExcelParser(source_path, config_path)
    variables = parser.extract_variables()
    societes = parser.extract_societe_info()
    inpi_data = parser.enrich_from_inpi(variables)
    output_name_loi = parser.get_output_filename_loi(variables)
    output_name_bail = parser.get_output_filename_bail(variables)
    return variables, societes, inpi_data, source_path, output_name_loi, output_name_bail


def main():
    st.title("📄 Générateur de Documents Immobiliers")
    st.markdown("Génération automatique de LOI et BAIL à partir d'une Fiche de décision")

    st.markdown("---")

    st.markdown("""
Cette application génère automatiquement des documents LOI (Lettres d'Intention) et BAIL (Baux Commerciaux).

### Comment ça marche ?
1. **Uploadez** votre fichier Excel (Fiche de décision)
2. **Vérifiez** les données extraites et enrichies (INPI)
3. **Choisissez** : Générer LOI ou Générer BAIL (ou les deux !)
4. **Téléchargez** les fichiers DOCX générés
""")

    st.markdown("---")

    _check_required_files()

    # Upload
    st.header("1. Upload du fichier Excel")
    uploaded_file = st.file_uploader(
        "Choisissez votre fichier Excel (Fiche de décision)",
        type=["xlsx", "xls"],
    )

    if not uploaded_file:
        st.info("👆 Uploadez un fichier Excel pour commencer")
        return

    try:
        st.success(f"✅ Fichier chargé: {uploaded_file.name}")

        file_content = uploaded_file.getbuffer().tobytes()
        cache_key = f"{uploaded_file.name}_{datetime.now().strftime('%Y-%m-%d')}"

        with st.spinner("Extraction des données et enrichissement INPI..."):
            variables, societes, inpi_data, source_path, output_name_loi, output_name_bail = _parse_excel(
                file_content, uploaded_file.name, str(CONFIG_LOI), cache_key,
            )

        # Ensure source file exists on disk (may have been lost between reruns)
        source_path = _ensure_source_file(file_content, uploaded_file.name)

        # Normalize + derive
        variables = normaliser_noms_variables(variables)
        variables_derivees = calculer_variables_derivees(variables, inpi_data)

        # Merge all variables for display
        all_vars = {**variables, **variables_derivees}

        st.success(f"✅ {len(variables)} variables extraites et enrichies (données en cache)")

        # =============================================
        # Section 2: Data display
        # =============================================
        st.header("2. Données extraites et enrichies")

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Nom Preneur", all_vars.get("Nom Preneur", "Non défini"))
            st.metric("Société Bailleur", all_vars.get("Société Bailleur", all_vars.get("Societe Bailleur", "Non défini")))
        with col2:
            st.metric("Date LOI", all_vars.get("Date LOI", "Non défini"))
            montant = all_vars.get("Montant du loyer", "Non défini")
            st.metric("Montant du loyer", f"{montant} €" if montant != "Non défini" else "Non défini")
        with col3:
            duree = all_vars.get("Durée Bail", all_vars.get("Duree Bail", "Non défini"))
            st.metric("Durée Bail", f"{duree} ans" if duree != "Non défini" else "Non défini")
            st.metric("Enseigne", all_vars.get("Enseigne", "Non défini"))

        # INPI section
        siret = all_vars.get("N° DE SIRET", all_vars.get("SIRET", ""))
        if siret:
            st.markdown("---")
            inpi_ok = inpi_data and inpi_data.status == "success"

            if inpi_ok:
                st.success("🏢 Données INPI enrichies automatiquement ✅")
            else:
                st.warning("⚠️ Enrichissement INPI échoué")

            with st.expander("📊 Informations INPI", expanded=inpi_ok):
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**SIRET**")
                    st.text(siret)
                    st.markdown("**Nom de la société**")
                    st.text(all_vars.get("NOM DE LA SOCIETE", "Non disponible"))
                    st.markdown("**Type de société**")
                    st.text(all_vars.get("TYPE DE SOCIETE", "Non disponible"))
                with col2:
                    st.markdown("**Capital social**")
                    st.text(all_vars.get("CAPITAL SOCIAL", "Non disponible"))
                    st.markdown("**Localité RCS**")
                    st.text(all_vars.get("LOCALITE RCS", "Non disponible"))
                st.markdown("**Adresse de domiciliation**")
                st.text(all_vars.get("ADRESSE DE DOMICILIATION", "Non disponible"))
                st.markdown("**Président / Gérant**")
                st.text(all_vars.get("PRESIDENT DE LA SOCIETE", "Non disponible"))

        # All variables expander
        with st.expander("📋 Voir toutes les variables extraites", expanded=False):
            display_vars = {k: v for k, v in all_vars.items() if not k.startswith("_")}
            sorted_vars = dict(sorted(display_vars.items()))

            missing_count = sum(1 for v in display_vars.values() if not v or str(v).strip() == "")
            total_count = len(display_vars)

            if missing_count > 0:
                st.warning(f"⚠️ {missing_count}/{total_count} variables manquantes")
            else:
                st.success(f"✅ Toutes les {total_count} variables sont renseignées")

            for key, value in sorted_vars.items():
                c1, c2, c3 = st.columns([2, 3, 1])
                with c1:
                    st.markdown(f"**{key}**")
                with c2:
                    if value and str(value).strip():
                        st.text(str(value))
                    else:
                        st.markdown("*Non défini*")
                with c3:
                    st.markdown("✅" if value and str(value).strip() else "⚠️")

        st.markdown("---")

        # =============================================
        # Section 3: Generation
        # =============================================
        st.header("3. Génération des documents")
        st.info("💡 **Info**: Grâce au cache, après la première génération, les suivantes seront quasi-instantanées ! La barre de chargement indique la progression.")

        col_loi, col_bail = st.columns(2)

        # LOI
        with col_loi:
            st.markdown("### 📄 Lettre d'Intention")
            st.markdown("""
            - Enrichissement INPI automatique
            - Sections optionnelles
            - Headers/Footers personnalisés
            """)

            if st.button("🚀 Générer LOI", type="primary", use_container_width=True, key="btn_gen_loi"):
                try:
                    with st.spinner("⏳ Génération en cours..."):
                        dossier = DossierData(
                            variables=variables,
                            variables_derivees=variables_derivees,
                            inpi_data=inpi_data,
                            source_file=Path(source_path),
                        )
                        generator = LOIGenerator(dossier)
                        renderer = LOIRenderer(str(TEMPLATE_LOI))
                        output_path = OUTPUT_DIR / output_name_loi
                        OUTPUT_DIR.mkdir(exist_ok=True)
                        renderer.render(generator, societes, str(output_path))

                    st.success("✅ Document LOI généré avec succès!")

                    with open(output_path, "rb") as f:
                        st.download_button(
                            "📥 Télécharger le document LOI",
                            f.read(),
                            file_name=output_name_loi,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                            key="download_loi",
                            type="primary",
                        )

                except Exception as e:
                    st.error(f"❌ Erreur lors de la génération LOI: {str(e)}")
                    with st.expander("Détails de l'erreur"):
                        st.code(traceback.format_exc())

        # BAIL
        with col_bail:
            st.markdown("### 📜 Bail Commercial")
            st.markdown("""
            - 16 articles conditionnels
            - Variables dérivées automatiques
            - Logique complexe
            """)

            if st.button("🚀 Générer BAIL", type="primary", use_container_width=True, key="btn_gen_bail"):
                try:
                    with st.spinner("⏳ Génération en cours..."):
                        bail_gen = BailGenerator(
                            str(CONFIG_BAIL), source_workbook_path=source_path
                        )
                        articles = bail_gen.generer_bail(all_vars)

                    ok_count = sum(1 for a in articles if a.contenu.strip())
                    st.success(f"✅ {ok_count}/{len(articles)} articles générés avec contenu")

                    # Show article details
                    with st.expander("📝 Détail des articles générés", expanded=True):
                        for art in articles:
                            has_content = bool(art.contenu.strip())
                            icon = "✅" if has_content else "⚠️"
                            c1, c2, c3 = st.columns([3, 5, 1])
                            with c1:
                                st.markdown(f"**{art.designation}**")
                            with c2:
                                if has_content:
                                    preview = art.contenu[:80].replace("\n", " ")
                                    st.text(preview + ("..." if len(art.contenu) > 80 else ""))
                                else:
                                    st.markdown("*Pas de contenu (conditions non remplies)*")
                            with c3:
                                st.markdown(icon)
                            if art.placeholders_manquants:
                                st.caption(f"  Placeholders manquants: {', '.join(art.placeholders_manquants[:5])}")

                    with st.spinner("⏳ Finalisation du document Word..."):
                        renderer = BailRenderer(str(TEMPLATE_BAIL))
                        output_path = OUTPUT_DIR / output_name_bail
                        OUTPUT_DIR.mkdir(exist_ok=True)
                        renderer.render(articles, all_vars, str(output_path))

                    st.success("✅ Document BAIL généré avec succès!")

                    with open(output_path, "rb") as f:
                        st.download_button(
                            "📥 Télécharger le document BAIL",
                            f.read(),
                            file_name=output_name_bail,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                            key="download_bail",
                            type="primary",
                        )

                except Exception as e:
                    st.error(f"❌ Erreur lors de la génération BAIL: {str(e)}")
                    with st.expander("Détails de l'erreur"):
                        st.code(traceback.format_exc())

    except Exception as e:
        st.error(f"❌ Erreur lors du traitement du fichier: {str(e)}")
        with st.expander("Détails de l'erreur"):
            st.code(traceback.format_exc())

    # Footer
    st.markdown("---")
    st.markdown("""
<div style='text-align: center; color: gray; padding: 20px;'>
    <p>Générateur automatique de LOI et BAIL v2.0</p>
    <p>Développé par Xavier Kain</p>
</div>
""", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
