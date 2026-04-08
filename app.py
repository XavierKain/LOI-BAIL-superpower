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
    """Ensure the uploaded Excel file exists on disk. Always re-writes."""
    file_hash = hashlib.sha256(file_content).hexdigest()[:12]
    tmp_path = Path(tempfile.gettempdir()) / f"loi_bail_{file_hash}.xlsx"
    # Always write to ensure file is fresh (Streamlit Cloud may clean /tmp)
    tmp_path.write_bytes(file_content)
    return str(tmp_path)


def _parse_excel(file_content: bytes, file_name: str, config_path: str):
    """Parse Excel file. No caching to avoid stale data issues."""
    source_path = _ensure_source_file(file_content, file_name)
    parser = ExcelParser(source_path, config_path)
    variables = parser.extract_variables()
    societes = parser.extract_societe_info()
    inpi_data = parser.enrich_from_inpi(variables)
    cond_mapping = parser.extract_conditions_suspensives_mapping()
    output_name_loi = parser.get_output_filename_loi(variables)
    output_name_bail = parser.get_output_filename_bail(variables)
    return variables, societes, inpi_data, source_path, output_name_loi, output_name_bail, cond_mapping


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
        file_content = uploaded_file.getbuffer().tobytes()
        file_hash = hashlib.sha256(file_content).hexdigest()[:12]

        # Only parse once per file — store in session state
        if st.session_state.get("_parsed_hash") != file_hash:
            with st.spinner("Extraction des données et enrichissement INPI..."):
                variables, societes, inpi_data, source_path, output_name_loi, output_name_bail, cond_mapping = _parse_excel(
                    file_content, uploaded_file.name, str(CONFIG_LOI),
                )
                st.session_state["_parsed_hash"] = file_hash
                st.session_state["_parsed_data"] = (variables, societes, inpi_data, source_path, output_name_loi, output_name_bail, cond_mapping)
                st.session_state["_file_content"] = file_content
                st.session_state["_file_name"] = uploaded_file.name

        variables, societes, inpi_data, source_path, output_name_loi, output_name_bail, cond_mapping = st.session_state["_parsed_data"]
        file_content = st.session_state["_file_content"]

        st.success(f"✅ Fichier chargé: {uploaded_file.name}")

        # Ensure source file exists on disk (may have been lost between reruns)
        source_path = _ensure_source_file(file_content, uploaded_file.name)

        # Normalize + derive + re-normalize to propagate aliases to derived vars
        variables = normaliser_noms_variables(variables)
        variables_derivees = calculer_variables_derivees(variables, inpi_data)
        # Re-normalize the merged dict to populate all alias variants
        _merged = {**variables, **variables_derivees}
        _merged = normaliser_noms_variables(_merged)
        # Update derivees with any new aliases
        for k, v in _merged.items():
            if k not in variables and k not in variables_derivees:
                variables_derivees[k] = v

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
            montant_raw = all_vars.get("Montant du loyer", "")
            try:
                montant_fmt = f"{int(float(str(montant_raw).replace(' ', ''))):,}".replace(",", " ")
                st.metric("Montant du loyer", f"{montant_fmt} €")
            except (ValueError, TypeError):
                st.metric("Montant du loyer", montant_raw or "Non défini")
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

        # All variables expander — deduplicated
        with st.expander("📋 Voir toutes les variables extraites", expanded=False):
            # Deduplicate aggressively: collapse alias/casing/accent variants
            # AND collapse "stop word" variants (e.g. "Durée Bail"/"Durée du Bail",
            # "Date prise d'effet"/"Date de prise d'effet"/"Date de prise d'effet du bail")
            import unicodedata
            def _strip_acc(s):
                return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))

            _STOPWORDS = {"de", "du", "des", "le", "la", "les", "l", "d", "bail", "preneur"}

            def _canonical(name: str) -> str:
                """Aggressive normalization for dedup: lowercase, strip accents,
                remove stopwords, collapse spaces."""
                s = _strip_acc(name.lower())
                # Replace non-alphanumeric (except space and digit) with space
                import re as _re
                s = _re.sub(r"[^a-z0-9 ]+", " ", s)
                tokens = [t for t in s.split() if t and t not in _STOPWORDS]
                return " ".join(tokens)

            # Group variables by canonical key, prefer the version with a value
            groups: dict[str, tuple[str, str]] = {}
            for k, v in all_vars.items():
                if k.startswith("_"):
                    continue
                canon = _canonical(k)
                if not canon:
                    continue
                v_str = str(v).strip() if v else ""
                if canon in groups:
                    # Keep the one with a value (or the shorter name if both empty/equal)
                    existing_k, existing_v = groups[canon]
                    if v_str and not existing_v:
                        groups[canon] = (k, v_str)
                    elif v_str and existing_v and len(k) < len(existing_k):
                        groups[canon] = (k, v_str)
                else:
                    groups[canon] = (k, v_str)

            display_vars = {}
            for canon, (k, v_str) in groups.items():
                # Format numbers nicely for display
                display_val = v_str
                if display_val and display_val.replace(".", "").replace("-", "").isdigit():
                    try:
                        num = float(display_val)
                        if num == int(num):
                            display_val = f"{int(num):,}".replace(",", " ")
                        else:
                            display_val = f"{num:,.2f}".replace(",", " ").replace(".", ",")
                    except (ValueError, TypeError):
                        pass
                display_vars[k] = display_val

            sorted_vars = dict(sorted(display_vars.items()))
            missing_count = sum(1 for v in sorted_vars.values() if not v.strip())
            total_count = len(sorted_vars)

            if missing_count > 0:
                st.warning(f"⚠️ {missing_count}/{total_count} variables manquantes")
            else:
                st.success(f"✅ Toutes les {total_count} variables sont renseignées")

            for key, value in sorted_vars.items():
                c1, c2, c3 = st.columns([2, 3, 1])
                with c1:
                    st.markdown(f"**{key}**")
                with c2:
                    if value.strip():
                        st.text(value)
                    else:
                        st.markdown("*Non défini*")
                with c3:
                    st.markdown("✅" if value.strip() else "⚠️")

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
                        source_path_loi = _ensure_source_file(file_content, uploaded_file.name)
                        dossier = DossierData(
                            variables=variables,
                            variables_derivees=variables_derivees,
                            inpi_data=inpi_data,
                            source_file=Path(source_path_loi),
                        )
                        generator = LOIGenerator(dossier, conditions_mapping=cond_mapping)
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
                        # Ensure source file is on disk
                        source_path_bail = _ensure_source_file(file_content, uploaded_file.name)

                        # Always use config/redaction_bail.xlsx for BAIL rules:
                        # the config file contains the maintained bold/underline
                        # formatting tags (<b>, <u>) and heading markers (**/***).
                        # The source file is still used for data lookups (Hypothèses).
                        bail_config = str(CONFIG_BAIL)

                        bail_gen = BailGenerator(
                            bail_config, source_workbook_path=source_path_bail
                        )
                        articles = bail_gen.generer_bail(all_vars)

                    ok_count = sum(1 for a in articles if a.contenu.strip())
                    st.success(f"✅ {ok_count}/{len(articles)} articles générés avec contenu")

                    # Show key BAIL variables for debugging
                    with st.expander("🔍 Variables BAIL clés", expanded=False):
                        bail_keys = [
                            "Société Bailleur", "Type Preneur", "Nom Preneur",
                            "N° DE SIRET", "NOM DE LA SOCIETE", "TYPE DE SOCIETE",
                            "CAPITAL SOCIAL", "LOCALITE RCS", "ADRESSE DE DOMICILIATION",
                            "PRESIDENT DE LA SOCIETE", "FONCTION INPI",
                            "Durée Bail", "Durée ferme Bail", "Date de prise d'effet",
                            "Montant du loyer",
                        ]
                        # Dynamically add Loyer année / palier for all existing years
                        for _pi in range(1, 7):
                            ly = all_vars.get(f"Loyer année {_pi}", "")
                            mp = all_vars.get(f"Montant du palier {_pi}", "")
                            if ly or mp:
                                bail_keys.append(f"Loyer année {_pi}")
                                bail_keys.append(f"Montant du palier {_pi}")
                        bail_keys.extend([
                            "Periode paliers",
                            "Actualisation", "Paiement", "Accession",
                            "Droit d'entrée", "Durée DG", "Montant du DG",
                            "Durée Franchise", "Honoraires Preneur", "DPE",
                            "Destination", "Enseigne", "Restauration sans extraction",
                            "Condition suspensive 1", "Condition suspensive 2",
                            "Condition suspensive 3", "Condition suspensive 4",
                            "Adresse Locaux Loues", "Numero et rue",
                            "Ville ou arrondissement", "Surface totale", "Surface RDC",
                            "Taxe foncière", "Charges Copro", "Participation Travaux",
                            "Date de signature",
                        ])
                        # Deduplicate while preserving order
                        _seen_bk = set()
                        bail_keys_dedup = []
                        for bk in bail_keys:
                            if bk not in _seen_bk:
                                _seen_bk.add(bk)
                                bail_keys_dedup.append(bk)
                        for bk in bail_keys_dedup:
                            val = all_vars.get(bk, "")
                            icon = "✅" if val else "⚠️"
                            st.markdown(f"{icon} **{bk}** = {val if val else '*vide*'}")

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
