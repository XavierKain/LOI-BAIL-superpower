# tests/test_shared.py
import pytest
from generators.shared import (
    normaliser_noms_variables,
    calculer_variables_derivees,
    est_societe,
    formater_nombre,
)
from core.models import InpiData


def test_normaliser_duree_bail():
    variables = {"Duree du Bail": "9"}
    result = normaliser_noms_variables(variables)
    assert result["Duree Bail"] == "9"


def test_normaliser_montant_palier():
    variables = {"Montant Palier 1": "5000"}
    result = normaliser_noms_variables(variables)
    assert result["Montant du palier 1"] == "5000"


def test_normaliser_date_prise_effet():
    variables = {"Date prise d'effet": "01/01/2026"}
    result = normaliser_noms_variables(variables)
    assert result["Date de prise d'effet"] == "01/01/2026"


def test_normaliser_typo_gapd():
    variables = {"Dure GAPD": "6"}
    result = normaliser_noms_variables(variables)
    assert result["Duree GAPD"] == "6"


def test_normaliser_case_insensitive_dedup():
    variables = {"Nom Preneur": "Test", "nom preneur": ""}
    result = normaliser_noms_variables(variables)
    assert result["nom preneur"] == "Test"


def test_adresse_rue_ville():
    variables = {"Numero et rue": "10 rue de la Paix", "Ville ou arrondissement": "Paris 2ème"}
    result = calculer_variables_derivees(variables, None)
    assert result["Adresse Locaux Loues"] == "10 rue de la Paix, Paris 2ème"


def test_adresse_only_rue():
    variables = {"Numero et rue": "10 rue de la Paix"}
    result = calculer_variables_derivees(variables, None)
    assert result["Adresse Locaux Loues"] == "10 rue de la Paix"


def test_paliers():
    variables = {"Montant du loyer": "160 000", "Loyer annee 1": "140000"}
    result = calculer_variables_derivees(variables, None)
    assert result["Montant du palier 1"] == "20 000"


def test_type_bail_9():
    variables = {"Duree Bail": "9"}
    result = calculer_variables_derivees(variables, None)
    assert result["Type Bail"] == "3/6/9"


def test_type_bail_10():
    variables = {"Duree Bail": "10"}
    result = calculer_variables_derivees(variables, None)
    assert result["Type Bail"] == "6/9/10"


def test_type_bail_other():
    variables = {"Duree Bail": "12"}
    result = calculer_variables_derivees(variables, None)
    assert result["Type Bail"] == "12 ans"


def test_date_signature():
    variables = {"Date d'aujourd'hui": "01/04/2026"}
    result = calculer_variables_derivees(variables, None)
    assert result["Date de signature"] == "22/04/2026"


def test_date_offre_valable():
    variables = {"Date d'aujourd'hui": "01/04/2026"}
    result = calculer_variables_derivees(variables, None)
    assert result["Date offre valable"] == "08/04/2026"


def test_date_prise_effet_plus_9_ans():
    variables = {"Date de prise d'effet": "01/06/2026"}
    result = calculer_variables_derivees(variables, None)
    # relativedelta(years=9) -> 01/06/2035
    assert result["Date de prise d'effet + 9 ans"] == "01/06/2035"


def test_surface_r_moins_1():
    variables = {"Surface totale": "200", "Surface RDC": "150"}
    result = calculer_variables_derivees(variables, None)
    assert result["Surface R-1"] == "50"


def test_montant_dg():
    variables = {"Montant du loyer": "120000", "Duree DG": "3"}
    result = calculer_variables_derivees(variables, None)
    assert result["Montant du DG"] == "30 000"


def test_periode_dg():
    variables = {"Duree DG": "3"}
    result = calculer_variables_derivees(variables, None)
    assert result["Periode DG"] == "quart"


def test_periode_dg_6():
    variables = {"Duree DG": "6"}
    result = calculer_variables_derivees(variables, None)
    assert result["Periode DG"] == "moitié"


def test_est_societe_sas():
    assert est_societe("SAS") is True


def test_est_societe_sarl():
    assert est_societe("SARL") is True


def test_est_societe_sasu():
    assert est_societe("SASU") is True


def test_est_societe_individual():
    assert est_societe("Personne physique") is False


def test_est_societe_passage_not_matched():
    assert est_societe("PASSAGE") is False


def test_est_societe_case_insensitive():
    assert est_societe("sas") is True
    assert est_societe("Sarl") is True


def test_formater_nombre_entier():
    assert formater_nombre(160000) == "160 000"


def test_formater_nombre_decimal():
    assert formater_nombre(1234.56) == "1 234,56"


def test_formater_nombre_string():
    assert formater_nombre("160000") == "160 000"


def test_formater_nombre_with_spaces():
    assert formater_nombre("160 000") == "160 000"


def test_inpi_data_merged():
    inpi = InpiData(
        nom_societe="Test SAS",
        type_societe="SAS",
        capital_social="10 000 €",
        localite_rcs="Paris",
        adresse_domiciliation="10 rue de la Paix",
        president="Jean Dupont",
        fonction="Président",
        status="success",
        error_message=None,
    )
    result = calculer_variables_derivees({}, inpi)
    assert result["NOM DE LA SOCIETE"] == "Test SAS"
    assert result["PRESIDENT DE LA SOCIETE"] == "Jean Dupont"
    assert result["FONCTION INPI"] == "Président"
