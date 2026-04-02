from pathlib import Path
from core.models import DossierData, InpiData, SocieteInfo, ArticleResult


def test_dossier_data_creation():
    data = DossierData(
        variables={"Nom Preneur": "Test SAS"},
        variables_derivees={"Adresse Locaux Loues": "10 rue de Paris, Paris"},
        inpi_data=None,
        source_file=Path("/tmp/test.xlsx"),
    )
    assert data.variables["Nom Preneur"] == "Test SAS"
    assert data.inpi_data is None


def test_inpi_data_success():
    inpi = InpiData(
        nom_societe="Ma Societe", type_societe="SAS", capital_social="10 000 €",
        localite_rcs="Paris", adresse_domiciliation="10 rue de la Paix, 75002 Paris",
        president="Jean Dupont", fonction="Président", status="success", error_message=None,
    )
    assert inpi.status == "success"
    assert inpi.president == "Jean Dupont"


def test_inpi_data_failed():
    inpi = InpiData(
        nom_societe="", type_societe="", capital_social="", localite_rcs="",
        adresse_domiciliation="", president="", fonction="",
        status="failed", error_message="SIRET invalide",
    )
    assert inpi.status == "failed"


def test_societe_info():
    info = SocieteInfo(nom="Bailleur SCI", header_text="BAILLEUR SCI", footer_text="Ligne 1\nLigne 2")
    assert "\n" in info.footer_text


def test_article_result():
    article = ArticleResult(designation="Article 1", contenu="Le bail est conclu pour <b>9 ans</b>.", placeholders_manquants=["Montant du loyer"])
    assert len(article.placeholders_manquants) == 1
