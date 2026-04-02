"""Tests for core.inpi_client."""

import pytest
from unittest.mock import patch, MagicMock

from core.inpi_client import INPIClient, _format_dirigeant_name, _extract_siren
from core.models import InpiData


# ---------------------------------------------------------------------------
# _extract_siren
# ---------------------------------------------------------------------------

class TestExtractSiren:
    def test_from_siret_14_digits(self):
        assert _extract_siren("12345678901234") == "123456789"

    def test_from_siren_9_digits(self):
        assert _extract_siren("123456789") == "123456789"

    def test_with_spaces(self):
        assert _extract_siren("123 456 789 01234") == "123456789"

    def test_invalid_length(self):
        assert _extract_siren("12345") is None

    def test_empty(self):
        assert _extract_siren("") is None

    def test_none_input(self):
        assert _extract_siren(None) is None


# ---------------------------------------------------------------------------
# _format_dirigeant_name
# ---------------------------------------------------------------------------

class TestFormatDirigeantName:
    def test_all_uppercase(self):
        assert _format_dirigeant_name("DUPONT", "JEAN") == "Jean Dupont"

    def test_mixed_case(self):
        assert _format_dirigeant_name("Dupont", "Jean") == "Jean Dupont"

    def test_no_prenom(self):
        assert _format_dirigeant_name("DUPONT", "") == "Dupont"

    def test_no_prenom_default(self):
        assert _format_dirigeant_name("DUPONT") == "Dupont"

    def test_mixed_case_no_transform(self):
        # Non-fully-uppercase strings are kept as-is
        assert _format_dirigeant_name("De La Fontaine", "Marie") == "Marie De La Fontaine"


# ---------------------------------------------------------------------------
# get_company_info - empty / invalid SIRET
# ---------------------------------------------------------------------------

class TestGetCompanyInfoValidation:
    def test_empty_siret(self):
        client = INPIClient(username="test", password="test")
        result = client.get_company_info("")
        assert isinstance(result, InpiData)
        assert result.status == "failed"
        assert "manquant" in result.error_message.lower()

    def test_invalid_siret(self):
        client = INPIClient(username="test", password="test")
        result = client.get_company_info("123")
        assert isinstance(result, InpiData)
        assert result.status == "failed"
        assert "invalide" in result.error_message.lower()


# ---------------------------------------------------------------------------
# get_company_info - mocked successful API response
# ---------------------------------------------------------------------------

class TestGetCompanyInfoAPISuccess:
    @patch("core.inpi_client.requests.post")
    @patch("core.inpi_client.requests.get")
    def test_api_success(self, mock_get, mock_post):
        # Mock authentication
        mock_post.return_value = MagicMock(
            status_code=200,
            json=lambda: {"token": "fake-token"},
        )
        # Mock company search
        mock_get.return_value = MagicMock(
            status_code=200,
            json=lambda: {
                "results": [
                    {
                        "formality": {
                            "content": {
                                "personneMorale": {
                                    "identite": {
                                        "entreprise": {
                                            "denomination": "TEST SAS",
                                        },
                                        "description": {},
                                    },
                                    "etablissementPrincipal": {
                                        "descriptionEtablissement": {},
                                        "adresse": {
                                            "commune": "PARIS",
                                            "codePostal": "75002",
                                            "voie": "10 RUE DE LA PAIX",
                                        },
                                    },
                                    "composition": {
                                        "pouvoirs": [
                                            {
                                                "roleEntreprise": "30",
                                                "typeDePersonne": "INDIVIDU",
                                                "actif": True,
                                                "individu": {
                                                    "descriptionPersonne": {
                                                        "nom": "DUPONT",
                                                        "prenoms": ["JEAN"],
                                                    }
                                                },
                                            }
                                        ]
                                    },
                                },
                                "natureCreation": {
                                    "formeJuridique": "SAS",
                                },
                            }
                        }
                    }
                ]
            },
        )

        client = INPIClient(username="test", password="test")
        result = client.get_company_info("123456789")

        assert isinstance(result, InpiData)
        assert result.status == "success"
        assert result.nom_societe == "TEST SAS"
        assert result.president == "Jean Dupont"
        assert result.fonction == "President"

    @patch("core.inpi_client.requests.post")
    @patch("core.inpi_client.requests.get")
    def test_api_not_found_falls_through(self, mock_get, mock_post):
        """When API returns 404, client tries scraping fallbacks, then returns failed."""
        mock_post.return_value = MagicMock(
            status_code=200,
            json=lambda: {"token": "fake-token"},
        )
        mock_get.return_value = MagicMock(status_code=404)

        client = INPIClient(username="test", password="test")

        # Patch both scraping methods to return None so we reach the final failure
        with patch.object(client, "_scrape_inpi_beautifulsoup", return_value=None), \
             patch.object(client, "_scrape_inpi_playwright", return_value=None):
            result = client.get_company_info("12345678901234")

        assert isinstance(result, InpiData)
        assert result.status == "failed"
        assert result.error_message  # has a message

    @patch("core.inpi_client.requests.post")
    @patch("core.inpi_client.requests.get")
    def test_api_success_with_address(self, mock_get, mock_post):
        """Verify address and localite_rcs extraction from API."""
        mock_post.return_value = MagicMock(
            status_code=200,
            json=lambda: {"token": "fake-token"},
        )
        mock_get.return_value = MagicMock(
            status_code=200,
            json=lambda: {
                "results": [
                    {
                        "formality": {
                            "content": {
                                "personneMorale": {
                                    "identite": {
                                        "entreprise": {"denomination": "ACME"},
                                        "description": {"montantCapital": 50000},
                                    },
                                    "etablissementPrincipal": {
                                        "descriptionEtablissement": {},
                                        "adresse": {
                                            "numVoie": "12",
                                            "typeVoie": "RUE",
                                            "voie": "DE LA PAIX",
                                            "codePostal": "75001",
                                            "commune": "PARIS 1ER ARRONDISSEMENT",
                                        },
                                    },
                                    "adresseEntreprise": {
                                        "adresse": {
                                            "numVoie": "12",
                                            "typeVoie": "RUE",
                                            "voie": "DE LA PAIX",
                                            "codePostal": "75001",
                                            "commune": "PARIS 1ER ARRONDISSEMENT",
                                        }
                                    },
                                    "composition": {"pouvoirs": []},
                                },
                                "natureCreation": {"formeJuridique": "SAS"},
                            }
                        }
                    }
                ]
            },
        )

        client = INPIClient(username="test", password="test")
        # Patch BS scraping to avoid network call for dirigeant fallback
        with patch.object(client, "_scrape_inpi_beautifulsoup", return_value=None):
            result = client.get_company_info("123456789")

        assert result.status == "success"
        assert result.localite_rcs == "PARIS"
        assert result.capital_social == "50 000 \u20ac"
        assert "75001" in result.adresse_domiciliation


# ---------------------------------------------------------------------------
# get_company_info - BeautifulSoup fallback
# ---------------------------------------------------------------------------

class TestBeautifulSoupFallback:
    @patch("core.inpi_client.requests.post")
    @patch("core.inpi_client.requests.get")
    def test_bs_fallback_returns_inpi_data(self, mock_get, mock_post):
        """When API fails but BS scraping succeeds, return InpiData with success."""
        mock_post.return_value = MagicMock(
            status_code=200,
            json=lambda: {"token": "fake-token"},
        )
        # API call returns 404
        mock_get.return_value = MagicMock(status_code=404)

        client = INPIClient(username="test", password="test")

        scraped = {
            "NOM DE LA SOCIETE": "SCRAPED SAS",
            "TYPE DE SOCIETE": "SAS",
            "CAPITAL SOCIAL": "10 000 \u20ac",
            "LOCALITE RCS": "LYON",
            "ADRESSE DE DOMICILIATION": "1 RUE TEST 69001 LYON",
            "PRESIDENT DE LA SOCIETE": "Jean Dupont",
            "FONCTION INPI": "President",
        }

        with patch.object(client, "_scrape_inpi_beautifulsoup", return_value=scraped):
            result = client.get_company_info("123456789")

        assert isinstance(result, InpiData)
        assert result.status == "success"
        assert result.nom_societe == "SCRAPED SAS"
        assert result.president == "Jean Dupont"
