"""
Client pour l'API INPI (Institut National de la Propriete Industrielle).
Recupere les informations des entreprises francaises via le RNE.

Returns InpiData (core.models) - never None.
"""

import logging
import re
import time
from typing import Optional

import requests
from ratelimit import limits, sleep_and_retry

from config import settings
from core.models import InpiData

try:
    from bs4 import BeautifulSoup

    SCRAPING_AVAILABLE = True
except ImportError:
    SCRAPING_AVAILABLE = False

try:
    from playwright.sync_api import sync_playwright

    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Role code -> label mapping
# ---------------------------------------------------------------------------
ROLES_LIBELLES = {
    "30": "President",
    "71": "President",
    "50": "Gerant",
    "10": "Directeur general",
}

# Priority order when scanning pouvoirs
_ROLES_DIRIGEANTS = ["30", "71", "50", "10"]

# Qualities that identify a real dirigeant on the scraped page
_QUALITES_DIRIGEANT = [
    "Gerant",
    "Gerant",
    "President",
    "President",
    "Directeur general",
    "Directeur general",
    "President du conseil d'administration",
    "President du conseil de surveillance",
]


# ---------------------------------------------------------------------------
# Module-level helpers (public for testability)
# ---------------------------------------------------------------------------

def _extract_siren(siret: str) -> Optional[str]:
    """Extract a 9-digit SIREN from a SIRET (14) or SIREN (9), stripping spaces.

    Returns None when the input is empty or has an unexpected length.
    """
    if not siret:
        return None
    clean = re.sub(r"\s", "", str(siret).strip())
    if len(clean) == 14 and clean.isdigit():
        return clean[:9]
    if len(clean) == 9 and clean.isdigit():
        return clean
    return None


def _format_dirigeant_name(nom: str, prenom: str = "") -> str:
    """Return *Prenom Nom* (unified format).

    - All-uppercase parts are `.title()`-d.
    - Empty prenom -> just the formatted nom.
    """
    nom_fmt = nom.title() if nom.isupper() else nom
    if not prenom:
        return nom_fmt
    prenom_fmt = prenom.title() if prenom.isupper() else prenom
    return f"{prenom_fmt} {nom_fmt}"


# ---------------------------------------------------------------------------
# INPI Client
# ---------------------------------------------------------------------------

class INPIClient:
    """Client pour interroger l'API INPI RNE."""

    def __init__(self, username: str = "", password: str = ""):
        self.base_url = settings.INPI_BASE_URL
        if username and password:
            self.username = username
            self.password = password
        else:
            creds = settings.get_inpi_credentials()
            self.username = creds["username"]
            self.password = creds["password"]
        self.token: Optional[str] = None
        self._token_expiry: float = 0

    # ---- auth ----

    def _authenticate(self) -> bool:
        if not self.username or not self.password:
            return False
        if self.token and time.time() < self._token_expiry:
            return True
        try:
            resp = requests.post(
                f"{self.base_url}/sso/login",
                json={"username": self.username, "password": self.password},
                headers={"Content-Type": "application/json"},
                timeout=10,
            )
            if resp.status_code == 200:
                self.token = resp.json().get("token")
                self._token_expiry = time.time() + settings.INPI_CACHE_DURATION
                return True
            logger.error("Auth INPI failed: %s", resp.status_code)
            return False
        except Exception as exc:
            logger.error("Auth INPI error: %s", exc)
            return False

    # ---- rate-limited request ----

    @sleep_and_retry
    @limits(calls=settings.INPI_MAX_CALLS, period=settings.INPI_PERIOD)
    def _make_request(self, endpoint: str, params: Optional[dict] = None) -> Optional[dict]:
        if not self._authenticate():
            return None
        try:
            resp = requests.get(
                f"{self.base_url}/{endpoint}",
                headers={
                    "Authorization": f"Bearer {self.token}",
                    "Content-Type": "application/json",
                },
                params=params,
                timeout=10,
            )
            if resp.status_code == 200:
                return resp.json()
            logger.warning("INPI %s returned %s", endpoint, resp.status_code)
            return None
        except Exception as exc:
            logger.error("INPI request error: %s", exc)
            return None

    # ---- search ----

    def _search_by_siren(self, siren: str) -> Optional[dict]:
        result = self._make_request("companies", params={"siren[]": siren})
        if result and isinstance(result, dict):
            # New API wraps in "results" list
            items = result.get("results", [])
            if items:
                return items[0]
        if result and isinstance(result, list) and result:
            return result[0]
        return None

    # ---- extract dirigeant from API data ----

    @staticmethod
    def _extract_dirigeant_from_api(personne_morale: dict) -> tuple:
        """Return (nom_dirigeant, fonction) from composition.pouvoirs or (None, None)."""
        try:
            pouvoirs = personne_morale.get("composition", {}).get("pouvoirs", [])
            for pouvoir in pouvoirs:
                role = pouvoir.get("roleEntreprise")
                if (
                    pouvoir.get("actif")
                    and role in _ROLES_DIRIGEANTS
                    and pouvoir.get("typeDePersonne") == "INDIVIDU"
                ):
                    desc = pouvoir.get("individu", {}).get("descriptionPersonne", {})
                    nom = desc.get("nom", "")
                    prenoms = desc.get("prenoms", [])
                    if nom:
                        prenom = prenoms[0] if prenoms else ""
                        dirigeant = _format_dirigeant_name(nom, prenom)
                        fonction = ROLES_LIBELLES.get(role, "Dirigeant")
                        return dirigeant, fonction
        except Exception as exc:
            logger.error("Error extracting dirigeant from API: %s", exc)
        return None, None

    # ---- BeautifulSoup scraping (full port) ----

    def _scrape_inpi_beautifulsoup(self, siren: str) -> Optional[dict]:
        """Scrape data.inpi.fr using BeautifulSoup + requests.

        The page renders company info server-side, so no JS engine is needed.
        """
        if not SCRAPING_AVAILABLE:
            logger.warning("BeautifulSoup not available")
            return None

        url = f"https://data.inpi.fr/entreprises/{siren}"
        try:
            headers = {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36"
                )
            }
            resp = requests.get(url, headers=headers, timeout=30)
            if resp.status_code != 200:
                return None

            soup = BeautifulSoup(resp.content, "html.parser")
            result: dict[str, str] = {}

            # 1. Company name
            h1 = soup.find("h1")
            if h1:
                nom = h1.get_text().strip()
                if " - SIREN" in nom:
                    nom = nom.split(" - SIREN")[0]
                if nom.startswith("Entreprise : "):
                    nom = nom.replace("Entreprise : ", "")
                result["NOM DE LA SOCIETE"] = nom.strip()

            # 2. Legal form
            forme_elements = soup.find_all(string=lambda s: s and "Forme juridique" in s)
            if forme_elements:
                parent = forme_elements[0].parent
                sibling = parent.find_next_sibling()
                if sibling:
                    result["TYPE DE SOCIETE"] = sibling.get_text(strip=True)

            # 3. Capital social  -> "145 131 987 EUR"
            capital_elements = soup.find_all(string=lambda s: s and "Capital social" in s)
            if capital_elements:
                parent = capital_elements[0].parent
                sibling = parent.find_next_sibling()
                if sibling:
                    capital_text = sibling.get_text(strip=True)
                    match = re.search(r"([\d\s]+)", capital_text)
                    if match:
                        montant = match.group(1).replace(" ", "").replace("\xa0", "")
                        montant_fmt = "{:,}".format(int(montant)).replace(",", " ")
                        result["CAPITAL SOCIAL"] = f"{montant_fmt} \u20ac"
                    else:
                        result["CAPITAL SOCIAL"] = capital_text.replace("EUR", "\u20ac").strip()

            # 4. Address
            adresse_elements = soup.find_all(string=lambda s: s and "Adresse du siege" in s)
            # Also try with accent
            if not adresse_elements:
                adresse_elements = soup.find_all(
                    string=lambda s: s and "Adresse du si\u00e8ge" in s
                )
            if adresse_elements:
                parent = adresse_elements[0].parent
                sibling = parent.find_next_sibling()
                if sibling:
                    adresse = sibling.get_text(strip=True)
                    result["ADRESSE DE DOMICILIATION"] = adresse

                    # 5. RCS locality (extract city after postal code)
                    parts = adresse.split()
                    for i, part in enumerate(parts):
                        if part.isdigit() and len(part) == 5:
                            if i + 1 < len(parts):
                                ville = " ".join(parts[i + 1 :])
                                ville = ville.replace(" 1ER ARRONDISSEMENT", "")
                                ville = ville.replace(" 2E ARRONDISSEMENT", "")
                                for j in range(3, 21):
                                    ville = ville.replace(f" {j}E ARRONDISSEMENT", "")
                                result["LOCALITE RCS"] = (
                                    ville.replace(" FRANCE", "").strip()
                                )
                                break

            # 6. Dirigeant from representants section
            gestion_h3 = soup.find("h3", id="representants")
            if gestion_h3:
                section_row = gestion_h3.find_parent("div", class_="row")
                if section_row:
                    blocs = section_row.find_all("div", class_="bloc-dirigeant")
                    if blocs:
                        dirigeants: list[dict[str, str]] = []
                        current: dict[str, str] = {}
                        for bloc in blocs:
                            paras = bloc.find_all("p")
                            if len(paras) >= 2:
                                label = paras[0].get_text().strip()
                                value = paras[1].get_text().strip()
                                if label in ["Nom, Prenom(s)", "Nom, Pr\u00e9nom(s)", "Denomination", "D\u00e9nomination"]:
                                    if current:
                                        dirigeants.append(current)
                                    current = {}
                                current[label] = value
                        if current:
                            dirigeants.append(current)

                        qualites_ok = [
                            "gerant", "g\u00e9rant",
                            "president", "pr\u00e9sident",
                            "directeur general", "directeur g\u00e9n\u00e9ral",
                            "president du conseil d'administration",
                            "pr\u00e9sident du conseil d'administration",
                            "president du conseil de surveillance",
                            "pr\u00e9sident du conseil de surveillance",
                        ]

                        for d in dirigeants:
                            qualite = d.get("Qualit\u00e9", d.get("Qualite", ""))
                            if "commissaire" in qualite.lower():
                                continue
                            is_dirigeant = any(
                                q in qualite.lower() for q in qualites_ok
                            )
                            if not is_dirigeant:
                                continue

                            dirigeant_name: Optional[str] = None
                            for denom_key in ["D\u00e9nomination", "Denomination"]:
                                if denom_key in d:
                                    dirigeant_name = d[denom_key]
                                    break

                            if dirigeant_name is None:
                                for np_key in ["Nom, Pr\u00e9nom(s)", "Nom, Prenom(s)"]:
                                    if np_key in d:
                                        raw = d[np_key]
                                        parts_np = [
                                            p.strip()
                                            for p in raw.split()
                                            if p.strip()
                                        ]
                                        if len(parts_np) >= 2:
                                            # scraped page: "NOM Prenom" -> unified "Prenom Nom"
                                            nom_part = parts_np[0]
                                            prenom_part = " ".join(parts_np[1:])
                                            dirigeant_name = _format_dirigeant_name(
                                                nom_part, prenom_part
                                            )
                                        elif len(parts_np) == 1:
                                            dirigeant_name = _format_dirigeant_name(
                                                parts_np[0]
                                            )
                                        break

                            if dirigeant_name:
                                result["PRESIDENT DE LA SOCIETE"] = dirigeant_name
                                result["FONCTION INPI"] = qualite or ""
                                break

            return result if result else None

        except Exception as exc:
            logger.error("BeautifulSoup scraping error: %s", exc)
            return None

    # ---- Playwright fallback (conditional import) ----

    def _scrape_inpi_playwright(self, siren: str) -> Optional[dict]:
        """Scrape data.inpi.fr using Playwright (headless Chromium).

        Used as a last-resort fallback when the static HTML scraping fails
        because the page requires JS rendering.
        """
        try:
            from playwright.sync_api import sync_playwright as _sync_pw
        except ImportError:
            logger.warning("Playwright not available")
            return None

        url = f"https://data.inpi.fr/entreprises/{siren}"
        try:
            result: dict[str, str] = {}
            with _sync_pw() as p:
                browser = p.chromium.launch(headless=True)
                page = browser.new_page()
                try:
                    page.goto(url, wait_until="domcontentloaded", timeout=60000)
                except Exception:
                    pass
                page.wait_for_timeout(3000)

                # Company name
                try:
                    h1 = page.locator("h1").first
                    if h1:
                        nom = h1.text_content().strip()
                        if " - SIREN" in nom:
                            nom = nom.split(" - SIREN")[0]
                        if nom.startswith("Entreprise : "):
                            nom = nom.replace("Entreprise : ", "")
                        result["NOM DE LA SOCIETE"] = nom.strip()
                except Exception:
                    pass

                # Legal form
                try:
                    el = page.locator("text=/Forme juridique/").first
                    if el:
                        sibling = el.locator("xpath=following-sibling::*[1]")
                        result["TYPE DE SOCIETE"] = sibling.text_content().strip()
                except Exception:
                    pass

                # Capital
                try:
                    el = page.locator("text=/Capital/").first
                    if el:
                        sibling = el.locator("xpath=following-sibling::*[1]")
                        cap = sibling.text_content().strip()
                        match = re.search(r"([\d\s]+)", cap)
                        if match:
                            montant = match.group(1).replace(" ", "").replace("\xa0", "")
                            montant_fmt = "{:,}".format(int(montant)).replace(",", " ")
                            result["CAPITAL SOCIAL"] = f"{montant_fmt} \u20ac"
                        else:
                            result["CAPITAL SOCIAL"] = cap.replace("EUR", "\u20ac").strip()
                except Exception:
                    pass

                # Address + RCS locality
                try:
                    el = page.locator("text=/Adresse du si/").first
                    if el:
                        sibling = el.locator("xpath=following-sibling::*[1]")
                        adresse = sibling.text_content().strip()
                        result["ADRESSE DE DOMICILIATION"] = adresse
                        parts = adresse.split()
                        for i, part in enumerate(parts):
                            if part.isdigit() and len(part) == 5 and i + 1 < len(parts):
                                ville = " ".join(parts[i + 1 :])
                                ville = ville.replace(" 1ER ARRONDISSEMENT", "")
                                ville = ville.replace(" 2E ARRONDISSEMENT", "")
                                for j in range(3, 21):
                                    ville = ville.replace(f" {j}E ARRONDISSEMENT", "")
                                result["LOCALITE RCS"] = ville.replace(" FRANCE", "").strip()
                                break
                except Exception:
                    pass

                # Dirigeant
                try:
                    page.wait_for_selector("h3#representants", timeout=5000)
                    blocs = page.locator(".bloc-dirigeant").all()
                    if blocs:
                        premier = blocs[0]
                        paragraphes = premier.locator("p").all()
                        info: dict[str, str] = {}
                        for idx in range(0, len(paragraphes), 2):
                            if idx + 1 < len(paragraphes):
                                label = paragraphes[idx].text_content().strip()
                                val = paragraphes[idx + 1].text_content().strip()
                                info[label] = val
                        dirigeant = None
                        if "Denomination" in info or "D\u00e9nomination" in info:
                            dirigeant = info.get("D\u00e9nomination", info.get("Denomination"))
                        elif "Nom" in info and "Pr\u00e9nom" in info:
                            dirigeant = _format_dirigeant_name(info["Nom"], info["Pr\u00e9nom"])
                        elif "Nom" in info:
                            dirigeant = _format_dirigeant_name(info["Nom"])
                        if dirigeant:
                            result["PRESIDENT DE LA SOCIETE"] = dirigeant
                except Exception:
                    pass

                browser.close()

            return result if result else None
        except Exception as exc:
            logger.error("Playwright scraping error: %s", exc)
            return None

    # ---- main public method ----

    def get_company_info(self, siret: str) -> InpiData:
        """Fetch company info from INPI for a given SIRET or SIREN.

        Always returns an InpiData instance (never None).
        """
        empty = InpiData(
            nom_societe="",
            type_societe="",
            capital_social="",
            localite_rcs="",
            adresse_domiciliation="",
            president="",
            fonction="",
            status="failed",
            error_message="",
        )

        if not siret:
            empty.error_message = "SIRET manquant"
            return empty

        siren = _extract_siren(siret)
        if siren is None:
            empty.error_message = f"SIRET invalide (longueur: {len(str(siret).replace(' ', ''))})"
            return empty

        # --- attempt 1: INPI REST API ---
        try:
            company_data = self._search_by_siren(siren)
        except Exception as exc:
            logger.error("INPI API error: %s", exc)
            company_data = None

        if company_data:
            return self._parse_api_response(company_data, siren)

        # --- attempt 2: BeautifulSoup scraping ---
        try:
            scraped = self._scrape_inpi_beautifulsoup(siren)
            if scraped:
                return self._scraped_dict_to_inpi_data(
                    scraped,
                    note="Donnees recuperees via scraping BeautifulSoup",
                )
        except Exception as exc:
            logger.warning("BS scraping failed: %s", exc)

        # --- attempt 3: Playwright fallback ---
        try:
            scraped = self._scrape_inpi_playwright(siren)
            if scraped:
                return self._scraped_dict_to_inpi_data(
                    scraped,
                    note="Donnees recuperees via scraping Playwright",
                )
        except Exception as exc:
            logger.warning("Playwright scraping failed: %s", exc)

        empty.error_message = (
            "Entreprise non trouvee dans la base INPI "
            "(API et scraping echoues)"
        )
        return empty

    # ---- internal parsers ----

    def _parse_api_response(self, company_data: dict, siren: str) -> InpiData:
        """Convert raw API JSON into InpiData."""
        formality = company_data.get("formality", {})
        content = formality.get("content", {})
        pm = content.get("personneMorale", {})
        nature = content.get("natureCreation", {})

        # Name
        etab = pm.get("etablissementPrincipal", {})
        desc_etab = etab.get("descriptionEtablissement", {})
        identite = pm.get("identite", {})
        entreprise = identite.get("entreprise", {})
        nom_societe = (
            desc_etab.get("nomCommercial")
            or desc_etab.get("enseigne")
            or pm.get("denomination")
            or entreprise.get("denomination")
            or ""
        )

        # Legal form
        forme = nature.get("formeJuridique", "")

        # Address
        adresse_obj = (
            pm.get("adresseEntreprise", {}).get("adresse")
            or etab.get("adresse")
            or {}
        )
        adresse_str = ""
        commune = ""
        if isinstance(adresse_obj, dict):
            parts = []
            for key in ("numVoie", "indiceRepetition", "typeVoie", "voie", "codePostal", "commune"):
                val = adresse_obj.get(key)
                if val:
                    parts.append(str(val))
            adresse_str = " ".join(parts)
            commune = adresse_obj.get("commune", "")

        # RCS locality
        localite_rcs = ""
        if commune:
            localite_rcs = commune
            localite_rcs = localite_rcs.replace(" 1ER ARRONDISSEMENT", "")
            localite_rcs = localite_rcs.replace(" 2E ARRONDISSEMENT", "")
            for j in range(3, 21):
                localite_rcs = localite_rcs.replace(f" {j}E ARRONDISSEMENT", "")
            localite_rcs = localite_rcs.strip()

        # Capital
        desc = identite.get("description", {})
        montant = desc.get("montantCapital")
        capital = ""
        if montant is not None:
            try:
                capital = "{:,}".format(int(montant)).replace(",", " ") + " \u20ac"
            except (ValueError, TypeError):
                capital = str(montant)

        # Dirigeant (API)
        dirigeant, fonction = self._extract_dirigeant_from_api(pm)

        # Fallback: scrape dirigeant if API had none
        if not dirigeant:
            try:
                scraped = self._scrape_inpi_beautifulsoup(siren)
                if scraped:
                    dirigeant = scraped.get("PRESIDENT DE LA SOCIETE", "")
                    fonction = scraped.get("FONCTION INPI", "")
            except Exception:
                pass

        return InpiData(
            nom_societe=nom_societe,
            type_societe=forme,
            capital_social=capital,
            localite_rcs=localite_rcs,
            adresse_domiciliation=adresse_str,
            president=dirigeant or "",
            fonction=fonction or "",
            status="success",
        )

    @staticmethod
    def _scraped_dict_to_inpi_data(data: dict, note: str = "") -> InpiData:
        return InpiData(
            nom_societe=data.get("NOM DE LA SOCIETE", ""),
            type_societe=data.get("TYPE DE SOCIETE", ""),
            capital_social=data.get("CAPITAL SOCIAL", ""),
            localite_rcs=data.get("LOCALITE RCS", ""),
            adresse_domiciliation=data.get("ADRESSE DE DOMICILIATION", ""),
            president=data.get("PRESIDENT DE LA SOCIETE", ""),
            fonction=data.get("FONCTION INPI", ""),
            status="success",
            error_message=note or None,
        )


# ---------------------------------------------------------------------------
# Factory
# ---------------------------------------------------------------------------

def get_inpi_client() -> Optional[INPIClient]:
    """Return an INPIClient if credentials are configured, else None."""
    if not settings.validate_inpi_credentials():
        logger.warning("INPI credentials not configured")
        return None
    try:
        return INPIClient()
    except Exception as exc:
        logger.error("INPIClient init error: %s", exc)
        return None
