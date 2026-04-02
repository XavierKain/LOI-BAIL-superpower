"""Typed data structures for the LOI-BAIL generator."""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


@dataclass
class InpiData:
    """Company data retrieved from INPI."""
    nom_societe: str
    type_societe: str
    capital_social: str
    localite_rcs: str
    adresse_domiciliation: str
    president: str
    fonction: str
    status: str  # "success" | "failed"
    error_message: Optional[str] = None


@dataclass
class SocieteInfo:
    """Landlord company info for document headers/footers."""
    nom: str
    header_text: str
    footer_text: str


@dataclass
class ArticleResult:
    """Generated article content for BAIL documents."""
    designation: str
    contenu: str
    placeholders_manquants: list[str] = field(default_factory=list)


@dataclass
class DossierData:
    """All extracted variables for a dossier (Excel + INPI + derived)."""
    variables: dict[str, str]
    variables_derivees: dict[str, str]
    inpi_data: Optional[InpiData]
    source_file: Path
