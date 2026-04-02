"""Centralized configuration. Supports .env (local) and Streamlit Secrets (production)."""

import os
from pathlib import Path
from dotenv import load_dotenv

env_path = Path(__file__).parent.parent / ".env"
load_dotenv(dotenv_path=env_path)


def _get_secret(key: str, default: str = "") -> str:
    """Get secret from Streamlit Secrets (production) or .env (local)."""
    try:
        import streamlit as st
        if hasattr(st, "secrets") and key in st.secrets:
            return st.secrets[key]
    except (ImportError, FileNotFoundError, KeyError):
        pass
    return os.getenv(key, default)


# File paths (relative to project root)
PROJECT_ROOT = Path(__file__).parent.parent
TEMPLATE_LOI = PROJECT_ROOT / "templates" / "Template LOI avec placeholder.docx"
TEMPLATE_BAIL = PROJECT_ROOT / "templates" / "Template BAIL avec placeholder.docx"
CONFIG_LOI = PROJECT_ROOT / "config" / "redaction_loi.xlsx"
CONFIG_BAIL = PROJECT_ROOT / "config" / "redaction_bail.xlsx"
OUTPUT_DIR = PROJECT_ROOT / "output"

# INPI
INPI_BASE_URL = "https://registre-national-entreprises.inpi.fr/api"
INPI_MAX_CALLS = 5
INPI_PERIOD = 60  # seconds
INPI_CACHE_DURATION = 3600  # 1 hour


def get_inpi_credentials() -> dict:
    """Return INPI credentials dict."""
    return {
        "username": _get_secret("INPI_USERNAME", ""),
        "password": _get_secret("INPI_PASSWORD", ""),
    }


def validate_inpi_credentials() -> bool:
    """Check that INPI credentials are configured."""
    creds = get_inpi_credentials()
    return bool(creds["username"] and creds["password"])
