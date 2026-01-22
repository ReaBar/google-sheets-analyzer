"""
Configuration management for Google Sheets Analyzer
"""
import os
from pathlib import Path
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Paths
PROJECT_ROOT = Path(__file__).parent.parent
TRANSACTION_SHEET_PROJECT = PROJECT_ROOT.parent / "transaction-to-google-sheet"

# Google Sheets API Credentials
# Try to reuse credentials from transaction-to-google-sheet project
CLIENT_SECRET_PATH = os.getenv(
    "CLIENT_SECRET_PATH",
    str(TRANSACTION_SHEET_PROJECT / "client_secret.json")
)
TOKEN_PICKLE_PATH = os.getenv(
    "TOKEN_PICKLE_PATH",
    str(TRANSACTION_SHEET_PROJECT / "token.pickle")
)

# Google Sheets IDs
SUMMARY_SHEET_ID = os.getenv(
    "SUMMARY_SHEET_ID",
    "1p4cRifbq93yIx1dT145m_Qk4USVwwG7FWKQWia3qm-E"
)

# Google Sheets API Scopes
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/script.projects"
]

# HyperFormula License
HYPERFORMULA_LICENSE = os.getenv("HYPERFORMULA_LICENSE", "gpl-v3")

# Clasp Configuration
CLASP_CONFIG_PATH = PROJECT_ROOT / ".clasp.json"
