"""
Google Sheets API Reader
Reads sheets, formulas, and structure from Google Sheets
"""
import os
import logging
import pickle
from typing import Dict, List, Optional, Any
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.exceptions import RefreshError
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from .config import SCOPES, CLIENT_SECRET_PATH, TOKEN_PICKLE_PATH

logger = logging.getLogger(__name__)


class SheetReader:
    """Reads data and structure from Google Sheets"""
    
    def __init__(self, client_secret_path: str = None, token_pickle_path: str = None):
        self.client_secret_path = client_secret_path or CLIENT_SECRET_PATH
        self.token_pickle_path = token_pickle_path or TOKEN_PICKLE_PATH
        self.service = None
        self._authenticate()
    
    def _authenticate(self):
        """Authenticate with Google Sheets API"""
        creds = None
        
        # Try to load existing token
        if os.path.exists(self.token_pickle_path):
            with open(self.token_pickle_path, "rb") as token:
                creds = pickle.load(token)
        
        # If no valid credentials, get new ones
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except RefreshError:
                    logger.warning("Token expired or revoked. Re-authenticating...")
                    if os.path.exists(self.token_pickle_path):
                        os.remove(self.token_pickle_path)
                    creds = None
            
            if not creds:
                if not os.path.exists(self.client_secret_path):
                    raise FileNotFoundError(
                        f"Client secret file not found: {self.client_secret_path}\n"
                        f"Please download it from Google Cloud Console"
                    )
                
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.client_secret_path, SCOPES
                )
                creds = flow.run_local_server(port=0)
            
            # Save credentials for next run
            with open(self.token_pickle_path, "wb") as token:
                pickle.dump(creds, token)
        
        self.service = build("sheets", "v4", credentials=creds)
        logger.info("Successfully authenticated with Google Sheets API")
    
    def get_spreadsheet_info(self, spreadsheet_id: str) -> Dict[str, Any]:
        """Get basic information about a spreadsheet"""
        try:
            spreadsheet = self.service.spreadsheets().get(
                spreadsheetId=spreadsheet_id
            ).execute()
            
            return {
                "id": spreadsheet_id,
                "title": spreadsheet.get("properties", {}).get("title", ""),
                "sheets": [
                    {
                        "id": sheet.get("properties", {}).get("sheetId"),
                        "title": sheet.get("properties", {}).get("title", ""),
                        "index": sheet.get("properties", {}).get("index", 0),
                        "rowCount": sheet.get("properties", {}).get("gridProperties", {}).get("rowCount", 0),
                        "columnCount": sheet.get("properties", {}).get("gridProperties", {}).get("columnCount", 0),
                    }
                    for sheet in spreadsheet.get("sheets", [])
                ]
            }
        except HttpError as error:
            logger.error(f"Error getting spreadsheet info: {error}")
            raise
    
    def get_sheet_data(
        self,
        spreadsheet_id: str,
        sheet_name: str,
        range_name: Optional[str] = None
    ) -> List[List[Any]]:
        """Get data from a specific sheet"""
        try:
            range_str = f"{sheet_name}!{range_name}" if range_name else sheet_name
            result = self.service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=range_str
            ).execute()
            
            return result.get("values", [])
        except HttpError as error:
            logger.error(f"Error getting sheet data: {error}")
            raise
    
    def get_formulas(
        self,
        spreadsheet_id: str,
        sheet_name: str,
        range_name: Optional[str] = None
    ) -> Dict[str, str]:
        """Get formulas from a specific sheet (returns dict of A1 notation -> formula)"""
        try:
            range_str = f"{sheet_name}!{range_name}" if range_name else sheet_name
            
            # Get the sheet ID first
            spreadsheet_info = self.get_spreadsheet_info(spreadsheet_id)
            sheet_id = None
            for sheet in spreadsheet_info["sheets"]:
                if sheet["title"] == sheet_name:
                    sheet_id = sheet["id"]
                    break
            
            if sheet_id is None:
                raise ValueError(f"Sheet '{sheet_name}' not found")
            
            # Request formulas
            result = self.service.spreadsheets().get(
                spreadsheetId=spreadsheet_id,
                ranges=[range_str],
                includeGridData=True
            ).execute()
            
            formulas = {}
            sheet_data = result.get("sheets", [])[0]
            if "data" in sheet_data:
                for row_data in sheet_data["data"]:
                    if "rowData" in row_data:
                        for row_idx, row in enumerate(row_data["rowData"]):
                            if "values" in row:
                                for col_idx, cell in enumerate(row["values"]):
                                    if "userEnteredValue" in cell:
                                        user_value = cell["userEnteredValue"]
                                        if "formulaValue" in user_value:
                                            # Convert to A1 notation
                                            col_letter = chr(65 + col_idx)  # A, B, C, ...
                                            if col_idx >= 26:
                                                col_letter = chr(64 + (col_idx // 26)) + chr(65 + (col_idx % 26))
                                            a1_notation = f"{col_letter}{row_idx + 1}"
                                            formulas[a1_notation] = user_value["formulaValue"]
            
            return formulas
        except HttpError as error:
            logger.error(f"Error getting formulas: {error}")
            raise
    
    def get_sheet_structure(
        self,
        spreadsheet_id: str,
        sheet_name: str
    ) -> Dict[str, Any]:
        """Get complete structure of a sheet including headers, data, and formulas"""
        data = self.get_sheet_data(spreadsheet_id, sheet_name)
        formulas = self.get_formulas(spreadsheet_id, sheet_name)
        
        return {
            "data": data,
            "formulas": formulas,
            "rowCount": len(data),
            "columnCount": len(data[0]) if data else 0,
            "headers": data[0] if data else []
        }
