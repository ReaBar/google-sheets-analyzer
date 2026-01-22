"""
Sheet Fixer
Fixes identified issues in Google Sheets
"""
import logging
from typing import Dict, List, Any, Optional
from googleapiclient.errors import HttpError

from .sheet_reader import SheetReader
from .formula_engine import FormulaEngine

logger = logging.getLogger(__name__)


class SheetFixer:
    """Fixes issues in Google Sheets"""
    
    def __init__(self, sheet_reader: SheetReader, formula_engine: FormulaEngine = None):
        self.reader = sheet_reader
        self.formula_engine = formula_engine or FormulaEngine()
        self.service = sheet_reader.service
    
    def fix_formula(
        self,
        spreadsheet_id: str,
        sheet_name: str,
        cell_range: str,
        new_formula: str
    ) -> bool:
        """Fix a formula in a specific cell"""
        try:
            # Ensure formula starts with =
            if not new_formula.startswith("="):
                new_formula = "=" + new_formula
            
            # Test formula first with HyperFormula
            test_result = self.formula_engine.test_formula(new_formula)
            if not test_result["success"]:
                logger.warning(
                    f"Formula test failed: {test_result['error']}. "
                    f"Proceeding anyway..."
                )
            
            # Update the formula in Google Sheets
            range_str = f"{sheet_name}!{cell_range}"
            body = {
                "values": [[new_formula]]
            }
            
            result = self.service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range_str,
                valueInputOption="USER_ENTERED",
                body=body
            ).execute()
            
            logger.info(f"Fixed formula in {range_str}")
            return True
            
        except HttpError as error:
            logger.error(f"Error fixing formula: {error}")
            return False
    
    def fix_multiple_formulas(
        self,
        spreadsheet_id: str,
        sheet_name: str,
        formula_fixes: List[Dict[str, str]]
    ) -> Dict[str, bool]:
        """Fix multiple formulas in batch"""
        results = {}
        
        for fix in formula_fixes:
            cell_range = fix["cell"]
            new_formula = fix["formula"]
            
            success = self.fix_formula(
                spreadsheet_id,
                sheet_name,
                cell_range,
                new_formula
            )
            
            results[cell_range] = success
        
        return results
    
    def update_cell_value(
        self,
        spreadsheet_id: str,
        sheet_name: str,
        cell_range: str,
        value: Any
    ) -> bool:
        """Update a cell value"""
        try:
            range_str = f"{sheet_name}!{cell_range}"
            body = {
                "values": [[value]]
            }
            
            result = self.service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range_str,
                valueInputOption="USER_ENTERED",
                body=body
            ).execute()
            
            logger.info(f"Updated cell {range_str}")
            return True
            
        except HttpError as error:
            logger.error(f"Error updating cell: {error}")
            return False
    
    def batch_update(
        self,
        spreadsheet_id: str,
        updates: List[Dict[str, Any]]
    ) -> bool:
        """Batch update multiple cells"""
        try:
            data = []
            for update in updates:
                range_str = f"{update['sheet_name']}!{update['range']}"
                values = update["values"]
                data.append({
                    "range": range_str,
                    "values": values
                })
            
            body = {
                "valueInputOption": "USER_ENTERED",
                "data": data
            }
            
            result = self.service.spreadsheets().values().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=body
            ).execute()
            
            logger.info(f"Batch updated {len(updates)} ranges")
            return True
            
        except HttpError as error:
            logger.error(f"Error in batch update: {error}")
            return False
