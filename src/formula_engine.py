"""
HyperFormula Integration
Local formula evaluation and testing
"""
import logging
from typing import Dict, List, Any, Optional
from hyperformula import HyperFormula

from .config import HYPERFORMULA_LICENSE

logger = logging.getLogger(__name__)


class FormulaEngine:
    """HyperFormula engine for local formula evaluation"""
    
    def __init__(self, license_key: str = None):
        self.license_key = license_key or HYPERFORMULA_LICENSE
        self.hf = HyperFormula.buildEmpty({"licenseKey": self.license_key})
        self.sheets = {}  # Map sheet names to sheet IDs
        logger.info("HyperFormula engine initialized")
    
    def add_sheet(self, sheet_name: str, data: List[List[Any]] = None) -> int:
        """Add a sheet to the HyperFormula instance"""
        sheet_id = self.hf.addSheet(sheet_name)
        
        if data:
            # Convert data to HyperFormula format
            # HyperFormula expects data as list of lists
            self.hf.setCellContents(
                {"sheet": sheet_id, row: 0, col: 0},
                data
            )
        
        self.sheets[sheet_name] = sheet_id
        logger.info(f"Added sheet '{sheet_name}' with ID {sheet_id}")
        return sheet_id
    
    def set_cell_value(
        self,
        sheet_name: str,
        row: int,
        col: int,
        value: Any
    ):
        """Set a cell value in a sheet"""
        sheet_id = self.sheets.get(sheet_name)
        if sheet_id is None:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        
        self.hf.setCellContents(
            {"sheet": sheet_id, row: row, col: col},
            [[value]]
        )
    
    def set_cell_formula(
        self,
        sheet_name: str,
        row: int,
        col: int,
        formula: str
    ):
        """Set a formula in a cell (formula should start with =)"""
        if not formula.startswith("="):
            formula = "=" + formula
        
        self.set_cell_value(sheet_name, row, col, formula)
    
    def get_cell_value(
        self,
        sheet_name: str,
        row: int,
        col: int
    ) -> Any:
        """Get calculated value from a cell"""
        sheet_id = self.sheets.get(sheet_name)
        if sheet_id is None:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        
        try:
            return self.hf.getCellValue({"sheet": sheet_id, row: row, col: col})
        except Exception as e:
            logger.error(f"Error getting cell value: {e}")
            return None
    
    def evaluate_formula(
        self,
        formula: str,
        context_data: Optional[Dict[str, List[List[Any]]]] = None
    ) -> Any:
        """Evaluate a formula with optional context data"""
        # Create a temporary sheet for evaluation
        temp_sheet_name = "__temp_eval__"
        temp_sheet_id = self.hf.addSheet(temp_sheet_name)
        
        try:
            # Add context data if provided
            if context_data:
                for sheet_name, data in context_data.items():
                    if sheet_name != temp_sheet_name:
                        self.add_sheet(sheet_name, data)
            
            # Set formula in temp sheet
            self.hf.setCellContents(
                {"sheet": temp_sheet_id, row: 0, col: 0},
                [[formula if formula.startswith("=") else "=" + formula]]
            )
            
            # Get result
            result = self.hf.getCellValue({"sheet": temp_sheet_id, row: 0, col: 0})
            
            return result
        finally:
            # Clean up temp sheet
            try:
                self.hf.removeSheet(temp_sheet_id)
            except:
                pass
    
    def test_formula(
        self,
        formula: str,
        expected_result: Any = None,
        context_data: Optional[Dict[str, List[List[Any]]]] = None
    ) -> Dict[str, Any]:
        """Test a formula and return result with error info"""
        try:
            result = self.evaluate_formula(formula, context_data)
            
            test_result = {
                "formula": formula,
                "result": result,
                "success": True,
                "error": None
            }
            
            if expected_result is not None:
                test_result["matches_expected"] = result == expected_result
                test_result["expected"] = expected_result
            
            return test_result
        except Exception as e:
            return {
                "formula": formula,
                "result": None,
                "success": False,
                "error": str(e)
            }
    
    def batch_evaluate_formulas(
        self,
        formulas: List[Dict[str, Any]],
        context_data: Optional[Dict[str, List[List[Any]]]] = None
    ) -> List[Dict[str, Any]]:
        """Evaluate multiple formulas in batch"""
        results = []
        
        for formula_info in formulas:
            formula = formula_info.get("formula")
            sheet_name = formula_info.get("sheet_name", "__temp__")
            row = formula_info.get("row", 0)
            col = formula_info.get("col", 0)
            
            result = self.test_formula(
                formula,
                expected_result=formula_info.get("expected_result"),
                context_data=context_data
            )
            
            result.update({
                "sheet_name": sheet_name,
                "row": row,
                "col": col
            })
            
            results.append(result)
        
        return results
