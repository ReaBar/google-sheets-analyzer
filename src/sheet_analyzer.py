"""
Sheet Analyzer
Analyzes sheet structure, formulas, and identifies issues
"""
import logging
import re
from typing import Dict, List, Any, Optional
from datetime import datetime

from .sheet_reader import SheetReader
from .formula_engine import FormulaEngine

logger = logging.getLogger(__name__)


class SheetAnalyzer:
    """Analyzes Google Sheets for structure, formulas, and issues"""
    
    def __init__(self, sheet_reader: SheetReader, formula_engine: FormulaEngine = None):
        self.reader = sheet_reader
        self.formula_engine = formula_engine or FormulaEngine()
    
    def analyze_spreadsheet(self, spreadsheet_id: str) -> Dict[str, Any]:
        """Complete analysis of a spreadsheet"""
        logger.info(f"Analyzing spreadsheet: {spreadsheet_id}")
        
        # Get spreadsheet info
        info = self.reader.get_spreadsheet_info(spreadsheet_id)
        
        analysis = {
            "spreadsheet_id": spreadsheet_id,
            "title": info["title"],
            "sheets": [],
            "issues": [],
            "dependencies": []
        }
        
        # Analyze each sheet
        for sheet_info in info["sheets"]:
            sheet_analysis = self.analyze_sheet(
                spreadsheet_id,
                sheet_info["title"]
            )
            analysis["sheets"].append(sheet_analysis)
            
            # Collect issues
            if sheet_analysis.get("issues"):
                analysis["issues"].extend(sheet_analysis["issues"])
            
            # Collect dependencies
            if sheet_analysis.get("dependencies"):
                analysis["dependencies"].extend(sheet_analysis["dependencies"])
        
        return analysis
    
    def analyze_sheet(
        self,
        spreadsheet_id: str,
        sheet_name: str
    ) -> Dict[str, Any]:
        """Analyze a specific sheet"""
        logger.info(f"Analyzing sheet: {sheet_name}")
        
        structure = self.reader.get_sheet_structure(spreadsheet_id, sheet_name)
        
        analysis = {
            "sheet_name": sheet_name,
            "structure": {
                "rowCount": structure["rowCount"],
                "columnCount": structure["columnCount"],
                "headers": structure["headers"]
            },
            "formulas": structure["formulas"],
            "formula_analysis": [],
            "issues": [],
            "dependencies": []
        }
        
        # Analyze formulas
        if structure["formulas"]:
            formula_analysis = self._analyze_formulas(
                structure["formulas"],
                structure["data"]
            )
            analysis["formula_analysis"] = formula_analysis
            
            # Check for formula errors
            for formula_info in formula_analysis:
                if not formula_info.get("valid", True):
                    analysis["issues"].append({
                        "type": "invalid_formula",
                        "cell": formula_info["cell"],
                        "formula": formula_info["formula"],
                        "error": formula_info.get("error")
                    })
        
        # Find dependencies (IMPORTRANGE, QUERY, etc.)
        dependencies = self._find_dependencies(structure["formulas"], structure["data"])
        analysis["dependencies"] = dependencies
        
        # Check for common issues
        issues = self._check_common_issues(structure)
        analysis["issues"].extend(issues)
        
        return analysis
    
    def _analyze_formulas(
        self,
        formulas: Dict[str, str],
        data: List[List[Any]]
    ) -> List[Dict[str, Any]]:
        """Analyze formulas for validity and dependencies"""
        formula_analysis = []
        
        for cell, formula in formulas.items():
            analysis = {
                "cell": cell,
                "formula": formula,
                "valid": True,
                "error": None,
                "dependencies": []
            }
            
            # Extract dependencies from formula
            dependencies = self._extract_formula_dependencies(formula)
            analysis["dependencies"] = dependencies
            
            # Try to evaluate formula if we have data
            if self.formula_engine:
                try:
                    # Create context from data
                    context = {"Sheet1": data}
                    result = self.formula_engine.test_formula(formula, context_data=context)
                    
                    if not result["success"]:
                        analysis["valid"] = False
                        analysis["error"] = result["error"]
                except Exception as e:
                    logger.warning(f"Could not evaluate formula in {cell}: {e}")
            
            formula_analysis.append(analysis)
        
        return formula_analysis
    
    def _extract_formula_dependencies(self, formula: str) -> List[Dict[str, str]]:
        """Extract dependencies from a formula (IMPORTRANGE, QUERY, etc.)"""
        dependencies = []
        
        # IMPORTRANGE pattern
        importrange_pattern = r'IMPORTRANGE\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"\s*\)'
        matches = re.finditer(importrange_pattern, formula, re.IGNORECASE)
        for match in matches:
            dependencies.append({
                "type": "IMPORTRANGE",
                "spreadsheet_id": match.group(1),
                "range": match.group(2)
            })
        
        # QUERY pattern (may reference other sheets)
        query_pattern = r'QUERY\s*\(\s*([^,]+)\s*,'
        matches = re.finditer(query_pattern, formula, re.IGNORECASE)
        for match in matches:
            range_ref = match.group(1).strip()
            if "!" in range_ref:
                sheet_name = range_ref.split("!")[0].strip("'\"")
                dependencies.append({
                    "type": "QUERY",
                    "sheet": sheet_name,
                    "range": range_ref
                })
        
        return dependencies
    
    def _find_dependencies(
        self,
        formulas: Dict[str, str],
        data: List[List[Any]]
    ) -> List[Dict[str, Any]]:
        """Find all external dependencies"""
        dependencies = []
        
        # Check formulas
        for cell, formula in formulas.items():
            deps = self._extract_formula_dependencies(formula)
            dependencies.extend(deps)
        
        # Check data for references to other sheets
        for row_idx, row in enumerate(data):
            for col_idx, cell_value in enumerate(row):
                if isinstance(cell_value, str) and "!" in cell_value:
                    # Might be a sheet reference
                    parts = cell_value.split("!")
                    if len(parts) == 2:
                        dependencies.append({
                            "type": "cell_reference",
                            "sheet": parts[0].strip("'\""),
                            "range": parts[1]
                        })
        
        return dependencies
    
    def _check_common_issues(self, structure: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Check for common issues in sheet structure"""
        issues = []
        
        # Check for empty headers
        if structure.get("headers"):
            for idx, header in enumerate(structure["headers"]):
                if not header or str(header).strip() == "":
                    issues.append({
                        "type": "empty_header",
                        "column": idx + 1,
                        "severity": "warning"
                    })
        
        # Check for completely empty rows
        empty_rows = []
        for row_idx, row in enumerate(structure.get("data", [])):
            if row_idx == 0:  # Skip header
                continue
            if not any(cell for cell in row if cell):
                empty_rows.append(row_idx + 1)
        
        if empty_rows:
            issues.append({
                "type": "empty_rows",
                "rows": empty_rows,
                "severity": "info"
            })
        
        return issues
    
    def find_summary_sheet_logic(
        self,
        spreadsheet_id: str,
        sheet_name: str
    ) -> Dict[str, Any]:
        """Analyze summary sheet to understand monthly aggregation logic"""
        logger.info(f"Analyzing summary sheet logic: {sheet_name}")
        
        structure = self.reader.get_sheet_structure(spreadsheet_id, sheet_name)
        
        logic = {
            "sheet_name": sheet_name,
            "monthly_rows": [],
            "formulas": {},
            "source_sheets": [],
            "aggregation_pattern": None
        }
        
        # Find rows that represent months
        data = structure["data"]
        if len(data) < 2:
            return logic
        
        headers = data[0]
        
        # Look for date/month columns
        date_col_idx = None
        for idx, header in enumerate(headers):
            header_lower = str(header).lower()
            if any(keyword in header_lower for keyword in ["month", "date", "תאריך", "חודש"]):
                date_col_idx = idx
                break
        
        # Analyze each data row
        for row_idx, row in enumerate(data[1:], start=2):
            if date_col_idx is not None and date_col_idx < len(row):
                month_value = row[date_col_idx]
                if month_value:
                    logic["monthly_rows"].append({
                        "row": row_idx,
                        "month": month_value,
                        "data": row
                    })
        
        # Analyze formulas
        formulas = structure["formulas"]
        for cell, formula in formulas.items():
            logic["formulas"][cell] = formula
            
            # Check if formula references other sheets
            deps = self._extract_formula_dependencies(formula)
            for dep in deps:
                if dep["type"] == "IMPORTRANGE":
                    logic["source_sheets"].append({
                        "spreadsheet_id": dep["spreadsheet_id"],
                        "range": dep["range"]
                    })
        
        return logic
