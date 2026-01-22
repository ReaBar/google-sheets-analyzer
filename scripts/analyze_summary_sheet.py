#!/usr/bin/env python3
"""
Analyze Summary Sheet
Main script to analyze the summary sheet and identify issues
"""
import sys
import json
import logging
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.sheet_reader import SheetReader
from src.sheet_analyzer import SheetAnalyzer
from src.formula_engine import FormulaEngine
from src.config import SUMMARY_SHEET_ID

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def main():
    """Main analysis function"""
    logger.info("Starting summary sheet analysis...")
    logger.info(f"Target spreadsheet ID: {SUMMARY_SHEET_ID}")
    
    # Initialize components
    reader = SheetReader()
    formula_engine = FormulaEngine()
    analyzer = SheetAnalyzer(reader, formula_engine)
    
    # Analyze the spreadsheet
    logger.info("Analyzing spreadsheet structure...")
    analysis = analyzer.analyze_spreadsheet(SUMMARY_SHEET_ID)
    
    # Print summary
    print("\n" + "="*80)
    print("SPREADSHEET ANALYSIS SUMMARY")
    print("="*80)
    print(f"Title: {analysis['title']}")
    print(f"Spreadsheet ID: {analysis['spreadsheet_id']}")
    print(f"Number of sheets: {len(analysis['sheets'])}")
    print(f"Total issues found: {len(analysis['issues'])}")
    print(f"Total dependencies: {len(analysis['dependencies'])}")
    
    # Print sheet details
    print("\n" + "-"*80)
    print("SHEETS:")
    print("-"*80)
    for sheet in analysis['sheets']:
        print(f"\n  Sheet: {sheet['sheet_name']}")
        print(f"    Rows: {sheet['structure']['rowCount']}")
        print(f"    Columns: {sheet['structure']['columnCount']}")
        print(f"    Formulas: {len(sheet['formulas'])}")
        print(f"    Issues: {len(sheet['issues'])}")
        print(f"    Dependencies: {len(sheet['dependencies'])}")
        
        if sheet['issues']:
            print(f"\n    Issues found:")
            for issue in sheet['issues']:
                print(f"      - {issue['type']}: {issue.get('cell', 'N/A')}")
                if 'error' in issue:
                    print(f"        Error: {issue['error']}")
    
    # Analyze summary sheet logic specifically
    print("\n" + "-"*80)
    print("SUMMARY SHEET LOGIC ANALYSIS:")
    print("-"*80)
    
    # Find the summary sheet (usually the first sheet or one with "summary" in name)
    summary_sheet_name = None
    for sheet in analysis['sheets']:
        sheet_name_lower = sheet['sheet_name'].lower()
        if 'summary' in sheet_name_lower or sheet_name_lower == 'summary':
            summary_sheet_name = sheet['sheet_name']
            break
    
    if not summary_sheet_name and analysis['sheets']:
        summary_sheet_name = analysis['sheets'][0]['sheet_name']
    
    if summary_sheet_name:
        logger.info(f"Analyzing summary sheet logic: {summary_sheet_name}")
        logic = analyzer.find_summary_sheet_logic(SUMMARY_SHEET_ID, summary_sheet_name)
        
        print(f"\n  Summary Sheet: {summary_sheet_name}")
        print(f"    Monthly rows found: {len(logic['monthly_rows'])}")
        print(f"    Source sheets: {len(logic['source_sheets'])}")
        print(f"    Formulas: {len(logic['formulas'])}")
        
        if logic['monthly_rows']:
            print(f"\n    Monthly rows:")
            for row_info in logic['monthly_rows'][:5]:  # Show first 5
                print(f"      Row {row_info['row']}: {row_info['month']}")
            if len(logic['monthly_rows']) > 5:
                print(f"      ... and {len(logic['monthly_rows']) - 5} more")
        
        if logic['source_sheets']:
            print(f"\n    Source sheets:")
            for source in logic['source_sheets']:
                print(f"      - Spreadsheet ID: {source['spreadsheet_id']}")
                print(f"        Range: {source['range']}")
        
        if logic['formulas']:
            print(f"\n    Formulas found:")
            for cell, formula in list(logic['formulas'].items())[:5]:  # Show first 5
                print(f"      {cell}: {formula[:60]}...")
            if len(logic['formulas']) > 5:
                print(f"      ... and {len(logic['formulas']) - 5} more formulas")
    
    # Save detailed analysis to JSON
    output_file = Path(__file__).parent.parent / "analysis_output.json"
    with open(output_file, "w") as f:
        json.dump(analysis, f, indent=2, default=str)
    
    print(f"\n" + "="*80)
    print(f"Detailed analysis saved to: {output_file}")
    print("="*80)
    
    return analysis


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logger.info("Analysis interrupted by user")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Error during analysis: {e}", exc_info=True)
        sys.exit(1)
