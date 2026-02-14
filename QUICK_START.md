# Quick Start Guide

## Step 1: Install Dependencies

### Install Python Packages
```bash
cd /Users/reabar/Repos/google-sheets-analyzer
pip install -r requirements.txt
```

### Install Node.js Packages (for Clasp)
```bash
npm install
```

### Install Clasp Globally (if not already installed)
```bash
npm install -g @google/clasp
```

## Step 2: Verify Setup

The project is configured to automatically use credentials from your `transaction-to-google-sheet` project. Verify they exist:

```bash
ls ../transaction-to-google-sheet/client_secret.json
ls ../transaction-to-google-sheet/token.pickle
```

If these files exist, you're ready to go! If not, you'll need to authenticate on first run.

## Step 3: Run Your First Analysis

Analyze your summary sheet:

```bash
python scripts/analyze_summary_sheet.py
```

**What happens:**
1. Connects to your Google Sheet (ID: `1p4cRifbq93yIx1dT145m_Qk4USVwwG7FWKQWia3qm-E`)
2. Reads all sheets and formulas
3. Identifies issues and dependencies
4. Prints a summary to console
5. Saves detailed analysis to `analysis_output.json`

## Step 4: Review Results

After running, check:
- **Console output**: Quick summary of findings
- **`analysis_output.json`**: Detailed analysis with all formulas, issues, and dependencies

## Common Use Cases

### Use Case 1: Analyze a Specific Sheet

Edit `scripts/analyze_summary_sheet.py` and change the `SUMMARY_SHEET_ID` or create a new script:

```python
from src.sheet_reader import SheetReader
from src.sheet_analyzer import SheetAnalyzer

reader = SheetReader()
analyzer = SheetAnalyzer(reader)

# Analyze any spreadsheet
analysis = analyzer.analyze_spreadsheet("YOUR_SPREADSHEET_ID")
print(analysis)
```

### Use Case 2: Test a Formula Before Applying

```python
from src.formula_engine import FormulaEngine

engine = FormulaEngine()

# Test a formula
result = engine.test_formula("=SUM(A1:A10)")
print(f"Result: {result['result']}")
print(f"Success: {result['success']}")
```

### Use Case 3: Fix a Broken Formula

```python
from src.sheet_reader import SheetReader
from src.sheet_fixer import SheetFixer

reader = SheetReader()
fixer = SheetFixer(reader)

# Fix a formula
fixer.fix_formula(
    spreadsheet_id="YOUR_SHEET_ID",
    sheet_name="Sheet1",
    cell_range="A1",
    new_formula="=SUM(B1:B10)"
)
```

### Use Case 4: Manage Apps Script

```bash
# Login to clasp
clasp login

# Clone existing Apps Script
clasp clone <SCRIPT_ID> --rootDir apps-script

# Pull latest code
cd apps-script
clasp pull

# Edit files, then push
clasp push
```

## Troubleshooting

### First Run: Authentication

On first run, if you see authentication errors:
1. The script will open a browser window
2. Sign in with your Google account
3. Grant permissions
4. A `token.pickle` file will be created (already exists if using shared credentials)

### "Module not found" Error

Make sure you installed dependencies:
```bash
pip install -r requirements.txt
```

### "Clasp not found"

Install clasp:
```bash
npm install -g @google/clasp
```

### Permission Denied

If you get permission errors:
- Ensure you have access to the Google Sheet
- Check that the sheet ID is correct
- Re-authenticate by deleting `token.pickle` and running again

## Next Steps

1. **Run the analysis** on your summary sheet
2. **Review the output** in `analysis_output.json`
3. **Identify issues** that need fixing
4. **Test fixes** using HyperFormula
5. **Apply fixes** using `sheet_fixer.py` or Apps Script

## Getting Help

- Check `README.md` for detailed documentation
- Review `analysis_output.json` for detailed findings
- Check console output for error messages
