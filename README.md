# Google Sheets Analyzer

A comprehensive tool for analyzing, fixing, and managing Google Sheets with HyperFormula and Clasp integration.

## Features

- **Sheet Analysis**: Analyze sheet structure, formulas, and dependencies
- **Formula Testing**: Test formulas locally using HyperFormula before applying
- **Issue Detection**: Automatically identify broken formulas, missing data, and structural issues
- **Apps Script Management**: Manage Google Apps Script projects via Clasp
- **Batch Operations**: Fix multiple issues at once

## Prerequisites

1. **Python 3.7+**
   ```bash
   python --version
   ```

2. **Node.js 18+** (for Clasp)
   ```bash
   node --version
   ```

3. **Google Sheets API Credentials**
   - Reuses credentials from `transaction-to-google-sheet` project
   - Or set `CLIENT_SECRET_PATH` and `TOKEN_PICKLE_PATH` in `.env`

## Installation

### 1. Install Python Dependencies

```bash
cd google-sheets-analyzer
pip install -r requirements.txt
```

### 2. Install Node.js Dependencies (for Clasp)

```bash
npm install
```

### 3. Install Clasp Globally (if not already installed)

```bash
npm install -g @google/clasp
```

### 4. Enable Apps Script API

1. Visit: https://script.google.com/home/usersettings
2. Enable "Apps Script API"

### 5. Configure Environment

```bash
cp .env.example .env
# Edit .env with your settings (optional, defaults work if reusing credentials)
```

## Usage

### Analyze Summary Sheet

```bash
python scripts/analyze_summary_sheet.py
```

This will:
- Connect to your Google Sheet
- Analyze all sheets and formulas
- Identify issues and dependencies
- Save detailed analysis to `analysis_output.json`

### Using Clasp for Apps Script

#### Login to Clasp

```bash
clasp login
```

#### Clone Existing Apps Script

```bash
clasp clone <SCRIPT_ID> --rootDir apps-script
```

#### Clone Apps Scripts (into named folders)

Use the clone script to pull any Apps Script into a named folder. Each script gets its own directory with a self-contained `.clasp.json`.

**Hasolidit Overview Sheet** (spreadsheet ID `1p4cRifbq93yIx1dT145m_Qk4USVwwG7FWKQWia3qm-E`):
```bash
./scripts/clone_summary_sheet_script.sh 1ylc5EWSYG5sgsiQpsynBljmX6F3FN2_BtJeG2EHMibCrcf640bxe62DW hasolidit-overview
```

**Portfolio Sheet** (stocks, real estate, cash balance):
```bash
./scripts/clone_summary_sheet_script.sh 1M6B-GT3QyFry0tDmTdx9Q2BOeSzyqAMl2JYwyseYtMT4KAfn7N5zUKtj portfolio
```

To get a Script ID:
1. Open the spreadsheet
2. **Extensions → Apps Script**
3. In Apps Script: **Project Settings** (gear) → **Script ID**

#### Create New Apps Script

```bash
clasp create --type sheets --title "My Script" --parentId <SPREADSHEET_ID> --rootDir apps-script
```

#### Pull/Push Apps Script Code

Each script directory has its own `.clasp.json`, so you can work with any script independently:

```bash
# Work with Hasolidit Overview script
cd hasolidit-overview
clasp pull    # Get latest from Google
# Edit files...
clasp push    # Push changes to Google

# Work with Portfolio script
cd ../portfolio
clasp pull
clasp push
```

## Project Structure

```
google-sheets-analyzer/
├── src/
│   ├── sheet_reader.py          # Google Sheets API reader
│   ├── sheet_analyzer.py        # Analysis logic
│   ├── formula_engine.py         # HyperFormula integration
│   ├── sheet_fixer.py           # Fix broken formulas/data
│   ├── apps_script_manager.py   # Clasp integration
│   └── config.py                # Configuration
├── scripts/
│   ├── analyze_summary_sheet.py       # Main analysis script
│   └── clone_summary_sheet_script.sh  # Clone any Apps Script by name
├── hasolidit-overview/                # Hasolidit Overview Sheet Apps Script
│   └── .clasp.json                    # Self-contained config
├── portfolio/                          # Portfolio Sheet Apps Script (stocks, real estate, cash)
│   └── .clasp.json                    # Self-contained config
├── requirements.txt
├── package.json
└── README.md
```

## Configuration

### Environment Variables

Create a `.env` file (or use defaults):

```env
CLIENT_SECRET_PATH=../transaction-to-google-sheet/client_secret.json
TOKEN_PICKLE_PATH=../transaction-to-google-sheet/token.pickle
SUMMARY_SHEET_ID=1p4cRifbq93yIx1dT145m_Qk4USVwwG7FWKQWia3qm-E
HYPERFORMULA_LICENSE=gpl-v3
```

### Clasp Configuration

The `.clasp.json` file is automatically managed by clasp. It contains:

```json
{
  "scriptId": "your-script-id",
  "rootDir": "apps-script"
}
```

## How It Works

### 1. Sheet Reading
- Uses Google Sheets API to read data, formulas, and structure
- Reuses OAuth credentials from `transaction-to-google-sheet` project

### 2. Formula Analysis
- Extracts all formulas from sheets
- Identifies dependencies (IMPORTRANGE, QUERY, etc.)
- Tests formulas locally with HyperFormula

### 3. Issue Detection
- Invalid formulas
- Missing dependencies
- Empty rows/columns
- Structural issues

### 4. Formula Testing
- HyperFormula evaluates formulas locally
- Test formulas before applying to Google Sheets
- Catch errors before they break your sheets

### 5. Apps Script Management
- Clasp manages Apps Script projects
- Pull/push code changes
- Run functions remotely

## Example Workflow

1. **Analyze your sheet**:
   ```bash
   python scripts/analyze_summary_sheet.py
   ```

2. **Review analysis output**:
   - Check `analysis_output.json` for detailed findings
   - Identify what needs to be fixed

3. **Test fixes locally** (using HyperFormula):
   - Test new formulas before applying
   - Validate calculations

4. **Apply fixes**:
   - Use `sheet_fixer.py` to update formulas
   - Or use Apps Script via Clasp

5. **Verify**:
   - Re-run analysis to confirm fixes

## Troubleshooting

### "Client secret file not found"
- Ensure `client_secret.json` exists in the path specified in `.env`
- Or download from Google Cloud Console

### "Clasp not found"
- Install: `npm install -g @google/clasp`
- Or use: `npx @google/clasp`

### "Permission denied" (Google Sheets)
- Ensure you have access to the spreadsheet
- Re-authenticate: delete `token.pickle` and run again

### "Apps Script API not enabled"
- Visit: https://script.google.com/home/usersettings
- Enable "Apps Script API"

## License

This project uses HyperFormula under GPLv3 or commercial license.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request
