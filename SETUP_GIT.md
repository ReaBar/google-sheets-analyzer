# Git Repository Setup Instructions

The project is ready to be pushed to a new git repository. Follow these steps:

## 1. Create a New Repository

Create a new repository on your preferred git hosting service:
- **GitHub**: https://github.com/new
- **GitLab**: https://gitlab.com/projects/new
- **Bitbucket**: https://bitbucket.org/repo/create

**Important**: Do NOT initialize the repository with a README, .gitignore, or license (we already have these).

## 2. Initialize Git (if not already done)

If git hasn't been initialized yet, run:

```bash
cd /Users/reabar/Repos/google-sheets-analyzer
git init
```

## 3. Add All Files

```bash
git add .
```

## 4. Create Initial Commit

```bash
git commit -m "Initial commit: Google Sheets Analyzer with HyperFormula and Clasp integration

- Google Sheets API integration for reading/writing sheets
- HyperFormula integration for local formula evaluation
- Sheet analyzer to identify issues and dependencies
- Apps Script manager via Clasp
- Main analysis script for summary sheets"
```

## 5. Add Remote Repository

Replace `YOUR_USERNAME` and `YOUR_REPO_NAME` with your actual values:

**For GitHub:**
```bash
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
```

**For GitLab:**
```bash
git remote add origin https://gitlab.com/YOUR_USERNAME/YOUR_REPO_NAME.git
```

## 6. Push to Remote

```bash
git branch -M main
git push -u origin main
```

## Files Included

The following files are included in the repository:
- All Python source code (`src/`)
- Analysis scripts (`scripts/`)
- Configuration files (`.clasp.json`, `requirements.txt`, `package.json`)
- Documentation (`README.md`)
- `.gitignore` (excludes sensitive files like `.env`, `token.pickle`, `client_secret.json`)

## Files Excluded (via .gitignore)

The following sensitive files are **NOT** included:
- `.env` - Environment variables
- `token.pickle` - OAuth tokens
- `client_secret.json` - Google API credentials
- `node_modules/` - Node.js dependencies
- `__pycache__/` - Python cache files
- `analysis_output.json` - Analysis results

## After Pushing

Once pushed, you can:
1. Share the repository with others
2. Clone it on other machines
3. Set up CI/CD if needed
4. Continue development with version control

## Troubleshooting

### "Repository already exists"
If you see this error, the remote might already be set. Check with:
```bash
git remote -v
```

### "Permission denied"
Make sure you have:
- Write access to the repository
- Correct authentication (SSH keys or personal access tokens)

### "Large files"
If you have large files, consider using Git LFS or excluding them in `.gitignore`.
