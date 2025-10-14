# Repository Guidelines

## Project Structure & Module Organization
- `comparison_app.py` hosts the Streamlit entry point, Gmail integration, and data-processing helpers; treat it as the primary module for feature additions.
- `.streamlit/` holds UI configuration; adjust theme tweaks there rather than in code.
- `FileMau/` contains Excel templates and mapping files (e.g., `Tong hop _ Report.xlsx`, `BangKe.xlsx`) that the app expects at runtime—keep filenames stable.
- Spreadsheet and CSV samples under `report (4)` and root-level archives are useful for regression checks; avoid committing large personal datasets.
- OAuth artifacts (`credentials.json`, `token.json`) live in the root for local runs and should be excluded from shared storage.

## Build, Test, and Development Commands
- `python -m venv .venv` then `.\.venv\Scripts\activate` – create and activate an isolated environment (PowerShell syntax).
- `pip install -r requirements.txt` – install Streamlit, pandas, Google API, and Excel dependencies aligned with production expectations.
- `streamlit run comparison_app.py` – launch the comparison dashboard locally; use `--server.port` to avoid clashes.
- `streamlit run comparison_app.py --client.toolbarMode minimal` – mimic deployment settings when validating UI spacing.

## Coding Style & Naming Conventions
- Follow PEP 8: 4-space indentation, snake_case for functions/variables, and CapWords for classes (rare in this codebase).
- Prefer descriptive helper names (`load_mapping_data`, `find_col`) and reuse existing utility patterns when introducing new data transformations.
- Keep DataFrame column names consistent with the original reports; introduce remapped aliases only at display time.
- Add concise comments only where intent is non-obvious, especially around Gmail API payloads or Excel template handling.

## Testing Guidelines
- No automated test suite exists; rely on manual flows.
- Use provided sample spreadsheets (`FileMau/*.xlsx`, `report (4)/*.xls`) to exercise uploads, merges, and PDF checks.
- Validate new logic by running the Streamlit app, uploading transport/express reports, and confirming summary totals match known values.
- When touching email delivery, dry-run by mocking credentials (set `st.session_state.credentials_json_content`) before hitting Gmail quota.

## Commit & Pull Request Guidelines
- Write imperative, scoped commit subjects such as `feat: surface detail drill-down for summaries` or `fix: guard missing pdf filenames`.
- Include a brief body outlining data sources touched (`FileMau`, OAuth tokens, etc.) and note any manual validation steps performed.
- Pull requests should describe the user scenario, list test commands executed, link to tracking issues, and attach screenshots/GIFs for UI changes.
- Run linting or formatting tools you introduce before submission and call out any new dependencies in the PR description.

## Security & Configuration Tips
- Treat `credentials.json` as sensitive; share via secure channels and rotate when leaving the project.
- Remove `token.json` before committing to avoid leaking refresh tokens.
- Document any new environment variables or secrets in the PR and update `README.md` with setup instructions for future agents.
