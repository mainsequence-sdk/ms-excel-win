# Developer Guide

This project delivers Excel integrations via xlOil. Follow these conventions to keep changes predictable and reviewable.

## Workflow
- **Branching:** Start from `development`, create a feature branch named after the feature (e.g., `feature/ribbon-logging`).
- **PRs:** Open a pull request with the same feature name; keep scope focused.
- **Examples:** For every feature, add/update a dedicated Excel file under `examples/`. Use it to demonstrate and manually validate the new capability.
- **Sanity check:** Open the main example workbook in `examples/` and exercise the core flows as a general user to confirm nothing regressed.
- **Commits:** Prefer small, cohesive commits with clear messages; avoid mixing unrelated changes.

## Technical Standards
- **Python version:** Target `>=3.10` (per `pyproject.toml`).
- **Formatting/style:** PEP 8; keep code ASCII; add comments only when logic isn’t obvious.
- **Dependencies:** Declare in `pyproject.toml`. Regenerate `requirements.txt` (pinned) via `uv export --format requirements.txt --output-file requirements.txt` after dependency changes.
- **Auth/tokens:** Never commit real secrets. Tokens are stored user-local (`~/.mainsequence/tokens.json`); keep `.env` out of version control.
- **Logging:** Use the module logger (`ms_excel_win.ms_excel_wrappers`) for feature diagnostics; log at INFO for high-level events, DEBUG for detail.
- **Excel UX:** Status bar messages should be brief and reset to “Ready”. Ribbon XML must avoid broken image references. Functions should return Excel-friendly types (lists of lists for tables, scalars for single values).
- **Data handling:** Preserve index when exporting DataFrames; normalize datetime/timedelta to Python types before sending to Excel; avoid type conversions on empty DataFrames.

## Validation Checklist
- Create/update the feature-specific workbook under `examples/` showing how to use the new functions.
- Open the main example workbook and run the primary flows (sign in, data fetch, ribbon actions) to ensure no regressions.
- Verify ribbon loads (xlOil GUI available) and status bar updates work without errors.
- Run a quick import check or `python -m py_compile` on touched modules to catch syntax issues.
- Regenerate `requirements.txt` if dependencies changed, and ensure `pyproject.toml` matches.

## Release Prep
- Ensure branch is up to date with `development` before PR.
- Confirm `README.md`/`README_DEV.md` instructions are still accurate if setup changed.
- Attach the updated example workbook and include a short usage note in the PR description.
