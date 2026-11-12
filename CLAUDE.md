# Working rules for this repository

## Every change goes straight to `main`

The owner wants every change made in a session to land in git automatically:

1. After each completed change, run the tests: `python -m pytest -q`.
2. If they pass, commit and **push to `main`** (also push the current working branch, if different).
   The owner has given standing permission for this; no need to ask first.
3. Never leave finished work uncommitted or only on a side branch.
4. If tests fail, fix them before pushing; never push a broken app to `main`.

Pushing to `main` redeploys the Streamlit app (Streamlit Community Cloud watches `main`), so
`main` must always be in a working state.

## The app

- Streamlit app (`app.py`) that builds the Tag & Env Stock report from three uploads:
  Stock Summary, Order Status and Exploded BOM reports.
- The downloaded Excel must keep the layout of the manually prepared report: only the sheets
  **Tag & Env Stock** and **Ghana & AB & AG Stock**.
- Business configuration lives in `tag_env/config.py`; rules are documented in `README.md`.
