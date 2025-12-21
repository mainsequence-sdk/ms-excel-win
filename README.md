# MS Excel Win

Starter scaffold for the MainSequence Excel integration powered by xlOil.

## Getting started

1. Install [uv](https://github.com/astral-sh/uv) if you do not have it already.
2. Create and activate a virtual environment:
   ```sh
   uv venv
   source .venv/bin/activate
   ```
3. Sync dependencies (adds `xloil` and `mainsequence`):
   ```sh
   uv sync
   ```
4. Run the package (placeholder entry point):
   ```sh
   uv run python -m ms_excel_win
   ```

The project targets Python 3.10+ and uses a `src/` layout. Adjust versions or add optional dependencies in `pyproject.toml` as the implementation evolves.

## Excel wrapper capabilities

- Ribbon tab “Main Sequence” with **Sign In** / **Sign Out** (or run `MS.LOGIN_DIALOG` / `MS.SIGN_OUT` commands).
- Tokens are stored at `~/.mainsequence/tokens.json`, refreshed automatically, and injected into `mainsequence.client`.
- Excel function `MS.GET_DATA_NODE` wraps `DataNodeStorage.get_data_between_dates_from_node_identifier` and returns a spilled table (header + rows).

### Authentication endpoints

Defaults assume JWT-style endpoints that return `{access, refresh}`. Override with environment variables if your paths differ:

```sh
set MS_AUTH_CREATE_URL=https://main-sequence.app/auth/jwt/create/
set MS_AUTH_REFRESH_URL=https://main-sequence.app/auth/jwt/refresh/
```

### Excel usage

Minimal call:

```
=MS.GET_DATA_NODE("node_identifier", DATE(2024,1,1), DATE(2024,2,1))
```

With optional filters (comma-separated or ranges):

```
=MS.GET_DATA_NODE("node_identifier", DATE(2024,1,1), DATE(2024,1,15), A2:A10, B2:B6)
```

If you are not signed in, the function returns an error prompting you to use the ribbon Sign In button or run `MS.LOGIN_DIALOG`.
