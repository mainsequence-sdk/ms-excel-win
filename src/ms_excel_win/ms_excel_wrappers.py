from __future__ import annotations

import base64
import json
import logging
import os
import threading
from collections.abc import Iterable
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
import requests

try:
    import xloil as xl
except ImportError as exc:  # pragma: no cover - xlOil drives this module
    raise RuntimeError("ms_excel_wrappers requires xlOil to be installed") from exc

try:  # xlOil versions differ on where ExcelGUI lives
    from xloil import ExcelGUI  # type: ignore
except Exception:  # pragma: no cover
    try:
        from xloil.gui import ExcelGUI  # type: ignore
    except Exception:  # pragma: no cover
        ExcelGUI = None

try:
    from PIL import Image
except Exception:  # pragma: no cover - ribbon image is optional
    Image = None

_LOG = logging.getLogger(__name__)

AUTH_CREATE_URL = os.environ.get(
    "MS_AUTH_CREATE_URL", "https://main-sequence.app/auth/jwt/create/"
)
AUTH_REFRESH_URL = os.environ.get(
    "MS_AUTH_REFRESH_URL", "https://main-sequence.app/auth/jwt/refresh/"
)
TOKEN_PATH = Path(os.environ.get("MS_TOKEN_PATH", Path.home() / ".mainsequence" / "tokens.json"))

_TOKEN_LOCK = threading.Lock()


def _token_file() -> Path:
    TOKEN_PATH.parent.mkdir(parents=True, exist_ok=True)
    return TOKEN_PATH


def _load_tokens() -> Optional[Dict[str, str]]:
    token_file = _token_file()
    if not token_file.exists():
        return None
    try:
        with token_file.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
            if not isinstance(data, dict):
                return None
            access = data.get("access")
            refresh = data.get("refresh")
            if access and refresh:
                return {"access": access, "refresh": refresh}
    except Exception as exc:  # pragma: no cover
        _LOG.debug("Could not read token file: %s", exc)
    return None


def _save_tokens(tokens: Dict[str, str]) -> None:
    try:
        with _token_file().open("w", encoding="utf-8") as handle:
            json.dump(tokens, handle)
    except Exception as exc:  # pragma: no cover
        _LOG.debug("Could not persist tokens: %s", exc)


def _clear_tokens() -> None:
    token_file = _token_file()
    if token_file.exists():
        try:
            token_file.unlink()
        except Exception as exc:  # pragma: no cover
            _LOG.debug("Could not remove token file: %s", exc)


def _decode_jwt_expiry(token: str) -> Optional[datetime]:
    parts = token.split(".")
    if len(parts) < 2:
        return None
    payload_b64 = parts[1] + "=" * (-len(parts[1]) % 4)
    try:
        payload = base64.urlsafe_b64decode(payload_b64.encode("utf-8"))
        data = json.loads(payload)
        exp = data.get("exp")
        if exp is None:
            return None
        return datetime.utcfromtimestamp(int(exp))
    except Exception:
        return None


def _is_expired(token: str, skew_seconds: int = 60) -> bool:
    expiry = _decode_jwt_expiry(token)
    if expiry is None:
        return False
    return datetime.utcnow() >= (expiry - timedelta(seconds=skew_seconds))


def _apply_access_token(access_token: str) -> None:
    bearer = f"Bearer {access_token}"
    try:
        import mainsequence.client as msc  # type: ignore
    except Exception as exc:  # pragma: no cover
        _LOG.debug("mainsequence client not available: %s", exc)
        return

    try:
        session = getattr(msc, "session", None)
        if session is None:
            session = requests.Session()
            setattr(msc, "session", session)
        if hasattr(session, "headers"):
            session.headers["Authorization"] = bearer
    except Exception as exc:  # pragma: no cover
        _LOG.debug("Unable to update mainsequence.session headers: %s", exc)

    try:
        loaders = getattr(msc, "loaders", None)
        if loaders is not None:
            headers = getattr(loaders, "default_headers", None)
            if isinstance(headers, dict):
                headers["Authorization"] = bearer
            elif hasattr(loaders, "set_default_headers"):
                loaders.set_default_headers({"Authorization": bearer})
    except Exception as exc:  # pragma: no cover
        _LOG.debug("Unable to update mainsequence loaders: %s", exc)

    try:
        import mainsequence.client.base as base  # type: ignore

        base_session = getattr(base, "session", None)
        if base_session is not None and hasattr(base_session, "headers"):
            base_session.headers["Authorization"] = bearer
    except Exception:  # pragma: no cover
        # Older SDKs may not expose base.session
        pass


def _refresh_tokens(tokens: Dict[str, str]) -> Optional[Dict[str, str]]:
    refresh_token = tokens.get("refresh")
    if not refresh_token:
        return None

    try:
        response = requests.post(
            AUTH_REFRESH_URL, json={"refresh": refresh_token}, timeout=10
        )
        response.raise_for_status()
        data = response.json()
        access = data.get("access")
        new_refresh = data.get("refresh", refresh_token)
        if access:
            refreshed = {"access": access, "refresh": new_refresh}
            _save_tokens(refreshed)
            _apply_access_token(access)
            return refreshed
    except Exception as exc:  # pragma: no cover
        _LOG.debug("Token refresh failed: %s", exc)
    return None


def _get_valid_tokens(force_refresh: bool = False) -> Optional[Dict[str, str]]:
    with _TOKEN_LOCK:
        tokens = _load_tokens()
        if not tokens:
            return None
        access = tokens.get("access")
        if not access:
            return None
        if force_refresh or _is_expired(access):
            refreshed = _refresh_tokens(tokens)
            return refreshed
        _apply_access_token(access)
        return tokens


def _authenticate(email: str, password: str) -> Optional[Dict[str, str]]:
    try:
        response = requests.post(
            AUTH_CREATE_URL,
            json={"email": email, "password": password},
            timeout=10,
        )
        response.raise_for_status()
        data = response.json()
        access = data.get("access")
        refresh = data.get("refresh")
        if access and refresh:
            tokens = {"access": access, "refresh": refresh}
            _save_tokens(tokens)
            _apply_access_token(access)
            return tokens
    except Exception as exc:
        _LOG.warning("Authentication failed: %s", exc)
    return None


def _flatten_range(value: Any) -> List[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [v for v in (part.strip() for part in value.split(",")) if v]
    if isinstance(value, Iterable):
        collected: List[str] = []
        for item in value:
            if isinstance(item, (list, tuple)):
                collected.extend(_flatten_range(item))
            elif item is None:
                continue
            else:
                text = str(item).strip()
                if text:
                    collected.append(text)
        return collected
    return [str(value).strip()]


def _excel_date_to_datetime(value: Any) -> datetime:
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, time.min)
    if isinstance(value, (int, float)):
        excel_epoch = datetime(1899, 12, 30)
        return excel_epoch + timedelta(days=float(value))
    raise ValueError(f"Unsupported date type: {type(value).__name__}")


def _dataframe_to_excel(df: pd.DataFrame, max_rows: int) -> List[List[Any]]:
    limited = df.head(int(max_rows))
    data = limited.where(pd.notnull(limited), None)
    return [list(data.columns)] + data.to_numpy().tolist()


def _friendly_error(message: str) -> List[List[str]]:
    return [[message]]


def _prompt_for_credentials() -> Optional[Dict[str, str]]:
    try:
        import tkinter as tk
        from tkinter import simpledialog
    except Exception as exc:  # pragma: no cover
        _LOG.warning("tkinter unavailable for login dialog: %s", exc)
        return None

    root = tk.Tk()
    root.withdraw()
    email = simpledialog.askstring("Main Sequence", "Email:", parent=root)
    if not email:
        root.destroy()
        return None
    password = simpledialog.askstring(
        "Main Sequence", "Password:", parent=root, show="*"
    )
    root.destroy()
    if not password:
        return None
    return {"email": email.strip(), "password": password}


@xl.func(name="MS.LOGIN_DIALOG", command=True)
def login_dialog() -> str:
    creds = _prompt_for_credentials()
    if not creds:
        return "Sign-in canceled."
    tokens = _authenticate(creds["email"], creds["password"])
    if tokens:
        return "Signed in to Main Sequence."
    return "Sign-in failed. Check credentials or network."


@xl.func(name="MS.SIGN_OUT", command=True)
def sign_out() -> str:
    _clear_tokens()
    return "Signed out of Main Sequence."


@xl.func(name="MS.GET_DATA_NODE")
def get_data_between_dates_from_node_identifier(
    node_identifier: str,
    start_date: Any,
    end_date: Any,
    unique_identifiers: Any = None,
    columns: Any = None,
    max_rows: int = 5000,
) -> List[List[Any]]:
    tokens = _get_valid_tokens()
    if not tokens:
        return _friendly_error("ERROR: Not signed in. Use MS.LOGIN_DIALOG or the ribbon Sign In button.")

    try:
        start_dt = _excel_date_to_datetime(start_date)
        end_dt = _excel_date_to_datetime(end_date)
    except Exception as exc:
        return _friendly_error(f"ERROR: Invalid date input: {exc}")

    unique_ids = _flatten_range(unique_identifiers)
    column_list = _flatten_range(columns)

    try:
        import mainsequence.client.models_tdag as models_tdag  # type: ignore
    except Exception as exc:
        return _friendly_error(f"ERROR: Unable to import mainsequence client: {exc}")

    try:
        df = models_tdag.DataNodeStorage.get_data_between_dates_from_node_identifier(
            node_identifier,
            start_dt,
            end_dt,
            unique_identifiers=unique_ids or None,
            columns=column_list or None,
            max_rows=int(max_rows) if max_rows is not None else 5000,
        )
    except Exception as exc:
        return _friendly_error(f"ERROR: Data fetch failed: {exc}")

    try:
        dataframe = pd.DataFrame(df)
        return _dataframe_to_excel(dataframe, max_rows)
    except Exception:
        # If result is already Excel-friendly (e.g., list of lists), return it directly.
        if isinstance(df, list):
            return df  # type: ignore[return-value]
        return _friendly_error("ERROR: Unexpected data format returned.")


_RIBBON_XML = """
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" loadImage="loadLogo">
  <ribbon>
    <tabs>
      <tab id="msTab" label="Main Sequence">
        <group id="msAuthGroup" label="Authentication">
          <button id="msLogin" label="Sign In" size="large" onAction="onSignIn" getImage="loadLogo" />
          <button id="msLogout" label="Sign Out" size="large" onAction="onSignOut" imageMso="HappyFace" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
""".strip()


def _load_logo_image(image_id: str) -> Any:  # Ribbon callback signature
    if Image is None:
        return None
    logo_path = Path(__file__).with_name("mainsequence_logo.png")
    if not logo_path.exists():
        return None
    try:
        return Image.open(logo_path)
    except Exception as exc:  # pragma: no cover
        _LOG.debug("Unable to load logo image: %s", exc)
        return None


def _register_ribbon() -> None:
    if ExcelGUI is None:
        _LOG.debug("ExcelGUI not available; skipping ribbon registration.")
        return
    try:
        gui = ExcelGUI(
            "MSRibbon",
            ribbon_xml=_RIBBON_XML,
            funcmap={
                "onSignIn": login_dialog,
                "onSignOut": sign_out,
                "loadLogo": _load_logo_image,
            },
        )
        gui.register()
    except Exception as exc:  # pragma: no cover
        _LOG.debug("Ribbon registration failed: %s", exc)


def _bootstrap() -> None:
    # Try to hydrate the session on import so Excel functions work immediately.
    _get_valid_tokens()
    _register_ribbon()


_bootstrap()
