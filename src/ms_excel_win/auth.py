from __future__ import annotations

import base64
import json
import logging
import os
import threading
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Optional

import requests

_LOG = logging.getLogger("ms_excel_win.auth")

ROOT_URL = os.environ.get("MS_ROOT_URL", "https://main-sequence.app")


os.environ["TDAG_ENDPOINT"] = ROOT_URL

# Token auth endpoint (Django REST framework style)
AUTH_CREATE_URL = os.environ.get(
    "MS_AUTH_CREATE_URL", f"{ROOT_URL}/auth/rest-token-auth/"
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
            if access:
                return {"access": access}
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


# Track whether we've refreshed mainsequence logger bindings
_LOGGER_BOUND = False


def _apply_access_token(access_token: str) -> None:
    try:
        if access_token is None:
            os.environ.pop("MAINSEQUENCE_TOKEN", None)
        else:
            os.environ["MAINSEQUENCE_TOKEN"] = str(access_token)
        _LOG.debug("MAINSEQUENCE_TOKEN environment variable set")

        # Try to refresh mainsequence logger bindings once
        global _LOGGER_BOUND
        if access_token and (not _LOGGER_BOUND):
            try:
                import mainsequence

                refresh = getattr(mainsequence, "refresh_application_logger_bindings", None)
                if callable(refresh):
                    refresh()
                    _LOGGER_BOUND = True
                    _LOG.debug("Refreshed mainsequence logger bindings")
            except Exception as exc:  # pragma: no cover
                _LOG.debug("Failed to refresh logger bindings: %s", exc)
    except Exception as exc:  # pragma: no cover
        _LOG.debug("Unable to set MAINSEQUENCE_TOKEN env var: %s", exc)


def _get_valid_tokens(force_refresh: bool = False) -> Optional[Dict[str, str]]:
    with _TOKEN_LOCK:
        tokens = _load_tokens()
        if not tokens:
            return None
        access = tokens.get("access")
        if not access:
            return None
        _apply_access_token(access)
        return tokens


def _authenticate(email: str, password: str) -> Optional[Dict[str, str]]:
    response = requests.post(
        AUTH_CREATE_URL,
        json={"username": email, "password": password},
        timeout=10,
    )
    response.raise_for_status()
    data = response.json()
    token = data.get("token") or data.get("access")
    if token:
        tokens = {"access": token}
        _save_tokens(tokens)
        _apply_access_token(token)
        return tokens
    raise ValueError("Authentication response missing token")
