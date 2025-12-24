from __future__ import annotations

import ast
import base64
import gzip
import json
import logging
import os
from collections.abc import Iterable
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
import requests

import xloil as xl
from PIL import Image

from auth import (
    AUTH_CREATE_URL,
    _apply_access_token,
    _authenticate,
    _clear_tokens,
    _get_valid_tokens,
)

_LOG = logging.getLogger("ms_excel_win.ms_excel_wrappers")
if not _LOG.handlers:
    log_path = Path.home() / "ms_excel_win.log"
    handler = logging.FileHandler(log_path, encoding="utf-8")
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s")
    handler.setFormatter(formatter)
    handler.setLevel(logging.DEBUG)
    _LOG.addHandler(handler)
    _LOG.setLevel(logging.DEBUG)
_GUI: Optional[xl.ExcelGUI] = None
_ASSET_CACHE: Dict[str, Any] = {}


def _flatten_range(value: Any) -> List[str]:
    if value is None:
        return []
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return []
        # If Excel passed a stringified list (e.g., "['ABC']", "[1,2]"), parse then flatten.
        if text.startswith(("(", "[")) and text.endswith((")", "]")):
            try:
                parsed = ast.literal_eval(text)
                return _flatten_range(parsed)
            except Exception:
                # Fall back to comma splitting if parsing fails.
                pass
        return [v for v in (part.strip().strip("'\"") for part in text.split(",")) if v]
    if isinstance(value, Iterable):
        collected: List[str] = []
        for item in value:
            collected.extend(_flatten_range(item))
        return collected
    text = str(value).strip()
    return [text] if text else []


def _excel_date_to_datetime(value: Any) -> datetime:
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, time.min)
    if isinstance(value, (int, float)):
        excel_epoch = datetime(1899, 12, 30)
        return excel_epoch + timedelta(days=float(value))
    raise ValueError(f"Unsupported date type: {type(value).__name__}")


def _to_excel_value(value: Any) -> Any:
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if isinstance(value, date) and not isinstance(value, datetime):
        return datetime.combine(value, time.min)
    if isinstance(value, pd.Timedelta):
        return value.to_pytimedelta()
    return value


def _dataframe_to_excel(df: pd.DataFrame, max_rows: int) -> List[List[Any]]:
    limited = df.head(int(max_rows))
    data = limited.where(pd.notnull(limited), None).applymap(_to_excel_value)
    index_names = [name if name is not None else "" for name in data.index.names]
    index_as_cols = data.index.to_frame(index=False)
    index_as_cols.columns = index_names
    combined = pd.concat([index_as_cols.reset_index(drop=True), data.reset_index(drop=True)], axis=1)
    header = list(combined.columns)
    return [header] + combined.to_numpy().tolist()


def _friendly_error(message: str) -> List[List[str]]:
    return [[message]]


def _set_status(message: Optional[str]) -> None:
    """Update Excel status bar using the xlOil app wrapper."""
    try:
        app = xl.app()
        impl = getattr(app, "impl", None)
        if impl is None:
            raise AttributeError("xl.app().impl is None")
        impl.StatusBar = message or False
    except Exception as exc:
        _LOG.debug("Unable to set Excel status bar: %s", exc)


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


def _show_message(msg: str, title: str = "Main Sequence", error: bool = False) -> None:
    """Display a message to the user. Try xl.alert(), then tkinter.messagebox.

    This is non-fatal — failures to show a UI are logged and ignored.
    """
    try:
        alert = getattr(xl, "alert", None)
        if callable(alert):
            alert(msg)
            return
    except Exception:
        _LOG.debug("xl.alert not available or failed")

    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        if error:
            messagebox.showerror(title, msg)
        else:
            messagebox.showinfo(title, msg)
        root.destroy()
        return
    except Exception:
        _LOG.debug("No tkinter UI available for messages")


@xl.func(name="MS.LOGIN_DIALOG", command=True)
def login_dialog() -> str:
    creds = _prompt_for_credentials()
    if not creds:
        return "Sign-in canceled."
    try:
        tokens = _authenticate(creds["email"], creds["password"])
        if tokens:
            msg = "Signed in to Main Sequence."
            try:
                _show_message(msg, "Signed in", error=False)
            except Exception:
                _LOG.debug("_show_message failed for success dialog")
            return msg
        # Shouldn't reach here because _authenticate raises on failure
        return "Sign-in failed. Check credentials or network."
    except requests.exceptions.HTTPError as exc:
        # Try to extract useful text from the response
        text = None
        try:
            text = exc.response.text
        except Exception:
            text = str(exc)
        status = getattr(exc.response, "status_code", "?")
        msg = f"Sign-in failed ({status}): {text}"
        try:
            _show_message(msg, "Sign-in failed", error=True)
        except Exception:
            _LOG.debug("_show_message failed for error dialog")
        return msg
    except ValueError as exc:
        msg = f"Sign-in failed: {exc}"
        try:
            _show_message(msg, "Sign-in failed", error=True)
        except Exception:
            _LOG.debug("_show_message failed for error dialog")
        return msg
    except Exception as exc:  # pragma: no cover
        _LOG.warning("Authentication failed: %s", exc)
        msg = "Sign-in failed. Check credentials or network."
        try:
            _show_message(msg, "Sign-in failed", error=True)
        except Exception:
            _LOG.debug("_show_message failed for error dialog")
        return msg


@xl.func(name="MS.SIGN_OUT", command=True)
def sign_out() -> str:
    _clear_tokens()
    try:
        _apply_access_token(None)
    except Exception:
        _LOG.debug("Failed to clear MAINSEQUENCE_TOKEN env var on sign out")
    return "Signed out of Main Sequence."


@xl.func(name="MS.SHOW_ENDPOINT")
def show_endpoint() -> str:
    """Show the root backend endpoint that the add-in is pointing to."""
    base_url = AUTH_CREATE_URL.rsplit("/auth/", 1)[0] if "/auth/" in AUTH_CREATE_URL else AUTH_CREATE_URL
    return f"Backend: {base_url}"


@xl.func(name="MS.MODULE_PATH")
def ms_module_path() -> str:
    """Return the absolute file path of this module."""
    return os.path.abspath(__file__)


@xl.func(name="MS.DEBUGPY_START", command=True)
def ms_debugpy_start(port: int = 5678) -> str:
    """Start debugpy listener for remote debugging without blocking Excel."""
    try:
        import debugpy
        debugpy.listen(("127.0.0.1", int(port)))
        return f"debugpy listening on 127.0.0.1:{port}"
    except ImportError:
        return "debugpy not installed. Run: python -m pip install debugpy"
    except Exception as exc:
        return f"Failed to start debugpy: {exc}"




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

    _set_status("Preparing request...")
    try:
        start_dt = _excel_date_to_datetime(start_date)
        end_dt = _excel_date_to_datetime(end_date)
    except Exception as exc:
        _set_status("Ready")
        return _friendly_error(f"ERROR: Invalid date input: {exc}")

    unique_ids = _flatten_range(unique_identifiers)
    column_list = _flatten_range(columns)

    try:
        import mainsequence.client.models_tdag as models_tdag  # type: ignore
    except Exception as exc:
        return _friendly_error(f"ERROR: Unable to import mainsequence client: {exc}")

    try:
        _set_status("Requesting data from Main Sequence...")
        result = models_tdag.DataNodeStorage.get_data_between_dates_from_node_identifier(
            node_identifier,
            start_dt,
            end_dt,
            unique_identifier_list=unique_ids or None,
            columns=column_list or None,
            #max_rows=int(max_rows) if max_rows is not None else 5000,
        )
    except Exception as exc:
        _set_status("Ready")
        return _friendly_error(f"ERROR: Data fetch failed: {exc}")

    try:
        _set_status("Processing data...")
        dataframe, storage_node = result
        storage_config = storage_node.sourcetableconfiguration

        if not dataframe.empty:
            dataframe = models_tdag.DataNodeStorage.map_columns_to_df(
                dataframe,
                column_dtypes_map=storage_config.column_dtypes_map,
                time_index_name=storage_config.time_index_name,
                index_names=storage_config.index_names,
            )

        excel_data = _dataframe_to_excel(dataframe, max_rows)
        _set_status("Ready")
        return excel_data
    except Exception as e:
        _set_status("Ready")
        # If result is already Excel-friendly (e.g., list of lists), return it directly.
        if isinstance(result, list):
            return result  # type: ignore[return-value]
        return _friendly_error(f"ERROR: Unexpected data format returned.{e}")


@xl.func(name="MS.GET_ASSET")
def get_asset(unique_identifier: str, spill_rows: bool = True) -> List[List[Any]]:
    """
    Fetch a single asset by unique identifier.

    Args:
        unique_identifier: Asset unique_identifier to look up.
        spill_rows: If True (default), spill each field/value pair as rows.
                    If False, return a single cell with the full JSON dump.

    Returns:
        Excel-friendly table or single-cell JSON string.
    """
    tokens = _get_valid_tokens()
    if not tokens:
        return _friendly_error("ERROR: Not signed in. Use MS.LOGIN_DIALOG or the ribbon Sign In button.")

    _set_status("Requesting asset...")
    try:
        import mainsequence.client as msc  # type: ignore
    except Exception as exc:
        _set_status("Ready")
        return _friendly_error(f"ERROR: Unable to import mainsequence client: {exc}")

    asset = _ASSET_CACHE.get(unique_identifier)
    if asset is None:
        try:
            asset = msc.Asset.get_or_none(unique_identifier=unique_identifier)
            if asset is not None:
                _ASSET_CACHE[unique_identifier] = asset
        except Exception as exc:
            _set_status("Ready")
            return _friendly_error(f"ERROR: Asset lookup failed: {exc}")

    if asset is None:
        _set_status("Ready")
        return _friendly_error("Asset not found.")

    try:
        raw = asset.model_dump()
    except Exception as exc:
        _set_status("Ready")
        return _friendly_error(f"ERROR: Unable to serialize asset: {exc}")

    def _serialize_value(val: Any) -> Any:
        if isinstance(val, (dict, list)):
            return json.dumps(val, default=str)
        if isinstance(val, pd.Timestamp):
            return val.to_pydatetime()
        if isinstance(val, date) and not isinstance(val, datetime):
            return datetime.combine(val, time.min)
        if isinstance(val, pd.Timedelta):
            return val.to_pytimedelta()
        return val

    if spill_rows:
        rows: List[List[Any]] = [["field", "value"]]
        for key, val in raw.items():
            rows.append([key, _serialize_value(val)])
        _set_status("Ready")
        return rows

    # Single cell JSON dump
    payload = json.dumps(raw, default=_serialize_value)
    _set_status("Ready")
    return [[payload]]


@xl.func(name="MS.GET_ASSET_FIELD")
def get_asset_field(unique_identifier: str, field_path: str) -> List[List[Any]]:
    """
    Fetch a single asset by unique identifier and return the value at the given field path.

    Args:
        unique_identifier: Asset unique_identifier to look up.
        field_path: Dot-separated path to the field (e.g., "current_snapshot.ticker").

    Returns:
        A single-cell table containing the field value, JSON-serialized for nested objects.
    """
    tokens = _get_valid_tokens()
    if not tokens:
        return _friendly_error("ERROR: Not signed in. Use MS.LOGIN_DIALOG or the ribbon Sign In button.")

    _set_status("Requesting asset field...")
    try:
        import mainsequence.client as msc  # type: ignore
    except Exception as exc:
        _set_status("Ready")
        return _friendly_error(f"ERROR: Unable to import mainsequence client: {exc}")

    asset = _ASSET_CACHE.get(unique_identifier)
    if asset is None:
        try:
            asset = msc.Asset.get_or_none(unique_identifier=unique_identifier)
            if asset is not None:
                _ASSET_CACHE[unique_identifier] = asset
        except Exception as exc:
            _set_status("Ready")
            return _friendly_error(f"ERROR: Asset lookup failed: {exc}")

    if asset is None:
        _set_status("Ready")
        return _friendly_error("Asset not found.")

    # Traverse dot path
    try:
        current = asset
        for part in field_path.split("."):
            current = getattr(current, part)
    except Exception as exc:
        _set_status("Ready")
        return _friendly_error(f"ERROR: Unable to resolve field '{field_path}': {exc}")

    def _serialize_value(val: Any) -> Any:
        if isinstance(val, (dict, list)):
            return json.dumps(val, default=str)
        if isinstance(val, pd.Timestamp):
            return val.to_pydatetime()
        if isinstance(val, date) and not isinstance(val, datetime):
            return datetime.combine(val, time.min)
        if isinstance(val, pd.Timedelta):
            return val.to_pytimedelta()
        return val

    _set_status("Ready")
    return _serialize_value(current)


_RIBBON_XML = r"""
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="msTab" label="Main Sequence" insertAfterMso="TabHome">
        <group id="msAuthGroup" label="Authentication">
          <button id="msLogin"
                  label="Sign In"
                  size="large"
                  onAction="onSignIn" />
          <button id="msLogout"
                  label="Sign Out"
                  size="large"
                  onAction="onSignOut"
                  imageMso="HappyFace" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
""".strip()


def onSignIn(ctrl: Any) -> None:
    login_dialog()


def onSignOut(ctrl: Any) -> None:
    sign_out()


def _register_ribbon() -> None:
    global _GUI
    if _GUI is not None and getattr(_GUI, "connected", False):
        _LOG.debug("Ribbon already registered.")
        return

    try:
        _LOG.info("Registering Main Sequence ribbon...")
        _GUI = xl.ExcelGUI(
            name="MSRibbon",
            ribbon=_RIBBON_XML,
            funcmap={
                "onSignIn": onSignIn,
                "onSignOut": onSignOut,
            },
            connect=True,
        )
        _LOG.info("Ribbon registered.")
    except Exception as exc:  # pragma: no cover
        _LOG.exception("Ribbon registration failed: %s", exc)


def _bootstrap() -> None:
    # Try to hydrate the session on import so Excel functions work immediately.
    _get_valid_tokens()
    _register_ribbon()


_bootstrap()


@xl.func(name="MS.DECOMPRESS_CURVE")
def decompress_curve(curve: str) -> List[List[Any]]:
    """
    Decompress a base64-encoded curve string using msi.data_interface.MSInterface.

    Args:
        curve: Base64 string representing the compressed curve.

    Returns:
        Single-cell JSON string of the decompressed curve or an error message.
    """
    _set_status("Decompressing curve...")
    try:
        import mainsequence.instruments as msi  # type: ignore
    except Exception as exc:
        _set_status("Ready")
        return _friendly_error(f"ERROR: Unable to import mainsequence.instruments: {exc}")

    try:
        data = msi.data_interface.MSInterface.decompress_string_to_curve(curve)
        rows: List[List[Any]] = [["days_to_maturity", "rate"]]
        for k, v in data.items():
            rows.append([k, v])
        _set_status("Ready")
        return rows
    except Exception as exc:
        _set_status("Ready")
        return _friendly_error(f"ERROR: Failed to decompress curve: {exc}")
