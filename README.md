# MS Excel Win

Starter scaffold for the MainSequence Excel integration powered by xlOil.

## What This Is

**xlOil** is the Excel add-in (`xlOil.xll`) that Excel loads automatically.

**ms_excel_win** is your Python package that xlOil imports, allowing Excel to call your `MS.*` functions and display the ribbon.

Your Python code is located in:
- `src/ms_excel_win/ms_excel_wrappers.py` — function definitions
- `src/ms_excel_win/__init__.py` — imports wrappers so functions register

## Requirements

- **Windows desktop Excel** (xlOil does not work on Mac Excel or Excel Online)
- **Python 3.11 64-bit** (recommended)
- **Excel bitness must match Python bitness** (usually 64-bit Excel + 64-bit Python)
- **Internet access** (to install packages)

## Installation & Setup (Windows)

### A. One-time machine prerequisites

#### A1) Install Microsoft VC++ Runtime (IMPORTANT)

If you see errors like:
```
ImportError: DLL load failed while importing xlOil_Python311: The specified module could not be found
```

Download and install:
- **Microsoft Visual C++ Redistributable 2015–2022 (x64)**  
  https://support.microsoft.com/en-us/help/2977003/

After installing, **reboot Windows**. (Without this, xlOil's compiled `.pyd` often cannot load.)

### B. Setup the Python environment

Open **PowerShell** and run:

```powershell
cd C:\Users\<YOU>\code\ms-excel-win
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -U pip
```

If you get an execution policy error on the activation script, run:
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

Confirm you're in the venv:
```powershell
python -c "import sys,platform; print(sys.executable); print(platform.architecture()); print(sys.version)"
```

The path should contain `.venv\Scripts\` and platform should be `('64bit', 'WindowsPE')`.

### C. Install xlOil and register with Excel

#### C1) Install xlOil

```powershell
python -m pip install xloil
python -m pip show xloil
```

#### C2) Register xlOil with Excel

```powershell
xloil install
```

Expected output:
```
Installed ...\Microsoft\Excel\XLSTART\xlOil.xll
Edited ...\xlOil\xlOil.ini to point to ...\.venv python distribution
```

This means:
- Excel will load `xlOil.xll` automatically on startup (from XLSTART)
- xlOil is configured to use your `.venv` Python

### D. Install the ms_excel_win package

From the repo root (where `pyproject.toml` lives):

```powershell
python -m pip install -e .
```

Sanity check:
```powershell
python -c "import ms_excel_win; print(ms_excel_win.__version__)"
```

### E. Configure xlOil to auto-import your module

Edit the xlOil config file at:
```
C:\Users\<YOU>\AppData\Roaming\xlOil\xlOil.ini
```

Find or add the `LoadModules` setting to include your module:
```ini
LoadModules = ms_excel_win
```

**Notes:**
- `ms_excel_win/__init__.py` imports `ms_excel_wrappers`, so importing `ms_excel_win` is sufficient.
- If `LoadModules` already exists, add `ms_excel_win` to the comma-separated list.
- **Restart Excel** after editing.

### F. Verify it works in Excel

Open Excel and test a cell formula.

**Recommended:** Add a quick ping function for troubleshooting in `ms_excel_wrappers.py`:

```python
import xloil as xl

@xl.func(name="MS.PING")
def ping():
    return "loaded"
```

Then in Excel:
```
=MS.PING()
```

Expected result: `loaded`

If Excel shows `#NAME?`, then either:
- xlOil didn't load, or
- xlOil loaded but didn't import `ms_excel_win` (check `LoadModules` in `xlOil.ini`)

### G. Debugging with Visual Studio (optional)

#### G1) Start Excel first

Open Excel and load a workbook that calls your functions (or enter `=MS.PING()` in a cell).

#### G2) Attach Visual Studio to Excel

1. Open **Visual Studio** (with Python Tools installed)
2. **Debug** (using %debugpy.command.debugUsingLaunchConfig.title%) → **Attach to Process…**
3. Select `EXCEL.EXE` (if multiple, pick the one matching your Excel window)
4. For code type, select **Python (Python code only)**

#### G3) Set breakpoints

1. Open the file you're debugging (e.g., `src/ms_excel_win/ms_excel_wrappers.py`)
2. Set a breakpoint inside a function
3. Trigger the function from Excel (enter the formula or press a ribbon button)

**Tip:** Set Excel to **Manual calculation** while debugging to avoid repeated recalculation.

## Troubleshooting

| Problem | Solution |
|---------|----------|
| `xloil install` fails with DLL load error | Install VC++ 2015–2022 x64 (Section A1) and reboot. Ensure 64-bit Python and 64-bit Excel. |
| Excel shows `#NAME?` for `MS.*` functions | Confirm `xloil install` succeeded. Check `C:\Users\<YOU>\AppData\Roaming\xlOil\xlOil.ini` has `LoadModules = ms_excel_win`. Restart Excel. |
| Wrong Python environment | Always use `python -m pip ...` and verify: `python -c "import sys; print(sys.executable)"` points to `...\ms-excel-win\.venv\...` |

## Project structure

The project targets Python 3.11+ and uses a `src/` layout. Adjust versions or add optional dependencies in `pyproject.toml` as the implementation evolves.

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

#### Array Spill in Excel

Some functions like `MS.GET_DATE_NODE()` and `MS.GET_ASSET()` return arrays. Depending on your Excel version, these arrays may or may not automatically spill into adjacent cells.

**To handle this reliably:**

1.  Select the range of cells where you expect the data to appear.

2.  Type the formula in the formula bar, for example:

        =MS.GET_DATA_NODE(D6, D7, D8)

3.  Press **CTRL + SHIFT + ENTER** to force the array to fill the selected range.

**Note:**\
Make sure the selected range is large enough. If you see `#N/A`, it usually means you have reached the end of the returned data.

For more details in array spill in Excel: https://support.microsoft.com/en-us/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531

