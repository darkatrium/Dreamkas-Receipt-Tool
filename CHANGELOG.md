# Changelog

## v6.9

### Added

- Added robust Python dependency bootstrap in BAT launchers.
- Added automatic `pip` recovery:
  - first via `python -m ensurepip --upgrade`;
  - then via official `get-pip.py` bootstrap if `ensurepip` fails.
- Added `diagnose_python.bat` for Python and dependency diagnostics.
- Added dependency checks to:
  - `start_dreamkas.bat`;
  - `start_gui.bat`;
  - `install_dependencies.bat`;
  - `build_exe.bat`.

### Improved

- Startup is more stable on Windows systems where Python is installed without `pip`.
- EXE build process is more resilient because dependencies are checked before PyInstaller is launched.
- README updated for GitHub with Russian and English documentation.

---

## v6.6

### Added

- Embedded Excel template support for PyInstaller EXE builds.
- Automatic Excel template recovery from EXE resources.
- Automatic Excel template creation if no template file is found.
- GUI support for restoring or creating the Excel template.
- Updated `build_exe.bat` to include:

```bat
--add-data "dreamkas_receipt_template.xlsx;."
```

### Improved

- EXE can now be moved to a new folder and launched without manually copying `dreamkas_receipt_template.xlsx`.

---

## Current baseline functionality

- Dreamkas API shop selection.
- Dreamkas API cash register selection.
- Excel-based receipt preparation.
- Terminal precheck before fiscalization.
- Receipt fiscalization through Dreamkas.
- Operation polling.
- SQLite storage.
- Duplicate protection through precheck hash.
- Full and partial refunds.
- TXT, QR, and PDF receipt export.
- SMTP email sending.
- Encrypted SMTP password storage.
- Logging.
- GUI launcher.
- PyInstaller EXE build support.
