# GitHub commit / release text

## Suggested commit message

```text
feat: improve EXE packaging and dependency bootstrap
```

## Suggested commit body

```text
- Embed Excel template into PyInstaller EXE builds
- Restore Excel template from EXE resources if missing
- Create a default Excel template automatically when no template exists
- Add robust pip bootstrap through ensurepip and get-pip.py
- Add diagnose_python.bat for Python environment diagnostics
- Update start/build/install BAT files to check dependencies before execution
- Update Russian and English README documentation
- Add CHANGELOG.md
```

## Suggested GitHub release title

```text
Dreamkas Receipt Tool v6.7 — EXE template restore and pip bootstrap
```

## Suggested GitHub release notes

```markdown
## Dreamkas Receipt Tool v6.7

This release improves EXE portability and Windows dependency installation.

### Added

- Excel template is embedded into the EXE during PyInstaller build.
- If `dreamkas_receipt_template.xlsx` is missing, the program restores it from the embedded EXE resource.
- If the embedded resource is unavailable, the program creates a new Excel template automatically.
- Added `diagnose_python.bat` to inspect Python, `pip`, `ensurepip`, and dependency imports.
- BAT launchers now try to recover `pip` automatically:
  - first via `ensurepip`;
  - then via official `get-pip.py`.

### Improved

- `start_dreamkas.bat`, `start_gui.bat`, `install_dependencies.bat`, and `build_exe.bat` now check dependencies before running.
- EXE can be moved to a new folder without manually copying the Excel template.
- README was updated in Russian and English.

### Notes

Do not commit `settings.txt`, generated receipts, logs, or SQLite database files to GitHub.
```
