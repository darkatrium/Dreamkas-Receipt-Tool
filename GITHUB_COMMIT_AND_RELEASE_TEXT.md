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


## Suggested commit for v6.15

```text
feat: add legal entity lookup by INN
```

```text
- Add DaData-based legal entity / sole proprietor lookup by INN
- Allow Excel legal buyer name to be empty when INN is provided
- Add settings for lookup provider and DaData token
- Save looked-up buyer name in precheck, SQLite, TXT, PDF, and FFD tags
- Optionally write found buyer name back to Excel
```


## Suggested commit for v6.15

```text
docs: add DaData setup guide
```

```text
- Add DADATA_SETUP.txt in Russian and English
- Document DaData registration and API token setup
- Add DaData settings examples to README
- Add troubleshooting for missing/invalid DaData token
- Update settings.example.txt with DaData setup reference
```


## Suggested commit for v6.18

```text
feat: add T-Bank SBP payment integration
```

```text
- Add SBP payment type through T-Bank acquiring
- Create T-Bank payments with /v2/Init
- Generate SBP QR with /v2/GetQr
- Print QR code in terminal and save it as PNG
- Poll payment status with /v2/GetState
- Fiscalize Dreamkas receipt only after successful SBP payment
- Store T-Bank terminal password encrypted as TBANK_PASSWORD_ENC
- Add T-Bank setup menu and documentation
```


## Suggested commit for v6.30

```text
feat: add cashless payment provider selection
```

```text
- Add CASHLESS_PAYMENT_PROVIDER setting
- Support external bank terminal and T-Bank SBP as cashless providers
- Add manual payment confirmation for external terminal
- Send Dreamkas receipt only after external terminal confirmation
- Add settings menu for cashless provider selection
```


## Suggested commit for v6.30

```text
feat: add local Python installer
```

```text
- Add install_local_python.bat
- Install Python into local _python folder
- Prefer _python\python.exe in all launchers
- Avoid relying on PATH, py launcher, winget detection, or registry
```


## Suggested commit for v6.30

```text
fix: install local Python into LocalAppData
```

```text
- Install local Python into LocalAppData instead of project folder
- Add embedded Python ZIP fallback
- Save detected local Python path to python_path.txt
- Avoid OneDrive and Cyrillic path installer issues
```


## Suggested commit for v6.30

```text
fix: use fixed Python path C:\Python314
```

```text
- Use C:\Python314\python.exe as primary interpreter
- Install Python directly into C:\Python314
- Simplify BAT launchers and dependency setup
- Document administrator requirement if C:\ installation is blocked
```


## Suggested commit for v6.30

```text
fix: force fresh Python installer download
```

```text
- Delete cached Python installer before download
- Force fresh Python installer download
- Use InstallAllUsers=1 for C:\Python314 installation
- Print installer exit code and log tail on failure
```


## Suggested commit for v6.30

```text
fix: repair broken C:\Python314 installation
```

```text
- Detect broken Python by importing encodings
- Remove incomplete C:\Python314 before reinstall
- Explicitly enable Include_exe and Include_lib in Python installer
- Add embedded Python ZIP fallback to C:\Python314
```


## Suggested commit for v6.30

```text
fix: require buyer contact selection
```

```text
- Remove "Do not specify" buyer contact option
- Require email, phone, or both for electronic receipt flow
- Update prompt text and README
```


## Suggested commit for v6.30

```text
feat: add tax mode menu and inline SBP status
```

```text
- Replace manual tax mode input with Russian numbered menu
- Map numeric choices to Dreamkas tax mode codes
- Display selected tax mode with Russian label in precheck
- Update SBP polling status in one terminal line
- Prevent QR code from shifting down during SBP waiting
```


## Suggested commit for v6.30

```text
feat: restore Excel dropdowns and add SBP refund offer
```

```text
- Restore Excel data validation for item type and VAT rate columns
- Add dropdowns to auto-generated Excel template
- Offer T-Bank SBP money refund after successful fiscal refund
- Send T-Bank /v2/Cancel using original PaymentId and refund amount
- Save money refund result to SQLite
```


## Suggested commit for v6.30

```text
feat: add T-Bank SBP B2B I2I mode for legal buyers
```

```text
- Ask legal entity buyers whether to use regular SBP or B2B/I2I SBP
- Add T-API settings and encrypted B2B token storage
- Create one-time B2B SBP link/QR through T-Bank T-API
- Poll B2B payment status before Dreamkas fiscalization
```


## Suggested commit for v6.30

```text
feat: allow cancel during payment and fiscalization waits
```

```text
- Add ESC key cancellation during SBP payment polling
- Add ESC key cancellation during B2B/I2I payment polling
- Add ESC key cancellation during Dreamkas operation polling
- Save cancelled waits as CANCELLED_BY_OPERATOR in SQLite
- Return to main menu after cancellation
```
