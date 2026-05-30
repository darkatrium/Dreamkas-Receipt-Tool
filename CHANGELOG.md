# Changelog

## v6.33

### Fixed

- Fixed buyer phone transfer to Dreamkas attributes.
- Phone number is now sent as `attributes.phone`.
- Email is sent only as `attributes.email`.
- Fixed the issue where phone number could appear in OFD as:

```text
Эл. адрес покупателя: +7...
```

### Notes

Correct mapping:

```text
buyer_email → attributes.email
buyer_phone → attributes.phone
```


## v6.32

### Fixed

- Fixed Dreamkas validation error when the operator selected "Общая система налогообложения — ОСН".
- Dreamkas API `/api/receipts` does not accept `taxMode = OSN`.
- The utility now sends `DEFAULT` for general taxation / OSN.
- Removed unsupported `ENVD` option from the tax mode menu.
- Added final safety normalization before sending `taxMode` to Dreamkas:
  - `OSN` → `DEFAULT`
  - unsupported/unknown values → `DEFAULT`

### Notes

Dreamkas API validation accepts only:

```text
DEFAULT
SIMPLE
SIMPLE_WO
AGRICULT
PATENT
```


All notable changes to **Dreamkas Receipt Tool** are documented in this file.

Current version: **v6.33**

---

## v6.30

### Added

- Added `ESC` cancellation during long waiting operations:
  - T-Bank SBP payment polling;
  - T-Bank SBP B2B/I2I polling;
  - Dreamkas fiscalization polling.
- Added `OperatorCancelled` flow.
- Cancelled waits are saved to SQLite as:

```text
CANCELLED_BY_OPERATOR
```

### Notes

- `ESC` cancels local waiting and returns the operator to the main menu.
- If a Dreamkas fiscalization task was already sent, it may still complete on the cash register side.
- The operator should check the saved record/status before retrying the same receipt to avoid duplicates.

---

## v6.29

### Added

- Added SBP payment mode choice for legal entity / sole proprietor buyers:
  - regular T-Bank EACQ SBP QR;
  - T-Bank B2B/I2I SBP via T-API.
- Added T-Bank B2B/I2I settings:
  - `TBANK_B2B_SBP_ENABLED`
  - `TBANK_B2B_API_TOKEN_ENC`
  - `TBANK_B2B_ACCOUNT_NUMBER`
  - `TBANK_B2B_TTL_DAYS`
  - `TBANK_B2B_VAT`
  - `TBANK_B2B_REDIRECT_URL`
  - `TBANK_B2B_PURPOSE_PREFIX`
- Added settings menu item:

```text
11. Настроить СБП B2B/I2I через T-Банк
```

- Added B2B/I2I one-time QR/link creation through T-API endpoint:

```text
/api/v1/b2b/qr/onetime
```

- Added B2B/I2I status polling through:

```text
/api/v1/b2b/qr/{qrId}/info
```

### Notes

- B2B/I2I uses T-API Bearer token, not EACQ `TerminalKey` / password.
- For legal entity buyers, the operator can still choose regular SBP if needed.

---

## v6.28

### Added

- Restored Excel dropdowns:
  - item type: `1` service, `0` product;
  - VAT: `0`, `5`, `7`, `10`, `22`, `5/105`, `7/107`, `10/110`, `22/122`.
- Programmatically created Excel templates now also include the same dropdowns.
- Added optional T-Bank SBP money refund flow after successful fiscal refund receipt.
- If the original sale was paid through T-Bank SBP, the tool offers to send `/v2/Cancel` for the original `PaymentId`.

### Notes

- T-Bank money refund is offered only after the Dreamkas refund receipt is successfully fiscalized.
- If the operator skips the money refund, the fiscal refund remains saved and the skip is logged in SQLite.
- If T-Bank refund fails, the fiscal refund remains completed and the error is saved in SQLite.

---

## v6.27

### Changed

- Tax mode input is now a numbered Russian-language menu.
- Operator no longer needs to type internal Dreamkas tax mode codes manually.
- Added choices:
  - `DEFAULT` — by cash register / Dreamkas default;
  - `OSN` — general taxation system;
  - `SIMPLE` — simplified tax system, income;
  - `SIMPLE_WO` — simplified tax system, income minus expenses;
  - `AGRICULT` — unified agricultural tax;
  - `PATENT` — patent tax system;
  - `ENVD` — archive imputed income tax mode.
- Precheck now displays tax mode with a Russian label.
- SBP payment polling status now overwrites the same terminal line instead of printing a new line every time.

---

## v6.26

### Changed

- Removed buyer contact option:

```text
0. Не указывать
```

- Buyer contact is now required for electronic receipt flow.
- Operator must choose one of:
  - email;
  - phone;
  - email + phone.
- Empty input no longer skips buyer contact selection.

---

## v6.25

### Fixed

- Detects broken `C:\Python314\python.exe` installations by checking:

```python
import encodings
```

- Removes incomplete `C:\Python314` before reinstalling.
- Explicitly passes Python installer options:
  - `Include_exe=1`
  - `Include_lib=1`
  - `Include_pip=1`
- Adds embedded Python ZIP fallback into `C:\Python314`.
- Fixes the case where installer exits with code `0` but Python fails with:

```text
ModuleNotFoundError: No module named 'encodings'
```

---

## v6.24

### Fixed

- Python installer is now always downloaded fresh.
- Old cached installer in `%TEMP%` is deleted before download.
- Installation to `C:\Python314` now uses `InstallAllUsers=1`.
- BAT files print Python installer exit code.
- If installation fails, BAT files print the last lines of:

```text
dreamkas-python-install.log
```

### Notes

- Installing to `C:\Python314` may require Administrator rights.
- Right-click `install_python.bat` and select `Run as administrator` if installation fails.

---

## v6.23

### Changed

- Python installation path is now fixed to:

```text
C:\Python314
```

- BAT launchers now prefer:

```text
C:\Python314\python.exe
```

- Removed complex LocalAppData / embedded Python fallback logic from launchers.
- If installation into `C:\Python314` fails, the user should run the BAT as Administrator.

---

## v6.22

### Fixed

- Local Python installation no longer uses the project folder as `TargetDir`.
- Python is installed into:

```text
%LOCALAPPDATA%\DreamkasReceiptTool\Python313
```

- This avoids OneDrive, Desktop, Cyrillic path, and installer `TargetDir` issues.

### Added

- Embedded Python ZIP fallback:

```text
%LOCALAPPDATA%\DreamkasReceiptTool\Python313_embed
```

- If the normal installer does not create `python.exe`, the BAT downloads and extracts Python embeddable ZIP.
- `python_path.txt` is automatically written after a valid local Python is found.

---

## v6.21

### Added

- Added local Python installation into the tool folder:

```text
_python\python.exe
```

- Added `install_local_python.bat`.
- BAT launchers now prefer `_python\python.exe` before any system Python.
- Local Python installation does not modify Windows PATH.

### Improved

- The tool can run on computers where:
  - Python is not installed;
  - `python` points to a broken Windows alias;
  - `py` is missing;
  - `winget` installs Python but it is not visible in the current terminal.

---

## v6.20

### Fixed

- Fixed BAT files where multiline PowerShell commands with `^` could be passed into PowerShell incorrectly.
- Python registry / recursive detection command was generated as a single PowerShell command line.
- Fixed errors like:

```text
^ : The term '^' is not recognized as the name of a cmdlet
```

### Notes

- If the terminal header shows an older version, for example `v6.12`, the old folder is still being launched.
- Extract the new version into a clean folder and run `start_dreamkas.bat` from that folder.

---

## v6.19

### Added

- Added cashless payment provider setting:
  - `EXTERNAL_TERMINAL`
  - `TBANK_SBP`
- Added settings menu item to choose cashless payment provider.
- Added external bank terminal confirmation flow.
- If external terminal is selected, the operator must confirm that payment was accepted before the Dreamkas receipt is sent.

### Changed

- Cashless payment is now a general payment type.
- Specific cashless processing method is selected in settings.
- T-Bank SBP remains available as one of the cashless providers.

---

## v6.18

### Added

- Added T-Bank SBP payment integration.
- Added creation of payment through T-Bank EACQ.
- Added QR retrieval through T-Bank.
- Added terminal QR display.
- Added SBP QR image saving.
- Added polling of payment status before fiscalization.
- Added SQLite fields and storage for SBP payment metadata.
- Added settings menu for T-Bank SBP.
- Added `TBANK_SBP_SETUP.txt`.

### Changed

- If payment type is T-Bank SBP, the Dreamkas receipt is sent only after successful payment.
- Dreamkas still receives the receipt as cashless payment.

---

## v6.17

### Changed

- Buyer contact input now starts with a choice:
  - email;
  - phone;
  - email + phone;
  - no contact.
- Program asks only the selected contact fields.
- Email and phone inputs are validated before continuing.
- Russian phone numbers starting with `8` are normalized to `+7`.

---

## v6.16

### Changed

- Excel template now contains only receipt line items.
- Receipt metadata is now entered in the terminal:
  - cashier name;
  - payment type;
  - buyer email;
  - buyer phone;
  - buyer type;
  - legal entity / sole proprietor INN and name;
  - tax mode.
- Cashier name is requested once at startup and reused during the whole session.
- After fiscalization or refund, the program returns to the main menu instead of closing.

### Improved

- Excel cleanup now clears only line item rows.
- New default Excel template is simpler and safer for operators.
- Legal entity lookup by INN now works from terminal input.

---

## v6.15

### Added

- Added `DADATA_SETUP.txt` with Russian and English setup instructions.
- Added DaData registration and token setup instructions to README.
- Added DaData troubleshooting notes:
  - missing token;
  - invalid token;
  - organization not found;
  - Dreamkas FFD tag validation issues.
- Updated `settings.example.txt` with a reference to `DADATA_SETUP.txt`.

---

## v6.14

### Added

- Legal entity / sole proprietor name auto-fill by INN.
- Added DaData provider for company lookup by INN.
- Added settings:
  - `LEGAL_ENTITY_LOOKUP_ENABLED`
  - `LEGAL_ENTITY_LOOKUP_PROVIDER`
  - `DADATA_TOKEN`
  - `DADATA_BRANCH_TYPE`
  - `DADATA_UPDATE_EXCEL_NAME`
- If buyer INN is filled and legal name is empty, the tool fetches the name automatically.
- Optionally writes the found legal name back to Excel.
- Settings menu includes legal entity lookup configuration.

### Notes

- Kontur.Focus can be added as another provider later, but it requires commercial API access.
- If lookup fails, legal entity name must be filled manually.

---

## v6.13

### Added

- Added legal entity / sole proprietor buyer support.
- Added buyer type field.
- Added legal entity / sole proprietor name.
- Added legal entity / sole proprietor INN.
- Added INN validation for 10 or 12 digits.
- Added local saving of legal buyer details.
- Added TXT/PDF output of legal buyer details.
- Added optional FFD tags:
  - `1227` — buyer / client;
  - `1228` — buyer / client INN.
- Added setting:

```text
SEND_LEGAL_ENTITY_TAGS_TO_DREAMKAS
```

---

## v6.12

### Added

- Added receipt-style PDF output.
- PDF is generated as narrow cashier tape instead of A4-style report.
- Added larger QR-code in PDF.
- Added receipt details to PDF:
  - seller;
  - cashier;
  - payment type;
  - tax mode;
  - FN / FD / FP where available;
  - receipt check link;
  - QR code.

---

## Earlier changes

### Added

- Initial Dreamkas fiscalization workflow.
- Excel parsing.
- TXT receipt saving.
- QR code generation.
- SQLite receipt storage.
- Full and partial refund flow.
- Receipt duplicate warning.
- SMTP setup and test tool.
- BAT launchers.
- Python auto-detection.
- Dependency installation.
- README and GitHub release helper files.
