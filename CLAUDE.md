# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

Masar (مسار) is a PyQt5 desktop app (Arabic, RTL) for managing civilian employee records, attachments, and correspondence/fax logs, with WeasyPrint PDF report export. Data lives in a local SQLite file. Deployed to Windows as a PyInstaller one-file exe.

## Commands

```sh
pip install -r requirements.txt   # PyQt5 (note: weasyprint and Pillow are NOT listed but are required: pip install weasyprint pillow)
python main.py                    # run the app
pytest test_masar.py -v           # all tests
pytest test_masar.py::TestNormalizeArabic::test_normalize_hamza_variants -v   # single test
```

The full PyInstaller build command is in `readme.md` and at the bottom of `main.py` (must `--add-data` assets, config.json, license.json, utils/pdf_bg_utils.py, attachments). There is no linter configured.

## Architecture

- **Entry flow (`main.py`)**: `init_db()` → on Windows, install the bundled Amiri font into `C:\Windows\Fonts` if missing (needed for PDF Arabic rendering) → `LicenseManager.validate_license()`; if invalid, show `LicenseDialog` and exit unless activated → `MasarMainWindow`.
- **`ui/`**: `main_window.py` creates one shared `sqlite3` connection and passes it to each tab (`DashboardTab`, `EmployeeTab`, `CorrespondenceTab`). Tabs hold the UI *and* the SQL directly; there is no real model layer. `ui/dialogs/` holds helper dialogs (column selection, print options, multi-column sort, license).
- **`database/db_manager.py`**: `init_db()` creates tables and applies schema upgrades by running `ALTER TABLE ... ADD COLUMN` and swallowing `OperationalError`. The `employee` table columns are generated from `EMPLOYEE_FIELDS`. `database/employee_model.py`, `correspondence_model.py`, `migrations.py`, `ui/retire_tab.py`, and `reports/pdf_exporter.py` are empty placeholders.
- **`utils/constants.py`**: single source of truth for `EMPLOYEE_FIELDS` (DB columns, in order), `AR_LABELS` (Arabic display names), `DB_FILE` and `ATTACHMENTS_DIR` (relative paths — the app depends on the working directory). Adding an employee field means updating both here and the UI/export code; `init_db` only auto-adds a few legacy columns via explicit ALTERs, so new fields also need an ALTER there for existing databases.
- **`reports/`**: `employee_export.py`, `retire_export.py`, `correspondence_export.py` build HTML and render it with WeasyPrint. They import `PyQt5.QtWidgets.QMessageBox` for error dialogs, and use `utils/pdf_bg_utils.process_bg_image` to apply the watermark background (`assets/masar-bg.png`) with contrast/brightness/opacity from `config.json` (which also holds report header lines and font size).
- **`utils/arabic_normalizer.py`**: `normalize_arabic` (unifies hamza forms, ة/ه, etc.) is applied on save and search so Arabic matching is tolerant.
- **Licensing (`utils/license_core.py`)**: `get_hardware_id()` hashes a hardware fingerprint (MAC, OS machine GUID, CPU) with SHA-256, truncates it, and encodes it as a short human-readable code (Crockford Base32, grouped like `XK3F-9QPL-7ZRT-M2CD`) shown in `LicenseDialog`. Activation codes are generated the same way by HMAC-signing that hardware code (`generate_activation_code`); the HMAC key is embedded in the source. `license.json` stores `hardware_id`/`activation_code` and is revalidated by recomputing and comparing on every startup. Be careful when editing: a broken license check blocks app startup. This module is shared (kept byte-identical) with the separate `encryption` project's `license_core.py`.
- **`utils/file_utils.resource_path`**: resolves asset paths for both source runs and the PyInstaller bundle (`sys._MEIPASS`). Use it for bundled resources.

## Testing notes

`conftest.py` replaces `PyQt5`, its submodules, and `pdf_bg_utils` with `MagicMock` before collection, so tests run headless and can only cover non-GUI logic (normalizer, `init_db`, constants, SQLite CRUD). `COVERAGE_*.md` files describe the (low) coverage; they still refer to a legacy single-file `masar.py`, which no longer exists.

## Repo hygiene

`masar.db`, `attachments/`, `*.pdf`, and `*.zip` are gitignored (real employee data) — don't commit them. Stray generated PDFs in the repo root are outputs, not fixtures.
