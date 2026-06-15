# Zoho Data

Local importer and browser for Zoho CRM exports, focused on fast internal account, contact, contract, and activity lookup.

## What This App Does

The app is account-centered and built for navigating large CRM exports locally without changing the source data. It currently supports:

- Account search with multi-select filters for account type, state, MSA, and contract status
- Saved filter views in the left rail
- A richer results list with visible owner, contract, updated date, and bed count
- A `Current view` strip so active search, filters, and sort can be removed one by one
- Account detail with summary cards, progressive disclosure, and contract-focused attachment browsing
- A unified `Timeline` tab that merges notes, contacts, deals, cases, emails, meetings, tasks, and attachments
- Interactive timeline items that can jump to source tabs, plus direct open/download actions for attachment events
- Dark mode with system-aware default behavior
- Filter-aware Excel and contact exports

## Requirements

- Python 3
- SQLite

## Local Environment

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
```

## Import the Zoho Export

```sh
.venv/bin/python scripts/import_zoho.py
```

This creates local generated data such as:

- `data/zoho.sqlite3`
- `data/zoho.import-report.json`

## Run the UI

```sh
.venv/bin/python app/server.py
```

Then open:

```text
http://127.0.0.1:8765
```

## Exports

- `Export Excel` downloads the currently filtered account set plus supporting workbook tabs
- `Export Contacts` downloads a flat contact list for the currently filtered account set
- Contact export excludes contacts whose `Title` or `Type` contains `former`
- Contact export only keeps contacts with at least one usable email, phone, or mobile value

## Data And Git

This repository is intended to track code and documentation only.

Ignored local-only content includes:

- `data/`
- `zoho backup/`
- `zoho export/`
- `.venv/`
- Python cache directories and local system files

That keeps the Git repo free of the CRM database, import artifacts, and attachment archives.
