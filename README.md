# Zoho Data

Local exploratory importer and browser for the Zoho CRM export in `zoho export/`.

## Requirements

- Python 3
- SQLite

## Set up the local environment

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
```

## Import the export

```sh
.venv/bin/python scripts/import_zoho.py
```

This creates:

- `data/zoho.sqlite3`
- `data/zoho.import-report.json`

## Run the local UI

```sh
.venv/bin/python app/server.py
```

Then open:

```text
http://127.0.0.1:8765
```

The first version is account-centered. It supports account search, account filters, and account detail views with related contacts, deals, cases, notes, and raw Zoho fields.

The results list also shows a `Current view` strip so users can see the active search, filters, and sort at a glance and remove them one by one without reopening the left-side filters.

Account detail now includes a `Timeline` tab that merges notes, contact changes, deals, cases, emails, meetings, tasks, and attachments into one chronological feed. The timeline can be narrowed by activity type and searched inline from the detail pane.

## Exports

- `Export Excel` downloads the currently filtered account set plus supporting tabs.
- `Export Contacts` downloads a flat contact list for the currently filtered account set.
- The contacts export excludes contacts whose Title or Type contains `former` and only keeps contacts with an email, phone, or mobile value.
