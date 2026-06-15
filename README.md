# CDX CRM

Local importer and browser for Zoho CRM exports, optimized for fast internal lookup of accounts, contacts, notes, contracts, attachments, and activity history.

## Overview

This project now supports two equally simple ways to run the app:

- Local Python development for day-to-day code changes
- Docker Compose for a self-contained packaged local service

The packaged service keeps the application code and runtime dependencies together, while leaving the large Zoho data files outside the container. That means:

- no manual Python environment management on deployment machines
- easy rebuild/redeploy when the code changes
- no duplication of the SQLite database or attachment archives
- continued access from other devices on the same home network

## Current Features

- Account search with multi-select filters for account type, state, MSA, and contract status
- Saved views in the left rail
- Results list with owner, contract, updated date, and bed count
- A `Current view` strip that shows active search, filters, and sort
- Account detail summary cards and progressive disclosure sections
- Contract-focused attachment browsing and filtering
- A unified `Timeline` tab that merges notes, contacts, deals, cases, emails, meetings, tasks, and attachments
- Interactive timeline items that can jump to their source tabs
- Direct `Open` and `Download` actions for attachments
- Dark mode with system-aware default behavior
- Filter-aware Excel and contact exports

## Project Structure

```text
app/
  server.py          HTTP server and API
  app.js             frontend behavior
  index.html         UI shell
  styles.css         UI styles

scripts/
  import_zoho.py                     main import workflow
  import_zoho_rollup_modules.py      rollup/import helpers
  index_zoho_attachments.py          attachment indexing
  inspect_zoho_backup.py             backup inspection
  report_zoho_attachments.py         attachment reporting

Dockerfile            container image definition
compose.yaml          packaged local service definition
.env.example          example deployment configuration
.dockerignore         excludes local data from image builds
.gitignore            excludes local data and dev artifacts from Git
requirements.txt      Python dependencies
README.md             project documentation
```

## Configuration

The application supports externalized runtime paths through environment variables:

- `ZOHO_DB_PATH`
  Absolute or project-relative path to the SQLite database file
- `ZOHO_BACKUP_DIR`
  Absolute or project-relative path to the directory containing attachment ZIP archives
- `HOST`
  Bind address for the HTTP service
- `PORT`
  Port for the HTTP service

Defaults if nothing is configured:

- database: `data/zoho.sqlite3`
- attachment archive directory: `zoho backup/`
- host: `0.0.0.0`
- port: `8765`

## Development Workflow

### 1. Set up a new development environment

Requirements:

- Python 3
- SQLite
- Docker Desktop or Docker Engine with Compose plugin, if you want to use the packaged workflow locally

Create the Python environment:

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
```

### 2. Import data for local development

If you need to rebuild the SQLite database from Zoho exports:

```sh
.venv/bin/python scripts/import_zoho.py
```

This creates local generated files such as:

- `data/zoho.sqlite3`
- `data/zoho.import-report.json`

### 3. Run the application for development

If you are using the default local paths:

```sh
.venv/bin/python app/server.py
```

If your database or attachment archives live somewhere else:

```sh
ZOHO_DB_PATH=/absolute/path/to/zoho.sqlite3 \
ZOHO_BACKUP_DIR="/absolute/path/to/zoho backup" \
.venv/bin/python app/server.py
```

Then open:

```text
http://127.0.0.1:8765
```

Because the app binds to `0.0.0.0` by default, it is also reachable from another machine on your LAN using this Mac's local IP address.

### 4. Make code changes and test them

Typical loop:

1. Edit files in `app/`
2. Restart the server
3. Reload the browser
4. Re-run any export or detail-flow checks you changed

Fast syntax check for the backend:

```sh
.venv/bin/python -m py_compile app/server.py
```

### 5. Rebuild artifacts after code changes

You do not need to rebuild the data for ordinary UI/API changes.

Re-run import or attachment indexing only when source data or import logic changes:

```sh
.venv/bin/python scripts/import_zoho.py
.venv/bin/python scripts/index_zoho_attachments.py
```

### 6. Common development troubleshooting

`Database not found. Run scripts/import_zoho.py first.`
: The configured `ZOHO_DB_PATH` is wrong or the database has not been built.

Attachments show as unavailable or missing
: Check that `ZOHO_BACKUP_DIR` points to the directory that actually contains `Attachments_*.zip`.

Port `8765` is already in use
: Stop the old process or run with a different port:

```sh
PORT=8766 .venv/bin/python app/server.py
```

The app loads but another device cannot reach it
: Check the Mac firewall and confirm the other device is using the correct local IP and port.

## Deployment Workflow

## Recommended Packaging Approach

The simplest maintainable deployment approach for this project is Docker Compose with bind-mounted host data.

Why this was chosen:

- it packages the code and Python dependencies together
- it avoids per-machine virtualenv maintenance
- it keeps the large Zoho files outside the image
- it preserves the existing `0.0.0.0` network behavior
- it is easy for one developer to rebuild, replace, and troubleshoot

The container image contains:

- application code
- Python runtime
- Python dependencies

The container does not contain:

- the SQLite database
- the Zoho export ZIP files
- the attachment archive ZIP files

## Build a deployable version

Copy the example environment file:

```sh
cp .env.example .env
```

Edit `.env` and set the real host paths:

```dotenv
ZOHO_APP_PORT=8765
ZOHO_DB_PATH_HOST="/absolute/path/to/zoho_data_work/data/zoho.sqlite3"
ZOHO_BACKUP_DIR_HOST="/absolute/path/to/zoho_data_work/zoho backup"
```

Quoted paths are recommended, especially when the attachment directory path contains spaces.

Build the container image:

```sh
docker compose build
```

## Start the application

Foreground:

```sh
docker compose up
```

Background:

```sh
docker compose up -d
```

The service publishes container port `8765` to the host port defined by `ZOHO_APP_PORT`, preserving local-network access.

## Stop the application

```sh
docker compose down
```

## Restart after changes

If only configuration changed:

```sh
docker compose up -d
```

If code changed and you want a fresh image:

```sh
docker compose up -d --build
```

## Update an existing deployment

1. Pull or copy the updated code
2. Rebuild the image
3. Restart the service

```sh
docker compose up -d --build
```

No dependency reinstall is needed on the host.

## Verify the application is running

Check container status:

```sh
docker compose ps
```

Open locally:

```text
http://127.0.0.1:8765
```

Or, if you changed the published port:

```text
http://127.0.0.1:<your-port>
```

## View logs and diagnose startup issues

Follow logs:

```sh
docker compose logs -f
```

Look at the most recent logs:

```sh
docker compose logs --tail=200
```

Useful things to check in startup output:

- reported database path
- reported attachment archive path
- the port the app bound to
- missing file or permission errors

## Data Management

### How the application locates data

The running service needs two external inputs:

- the SQLite database file
- the attachment archive directory

In containerized mode, these are provided through bind mounts configured in `compose.yaml` and sourced from `.env`.

Inside the container they are exposed as:

- `/external-data/zoho.sqlite3`
- `/external-attachments`

### How to change the data location

For local Python runs:

- set `ZOHO_DB_PATH`
- set `ZOHO_BACKUP_DIR`

For Docker Compose runs:

- edit `.env`
- update `ZOHO_DB_PATH_HOST`
- update `ZOHO_BACKUP_DIR_HOST`
- restart the service

### How to migrate to another machine

You do not need to copy the whole packaged application plus data together.

Move only:

- the code repository
- the external SQLite database
- the external attachment archive directory, if attachment access is needed

Then update `.env` on the new machine with the correct absolute paths.

### What should and should not be backed up

Back up:

- application code
- `.env` for deployment machines
- the SQLite database
- the attachment archive directory if you need document access
- original Zoho exports if you want to preserve raw import sources

Do not treat these as important backup targets:

- `.venv/`
- Python cache directories
- built container layers
- temporary exported workbooks

## Network Access

### Local access

Open:

```text
http://127.0.0.1:8765
```

Or use your configured published port.

### Access from another device on the same network

Because the service binds to `0.0.0.0` and Docker publishes the port to the host, another device on the same LAN can open:

```text
http://<mac-local-ip>:8765
```

For example:

```text
http://192.168.1.25:8765
```

To find the Mac's local IP:

```sh
ipconfig getifaddr en0
```

If needed, check `en1` or your active network interface instead.

### Firewall and networking considerations

- Docker must be running
- the published host port must not be blocked by the Mac firewall
- the client device must be on the same local network
- some routers isolate guest networks from main LAN devices

## Files Added For Packaging

- `Dockerfile`
  Builds a self-contained runtime image with the app and Python dependencies
- `compose.yaml`
  Runs the app as a local service, publishes the port, and mounts external data
- `.env.example`
  Template for deployment-specific host paths and published port
- `.dockerignore`
  Keeps local databases, exports, and archives out of the build context

## Architectural Changes Made

- The server now reads `ZOHO_DB_PATH` instead of requiring a fixed in-repo SQLite path
- The server now reads `ZOHO_BACKUP_DIR` instead of requiring a fixed in-repo attachment directory
- The server bind host is configurable through `HOST`, while still defaulting to `0.0.0.0`
- Docker Compose publishes the app port to the host so LAN access continues to work after containerization
- Large Zoho data files remain outside the container and are mounted read-only, avoiding wasteful duplication

This keeps the deployment model simple: rebuild the package when code changes, keep the data where it already lives, and point the packaged service at those existing files.
