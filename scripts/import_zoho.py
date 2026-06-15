#!/usr/bin/env python3
import csv
import json
import re
import sqlite3
import sys
import zipfile
from collections import Counter
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
EXPORT_DIR = ROOT / "zoho export"
DB_PATH = ROOT / "data" / "zoho.sqlite3"

EXPORTS = {
    "accounts": "Accounts_2026_05_25.zip",
    "contacts": "Contacts_2026_05_25.zip",
    "deals": "Deals_2026_05_25.zip",
    "cases": "Cases_2026_05_25.zip",
    "notes_accounts": "Notes_Accounts_2026_05_25.zip",
    "notes_contacts": "Notes_Contacts_2026_05_25.zip",
    "msa_mappings": "MSAMappings_2026_05_25.zip",
}


def clean_name(name, existing):
    cleaned = name.strip().lower()
    cleaned = cleaned.replace("%", "pct")
    cleaned = re.sub(r"[^a-z0-9]+", "_", cleaned)
    cleaned = cleaned.strip("_") or "field"
    if cleaned[0].isdigit():
        cleaned = f"field_{cleaned}"

    base = cleaned
    index = 2
    while cleaned in existing:
        cleaned = f"{base}_{index}"
        index += 1
    existing.add(cleaned)
    return cleaned


def q(identifier):
    return '"' + identifier.replace('"', '""') + '"'


def open_csv_from_zip(zip_path):
    archive = zipfile.ZipFile(zip_path)
    names = [name for name in archive.namelist() if name.lower().endswith(".csv")]
    if len(names) != 1:
        raise RuntimeError(f"{zip_path.name} should contain exactly one CSV, found {names}")
    raw = archive.open(names[0])
    text = (line.decode("utf-8-sig", errors="replace") for line in raw)
    return archive, names[0], csv.DictReader(text)


def create_table(conn, table, headers, mapping):
    conn.execute(f"DROP TABLE IF EXISTS {q(table)}")
    columns = [f"{q(mapping[column])} TEXT" for column in headers]
    columns.append('"raw_data" TEXT NOT NULL')
    conn.execute(f"CREATE TABLE {q(table)} ({', '.join(columns)})")


def create_index(conn, table, column):
    try:
        conn.execute(f"CREATE INDEX IF NOT EXISTS {q(f'idx_{table}_{column}')} ON {q(table)} ({q(column)})")
    except sqlite3.OperationalError:
        pass


def import_export(conn, table, zip_name):
    zip_path = EXPORT_DIR / zip_name
    archive, csv_name, reader = open_csv_from_zip(zip_path)
    headers = reader.fieldnames or []
    existing = set()
    mapping = {header: clean_name(header, existing) for header in headers}

    create_table(conn, table, headers, mapping)
    conn.execute(
        """
        INSERT INTO import_metadata (table_name, zip_file, csv_file, row_count, column_count, column_map)
        VALUES (?, ?, ?, 0, ?, ?)
        """,
        (table, zip_name, csv_name, len(headers), json.dumps(mapping, indent=2)),
    )

    db_columns = [mapping[header] for header in headers] + ["raw_data"]
    placeholders = ", ".join(["?"] * len(db_columns))
    insert_sql = f"INSERT INTO {q(table)} ({', '.join(q(column) for column in db_columns)}) VALUES ({placeholders})"

    rows = 0
    record_ids = Counter()
    batch = []
    for row in reader:
        rows += 1
        record_id = row.get("Record Id", "")
        if record_id:
            record_ids[record_id] += 1
        values = [row.get(header, "") for header in headers]
        values.append(json.dumps(row, ensure_ascii=False, separators=(",", ":")))
        batch.append(values)
        if len(batch) >= 1000:
            conn.executemany(insert_sql, batch)
            batch = []
    if batch:
        conn.executemany(insert_sql, batch)

    conn.execute("UPDATE import_metadata SET row_count = ? WHERE table_name = ?", (rows, table))
    archive.close()

    for preferred in (
        "record_id",
        "account_name_id",
        "parent_id_id",
        "contact_name_id",
        "chain_account_id",
        "network_parent_account_id",
        "msa",
        "account_type",
        "contract_status",
        "stage",
        "status",
        "modified_time",
        "created_time",
    ):
        if preferred in db_columns:
            create_index(conn, table, preferred)

    duplicate_count = sum(count - 1 for count in record_ids.values() if count > 1)
    return {"table": table, "rows": rows, "columns": len(headers), "record_id_duplicates": duplicate_count}


def scalar(conn, sql, params=()):
    return conn.execute(sql, params).fetchone()[0]


def relationship_report(conn):
    checks = [
        ("contacts", "account_name_id", "accounts", "record_id"),
        ("deals", "account_name_id", "accounts", "record_id"),
        ("deals", "contact_name_id", "contacts", "record_id"),
        ("cases", "account_name_id", "accounts", "record_id"),
        ("notes_accounts", "parent_id_id", "accounts", "record_id"),
        ("notes_contacts", "parent_id_id", "contacts", "record_id"),
        ("accounts", "chain_account_id", "accounts", "record_id"),
        ("accounts", "network_parent_account_id", "accounts", "record_id"),
    ]
    results = []
    for source_table, source_column, target_table, target_column in checks:
        source_cols = {row[1] for row in conn.execute(f"PRAGMA table_info({q(source_table)})")}
        target_cols = {row[1] for row in conn.execute(f"PRAGMA table_info({q(target_table)})")}
        if source_column not in source_cols or target_column not in target_cols:
            continue
        total = scalar(
            conn,
            f"SELECT COUNT(*) FROM {q(source_table)} WHERE COALESCE({q(source_column)}, '') <> ''",
        )
        matched = scalar(
            conn,
            f"""
            SELECT COUNT(*)
            FROM {q(source_table)} source
            JOIN {q(target_table)} target ON source.{q(source_column)} = target.{q(target_column)}
            WHERE COALESCE(source.{q(source_column)}, '') <> ''
            """,
        )
        results.append(
            {
                "source": source_table,
                "column": source_column,
                "target": target_table,
                "values": total,
                "matched": matched,
                "unmatched": total - matched,
            }
        )
    return results


def create_support_tables(conn):
    conn.executescript(
        """
        DROP TABLE IF EXISTS import_metadata;
        CREATE TABLE import_metadata (
            table_name TEXT PRIMARY KEY,
            zip_file TEXT NOT NULL,
            csv_file TEXT NOT NULL,
            row_count INTEGER NOT NULL,
            column_count INTEGER NOT NULL,
            column_map TEXT NOT NULL,
            imported_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );

        DROP VIEW IF EXISTS account_search;
        """
    )


def create_views(conn):
    conn.executescript(
        """
        DROP VIEW IF EXISTS account_search;
        CREATE VIEW account_search AS
        SELECT
            record_id,
            account_name,
            account_type,
            facility_type,
            billing_city,
            billing_state,
            billing_county,
            msa,
            territories,
            contract_status,
            contract_type,
            products_available,
            account_owner,
            modified_time,
            COALESCE(total_beds, certified_medicare_beds) AS beds
        FROM accounts;
        """
    )


def main():
    if not EXPORT_DIR.exists():
        print(f"Missing export directory: {EXPORT_DIR}", file=sys.stderr)
        return 1

    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    if DB_PATH.exists():
        DB_PATH.unlink()

    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA synchronous = NORMAL")
    create_support_tables(conn)

    summaries = []
    with conn:
        for table, zip_name in EXPORTS.items():
            print(f"Importing {zip_name} -> {table}")
            summaries.append(import_export(conn, table, zip_name))
        create_views(conn)

    relationships = relationship_report(conn)
    report = {"database": str(DB_PATH), "tables": summaries, "relationships": relationships}
    report_path = DB_PATH.with_suffix(".import-report.json")
    report_path.write_text(json.dumps(report, indent=2), encoding="utf-8")

    print(json.dumps(report, indent=2))
    conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
