#!/usr/bin/env python3
import argparse
import csv
import json
import sqlite3
import sys
import zipfile
from collections import Counter
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
BACKUP_DIR = ROOT / "zoho backup"
DB_PATH = ROOT / "data" / "zoho.sqlite3"
DATA_ZIP_NAME = "Data_001.zip"
BATCH_SIZE = 1000


MODULES = {
    "zohosign_documents_c": {
        "csv": "Data/ZohoSign Documents_C_001.csv",
        "label": "ZohoSign Documents_C",
        "columns": [
            "record_id",
            "document_name",
            "account_id",
            "contact_id",
            "deal_id",
            "lead_id",
            "referral_id",
            "module_name",
            "module_record_id",
            "document_status",
            "date_sent",
            "date_completed",
            "created_time",
            "modified_time",
            "record_status",
            "rollup_account_id",
        ],
    },
    "referrals_c": {
        "csv": "Data/Referrals_C_001.csv",
        "label": "Referrals_C",
        "columns": [
            "record_id",
            "referral_id",
            "facility_account_id",
            "account_id_text",
            "contact_id",
            "lead_id",
            "connected_to_id",
            "connected_to_module",
            "patient_first_name",
            "patient_last_name",
            "first_name",
            "last_name",
            "email",
            "referral_type",
            "study_type",
            "scheduling_status",
            "created_time",
            "modified_time",
            "record_status",
            "rollup_account_id",
        ],
    },
    "emails": {
        "csv": "Data/Emails_001.csv",
        "label": "Emails",
        "columns": [
            "record_id",
            "subject",
            "module",
            "record_name_id",
            "account_id",
            "contact_id",
            "deal_id",
            "case_id",
            "referral_id",
            "sender",
            "sent_to",
            "sent_on",
            "status",
            "has_attachment",
            "attachment_count",
            "attachment_name",
            "created_time",
            "modified_time",
            "rollup_account_id",
        ],
    },
    "tasks": {
        "csv": "Data/Tasks_001.csv",
        "label": "Tasks",
        "columns": [
            "record_id",
            "subject",
            "contact_id",
            "related_to_id",
            "related_to_module",
            "status",
            "priority",
            "due_date",
            "created_time",
            "modified_time",
            "record_status",
            "rollup_account_id",
        ],
    },
    "meetings": {
        "csv": "Data/Meetings_001.csv",
        "label": "Meetings",
        "columns": [
            "record_id",
            "title",
            "contact_id",
            "related_to_id",
            "related_to_module",
            "from_time",
            "to_time",
            "location",
            "meeting_type",
            "created_time",
            "modified_time",
            "record_status",
            "rollup_account_id",
        ],
    },
    "notes_unified": {
        "csv": "Data/Notes_001.csv",
        "label": "Notes",
        "columns": [
            "record_id",
            "parent_id",
            "parent_module_raw",
            "parent_module",
            "note_title",
            "note_content_preview",
            "associated_id",
            "created_time",
            "modified_time",
            "record_status",
            "rollup_account_id",
        ],
    },
}


def q(identifier):
    return '"' + identifier.replace('"', '""') + '"'


def safe_decode_lines(binary_file):
    for line in binary_file:
        yield line.decode("utf-8-sig", errors="replace")


def load_existing_relationship_maps(conn):
    maps = {
        "accounts": set(),
        "contact_to_account": {},
        "deal_to_account": {},
        "case_to_account": {},
    }
    if table_exists(conn, "accounts"):
        maps["accounts"] = {row[0] for row in conn.execute("SELECT record_id FROM accounts WHERE record_id <> ''")}
    if table_exists(conn, "contacts"):
        maps["contact_to_account"] = dict(
            conn.execute(
                "SELECT record_id, account_name_id FROM contacts WHERE record_id <> '' AND COALESCE(account_name_id, '') <> ''"
            )
        )
    if table_exists(conn, "deals"):
        maps["deal_to_account"] = dict(
            conn.execute(
                "SELECT record_id, account_name_id FROM deals WHERE record_id <> '' AND COALESCE(account_name_id, '') <> ''"
            )
        )
    if table_exists(conn, "cases"):
        maps["case_to_account"] = dict(
            conn.execute(
                "SELECT record_id, account_name_id FROM cases WHERE record_id <> '' AND COALESCE(account_name_id, '') <> ''"
            )
        )
    return maps


def load_rollup_map(conn, table):
    if not table_exists(conn, table):
        return {}
    columns = {row[1] for row in conn.execute(f"PRAGMA table_info({q(table)})")}
    if "record_id" not in columns or "rollup_account_id" not in columns:
        return {}
    return dict(
        conn.execute(
            f"""
            SELECT record_id, rollup_account_id
            FROM {q(table)}
            WHERE record_id <> '' AND COALESCE(rollup_account_id, '') <> ''
            """
        )
    )


def refresh_imported_rollup_maps(conn, maps):
    maps["referral_to_account"] = load_rollup_map(conn, "referrals_c")
    maps["zohosign_to_account"] = load_rollup_map(conn, "zohosign_documents_c")
    maps["email_to_account"] = load_rollup_map(conn, "emails")
    maps["task_to_account"] = load_rollup_map(conn, "tasks")
    maps["meeting_to_account"] = load_rollup_map(conn, "meetings")
    return maps


def table_exists(conn, table):
    return conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type IN ('table', 'view') AND name = ?",
        (table,),
    ).fetchone() is not None


def infer_module_and_account(record_id, maps, referral_to_account=None):
    referral_to_account = referral_to_account or maps.get("referral_to_account", {})
    if not record_id:
        return "", ""
    if record_id in maps["accounts"]:
        return "Accounts", record_id
    if record_id in maps["contact_to_account"]:
        return "Contacts", maps["contact_to_account"].get(record_id, "")
    if record_id in maps["deal_to_account"]:
        return "Deals", maps["deal_to_account"].get(record_id, "")
    if record_id in maps["case_to_account"]:
        return "Cases", maps["case_to_account"].get(record_id, "")
    if record_id in referral_to_account:
        return "Referrals_C", referral_to_account.get(record_id, "")
    for module, key in (
        ("ZohoSign Documents_C", "zohosign_to_account"),
        ("Emails", "email_to_account"),
        ("Tasks", "task_to_account"),
        ("Meetings", "meeting_to_account"),
    ):
        lookup = maps.get(key, {})
        if record_id in lookup:
            return module, lookup.get(record_id, "")
    return "", ""


def create_table(conn, table, columns):
    conn.execute(f"DROP TABLE IF EXISTS {q(table)}")
    defs = ["record_id TEXT PRIMARY KEY"]
    defs.extend(f"{q(column)} TEXT" for column in columns if column != "record_id")
    defs.append("raw_data TEXT NOT NULL")
    conn.execute(f"CREATE TABLE {q(table)} ({', '.join(defs)})")


def create_indexes(conn, table, columns):
    table_columns = {row[1] for row in conn.execute(f"PRAGMA table_info({q(table)})")}
    for column in columns:
        if column not in table_columns:
            continue
        conn.execute(f"CREATE INDEX IF NOT EXISTS {q(f'idx_{table}_{column}')} ON {q(table)} ({q(column)})")


def read_csv_rows(data_zip_path, csv_name):
    with zipfile.ZipFile(data_zip_path) as archive:
        with archive.open(csv_name) as raw:
            reader = csv.DictReader(safe_decode_lines(raw))
            yield reader.fieldnames or []
            for row in reader:
                yield row


def collect_referral_to_account(data_zip_path):
    referral_to_account = {}
    for item in read_csv_rows(data_zip_path, MODULES["referrals_c"]["csv"]):
        if isinstance(item, list):
            continue
        record_id = item.get("Record Id", "")
        account_id = item.get("Facility Name.id", "")
        if record_id and account_id:
            referral_to_account[record_id] = account_id
    return referral_to_account


def normalize_zohosign(row, maps, referral_to_account):
    account_id = row.get("Account.id", "")
    if not account_id and row.get("Deal.id", ""):
        account_id = maps["deal_to_account"].get(row.get("Deal.id", ""), "")
    if not account_id and row.get("Contact.id", ""):
        account_id = maps["contact_to_account"].get(row.get("Contact.id", ""), "")
    if not account_id and row.get("Referrals.id", ""):
        account_id = referral_to_account.get(row.get("Referrals.id", ""), "")
    return {
        "record_id": row.get("Record Id", ""),
        "document_name": row.get("ZohoSign Documents Name", ""),
        "account_id": row.get("Account.id", ""),
        "contact_id": row.get("Contact.id", ""),
        "deal_id": row.get("Deal.id", ""),
        "lead_id": row.get("Lead.id", ""),
        "referral_id": row.get("Referrals.id", ""),
        "module_name": row.get("Module Name", ""),
        "module_record_id": row.get("Module Record ID", ""),
        "document_status": row.get("Document Status", ""),
        "date_sent": row.get("Date Sent", ""),
        "date_completed": row.get("Date Completed", ""),
        "created_time": row.get("Created Time", ""),
        "modified_time": row.get("Modified Time", ""),
        "record_status": row.get("Record Status", ""),
        "rollup_account_id": account_id,
    }


def normalize_referral(row, maps, referral_to_account):
    return {
        "record_id": row.get("Record Id", ""),
        "referral_id": row.get("Referral_ID", ""),
        "facility_account_id": row.get("Facility Name.id", ""),
        "account_id_text": row.get("Account Id", ""),
        "contact_id": row.get("Contact Name.id", ""),
        "lead_id": row.get("Lead Name.id", ""),
        "connected_to_id": row.get("Connected To.id", ""),
        "connected_to_module": row.get("Connected To.Module", ""),
        "patient_first_name": row.get("Patient First Name", ""),
        "patient_last_name": row.get("Patient Last Name", ""),
        "first_name": row.get("First Name", ""),
        "last_name": row.get("Last Name", ""),
        "email": row.get("Email", ""),
        "referral_type": row.get("Referral Type", ""),
        "study_type": row.get("Study Type", ""),
        "scheduling_status": row.get("Scheduling Status", ""),
        "created_time": row.get("Created Time", ""),
        "modified_time": row.get("Modified Time", ""),
        "record_status": row.get("Record Status", ""),
        "rollup_account_id": row.get("Facility Name.id", ""),
    }


def normalize_email(row, maps, referral_to_account):
    record_name_id = row.get("Record Name.id", "")
    module, account_id = infer_module_and_account(record_name_id, maps, referral_to_account)
    email_module = row.get("Module", "")
    if email_module == "Accounts":
        module = "Accounts"
        account_id = record_name_id
    elif email_module == "Contacts":
        module = "Contacts"
        account_id = maps["contact_to_account"].get(record_name_id, account_id)
    elif email_module == "Potentials":
        module = "Deals"
        account_id = maps["deal_to_account"].get(record_name_id, account_id)
    elif email_module == "Cases":
        module = "Cases"
        account_id = maps["case_to_account"].get(record_name_id, account_id)
    elif email_module == "CustomModule23":
        module = "Referrals_C"
        account_id = referral_to_account.get(record_name_id, account_id)

    return {
        "record_id": row.get("Record Id", ""),
        "subject": row.get("Subject", ""),
        "module": email_module,
        "record_name_id": record_name_id,
        "account_id": record_name_id if module == "Accounts" else "",
        "contact_id": record_name_id if module == "Contacts" else "",
        "deal_id": record_name_id if module == "Deals" else "",
        "case_id": record_name_id if module == "Cases" else "",
        "referral_id": record_name_id if module == "Referrals_C" else "",
        "sender": row.get("Sender", ""),
        "sent_to": row.get("Sent To", ""),
        "sent_on": row.get("Sent On", ""),
        "status": row.get("Status", ""),
        "has_attachment": row.get("Has Attachment", ""),
        "attachment_count": row.get("Attachment count", ""),
        "attachment_name": row.get("Attachment Name", ""),
        "created_time": row.get("Created Time", ""),
        "modified_time": row.get("Modified Time", ""),
        "rollup_account_id": account_id,
    }


def normalize_task(row, maps, referral_to_account):
    related_to_id = row.get("Related To.id", "")
    related_module, account_id = infer_module_and_account(related_to_id, maps, referral_to_account)
    if not account_id and row.get("Contact Name.id", ""):
        account_id = maps["contact_to_account"].get(row.get("Contact Name.id", ""), "")
    return {
        "record_id": row.get("Record Id", ""),
        "subject": row.get("Subject", ""),
        "contact_id": row.get("Contact Name.id", ""),
        "related_to_id": related_to_id,
        "related_to_module": related_module,
        "status": row.get("Status", ""),
        "priority": row.get("Priority", ""),
        "due_date": row.get("Due Date", ""),
        "created_time": row.get("Created Time", ""),
        "modified_time": row.get("Modified Time", ""),
        "record_status": row.get("Record Status", ""),
        "rollup_account_id": account_id,
    }


def normalize_meeting(row, maps, referral_to_account):
    related_to_id = row.get("Related To.id", "")
    related_module, account_id = infer_module_and_account(related_to_id, maps, referral_to_account)
    if not account_id and row.get("Contact Name.id", ""):
        account_id = maps["contact_to_account"].get(row.get("Contact Name.id", ""), "")
    return {
        "record_id": row.get("Record Id", ""),
        "title": row.get("Title", ""),
        "contact_id": row.get("Contact Name.id", ""),
        "related_to_id": related_to_id,
        "related_to_module": related_module,
        "from_time": row.get("From", ""),
        "to_time": row.get("To", ""),
        "location": row.get("Location", ""),
        "meeting_type": row.get("Meeting Type", ""),
        "created_time": row.get("Created Time", ""),
        "modified_time": row.get("Modified Time", ""),
        "record_status": row.get("Record Status", ""),
        "rollup_account_id": account_id,
    }


def normalize_note(row, maps, referral_to_account):
    parent_id = row.get("Parent.id", "")
    raw_module = row.get("Parent Id.Module", "")
    module_aliases = {
        "Accounts": "Accounts",
        "Contacts": "Contacts",
        "Potentials": "Deals",
        "Deals": "Deals",
        "Cases": "Cases",
        "CustomModule23": "Referrals_C",
        "Referrals": "Referrals_C",
        "CustomModule16": "ZohoSign Documents_C",
        "ZohoSign Documents": "ZohoSign Documents_C",
        "Emails": "Emails",
        "Tasks": "Tasks",
        "Events": "Meetings",
        "Meetings": "Meetings",
    }
    parent_module = module_aliases.get(raw_module, raw_module)

    account_id = ""
    if parent_module == "Accounts":
        account_id = parent_id
    elif parent_module == "Contacts":
        account_id = maps["contact_to_account"].get(parent_id, "")
    elif parent_module == "Deals":
        account_id = maps["deal_to_account"].get(parent_id, "")
    elif parent_module == "Cases":
        account_id = maps["case_to_account"].get(parent_id, "")
    elif parent_module == "Referrals_C":
        account_id = referral_to_account.get(parent_id, "") or maps.get("referral_to_account", {}).get(parent_id, "")
    elif parent_module == "ZohoSign Documents_C":
        account_id = maps.get("zohosign_to_account", {}).get(parent_id, "")
    elif parent_module == "Emails":
        account_id = maps.get("email_to_account", {}).get(parent_id, "")
    elif parent_module == "Tasks":
        account_id = maps.get("task_to_account", {}).get(parent_id, "")
    elif parent_module == "Meetings":
        account_id = maps.get("meeting_to_account", {}).get(parent_id, "")

    content = row.get("Note Content", "")
    return {
        "record_id": row.get("Record Id", ""),
        "parent_id": parent_id,
        "parent_module_raw": raw_module,
        "parent_module": parent_module,
        "note_title": row.get("Note Title", ""),
        "note_content_preview": content[:500],
        "associated_id": row.get("Associated_Id", ""),
        "created_time": row.get("Created Time", ""),
        "modified_time": row.get("Modified Time", ""),
        "record_status": row.get("Record Status", ""),
        "rollup_account_id": account_id,
    }


NORMALIZERS = {
    "zohosign_documents_c": normalize_zohosign,
    "referrals_c": normalize_referral,
    "emails": normalize_email,
    "tasks": normalize_task,
    "meetings": normalize_meeting,
    "notes_unified": normalize_note,
}


def import_module(conn, data_zip_path, table, spec, maps, referral_to_account):
    columns = spec["columns"]
    create_table(conn, table, columns)
    create_indexes(
        conn,
        table,
        [
            "record_id",
            "rollup_account_id",
            "account_id",
            "contact_id",
            "deal_id",
            "case_id",
            "referral_id",
            "related_to_id",
            "related_to_module",
            "record_name_id",
            "module",
            "parent_id",
            "parent_module",
            "parent_module_raw",
            "created_time",
            "modified_time",
        ],
    )

    placeholders = ", ".join(["?"] * (len(columns) + 1))
    sql = f"INSERT INTO {q(table)} ({', '.join(q(c) for c in columns)}, raw_data) VALUES ({placeholders})"
    normalizer = NORMALIZERS[table]
    batch = []
    row_count = 0
    rollup_count = 0
    relationship_counts = Counter()
    headers = []

    for item in read_csv_rows(data_zip_path, spec["csv"]):
        if isinstance(item, list):
            headers = item
            continue
        normalized = normalizer(item, maps, referral_to_account)
        values = [normalized.get(column, "") for column in columns]
        values.append(json.dumps(item, ensure_ascii=False, separators=(",", ":")))
        batch.append(values)
        row_count += 1
        if normalized.get("rollup_account_id"):
            rollup_count += 1
        for field in (
            "account_id",
            "contact_id",
            "deal_id",
            "case_id",
            "referral_id",
            "related_to_id",
            "record_name_id",
            "facility_account_id",
            "parent_id",
            "parent_module",
        ):
            if normalized.get(field):
                relationship_counts[field] += 1
        if len(batch) >= BATCH_SIZE:
            conn.executemany(sql, batch)
            batch = []
    if batch:
        conn.executemany(sql, batch)

    return {
        "table": table,
        "label": spec["label"],
        "csv": spec["csv"],
        "rows": row_count,
        "columns": columns,
        "csv_column_count": len(headers),
        "rollup_account_rows": rollup_count,
        "relationship_counts": relationship_counts,
    }


def account_rollup_attachment_count(conn):
    direct = conn.execute(
        """
        SELECT COUNT(*)
        FROM attachments
        WHERE parent_module = 'Accounts' AND COALESCE(parent_record_id, '') <> ''
        """
    ).fetchone()[0]
    indirect = 0
    for table, module in (
        ("zohosign_documents_c", "ZohoSign Documents_C"),
        ("referrals_c", "Referrals_C"),
        ("emails", "Emails"),
        ("tasks", "Tasks"),
        ("meetings", "Meetings"),
        ("contacts", "Contacts"),
        ("deals", "Deals"),
        ("cases", "Cases"),
    ):
        if not table_exists(conn, table):
            continue
        columns = {row[1] for row in conn.execute(f"PRAGMA table_info({q(table)})")}
        if "rollup_account_id" in columns:
            account_expr = "related.rollup_account_id"
        elif table in {"contacts", "deals", "cases"} and "account_name_id" in columns:
            account_expr = "related.account_name_id"
        else:
            continue
        indirect += conn.execute(
            f"""
            SELECT COUNT(*)
            FROM attachments attachment
            JOIN {q(table)} related ON attachment.parent_record_id = related.record_id
            WHERE attachment.parent_module = ?
              AND COALESCE({account_expr}, '') <> ''
            """,
            (module,),
        ).fetchone()[0]
    if table_exists(conn, "notes_unified"):
        indirect += conn.execute(
            """
            SELECT COUNT(*)
            FROM attachments attachment
            JOIN notes_unified note ON attachment.parent_record_id = note.record_id
            WHERE attachment.parent_module = 'Notes'
              AND COALESCE(note.rollup_account_id, '') <> ''
            """
        ).fetchone()[0]
    elif table_exists(conn, "notes_accounts"):
        indirect += conn.execute(
            """
            SELECT COUNT(*)
            FROM attachments attachment
            JOIN notes_accounts note ON attachment.parent_record_id = note.record_id
            WHERE attachment.parent_module = 'Notes'
              AND COALESCE(note.parent_id_id, '') <> ''
            """
        ).fetchone()[0]
    if not table_exists(conn, "notes_unified") and table_exists(conn, "notes_contacts") and table_exists(conn, "contacts"):
        indirect += conn.execute(
            """
            SELECT COUNT(*)
            FROM attachments attachment
            JOIN notes_contacts note ON attachment.parent_record_id = note.record_id
            JOIN contacts contact ON note.parent_id_id = contact.record_id
            WHERE attachment.parent_module = 'Notes'
              AND COALESCE(contact.account_name_id, '') <> ''
            """
        ).fetchone()[0]
    return direct, indirect, direct + indirect


def main():
    parser = argparse.ArgumentParser(description="Import minimal Zoho backup relationship modules for attachment rollups.")
    parser.add_argument("--backup-dir", type=Path, default=BACKUP_DIR)
    parser.add_argument("--db", type=Path, default=DB_PATH)
    args = parser.parse_args()

    data_zip_path = args.backup_dir / DATA_ZIP_NAME
    if not data_zip_path.exists():
        print(f"Missing backup data zip: {data_zip_path}", file=sys.stderr)
        return 1
    if not args.db.exists():
        print(f"Missing SQLite database: {args.db}", file=sys.stderr)
        return 1

    conn = sqlite3.connect(args.db)
    try:
        conn.execute("PRAGMA journal_mode = WAL")
        conn.execute("PRAGMA synchronous = NORMAL")
        before = account_rollup_attachment_count(conn)
        maps = load_existing_relationship_maps(conn)
        referral_to_account = collect_referral_to_account(data_zip_path)

        summaries = []
        with conn:
            for table, spec in MODULES.items():
                if table == "notes_unified":
                    refresh_imported_rollup_maps(conn, maps)
                summaries.append(import_module(conn, data_zip_path, table, spec, maps, referral_to_account))

        after = account_rollup_attachment_count(conn)
    finally:
        conn.close()

    print("ZOHO ROLLUP MODULE IMPORT SUMMARY")
    print(f"- Database: {args.db}")
    print(f"- Source: {data_zip_path}")
    print(f"- Account-rollup attachment coverage before: direct={before[0]:,}, indirect={before[1]:,}, total={before[2]:,}")
    print(f"- Account-rollup attachment coverage after: direct={after[0]:,}, indirect={after[1]:,}, total={after[2]:,}")
    print()
    for summary in summaries:
        print(f"{summary['table']} ({summary['label']})")
        print(f"- Source CSV: {summary['csv']} ({summary['csv_column_count']} columns)")
        print(f"- Rows imported: {summary['rows']:,}")
        print(f"- Rows with rollup_account_id: {summary['rollup_account_rows']:,}")
        print(f"- Normalized columns: {', '.join(summary['columns'])}")
        print(f"- Detected relationship fields: {dict(summary['relationship_counts'].most_common())}")
        if summary["rollup_account_rows"] == 0:
            print("- Relationship note: unclear account relationship; no rows resolved to accounts.")
        print()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
