#!/usr/bin/env python3
import argparse
import json
import sqlite3
import sys
from collections import Counter
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DB_PATH = ROOT / "data" / "zoho.sqlite3"

JOIN_REPORTS = [
    {
        "label": "Accounts",
        "table": "accounts",
        "module": "Accounts",
        "id_column": "record_id",
        "display_columns": ["account_name", "billing_city", "billing_state"],
    },
    {
        "label": "Contacts",
        "table": "contacts",
        "module": "Contacts",
        "id_column": "record_id",
        "display_columns": ["contact_name", "email", "account_name"],
    },
    {
        "label": "Deals",
        "table": "deals",
        "module": "Deals",
        "id_column": "record_id",
        "display_columns": ["deal_name", "stage", "account_name"],
    },
    {
        "label": "Cases",
        "table": "cases",
        "module": "Cases",
        "id_column": "record_id",
        "display_columns": ["case_number", "subject", "account_name"],
    },
    {
        "label": "Notes",
        "table": "notes_unified",
        "module": "Notes",
        "id_column": "record_id",
        "display_columns": ["note_title", "parent_module", "rollup_account_id"],
    },
    {
        "label": "Notes",
        "table": "notes_accounts",
        "module": "Notes",
        "id_column": "record_id",
        "display_columns": ["note_title", "parent_id", "modified_time"],
    },
    {
        "label": "Notes",
        "table": "notes_contacts",
        "module": "Notes",
        "id_column": "record_id",
        "display_columns": ["note_title", "parent_id", "modified_time"],
    },
    {
        "label": "Referrals_C",
        "table": "referrals_c",
        "module": "Referrals_C",
        "id_column": "record_id",
        "display_columns": ["referral_id", "patient_first_name", "patient_last_name", "rollup_account_id"],
    },
    {
        "label": "ZohoSign Documents_C",
        "table": "zohosign_documents_c",
        "module": "ZohoSign Documents_C",
        "id_column": "record_id",
        "display_columns": ["document_name", "document_status", "rollup_account_id"],
    },
    {
        "label": "Emails",
        "table": "emails",
        "module": "Emails",
        "id_column": "record_id",
        "display_columns": ["subject", "module", "record_name_id", "rollup_account_id"],
    },
    {
        "label": "Tasks",
        "table": "tasks",
        "module": "Tasks",
        "id_column": "record_id",
        "display_columns": ["subject", "related_to_module", "rollup_account_id"],
    },
    {
        "label": "Meetings",
        "table": "meetings",
        "module": "Meetings",
        "id_column": "record_id",
        "display_columns": ["title", "related_to_module", "rollup_account_id"],
    },
]

ROLLUP_MODULE_PRIORITY = [
    "ZohoSign Documents_C",
    "Notes",
    "Emails",
    "Referrals_C",
    "Leads",
    "Tasks",
    "Meetings",
    "Contacts_002",
]


def q(identifier):
    return '"' + identifier.replace('"', '""') + '"'


def human_size(size):
    value = float(size or 0)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if value < 1024 or unit == "TB":
            if unit == "B":
                return f"{int(value)} {unit}"
            return f"{value:.1f} {unit}"
        value /= 1024
    return f"{size} B"


def connect(db_path):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def table_exists(conn, table):
    return (
        conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type IN ('table', 'view') AND name = ?",
            (table,),
        ).fetchone()
        is not None
    )


def table_columns(conn, table):
    if not table_exists(conn, table):
        return set()
    return {row["name"] for row in conn.execute(f"PRAGMA table_info({q(table)})")}


def one(conn, sql, params=()):
    return conn.execute(sql, params).fetchone()[0]


def grouped_counts(conn, sql):
    return [(row[0] if row[0] else "[unknown]", row[1]) for row in conn.execute(sql)]


def extension_counts(conn):
    counts = Counter()
    for row in conn.execute("SELECT original_filename FROM attachments"):
        counts[Path(row["original_filename"] or "").suffix.lower() or "[none]"] += 1
    return counts.most_common()


def largest_attachments(conn, limit=25):
    return conn.execute(
        """
        SELECT
            attachment_record_id,
            parent_module,
            parent_record_id,
            original_filename,
            COALESCE(zip_file_size, metadata_file_size, 0) AS size,
            zip_file,
            mapping_confidence
        FROM attachments
        ORDER BY size DESC, original_filename COLLATE NOCASE
        LIMIT ?
        """,
        (limit,),
    ).fetchall()


def top_parents(conn, limit=25):
    return conn.execute(
        """
        SELECT
            COALESCE(NULLIF(parent_module, ''), '[unknown]') AS parent_module,
            parent_record_id,
            COUNT(*) AS attachment_count,
            SUM(COALESCE(zip_file_size, metadata_file_size, 0)) AS total_size
        FROM attachments
        GROUP BY parent_module, parent_record_id
        ORDER BY attachment_count DESC, total_size DESC
        LIMIT ?
        """,
        (limit,),
    ).fetchall()


def select_display_exprs(conn, table, display_columns):
    columns = table_columns(conn, table)
    exprs = []
    aliases = []
    for column in display_columns:
        if column in columns:
            exprs.append(f"record.{q(column)} AS {q(column)}")
            aliases.append(column)
    return exprs, aliases


def joined_samples(conn, report, limit=8):
    table = report["table"]
    if not table_exists(conn, table):
        return None
    columns = table_columns(conn, table)
    if report["id_column"] not in columns:
        return None

    exprs, aliases = select_display_exprs(conn, table, report["display_columns"])
    selected = ", ".join([f"record.{q(report['id_column'])} AS record_id"] + exprs + [
        "attachment.original_filename AS attachment_filename",
        "attachment.zip_file AS zip_file",
        "COALESCE(attachment.zip_file_size, attachment.metadata_file_size, 0) AS attachment_size",
    ])
    rows = conn.execute(
        f"""
        SELECT {selected}
        FROM attachments attachment
        JOIN {q(table)} record
            ON attachment.parent_record_id = record.{q(report['id_column'])}
        WHERE attachment.parent_module = ?
        ORDER BY attachment.modified_time DESC, attachment.original_filename COLLATE NOCASE
        LIMIT ?
        """,
        (report["module"], limit),
    ).fetchall()
    return aliases, rows


def current_app_modules(conn):
    modules = set()
    if table_exists(conn, "accounts"):
        modules.add("Accounts")
    if table_exists(conn, "contacts"):
        modules.add("Contacts")
    if table_exists(conn, "deals"):
        modules.add("Deals")
    if table_exists(conn, "cases"):
        modules.add("Cases")
    if table_exists(conn, "notes_unified") or table_exists(conn, "notes_accounts") or table_exists(conn, "notes_contacts"):
        modules.add("Notes")
    if table_exists(conn, "msa_mappings"):
        modules.add("MSAMappings")
    if table_exists(conn, "zohosign_documents_c"):
        modules.add("ZohoSign Documents_C")
    if table_exists(conn, "referrals_c"):
        modules.add("Referrals_C")
    if table_exists(conn, "emails"):
        modules.add("Emails")
    if table_exists(conn, "tasks"):
        modules.add("Tasks")
    if table_exists(conn, "meetings"):
        modules.add("Meetings")
    return modules


def account_rollup_coverage(conn):
    direct = one(
        conn,
        """
        SELECT COUNT(*)
        FROM attachments
        WHERE parent_module = 'Accounts' AND COALESCE(parent_record_id, '') <> ''
        """,
    )
    indirect_rows = []
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
        columns = table_columns(conn, table)
        if "rollup_account_id" in columns:
            account_expr = "related.rollup_account_id"
        elif table == "contacts" and "account_name_id" in columns:
            account_expr = "related.account_name_id"
        elif table in {"deals", "cases"} and "account_name_id" in columns:
            account_expr = "related.account_name_id"
        else:
            continue
        count = one(
            conn,
            f"""
            SELECT COUNT(*)
            FROM attachments attachment
            JOIN {q(table)} related ON attachment.parent_record_id = related.record_id
            WHERE attachment.parent_module = ?
              AND COALESCE({account_expr}, '') <> ''
            """,
            (module,),
        )
        if count:
            indirect_rows.append((module, count))
    if table_exists(conn, "notes_unified"):
        count = one(
            conn,
            """
            SELECT COUNT(*)
            FROM attachments attachment
            JOIN notes_unified note ON attachment.parent_record_id = note.record_id
            WHERE attachment.parent_module = 'Notes'
              AND COALESCE(note.rollup_account_id, '') <> ''
            """,
        )
        if count:
            indirect_rows.append(("Notes", count))
    elif table_exists(conn, "notes_accounts"):
        count = one(
            conn,
            """
            SELECT COUNT(*)
            FROM attachments attachment
            JOIN notes_accounts note ON attachment.parent_record_id = note.record_id
            WHERE attachment.parent_module = 'Notes'
              AND COALESCE(note.parent_id_id, '') <> ''
            """,
        )
        if count:
            indirect_rows.append(("Notes via accounts", count))
    if not table_exists(conn, "notes_unified") and table_exists(conn, "notes_contacts") and table_exists(conn, "contacts"):
        count = one(
            conn,
            """
            SELECT COUNT(*)
            FROM attachments attachment
            JOIN notes_contacts note ON attachment.parent_record_id = note.record_id
            JOIN contacts contact ON note.parent_id_id = contact.record_id
            WHERE attachment.parent_module = 'Notes'
              AND COALESCE(contact.account_name_id, '') <> ''
            """,
        )
        if count:
            indirect_rows.append(("Notes via contacts", count))
    return direct, indirect_rows, direct + sum(count for _, count in indirect_rows)


def unrollable_by_module(parent_counts, direct_rollup, indirect_rollups):
    rollable = Counter()
    rollable["Accounts"] = direct_rollup
    for module, count in indirect_rollups:
        if module.startswith("Notes"):
            rollable["Notes"] += count
        else:
            rollable[module] += count

    remaining = []
    for module, total in parent_counts:
        if module == "[unknown]":
            remaining.append((module, total))
            continue
        unrollable = total - rollable.get(module, 0)
        if unrollable > 0:
            remaining.append((module, unrollable))
    return sorted(remaining, key=lambda item: item[1], reverse=True)


def missing_modules(conn):
    attachment_modules = {
        row[0]
        for row in conn.execute(
            "SELECT DISTINCT parent_module FROM attachments WHERE COALESCE(parent_module, '') <> ''"
        )
    }
    imported = current_app_modules(conn)
    missing = attachment_modules - imported
    counts = dict(
        conn.execute(
            """
            SELECT parent_module, COUNT(*)
            FROM attachments
            WHERE COALESCE(parent_module, '') <> ''
            GROUP BY parent_module
            """
        ).fetchall()
    )
    return sorted(missing, key=lambda module: counts.get(module, 0), reverse=True), imported, counts


def recommendations(missing, counts):
    ordered = [module for module in ROLLUP_MODULE_PRIORITY if module in missing]
    ordered.extend(module for module in missing if module not in ordered)

    recs = []
    for module in ordered:
        count = counts.get(module, 0)
        if module == "ZohoSign Documents_C":
            recs.append((module, count, "High value: many contract/signature attachments can roll up through Account.id and Deal.id."))
        elif module == "Notes":
            recs.append((module, count, "High value: unified backup Notes has parent module data and covers attachments not present in the split notes import."))
        elif module == "Emails":
            recs.append((module, count, "High value: many attachments are email-owned; import minimal email parent fields and related record fields."))
        elif module == "Referrals_C":
            recs.append((module, count, "High value: referral attachments can likely roll up to accounts via Facility Name.id or related account fields."))
        elif module in {"Tasks", "Meetings"}:
            recs.append((module, count, "Medium value: activity attachments may roll up through Related To.id or contact/account fields."))
        elif module == "Leads":
            recs.append((module, count, "Lower immediate account-detail value unless lead conversion links are needed."))
        else:
            recs.append((module, count, "Import minimally if these attachments need search or rollup coverage."))
    return recs


def print_counts(title, rows, limit=None):
    print(title)
    selected = rows[:limit] if limit else rows
    for key, count in selected:
        print(f"- {key}: {count:,}")
    if limit and len(rows) > limit:
        print(f"- ... {len(rows) - limit:,} more")
    print()


def print_joined_samples(conn):
    print("SAMPLE JOINED RECORDS")
    emitted = set()
    for report in JOIN_REPORTS:
        key = (report["label"], report["table"])
        aliases_rows = joined_samples(conn, report)
        if aliases_rows is None:
            print(f"- {report['label']}: table `{report['table']}` is not currently imported")
            continue
        aliases, rows = aliases_rows
        if not rows:
            print(f"- {report['label']} via `{report['table']}`: no joined attachment samples")
            continue
        title = report["label"]
        if report["label"] in emitted:
            title = f"{report['label']} via `{report['table']}`"
        emitted.add(report["label"])
        print(f"- {title}:")
        for row in rows:
            parts = [f"record_id={row['record_id']}"]
            for alias in aliases:
                value = row[alias]
                if value:
                    parts.append(f"{alias}={value}")
            parts.append(f"file={row['attachment_filename']}")
            parts.append(f"size={human_size(row['attachment_size'])}")
            print(f"  {', '.join(parts)}")
    print()


def main():
    parser = argparse.ArgumentParser(description="Report on the Zoho attachment metadata index.")
    parser.add_argument("--db", type=Path, default=DB_PATH)
    args = parser.parse_args()

    if not args.db.exists():
        print(f"Missing SQLite database: {args.db}", file=sys.stderr)
        return 1

    conn = connect(args.db)
    try:
        if not table_exists(conn, "attachments"):
            print("Missing attachments table. Run scripts/index_zoho_attachments.py first.", file=sys.stderr)
            return 1

        total = one(conn, "SELECT COUNT(*) FROM attachments")
        matched = one(conn, "SELECT COUNT(*) FROM attachments WHERE COALESCE(zip_file, '') <> ''")
        unmatched = total - matched
        parent_counts = grouped_counts(
            conn,
            """
            SELECT COALESCE(NULLIF(parent_module, ''), '[unknown]') AS parent_module, COUNT(*)
            FROM attachments
            GROUP BY parent_module
            ORDER BY COUNT(*) DESC, parent_module
            """,
        )
        confidence_counts = grouped_counts(
            conn,
            """
            SELECT mapping_confidence, COUNT(*)
            FROM attachments
            GROUP BY mapping_confidence
            ORDER BY COUNT(*) DESC, mapping_confidence
            """,
        )
        ext_counts = extension_counts(conn)
        missing, imported, module_count_lookup = missing_modules(conn)
        recs = recommendations(missing, module_count_lookup)
        direct_rollup, indirect_rollups, total_rollup = account_rollup_coverage(conn)
        unrollable = unrollable_by_module(parent_counts, direct_rollup, indirect_rollups)

        print("ZOHO ATTACHMENT REPORT")
        print(f"- Database: {args.db}")
        print(f"- Total attachment rows: {total:,}")
        print(f"- Matched attachments: {matched:,}")
        print(f"- Unmatched attachments: {unmatched:,}")
        print()

        print("ACCOUNT ROLLUP COVERAGE")
        print(f"- Direct Account attachments: {direct_rollup:,}")
        print(f"- Indirect account-rollup attachments: {sum(count for _, count in indirect_rollups):,}")
        print(f"- Total attachments currently rollable to Account: {total_rollup:,}")
        for module, count in indirect_rollups:
            print(f"- {module}: {count:,}")
        print()

        print_counts("COUNTS BY PARENT MODULE", parent_counts)
        print_counts("COUNTS BY FILE EXTENSION", ext_counts, limit=25)
        print_counts("COUNTS BY MAPPING CONFIDENCE", confidence_counts)
        print_counts("UNROLLABLE ATTACHMENTS BY MODULE", unrollable)

        print("LARGEST 25 ATTACHMENTS")
        for row in largest_attachments(conn):
            print(
                f"- {human_size(row['size'])}: {row['original_filename']} "
                f"[module={row['parent_module'] or '[unknown]'}, parent={row['parent_record_id']}, "
                f"zip={row['zip_file'] or '[unmatched]'}, confidence={row['mapping_confidence']}]"
            )
        print()

        print("TOP 25 PARENT RECORDS BY ATTACHMENT COUNT")
        for row in top_parents(conn):
            print(
                f"- {row['attachment_count']:,} files, {human_size(row['total_size'])}: "
                f"{row['parent_module']} {row['parent_record_id']}"
            )
        print()

        print_joined_samples(conn)

        print("MISSING MODULES")
        print(f"- Imported app modules with attachment coverage: {', '.join(sorted(imported))}")
        if missing:
            for module in missing:
                print(f"- {module}: {module_count_lookup.get(module, 0):,} attachment rows")
        else:
            print("- None")
        print()

        print("RECOMMENDATIONS")
        if recs:
            for module, count, reason in recs:
                print(f"- Import `{module}` minimally next ({count:,} attachment rows). {reason}")
        else:
            print("- No missing attachment parent modules found.")

    finally:
        conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
