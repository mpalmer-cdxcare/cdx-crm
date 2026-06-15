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
ATTACHMENTS_CSV = "Data/Attachments_001.csv"
BATCH_SIZE = 1000


def q(identifier):
    return '"' + identifier.replace('"', '""') + '"'


def safe_decode_lines(binary_file):
    for line in binary_file:
        yield line.decode("utf-8-sig", errors="replace")


def natural_key(path):
    return [int(part) if part.isdigit() else part.lower() for part in path.stem.replace("_", " ").split()]


def human_size(size):
    value = float(size or 0)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if value < 1024 or unit == "TB":
            if unit == "B":
                return f"{int(value)} {unit}"
            return f"{value:.1f} {unit}"
        value /= 1024
    return f"{size} B"


def parse_int(value):
    if value in (None, ""):
        return None
    try:
        return int(value)
    except ValueError:
        return None


def module_name_from_data_csv(path):
    name = Path(path).name
    if name.endswith("_RecordAccess_001.csv"):
        return None
    for suffix in ("_001.csv", "_C_001.csv"):
        if name.endswith(suffix):
            return name[: -len(suffix)]
    return Path(path).stem


def attachment_zip_paths(backup_dir):
    return sorted(backup_dir.glob("Attachments_*.zip"), key=natural_key)


def read_attachment_metadata(data_zip_path):
    rows = []
    parent_ids = set()
    duplicate_metadata_rows = 0
    seen_ids = set()

    with zipfile.ZipFile(data_zip_path) as archive:
        if ATTACHMENTS_CSV not in archive.namelist():
            raise RuntimeError(f"{data_zip_path} does not contain {ATTACHMENTS_CSV}")

        with archive.open(ATTACHMENTS_CSV) as raw:
            reader = csv.DictReader(safe_decode_lines(raw))
            for row in reader:
                attachment_record_id = row.get("Record Id", "")
                if not attachment_record_id:
                    continue
                if attachment_record_id in seen_ids:
                    duplicate_metadata_rows += 1
                seen_ids.add(attachment_record_id)

                parent_id = row.get("Parent.id", "")
                if parent_id:
                    parent_ids.add(parent_id)

                rows.append(
                    {
                        "attachment_record_id": attachment_record_id,
                        "parent_record_id": parent_id,
                        "original_filename": row.get("File Name", ""),
                        "metadata_file_size": parse_int(row.get("Size", "")),
                        "created_time": row.get("Created Time", ""),
                        "modified_time": row.get("Modified Time", ""),
                        "record_status": row.get("Record Status", ""),
                        "raw_data": json.dumps(row, ensure_ascii=False, separators=(",", ":")),
                    }
                )

    return rows, parent_ids, duplicate_metadata_rows


def scan_attachment_zips(zip_paths, attachment_ids):
    entries = {}
    duplicate_entries_ignored = 0
    unmatched_zip_entries = 0

    for zip_path in zip_paths:
        with zipfile.ZipFile(zip_path) as archive:
            for info in archive.infolist():
                if info.is_dir():
                    continue
                basename = Path(info.filename).name
                if basename not in attachment_ids:
                    unmatched_zip_entries += 1
                    continue
                if basename in entries:
                    duplicate_entries_ignored += 1
                    continue
                entries[basename] = {
                    "zip_file": zip_path.name,
                    "zip_inner_path": info.filename,
                    "zip_file_size": info.file_size,
                }

    return entries, duplicate_entries_ignored, unmatched_zip_entries


def infer_parent_modules(data_zip_path, parent_ids):
    parent_modules = {}

    with zipfile.ZipFile(data_zip_path) as archive:
        csv_infos = [
            info
            for info in archive.infolist()
            if not info.is_dir()
            and info.filename.startswith("Data/")
            and info.filename.lower().endswith(".csv")
            and info.filename != ATTACHMENTS_CSV
            and not info.filename.endswith("_RecordAccess_001.csv")
        ]

        for info in csv_infos:
            module = module_name_from_data_csv(info.filename)
            if not module:
                continue
            with archive.open(info) as raw:
                reader = csv.DictReader(safe_decode_lines(raw))
                if "Record Id" not in (reader.fieldnames or []):
                    continue
                for row in reader:
                    record_id = row.get("Record Id", "")
                    if record_id in parent_ids and record_id not in parent_modules:
                        parent_modules[record_id] = module
            if len(parent_modules) == len(parent_ids):
                break

    return parent_modules


def create_attachments_table(conn):
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS attachments (
            attachment_record_id TEXT PRIMARY KEY,
            parent_record_id TEXT,
            parent_module TEXT,
            original_filename TEXT,
            zip_file TEXT,
            zip_inner_path TEXT,
            metadata_file_size INTEGER,
            zip_file_size INTEGER,
            created_time TEXT,
            modified_time TEXT,
            record_status TEXT,
            mapping_confidence TEXT NOT NULL,
            raw_data TEXT NOT NULL,
            indexed_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );

        CREATE INDEX IF NOT EXISTS idx_attachments_attachment_record_id
            ON attachments (attachment_record_id);
        CREATE INDEX IF NOT EXISTS idx_attachments_parent_record_id
            ON attachments (parent_record_id);
        CREATE INDEX IF NOT EXISTS idx_attachments_parent_module
            ON attachments (parent_module);
        CREATE INDEX IF NOT EXISTS idx_attachments_original_filename
            ON attachments (original_filename);
        """
    )


def build_insert_rows(metadata_rows, zip_entries, parent_modules):
    for row in metadata_rows:
        attachment_id = row["attachment_record_id"]
        zip_entry = zip_entries.get(attachment_id, {})
        parent_id = row["parent_record_id"]
        parent_module = parent_modules.get(parent_id, "")

        if zip_entry and parent_module:
            confidence = "metadata_zip_parent_module"
        elif zip_entry:
            confidence = "metadata_zip"
        elif parent_module:
            confidence = "metadata_parent_module"
        else:
            confidence = "metadata_only"

        yield (
            attachment_id,
            parent_id,
            parent_module,
            row["original_filename"],
            zip_entry.get("zip_file", ""),
            zip_entry.get("zip_inner_path", ""),
            row["metadata_file_size"],
            zip_entry.get("zip_file_size"),
            row["created_time"],
            row["modified_time"],
            row["record_status"],
            confidence,
            row["raw_data"],
        )


def rebuild_attachments(conn, insert_rows):
    create_attachments_table(conn)
    conn.execute("DELETE FROM attachments")

    sql = """
        INSERT INTO attachments (
            attachment_record_id,
            parent_record_id,
            parent_module,
            original_filename,
            zip_file,
            zip_inner_path,
            metadata_file_size,
            zip_file_size,
            created_time,
            modified_time,
            record_status,
            mapping_confidence,
            raw_data
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(attachment_record_id) DO UPDATE SET
            parent_record_id = excluded.parent_record_id,
            parent_module = excluded.parent_module,
            original_filename = excluded.original_filename,
            zip_file = excluded.zip_file,
            zip_inner_path = excluded.zip_inner_path,
            metadata_file_size = excluded.metadata_file_size,
            zip_file_size = excluded.zip_file_size,
            created_time = excluded.created_time,
            modified_time = excluded.modified_time,
            record_status = excluded.record_status,
            mapping_confidence = excluded.mapping_confidence,
            raw_data = excluded.raw_data,
            indexed_at = CURRENT_TIMESTAMP
    """

    batch = []
    inserted = 0
    for values in insert_rows:
        batch.append(values)
        if len(batch) >= BATCH_SIZE:
            conn.executemany(sql, batch)
            inserted += len(batch)
            batch = []
    if batch:
        conn.executemany(sql, batch)
        inserted += len(batch)
    return inserted


def summarize(conn):
    def one(sql, params=()):
        return conn.execute(sql, params).fetchone()[0]

    parent_modules = Counter(
        dict(
            conn.execute(
                """
                SELECT COALESCE(NULLIF(parent_module, ''), '[unknown]') AS module, COUNT(*)
                FROM attachments
                GROUP BY module
                """
            ).fetchall()
        )
    )
    confidence = Counter(
        dict(
            conn.execute(
                """
                SELECT mapping_confidence, COUNT(*)
                FROM attachments
                GROUP BY mapping_confidence
                """
            ).fetchall()
        )
    )
    extensions = Counter()
    for (filename,) in conn.execute("SELECT original_filename FROM attachments"):
        extensions[Path(filename or "").suffix.lower() or "[none]"] += 1

    return {
        "rows": one("SELECT COUNT(*) FROM attachments"),
        "matched": one("SELECT COUNT(*) FROM attachments WHERE zip_file <> ''"),
        "unmatched": one("SELECT COUNT(*) FROM attachments WHERE zip_file = ''"),
        "parent_modules": parent_modules,
        "confidence": confidence,
        "extensions": extensions,
    }


def main():
    parser = argparse.ArgumentParser(description="Index Zoho attachment metadata without extracting attachment files.")
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

    metadata_rows, parent_ids, duplicate_metadata_rows = read_attachment_metadata(data_zip_path)
    attachment_ids = {row["attachment_record_id"] for row in metadata_rows}
    zip_entries, duplicate_zip_entries_ignored, unmatched_zip_entries = scan_attachment_zips(
        attachment_zip_paths(args.backup_dir),
        attachment_ids,
    )
    parent_modules = infer_parent_modules(data_zip_path, parent_ids)

    conn = sqlite3.connect(args.db)
    try:
        conn.execute("PRAGMA journal_mode = WAL")
        conn.execute("PRAGMA synchronous = NORMAL")
        with conn:
            processed_rows = rebuild_attachments(conn, build_insert_rows(metadata_rows, zip_entries, parent_modules))
        summary = summarize(conn)
    finally:
        conn.close()

    print("ZOHO ATTACHMENT INDEX SUMMARY")
    print(f"- Database: {args.db}")
    print(f"- Metadata source: {data_zip_path.name}:{ATTACHMENTS_CSV}")
    print(f"- Metadata rows read: {len(metadata_rows):,}")
    print(f"- Unique attachment Record Ids: {len(attachment_ids):,}")
    print(f"- Duplicate metadata rows updated by primary key: {duplicate_metadata_rows:,}")
    print(f"- Rows processed into attachments table: {processed_rows:,}")
    print(f"- Attachments table rows: {summary['rows']:,}")
    print(f"- Matched attachments: {summary['matched']:,}")
    print(f"- Unmatched metadata rows: {summary['unmatched']:,}")
    print(f"- Duplicate zip entries ignored: {duplicate_zip_entries_ignored:,}")
    print(f"- Zip entries with no metadata row ignored: {unmatched_zip_entries:,}")
    print(f"- Metadata file size total: {human_size(sum(row['metadata_file_size'] or 0 for row in metadata_rows))}")
    print(f"- Zip file size total for matched rows: {human_size(sum(entry['zip_file_size'] for entry in zip_entries.values()))}")
    print(f"- Mapping confidence: {dict(summary['confidence'].most_common())}")
    print(f"- Counts by parent module: {dict(summary['parent_modules'].most_common())}")
    print(f"- Counts by file extension: {dict(summary['extensions'].most_common(20))}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
