#!/usr/bin/env python3
import argparse
import csv
import json
import sqlite3
import sys
import zipfile
from collections import Counter, defaultdict
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
BACKUP_DIR = ROOT / "zoho backup"
DB_PATH = ROOT / "data" / "zoho.sqlite3"

STRUCTURED_EXTENSIONS = {".csv", ".json", ".xml", ".properties", ".txt"}
SAMPLE_ROWS = 3
SAMPLE_ENTRIES = 8
IMPORTANT_DATA_CSVS = {
    "Data/Attachments_001.csv",
    "Data/Accounts_001.csv",
    "Data/Contacts_001.csv",
    "Data/Deals_001.csv",
    "Data/Cases_001.csv",
    "Data/Notes_001.csv",
    "Data/Referrals_C_001.csv",
    "Data/ZohoSign Documents_C_001.csv",
    "Metadata/Modules_001.csv",
    "Metadata/Fields_001.csv",
}


def human_size(size):
    value = float(size)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if value < 1024 or unit == "TB":
            if unit == "B":
                return f"{int(value)} {unit}"
            return f"{value:.1f} {unit}"
        value /= 1024
    return f"{size} B"


def natural_key(path):
    return [int(part) if part.isdigit() else part.lower() for part in path.stem.replace("_", " ").split()]


def zip_files(backup_dir):
    return sorted(backup_dir.glob("*.zip"), key=natural_key)


def safe_decode_lines(binary_file):
    for line in binary_file:
        yield line.decode("utf-8-sig", errors="replace")


def is_file(info):
    return not info.is_dir()


def folder_key(name, depth=2):
    parts = [part for part in name.split("/") if part]
    if not parts:
        return "[root]"
    return "/".join(parts[:depth])


def sample_zip_entries(infos, limit=SAMPLE_ENTRIES):
    files = [info for info in infos if is_file(info)]
    return files[:limit]


def largest_zip_entries(infos, limit=5):
    files = [info for info in infos if is_file(info)]
    return sorted(files, key=lambda info: info.file_size, reverse=True)[:limit]


def structured_entries(infos):
    return [
        info
        for info in infos
        if is_file(info) and Path(info.filename).suffix.lower() in STRUCTURED_EXTENSIONS
    ]


def read_csv_sample(archive, info, max_rows=SAMPLE_ROWS):
    with archive.open(info) as raw:
        reader = csv.DictReader(safe_decode_lines(raw))
        rows = []
        for _, row in zip(range(max_rows), reader):
            rows.append(row)
    return reader.fieldnames or [], rows


def compact_row(row, interesting_fields=None):
    if interesting_fields:
        keys = [key for key in interesting_fields if key in row]
    else:
        keys = list(row)[:8]
    return {key: row.get(key, "") for key in keys}


def inspect_zip(zip_path):
    with zipfile.ZipFile(zip_path) as archive:
        infos = archive.infolist()
        files = [info for info in infos if is_file(info)]
        dirs = Counter(folder_key(info.filename, 1) for info in infos)
        structures = Counter(folder_key(info.filename, 2) for info in infos)
        extensions = Counter(Path(info.filename).suffix.lower() or "[none]" for info in files)
        uncompressed = sum(info.file_size for info in files)
        compressed = sum(info.compress_size for info in files)

        return {
            "path": zip_path,
            "entries": len(infos),
            "files": len(files),
            "compressed_size": zip_path.stat().st_size,
            "zip_compressed_members": compressed,
            "uncompressed_size": uncompressed,
            "top_dirs": dirs,
            "structures": structures,
            "extensions": extensions,
            "structured": structured_entries(infos),
            "samples": sample_zip_entries(infos),
            "largest": largest_zip_entries(infos),
        }


def collect_attachment_zip_entries(zip_paths):
    names = {}
    extensions = Counter()
    duplicates = Counter()
    prefix_counts = Counter()
    zip_counts = Counter()
    total_entries = 0

    for zip_path in zip_paths:
        with zipfile.ZipFile(zip_path) as archive:
            for info in archive.infolist():
                if not is_file(info):
                    continue
                total_entries += 1
                base_name = Path(info.filename).name
                if base_name in names:
                    duplicates[base_name] += 1
                names[base_name] = {
                    "zip": zip_path.name,
                    "path": info.filename,
                    "size": info.file_size,
                    "compressed_size": info.compress_size,
                }
                zip_counts[zip_path.name] += 1
                extensions[Path(base_name).suffix.lower() or "[none]"] += 1
                prefix_counts[base_name.split("_", 1)[0] if "_" in base_name else "[no underscore]"] += 1

    return {
        "names": names,
        "extensions": extensions,
        "duplicates": duplicates,
        "prefix_counts": prefix_counts,
        "zip_counts": zip_counts,
        "total_entries": total_entries,
    }


def read_attachment_metadata(data_zip_path):
    rows = []
    if not data_zip_path.exists():
        return None

    with zipfile.ZipFile(data_zip_path) as archive:
        candidates = [name for name in archive.namelist() if name.lower().endswith("/attachments_001.csv")]
        if not candidates:
            return None

        name = candidates[0]
        with archive.open(name) as raw:
            reader = csv.DictReader(safe_decode_lines(raw))
            fieldnames = reader.fieldnames or []
            status_counts = Counter()
            document_counts = Counter()
            parent_counts = Counter()
            field_counts = Counter()
            link_url_count = 0
            total_size = 0
            sample_rows = []

            for index, row in enumerate(reader, 1):
                record_id = row.get("Record Id", "")
                parent_id = row.get("Parent.id", "")
                size_text = row.get("Size", "") or "0"
                try:
                    total_size += int(size_text)
                except ValueError:
                    pass

                rows.append(
                    {
                        "record_id": record_id,
                        "file_name": row.get("File Name", ""),
                        "size": row.get("Size", ""),
                        "parent_id": parent_id,
                        "status": row.get("Record Status", ""),
                        "documents": row.get("Documents", ""),
                        "created_time": row.get("Created Time", ""),
                        "modified_time": row.get("Modified Time", ""),
                        "old_attachment_id": row.get("Old Attachment Id", ""),
                        "field_id": row.get("Field.id", ""),
                    }
                )
                status_counts[row.get("Record Status", "")] += 1
                document_counts[row.get("Documents", "")] += 1
                parent_counts[parent_id or "[blank]"] += 1
                field_counts[row.get("Field.id", "") or "[blank]"] += 1
                if row.get("Link URL", ""):
                    link_url_count += 1
                if len(sample_rows) < SAMPLE_ROWS:
                    sample_rows.append(compact_row(row, [
                        "Record Id",
                        "File Name",
                        "Size",
                        "Parent.id",
                        "Record Status",
                        "Documents",
                        "Created Time",
                        "Modified Time",
                    ]))

    return {
        "csv_name": name,
        "fieldnames": fieldnames,
        "rows": rows,
        "row_count": len(rows),
        "status_counts": status_counts,
        "document_counts": document_counts,
        "parent_counts": parent_counts,
        "field_counts": field_counts,
        "link_url_count": link_url_count,
        "total_size": total_size,
        "sample_rows": sample_rows,
    }


def compare_attachment_files(metadata, attachment_files):
    if not metadata:
        return None

    file_by_name = attachment_files["names"]
    metadata_ids = {row["record_id"] for row in metadata["rows"] if row["record_id"]}
    file_names = set(file_by_name)
    matched = metadata_ids & file_names
    missing_files = sorted(metadata_ids - file_names)[:20]
    extra_files = sorted(file_names - metadata_ids)[:20]

    size_mismatches = []
    located_zip_counts = Counter()
    for row in metadata["rows"]:
        record_id = row["record_id"]
        found = file_by_name.get(record_id)
        if not found:
            continue
        located_zip_counts[found["zip"]] += 1
        try:
            metadata_size = int(row["size"] or 0)
        except ValueError:
            continue
        if metadata_size != found["size"] and len(size_mismatches) < 10:
            size_mismatches.append({
                "record_id": record_id,
                "metadata_size": metadata_size,
                "zip_size": found["size"],
                "zip": found["zip"],
            })

    return {
        "metadata_ids": len(metadata_ids),
        "file_names": len(file_names),
        "matched": len(matched),
        "missing_file_samples": missing_files,
        "extra_file_samples": extra_files,
        "size_mismatches": size_mismatches,
        "located_zip_counts": located_zip_counts,
    }


def sqlite_table_columns(conn, table):
    return {row[1] for row in conn.execute(f'PRAGMA table_info("{table}")')}


def compare_parents_to_db(metadata, db_path):
    if not metadata or not db_path.exists():
        return None

    conn = sqlite3.connect(db_path)
    try:
        tables = [
            row[0]
            for row in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
            )
        ]
        record_tables = [table for table in tables if "record_id" in sqlite_table_columns(conn, table)]
        parent_ids = {row["parent_id"] for row in metadata["rows"] if row["parent_id"]}
        parent_to_tables = defaultdict(list)

        for table in record_tables:
            rows = conn.execute(f'SELECT record_id FROM "{table}" WHERE record_id IS NOT NULL AND record_id <> ""')
            ids = {row[0] for row in rows}
            matched = parent_ids & ids
            for parent_id in matched:
                parent_to_tables[parent_id].append(table)

        table_counts = Counter()
        unresolved = 0
        for row in metadata["rows"]:
            parent_id = row["parent_id"]
            matches = parent_to_tables.get(parent_id, [])
            if matches:
                for table in matches:
                    table_counts[table] += 1
            else:
                unresolved += 1

        return {
            "db_path": db_path,
            "tables_checked": record_tables,
            "unique_parent_ids": len(parent_ids),
            "matched_unique_parent_ids": len(parent_to_tables),
            "attachment_counts_by_table": table_counts,
            "unresolved_attachment_rows": unresolved,
            "unresolved_parent_samples": [
                parent_id
                for parent_id in parent_ids
                if parent_id not in parent_to_tables
            ][:20],
        }
    finally:
        conn.close()


def module_name_from_data_csv(path):
    name = Path(path).name
    if name.endswith("_RecordAccess_001.csv"):
        return None
    for suffix in ("_001.csv", "_C_001.csv"):
        if name.endswith(suffix):
            return name[: -len(suffix)]
    return Path(path).stem


def compare_parents_to_backup_data(metadata, data_zip_path):
    if not metadata or not data_zip_path.exists():
        return None

    parent_ids = {row["parent_id"] for row in metadata["rows"] if row["parent_id"]}
    parent_to_module = {}

    with zipfile.ZipFile(data_zip_path) as archive:
        infos = [
            info
            for info in archive.infolist()
            if is_file(info)
            and info.filename.startswith("Data/")
            and info.filename.lower().endswith(".csv")
            and not info.filename.endswith("_RecordAccess_001.csv")
            and info.filename != "Data/Attachments_001.csv"
        ]

        for info in infos:
            module = module_name_from_data_csv(info.filename)
            if not module:
                continue
            with archive.open(info) as raw:
                reader = csv.DictReader(safe_decode_lines(raw))
                if "Record Id" not in (reader.fieldnames or []):
                    continue
                for row in reader:
                    record_id = row.get("Record Id", "")
                    if record_id in parent_ids and record_id not in parent_to_module:
                        parent_to_module[record_id] = module

    module_counts = Counter()
    unresolved = 0
    for row in metadata["rows"]:
        module = parent_to_module.get(row["parent_id"])
        if module:
            module_counts[module] += 1
        else:
            unresolved += 1

    return {
        "unique_parent_ids": len(parent_ids),
        "matched_unique_parent_ids": len(parent_to_module),
        "attachment_counts_by_module": module_counts,
        "unresolved_attachment_rows": unresolved,
        "unresolved_parent_samples": [
            parent_id
            for parent_id in parent_ids
            if parent_id not in parent_to_module
        ][:20],
    }


def inspect_data_zip_csvs(data_zip_path):
    if not data_zip_path.exists():
        return []

    csv_reports = []
    with zipfile.ZipFile(data_zip_path) as archive:
        infos = [info for info in archive.infolist() if is_file(info) and info.filename.lower().endswith(".csv")]
        for info in infos:
            if info.filename not in IMPORTANT_DATA_CSVS:
                continue
            fields, rows = read_csv_sample(archive, info)
            interesting = [
                field
                for field in (fields or [])
                if any(token in field.lower() for token in ("record", "parent", "module", "attach", "file", "document", ".id"))
            ][:12]
            csv_reports.append({
                "name": info.filename,
                "size": info.file_size,
                "fields": fields,
                "interesting_fields": interesting,
                "sample_rows": [compact_row(row, interesting or None) for row in rows[:1]],
            })
    return csv_reports


def print_zip_report(reports):
    print("ZIP INVENTORY")
    print(f"- Zip files: {len(reports)}")
    print(f"- Compressed on disk: {human_size(sum(report['compressed_size'] for report in reports))}")
    print(f"- Uncompressed member bytes: {human_size(sum(report['uncompressed_size'] for report in reports))}")
    print()

    for report in reports:
        print(f"{report['path'].name}")
        print(
            f"  files={report['files']:,} entries={report['entries']:,} "
            f"zip_size={human_size(report['compressed_size'])} "
            f"uncompressed={human_size(report['uncompressed_size'])}"
        )
        print(f"  top folders: {', '.join(f'{k} ({v:,})' for k, v in report['top_dirs'].most_common(5))}")
        print(f"  extensions: {', '.join(f'{k} ({v:,})' for k, v in report['extensions'].most_common(8))}")
        structured = report["structured"][:8]
        if structured:
            print("  structured files:")
            for info in structured:
                print(f"    - {info.filename} ({human_size(info.file_size)})")
        print("  representative entries:")
        for info in report["samples"]:
            print(f"    - {info.filename} ({human_size(info.file_size)})")
        print("  largest entries:")
        for info in report["largest"]:
            print(f"    - {info.filename} ({human_size(info.file_size)})")
        print()


def print_data_csv_report(csv_reports):
    if not csv_reports:
        return

    print("DATA ZIP CSV SAMPLES")
    print(f"- Key CSV files sampled: {len(csv_reports)}")
    for report in csv_reports:
        print(f"- {report['name']} ({human_size(report['size'])})")
        print(f"  columns={len(report['fields'])}; interesting={', '.join(report['interesting_fields'][:10]) or 'none detected'}")
        for row in report["sample_rows"]:
            print(f"  sample={json.dumps(row, ensure_ascii=False)}")
    print()


def print_attachment_summary(metadata, file_compare, db_compare, attachment_files):
    print("ATTACHMENT LINKAGE SUMMARY")
    if not metadata:
        print("- No Data/Attachments_001.csv metadata file found.")
        return

    print(f"- Metadata CSV: {metadata['csv_name']}")
    print(f"- Attachment metadata rows: {metadata['row_count']:,}")
    print(f"- Metadata-reported attachment bytes: {human_size(metadata['total_size'])}")
    print(f"- Record Status: {dict(metadata['status_counts'].most_common())}")
    print(f"- Documents values: {dict(metadata['document_counts'].most_common())}")
    print(f"- Non-empty Link URL rows: {metadata['link_url_count']:,}")
    print(f"- Unique Parent.id values: {len([k for k in metadata['parent_counts'] if k != '[blank]']):,}")
    print("- Attachment metadata samples:")
    for row in metadata["sample_rows"]:
        print(f"  {json.dumps(row, ensure_ascii=False)}")

    print("- Attachment zip file extensions:")
    print(f"  {', '.join(f'{k} ({v:,})' for k, v in attachment_files['extensions'].most_common(12))}")
    print(
        f"- Attachment zip entries: {attachment_files['total_entries']:,}; "
        f"unique basenames: {len(attachment_files['names']):,}; "
        f"duplicated basenames: {len(attachment_files['duplicates']):,}"
    )

    if file_compare:
        print("- Metadata-to-zip filename comparison:")
        print(
            f"  metadata ids={file_compare['metadata_ids']:,}; "
            f"zip basenames={file_compare['file_names']:,}; "
            f"matched={file_compare['matched']:,}"
        )
        print(f"  missing file samples={file_compare['missing_file_samples'][:5]}")
        print(f"  extra file samples={file_compare['extra_file_samples'][:5]}")
        print(f"  size mismatch samples={file_compare['size_mismatches'][:3]}")

    if db_compare:
        print("- Existing SQLite parent match:")
        print(f"  db={db_compare['db_path']}")
        print(f"  tables checked={', '.join(db_compare['tables_checked'])}")
        print(
            f"  unique parents matched={db_compare['matched_unique_parent_ids']:,} "
            f"of {db_compare['unique_parent_ids']:,}"
        )
        print(f"  attachment rows by matched table={dict(db_compare['attachment_counts_by_table'].most_common())}")
        print(f"  unresolved attachment rows={db_compare['unresolved_attachment_rows']:,}")
        print(f"  unresolved parent samples={db_compare['unresolved_parent_samples'][:5]}")
    print()


def print_backup_parent_summary(backup_compare):
    if not backup_compare:
        return

    print("BACKUP DATA PARENT MATCH")
    print(
        f"- Unique attachment parents matched to Data_001 module CSVs: "
        f"{backup_compare['matched_unique_parent_ids']:,} of {backup_compare['unique_parent_ids']:,}"
    )
    print(
        "- Attachment rows by backup module: "
        f"{dict(backup_compare['attachment_counts_by_module'].most_common(15))}"
    )
    print(f"- Unresolved attachment rows after scanning Data_001 module CSVs: {backup_compare['unresolved_attachment_rows']:,}")
    print(f"- Unresolved parent samples: {backup_compare['unresolved_parent_samples'][:5]}")
    print()


def main():
    parser = argparse.ArgumentParser(description="Inspect a Zoho backup without extracting attachment contents.")
    parser.add_argument("--backup-dir", type=Path, default=BACKUP_DIR)
    parser.add_argument("--db", type=Path, default=DB_PATH)
    args = parser.parse_args()

    if not args.backup_dir.exists():
        print(f"Missing backup directory: {args.backup_dir}", file=sys.stderr)
        return 1

    zips = zip_files(args.backup_dir)
    data_zip = args.backup_dir / "Data_001.zip"
    attachment_zips = [path for path in zips if path.name.startswith("Attachments_")]

    reports = [inspect_zip(path) for path in zips]
    attachment_files = collect_attachment_zip_entries(attachment_zips)
    attachment_metadata = read_attachment_metadata(data_zip)
    file_compare = compare_attachment_files(attachment_metadata, attachment_files)
    db_compare = compare_parents_to_db(attachment_metadata, args.db)
    backup_compare = compare_parents_to_backup_data(attachment_metadata, data_zip)
    data_csv_reports = inspect_data_zip_csvs(data_zip)

    print_zip_report(reports)
    print_data_csv_report(data_csv_reports)
    print_attachment_summary(attachment_metadata, file_compare, db_compare, attachment_files)
    print_backup_parent_summary(backup_compare)

    print("RECOMMENDATION SIGNALS")
    print("- Attachment zip paths use basename == Data/Attachments_001.csv Record Id.")
    print("- Data/Attachments_001.csv Parent.id is the record-level foreign key to Zoho module records.")
    print("- Keep attachment binaries external to SQLite; store zip name, inner path, size, filename, and parent linkage metadata only.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
