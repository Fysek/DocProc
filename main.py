import argparse
from pathlib import Path

import tasks


def main():
    parser = argparse.ArgumentParser(
        description="CLI Application for managing employee documents"
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # Command: check
    parser_check = subparsers.add_parser(
        "check", help="Validates document completeness and standardizes filenames"
    )
    parser_check.add_argument(
        "--path", required=True, type=Path, help="Path to the main database folder"
    )
    parser_check.add_argument(
        "--hide-complete",
        action="store_true",
        help="Hide employees with complete documentation",
    )

    # Command: pack
    parser_pack = subparsers.add_parser(
        "pack", help="Creates a master ZIP archive based on a JSON list"
    )
    parser_pack.add_argument(
        "--path", required=True, type=Path, help="Path to the main database folder"
    )
    parser_pack.add_argument(
        "json_file", type=Path, help="Path to the .json file containing employee names"
    )
    parser_pack.add_argument(
        "--allow-incomplete",
        action="store_true",
        help="Pack available documents even if incomplete",
    )

    # Command: excel (stara wersja - skanuje i generuje na żywo)
    parser_excel = subparsers.add_parser(
        "excel", help="Generates an XLSX report directly from folders (legacy)"
    )
    parser_excel.add_argument(
        "--path", required=True, type=Path, help="Path to the main database folder"
    )

    # Command: compress
    parser_compress = subparsers.add_parser(
        "compress", help="Finds and compresses PDF files"
    )
    parser_compress.add_argument(
        "--path", required=True, type=Path, help="Path to the main database folder"
    )

    # --- NOWE KOMENDY ---

    # Command: updatedb (Generuje/Aktualizuje JSONa)
    parser_updatedb = subparsers.add_parser(
        "updatedb",
        help="Scans folders and updates database.json with document statuses and dates",
    )
    parser_updatedb.add_argument(
        "--path", required=True, type=Path, help="Path to the main database folder"
    )
    parser_updatedb.add_argument(
        "--db-file",
        type=Path,
        default=Path("database.json"),
        help="Path to the JSON database file",
    )
    parser_updatedb.add_argument(
        "--force",
        action="store_true",
        help="Force OCR to run again and overwrite existing dates",
    )

    # Command: excel-db (Generuje Excela z JSONa)
    parser_excel_db = subparsers.add_parser(
        "excel-db", help="Generates an XLSX report from the database.json file"
    )
    parser_excel_db.add_argument(
        "--db-file",
        type=Path,
        default=Path("database.json"),
        help="Path to the JSON database file",
    )

    args = parser.parse_args()

    # Routing
    if args.command in ["check", "pack", "excel", "compress", "updatedb"]:
        if not args.path.exists() or not args.path.is_dir():
            print(
                f"[!] Błąd: Ścieżka '{args.path}' nie istnieje lub nie jest folderem."
            )
            return

    if args.command == "check":
        tasks.run_check(args.path, args.hide_complete)
    elif args.command == "pack":
        if not args.json_file.exists():
            print(f"[!] Błąd: Plik '{args.json_file}' nie istnieje.")
            return
        tasks.run_pack(args.path, args.json_file, args.allow_incomplete)
    elif args.command == "excel":
        tasks.run_excel(args.path)
    elif args.command == "compress":
        tasks.run_compress(args.path)
    elif args.command == "updatedb":
        tasks.run_update_db(args.path, args.db_file, args.force)
    elif args.command == "excel-db":
        tasks.run_excel_from_db(args.db_file)


if __name__ == "__main__":
    main()
