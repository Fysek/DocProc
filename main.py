import argparse
from pathlib import Path

import tasks


def main():
    parser = argparse.ArgumentParser(
        description="Aplikacja CLI do zarządzania dokumentami pracowników"
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    parser_check = subparsers.add_parser(
        "check", help="Weryfikuje kompletność dokumentów i ujednolica nazwy"
    )
    parser_check.add_argument(
        "--path", required=True, type=Path, help="Ścieżka do głównego folderu z bazą"
    )

    parser_check.add_argument(
        "--hide-complete",
        action="store_true",
        help="Nie wypisuj pracowników, którzy mają komplet (pokaż tylko braki)",
    )

    parser_pack = subparsers.add_parser(
        "pack", help="Tworzy paczki ZIP z dokumentami wg listy JSON"
    )
    parser_pack.add_argument(
        "--path", required=True, type=Path, help="Ścieżka do głównego folderu z bazą"
    )
    parser_pack.add_argument("json_file", type=Path, help="Plik .json z listą nazwisk")

    parser_pack.add_argument(
        "--allow-incomplete",
        action="store_true",
        help="Ignoruj braki i pakuj dokumenty, które są dostępne",
    )

    parser_excel = subparsers.add_parser(
        "excel", help="Generuje raport XLSX ze statusem dokumentów"
    )
    parser_excel.add_argument(
        "--path", required=True, type=Path, help="Ścieżka do głównego folderu z bazą"
    )

    parser_compress = subparsers.add_parser(
        "compress", help="Wyszukuje i kompresuje pliki PDF > 1MB"
    )
    parser_compress.add_argument(
        "--path", required=True, type=Path, help="Ścieżka do głównego folderu z bazą"
    )

    args = parser.parse_args()

    if not args.path.exists() or not args.path.is_dir():
        print(f"[!] Błąd: Ścieżka '{args.path}' nie istnieje lub nie jest folderem.")
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


if __name__ == "__main__":
    main()
