import argparse
from pathlib import Path
import tasks

def main():
    parser = argparse.ArgumentParser(description="Aplikacja CLI do zarządzania dokumentami pracowników")
    subparsers = parser.add_subparsers(dest="command", required=True)
    
    # Komenda: check
    parser_check = subparsers.add_parser("check", help="Weryfikuje kompletność dokumentów i ujednolica nazwy")
    parser_check.add_argument("--path", required=True, type=Path, help="Ścieżka do głównego folderu z bazą")
    
    # Komenda: pack
    parser_pack = subparsers.add_parser("pack", help="Tworzy paczki ZIP z dokumentami wg listy JSON")
    parser_pack.add_argument("--path", required=True, type=Path, help="Ścieżka do głównego folderu z bazą")
    parser_pack.add_argument("json_file", type=Path, help="Plik .json z listą nazwisk")
    
    # Komenda: excel
    parser_excel = subparsers.add_parser("excel", help="Generuje raport XLSX ze statusem dokumentów")
    parser_excel.add_argument("--path", required=True, type=Path, help="Ścieżka do głównego folderu z bazą")
    
    # Komenda: compress
    parser_compress = subparsers.add_parser("compress", help="Wyszukuje pliki PDF > 1MB do kompresji")
    parser_compress.add_argument("--path", required=True, type=Path, help="Ścieżka do głównego folderu z bazą")

    args = parser.parse_args()

    # Walidacja ścieżki
    if not args.path.exists() or not args.path.is_dir():
        print(f"[!] Błąd: Ścieżka '{args.path}' nie istnieje lub nie jest folderem.")
        return

    # Uruchamianie odpowiedniego zadania
    if args.command == "check":
        tasks.run_check(args.path)
    elif args.command == "pack":
        if not args.json_file.exists():
            print(f"[!] Błąd: Plik '{args.json_file}' nie istnieje.")
            return
        tasks.run_pack(args.path, args.json_file)
    elif args.command == "excel":
        tasks.run_excel(args.path)
    elif args.command == "compress":
        tasks.run_compress(args.path)

if __name__ == "__main__":
    main()