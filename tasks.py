import json
import re
import shutil
import subprocess
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl
import pytesseract
from pdf2image import convert_from_path
from PIL import Image

from core import REQUIRED_DOCS, Worker


def get_workers(base_path: Path):
    """Returns a list of Worker objects from the given base directory."""
    workers = []
    for p in base_path.iterdir():
        # no _archive
        if p.is_dir() and p.name != "_archive" and not p.name.startswith("."):
            workers.append(Worker(p))
    return workers


def run_check(base_path: Path, hide_complete: bool = False):
    """Validates document completeness, archives old files, and standardizes filenames."""
    workers = get_workers(base_path)

    # Define the archive directory path
    archive_dir = base_path / "_archive"

    for w in workers:
        safe_name = w.name.title().replace(" ", "_")

        # 1. Archive obsolete documents first
        if w.obsolete_docs:
            archive_dir.mkdir(
                exist_ok=True
            )  # Create _archive folder if it doesn't exist

            for old_file in w.obsolete_docs:
                # Add timestamp to avoid filename collisions in the shared archive
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                archived_name = f"{safe_name}_{timestamp}_{old_file.name}"
                new_path = archive_dir / archived_name

                old_file.rename(new_path)
                print(
                    f"    [A] Zarchiwizowano starszy dokument: {old_file.name} -> _archive/{archived_name}"
                )

        # 2. Print status
        if w.is_complete:
            if not hide_complete:
                print(f"[✓] {w.name} - Komplet")
        else:
            missing = ", ".join(w.missing_docs)
            print(f"[x] {w.name} - Braki: {missing}")

        # 3. Standardize filenames of the valid documents
        for doc_type, file_path in w.documents.items():
            ext = file_path.suffix.lower()
            doc_type_clean = doc_type.lower()
            expected_name = f"{safe_name}_{doc_type_clean}{ext}"

            if file_path.name != expected_name:
                new_path = file_path.with_name(expected_name)
                file_path.rename(new_path)
                w.documents[doc_type] = new_path


def run_pack(base_path: Path, json_path: Path, allow_incomplete: bool = False):
    with open(json_path, "r", encoding="utf-8") as f:
        target_names = json.load(f)

    workers_dict = {w.name: w for w in get_workers(base_path)}
    now_str = datetime.now().strftime("%d_%m_%Y_%H_%M")
    master_zip_name = f"paczka_pracownicy_{now_str}.zip"

    added_any = False

    with zipfile.ZipFile(master_zip_name, "w") as zf:
        for name in target_names:
            if name not in workers_dict:
                print(f"[!] Nie znaleziono folderu dla: {name}")
                continue

            w = workers_dict[name]

            # Jeśli pracownik ma braki i NIE użyliśmy flagi pozwalającej na pakowanie z brakami
            if not w.is_complete and not allow_incomplete:
                print(f"[x] {w.name} - Pominięto, braki: {', '.join(w.missing_docs)}")
                continue

            safe_folder_name = name.replace(" ", "_")
            for doc_path in w.documents.values():
                arcname = f"{safe_folder_name}/{doc_path.name}"
                zf.write(doc_path, arcname)

            # Zmieniamy komunikat, jeśli pakujemy osobę z brakami
            if w.is_complete:
                print(f"[+] Dodano do paczki (komplet): {w.name}")
            else:
                print(
                    f"[~] Dodano do paczki (CZĘŚCIOWO): {w.name} (Brakuje: {', '.join(w.missing_docs)})"
                )

            added_any = True

    if added_any:
        print(f"\n[✓] Gotowe! Utworzono paczkę: {master_zip_name}")
    else:
        print(
            f"\n[!] Paczka nie została utworzona (brak spełniających warunki pracowników)."
        )
        Path(master_zip_name).unlink(missing_ok=True)


def run_excel(base_path: Path):
    """Generates an Excel report with document statuses and extracted dates."""
    workers = get_workers(base_path)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Status Dokumentów"

    headers = ["Imię i Nazwisko"] + [d.upper() for d in REQUIRED_DOCS]
    ws.append(headers)

    for w in workers:
        row = [w.name]
        for doc_type in REQUIRED_DOCS:
            if doc_type in w.documents:
                file_path = w.documents[doc_type]

                if doc_type == "foto":
                    row.append("Tak")
                elif doc_type == "bhp":
                    # Run OCR for BHP document
                    print(f"[*] Skanowanie daty BHP dla: {w.name}...")
                    date_val = extract_bhp_date(file_path)
                    print(date_val)
                    row.append(date_val)
                else:
                    row.append("Wymaga integracji OCR")  # Placeholder for other docs
            else:
                row.append("Brak")
        ws.append(row)

    output_file = "raport_dokumentow.xlsx"
    wb.save(output_file)
    print(f"[✓] Wygenerowano raport Excel: {output_file}")


def run_compress(base_path: Path):
    # Sprawdzamy, czy Ghostscript jest zainstalowany w systemie
    if not shutil.which("gs"):
        print("[!] Error: Cannot find Ghostscript.")
        print("    Install with: sudo apt install ghostscript")
        return

    workers = get_workers(base_path)
    for w in workers:
        for doc_type, file_path in w.documents.items():
            if file_path.suffix.lower() == ".pdf":
                size_mb = file_path.stat().st_size / (1024 * 1024)

                if size_mb > 1.0:
                    print(
                        f"[*] Compressing: {file_path.name} ({size_mb:.2f} MB) in folder: {w.name}"
                    )
                    temp_path = file_path.with_name(f"temp_{file_path.name}")

                    # Twoja komenda z agresywnym, ręcznym sterowaniem downsamplingiem
                    gs_cmd = [
                        "gs",
                        "-sDEVICE=pdfwrite",
                        "-dCompatibilityLevel=1.4",
                        "-dNOPAUSE",
                        "-dQUIET",
                        "-dBATCH",
                        # Wymuszenie downsamplingu
                        "-dDownsampleColorImages=true",
                        "-dDownsampleGrayImages=true",
                        "-dDownsampleMonoImages=true",
                        # Typ algorytmu
                        "-dColorImageDownsampleType=/Bicubic",
                        "-dGrayImageDownsampleType=/Bicubic",
                        "-dMonoImageDownsampleType=/Subsample",
                        # Rozdzielczość docelowa (DPI)
                        "-dColorImageResolution=110",
                        "-dGrayImageResolution=110",
                        "-dMonoImageResolution=300",
                        f"-sOutputFile={temp_path}",
                        str(file_path),
                    ]

                    try:
                        subprocess.run(gs_cmd, check=True)

                        new_size_mb = temp_path.stat().st_size / (1024 * 1024)

                        if new_size_mb < size_mb:
                            temp_path.replace(file_path)
                            zaoszczedzono = size_mb - new_size_mb
                            print(
                                f"    [✓] Sukces! Zmniejszono do {new_size_mb:.2f} MB (zaoszczędzono {zaoszczedzono:.2f} MB)"
                            )
                        else:
                            temp_path.unlink()
                            print(
                                "    [-] Plik jest już optymalnie skompresowany, pomijam."
                            )

                    except subprocess.CalledProcessError as e:
                        print(
                            f"    [!] Error podczas kompresji pliku {file_path.name}: {e}"
                        )
                        if temp_path.exists():
                            temp_path.unlink()


def extract_bhp_date(file_path: Path):
    """
    Extracts the training completion date via OCR, filters out dates before 2024
    (e.g., old laws), picks the latest valid date, and calculates expiration (+1 years).
    """
    try:
        # Load document depending on its extension
        if file_path.suffix.lower() == ".pdf":
            images = convert_from_path(file_path)
            text = pytesseract.image_to_string(images[0], lang="pol")
        else:
            text = pytesseract.image_to_string(Image.open(file_path), lang="pol")

        text_lower = text.lower()
        dates_found = []

        # Pattern 1: All dates in DD.MM.YYYY or DD-MM-YYYY format
        for match in re.finditer(r"(\d{2})[\.\-](\d{2})[\.\-](\d{4})", text_lower):
            day, month, year = match.groups()
            dates_found.append((int(year), int(month), int(day)))

        # Pattern 2: All textual dates (e.g., 28 stycznia 2026)
        months = {
            "stycznia": 1,
            "lutego": 2,
            "marca": 3,
            "kwietnia": 4,
            "maja": 5,
            "czerwca": 6,
            "lipca": 7,
            "sierpnia": 8,
            "września": 9,
            "października": 10,
            "listopada": 11,
            "grudnia": 12,
        }
        for match in re.finditer(
            r"(\d{1,2})\s+([a-ząćęłńóśźż]+)\s+(\d{4})", text_lower
        ):
            day, month_str, year = match.groups()
            if month_str in months:
                dates_found.append((int(year), months[month_str], int(day)))

        # Filter and validate dates
        valid_dates = []
        for y, m, d in dates_found:
            # Ignore anything before 2024 to bypass old laws/regulations
            if y >= 2024:
                try:
                    # Validate if it's a real calendar date (e.g., catching 31.02.2025 errors)
                    valid_dates.append(datetime(y, m, d))
                except ValueError:
                    pass

        if valid_dates:
            # Pick the most recent date found in the document
            completion_date = max(valid_dates)

            # Add 1 years for validity
            expiration_date = completion_date + timedelta(days=1 * 365)
            return expiration_date.strftime("%Y-%m-%d")
        else:
            return "No data > 2023"

    except Exception as e:
        print(f"    [!] Error OCR dla pliku {file_path.name}: {e}")
        return "Error odczytu"


def extract_g2e_date(file_path: Path):
    """
    Extracts the expiration date from a G2E qualification certificate using OCR.
    Searches for the specific phrase 'ważne do dnia' and captures the subsequent YYYY-MM-DD date.
    """
    try:
        # 1. Load document depending on its extension
        if file_path.suffix.lower() == ".pdf":
            images = convert_from_path(file_path)
            text = pytesseract.image_to_string(images[0], lang="pol")
        else:
            text = pytesseract.image_to_string(Image.open(file_path), lang="pol")

        # Convert text to lowercase to make regex matching case-insensitive
        text_lower = text.lower()

        # 2. Extract the specific expiration date
        # The Regex pattern looks for "ważne do dnia" (valid until), allows for newlines or spaces ([\s\S]*?),
        # and captures the first date in YYYY-MM-DD or YYYY.MM.DD format.
        # We use wa[zż]ne to account for potential OCR typos missing the Polish 'ż'.
        match = re.search(
            r"wa[zż]ne do dnia[\s\S]*?(\d{4}[-.]\d{2}[-.]\d{2})", text_lower
        )

        if match:
            raw_date = match.group(1)
            # Normalize dots to hyphens just in case OCR read 2031.06.17 instead of 2031-06-17
            expiration_date = raw_date.replace(".", "-")
            return expiration_date
        else:
            return "Brak daty ważności"

    except Exception as e:
        print(f"[!] OCR Error for file {file_path.name}: {e}")
        return "Błąd odczytu"


def run_update_db(base_path: Path, db_path: Path, force: bool = False):
    """Skanuje dokumenty i aktualizuje plik JSON (bazę danych)."""
    db_data = {}

    # 1. Ładowanie istniejącej bazy, jeśli istnieje (żeby nie nadpisywać ręcznych zmian)
    if db_path.exists():
        with open(db_path, "r", encoding="utf-8") as f:
            try:
                db_data = json.load(f)
            except json.JSONDecodeError:
                print(f"[!] Error odczytu {db_path.name}. Tworzę nową bazę.")
                db_data = {}

    workers = get_workers(base_path)

    for w in workers:
        # Jeśli pracownika nie ma w bazie, tworzymy dla niego pusty wpis
        if w.name not in db_data:
            db_data[w.name] = {doc: "Brak" for doc in REQUIRED_DOCS}

        print(f"[*] Aktualizacja danych dla: {w.name}")

        for doc_type in REQUIRED_DOCS:
            if doc_type in w.documents:
                file_path = w.documents[doc_type]
                current_val = db_data[w.name].get(doc_type, "Brak")

                if doc_type == "foto":
                    db_data[w.name][doc_type] = "Tak"
                elif doc_type == "bhp":
                    # Uruchamiamy OCR tylko gdy wymuszono flagą --force,
                    # albo gdy w bazie nie ma jeszcze poprawnej daty
                    if force or current_val in [
                        "Brak",
                        "Error odczytu",
                        "Brak daty > 2023",
                        "Wymaga integracji OCR",
                    ]:
                        print(f"    - Skanowanie daty BHP...")
                        date_val = extract_bhp_date(file_path)
                        db_data[w.name][doc_type] = date_val
                    else:
                        print(f"    - BHP: zachowano istniejącą datę ({current_val})")
                elif doc_type == "g2e":
                    # NEW BLOCK FOR G2E OCR
                    if force or current_val in [
                        "Brak",
                        "Błąd odczytu",
                        "Brak daty ważności",
                        "Wymaga integracji OCR",
                    ]:
                        print(f"    - Skanowanie daty G2E...")
                        db_data[w.name][doc_type] = extract_g2e_date(file_path)
                    else:
                        print(f"    - G2E: zachowano istniejącą datę ({current_val})")
                else:
                    # Inne dokumenty, dla których na razie nie mamy OCR
                    # Aktualizujemy status, chyba że ktoś wpisał datę ręcznie
                    if current_val in ["Brak", "Wymaga integracji OCR"]:
                        db_data[w.name][doc_type] = "Obecny (Wymaga OCR)"
            else:
                db_data[w.name][doc_type] = "Brak"

    # Zapis do pliku JSON
    with open(db_path, "w", encoding="utf-8") as f:
        json.dump(db_data, f, ensure_ascii=False, indent=4)

    print(f"\n[✓] Gotowe! Zaktualizowano bazę danych: {db_path.name}")


def run_excel_from_db(db_path: Path):
    """Generuje raport Excel na podstawie danych zapisanych w pliku JSON."""
    if not db_path.exists():
        print(
            f"[!] Error: Baza danych '{db_path}' nie istnieje. Uruchom najpierw komendę 'updatedb'."
        )
        return

    with open(db_path, "r", encoding="utf-8") as f:
        db_data = json.load(f)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Status Dokumentów"

    # Nagłówki
    headers = ["Imię i Nazwisko"] + [d.upper() for d in REQUIRED_DOCS]
    ws.append(headers)

    # Zrzucanie danych z JSONa do Excela
    for worker_name, docs in db_data.items():
        row = [worker_name]
        for doc_type in REQUIRED_DOCS:
            row.append(docs.get(doc_type, "Brak"))
        ws.append(row)

    output_file = "raport_z_bazy.xlsx"
    wb.save(output_file)
    print(f"[✓] Generated Excel file based on database JSON: {output_file}")
