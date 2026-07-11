import json
import zipfile
import subprocess
import shutil
from datetime import datetime
from pathlib import Path
import openpyxl
from core import Worker, REQUIRED_DOCS


def get_workers(base_path: Path):
    """Returns a list of Worker objects from the given base directory."""
    return [Worker(p) for p in base_path.iterdir() if p.is_dir()]


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
    workers = get_workers(base_path)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Status Dokumentów"

    # Nagłówki
    headers = ["Imię i Nazwisko"] + [d.upper() for d in REQUIRED_DOCS]
    ws.append(headers)

    # Szkielet funkcji do czytania daty (na razie zwraca stałą wartość)
    def extract_date(doc_path):
        return "2025-12-31"  # ZAŚLEPKA NA PRZYSZŁOŚĆ

    for w in workers:
        row = [w.name]
        for doc_type in REQUIRED_DOCS:
            if doc_type in w.documents:
                if doc_type == "foto":
                    row.append("Tak")
                else:
                    row.append(extract_date(w.documents[doc_type]))
            else:
                row.append("Brak")
        ws.append(row)

    output_file = "raport_dokumentow.xlsx"
    wb.save(output_file)
    print(f"[✓] Wygenerowano raport Excel: {output_file}")


def run_compress(base_path: Path):
    # Sprawdzamy, czy Ghostscript jest zainstalowany w systemie
    if not shutil.which("gs"):
        print("[!] Błąd: Nie znaleziono programu Ghostscript.")
        print("    Zainstaluj go komendą: sudo apt install ghostscript")
        return

    workers = get_workers(base_path)
    for w in workers:
        for doc_type, file_path in w.documents.items():
            if file_path.suffix.lower() == ".pdf":
                size_mb = file_path.stat().st_size / (1024 * 1024)

                if size_mb > 1.0:
                    print(
                        f"[*] Kompresja: {file_path.name} ({size_mb:.2f} MB) w folderze: {w.name}"
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
                        # Uruchamiamy proces kompresji
                        subprocess.run(gs_cmd, check=True)

                        # Sprawdzamy nowy rozmiar
                        new_size_mb = temp_path.stat().st_size / (1024 * 1024)

                        if new_size_mb < size_mb:
                            temp_path.replace(file_path)
                            zaoszczedzono = size_mb - new_size_mb
                            print(
                                f"    [✓] Sukces! Zmniejszono do {new_size_mb:.2f} MB (zaoszczędzono {zaoszczedzono:.2f} MB)"
                            )
                        else:
                            temp_path.unlink()  # usuwamy plik tymczasowy
                            print(
                                "    [-] Plik jest już optymalnie skompresowany, pomijam."
                            )

                    except subprocess.CalledProcessError as e:
                        print(
                            f"    [!] Błąd podczas kompresji pliku {file_path.name}: {e}"
                        )
                        if temp_path.exists():
                            temp_path.unlink()
