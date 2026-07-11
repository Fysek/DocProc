import json
import zipfile
from pathlib import Path
import openpyxl
from core import Worker, REQUIRED_DOCS

def get_workers(base_path: Path):
    """Zwraca listę obiektów Worker z podanego folderu głównego."""
    return [Worker(p) for p in base_path.iterdir() if p.is_dir()]

def run_check(base_path: Path):
    workers = get_workers(base_path)
    for w in workers:
        if w.is_complete:
            print(f"[✓] {w.name} - Komplet")
        else:
            missing = ", ".join(w.missing_docs)
            print(f"[x] {w.name} - Braki: {missing}")
        
        # Standaryzacja nazw (jeśli plik nie nazywa się poprawnie, zmieniamy to)
        safe_name = w.name.replace(" ", "_")
        for doc_type, file_path in w.documents.items():
            ext = file_path.suffix # rozszerzenie np. .pdf
            expected_name = f"{safe_name}_{doc_type}{ext}"
            
            if file_path.name != expected_name:
                new_path = file_path.with_name(expected_name)
                file_path.rename(new_path)
                w.documents[doc_type] = new_path # aktualizacja ścieżki po zmianie nazwy

def run_pack(base_path: Path, json_path: Path):
    with open(json_path, 'r', encoding='utf-8') as f:
        target_names = json.load(f)
    
    workers_dict = {w.name: w for w in get_workers(base_path)}
    
    for name in target_names:
        if name not in workers_dict:
            print(f"[!] Nie znaleziono folderu dla: {name}")
            continue
        
        w = workers_dict[name]
        if not w.is_complete:
            print(f"[x] {w.name} - Nie można zrobić paczki, braki: {', '.join(w.missing_docs)}")
            continue
        
        # Tworzenie pliku .zip
        zip_name = f"{name.replace(' ', '_')}_paczka.zip"
        with zipfile.ZipFile(zip_name, 'w') as zf:
            for doc_path in w.documents.values():
                zf.write(doc_path, doc_path.name)
        print(f"[✓] Utworzono paczkę: {zip_name}")

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
        return "2025-12-31" # ZAŚLEPKA NA PRZYSZŁOŚĆ
    
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
    workers = get_workers(base_path)
    for w in workers:
        for file_path in w.documents.values():
            if file_path.suffix.lower() == '.pdf':
                size_mb = file_path.stat().st_size / (1024 * 1024)
                if size_mb > 1.0:
                    print(f"[*] Do kompresji: {file_path.name} ({size_mb:.2f} MB) w folderze {w.name}")
                    # TUTAJ W PRZYSZŁOŚCI DODAMY LOGIKĘ KOMPRESJI