from pathlib import Path

# Definiujemy naszą Wielką Piątkę
REQUIRED_DOCS = ["bhp", "lekarskie", "dowod", "foto", "g2e"]


class Worker:
    def __init__(self, path: Path):
        self.path = path
        self.name = path.name  # np. "Jan Kowalski" z nazwy folderu
        self.documents = {}  # Słownik: typ_dokumentu -> ścieżka_do_pliku
        self.scan_documents()

    def scan_documents(self):
        """Przeszukuje folder pracownika i dopasowuje pliki do Wielkiej Piątki."""
        if not self.path.is_dir():
            return

        for file_path in self.path.iterdir():
            if file_path.is_file():
                fname = (
                    file_path.stem.lower()
                )  # nazwa pliku bez rozszerzenia małą literą

                # Prosta heurystyka: szukamy słowa kluczowego w nazwie pliku
                for doc_type in REQUIRED_DOCS:
                    if doc_type in fname:
                        self.documents[doc_type] = file_path
                        break

    @property
    def is_complete(self):
        """Zwraca True, jeśli pracownik ma wszystkie 5 dokumentów."""
        return len(self.documents) == len(REQUIRED_DOCS)

    @property
    def missing_docs(self):
        """Zwraca listę brakujących dokumentów."""
        return [doc for doc in REQUIRED_DOCS if doc not in self.documents]
