from pathlib import Path

# Define the "Big Five" required documents
REQUIRED_DOCS = ["bhp", "lekarskie", "dowod", "foto", "g2e"]


class Worker:
    def __init__(self, path: Path):
        self.path = path
        self.name = path.name  # e.g., "Jan Kowalski" from the folder name
        self.documents = (
            {}
        )  # Dictionary: document_type -> path_to_file (the newest one)
        self.obsolete_docs = []  # List of older files to be archived
        self.scan_documents()

    def scan_documents(self):
        """Scans the worker's folder, resolves duplicates by keeping the newest file."""
        if not self.path.is_dir():
            return

        # Temporary dictionary to group files by document type
        found_docs = {doc_type: [] for doc_type in REQUIRED_DOCS}

        for file_path in self.path.iterdir():
            if file_path.is_file():
                fname = file_path.stem.lower()  # Filename without extension, lowercase

                for doc_type in REQUIRED_DOCS:
                    if doc_type in fname:
                        found_docs[doc_type].append(file_path)
                        break  # File assigned to one doc type, move to next file

        # Process found documents: keep the newest, archive the rest
        for doc_type, paths in found_docs.items():
            if not paths:
                continue

            if len(paths) == 1:
                # Only one file found, it's the valid one
                self.documents[doc_type] = paths[0]
            else:
                # Multiple files found: sort by modification time (newest first)
                paths.sort(key=lambda p: p.stat().st_mtime, reverse=True)

                self.documents[doc_type] = paths[0]  # Keep the newest
                self.obsolete_docs.extend(paths[1:])  # Mark older files for archiving

    @property
    def is_complete(self):
        """Returns True if the worker has all required documents."""
        return len(self.documents) == len(REQUIRED_DOCS)

    @property
    def missing_docs(self):
        """Returns a list of missing document types."""
        return [doc for doc in REQUIRED_DOCS if doc not in self.documents]
