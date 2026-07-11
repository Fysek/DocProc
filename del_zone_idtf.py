import os


def remove_zone_identifier_files(root_folder):
    """
    Recursively remove Zone.Identifier files from the given folder and its subfolders.
    """
    removed = 0
    for dirpath, _, filenames in os.walk(root_folder):
        for filename in filenames:
            if "Zone.Identifier" in filename:
                file_path = os.path.join(dirpath, filename)
                try:
                    os.remove(file_path)
                    removed += 1
                    print(f"🗑️ Removed: {file_path}")
                except Exception as e:
                    print(f"❌ Failed to remove {file_path}: {e}")
    print(f"\n✅ Done. Removed {removed} Zone.Identifier files.")


if __name__ == "__main__":
    folder = os.path.expanduser("/home/mdyrdol/_all")  # change if needed
    if os.path.isdir(folder):
        remove_zone_identifier_files(folder)
    else:
        print(f"❌ Folder not found: {folder}")
