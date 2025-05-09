import os

EXTENSIONES_OFFICE = (".xlsm", ".docm", ".pptm")


def list_office_files_recursively(directory):
    archivos = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith(EXTENSIONES_OFFICE) and not file.startswith("~$"):
                full_path = os.path.join(root, file)
                relative_path = os.path.relpath(full_path, directory)
                archivos.append((full_path, relative_path))
    return archivos
