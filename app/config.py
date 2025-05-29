import os
import sys


def buscar_signtool():
    posibles_rutas = [
        r"C:\Program Files (x86)\Windows Kits\10\bin",
        r"C:\Program Files\Windows Kits\10\bin"
    ]

    versiones = [
        "10.0.22621.0",
        "10.0.22000.0",
        "10.0.19041.0",
        "10.0.18362.0"
    ]

    arquitecturas = ["x86", "x64"]

    for base in posibles_rutas:
        for version in versiones:
            for arch in arquitecturas:
                ruta = os.path.join(base, version, arch, "signtool.exe")
                if os.path.isfile(ruta):
                    return ruta

    return None


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Carpeta local del proyecto
default_directory = os.path.abspath(os.path.join(
    os.path.dirname(__file__), "..", "documentos"))

# signtool_path = r"C:\Program Files (x86)\Windows Kits\10\bin\10.0.22621.0\x86\signtool.exe"

timestamp_url = "http://timestamp.digicert.com"
whitelist = os.path.join(os.path.dirname(__file__), "whitelist_hashes.txt")
script_path = os.path.abspath(os.path.join(
    os.path.dirname(__file__), "firma_programada.py"))
logs_directory = os.path.abspath(os.path.join(
    os.path.dirname(__file__), "..", "logs"))
os.makedirs(logs_directory, exist_ok=True)
python_exe = sys.executable


# Ruta absoluta del ejecutable
signtool_path = buscar_signtool()
if not signtool_path:
    raise FileNotFoundError(
        "❌ No se encontró signtool.exe en el sistema.\n"
        "Asegúrate de tener instalado el Windows SDK y que contenga signtool.exe.\n"
        "Puedes instalarlo desde: https://developer.microsoft.com/en-us/windows/downloads/windows-sdk/"
    )
