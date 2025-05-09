import subprocess
import os


def sign_file(signtool_path, cert_name, timestamp_url, filepath):
    if not os.path.exists(signtool_path):
        print("❌ No se encontró signtool.exe en la ruta especificada.")
        return
    comando = [
        signtool_path,
        "sign",
        "/n", cert_name,
        "/fd", "SHA256",
        "/td", "SHA256",
        "/tr", timestamp_url,
        filepath
    ]
    resultado = subprocess.run(comando, capture_output=True, text=True)
    return resultado.returncode == 0, resultado.stdout + resultado.stderr
