import sys
import os
import csv
from datetime import datetime

from files import list_office_files_recursively
from signer import sign_file
from verify import verificar_firma
from config import default_directory, signtool_path, timestamp_url, logs_directory


def main():
    if len(sys.argv) < 2:
        print("❌ Uso: firma_programada.py \"Nombre del certificado\"")
        return

    cert_name = sys.argv[1]
    archivos = list_office_files_recursively(default_directory)

    fecha = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_file = os.path.join(logs_directory, f"log_firma_{fecha}.csv")

    with open(log_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Archivo", "Resultado", "Motivo"])

        for full_path, relative_path in archivos:
            firmado, expiracion, expirada, _, _ = verificar_firma(
                signtool_path, full_path)

            if not firmado or expirada:
                ok, salida = sign_file(
                    signtool_path, cert_name, timestamp_url, full_path)
                if ok:
                    print(f"✅ Firmado: {relative_path}")
                    writer.writerow([relative_path, "Firmado", "-"])
                else:
                    print(f"❌ Error: {relative_path}\n{salida}")
                    writer.writerow([relative_path, "Error", salida.strip()])
            else:
                print(f"🔍 Ya firmado: {relative_path}")
                writer.writerow([relative_path, "Ya firmado", "-"])

    print(f"\n📄 Log guardado en: {log_file}")


if __name__ == "__main__":
    main()
