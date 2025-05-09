import os
import subprocess
import shutil

# === CONFIGURACIÓN ===
signtool_path = r"C:\Program Files (x86)\Windows Kits\10\bin\10.0.22621.0\x86\signtool.exe"
nombre_certificado = "Cert1"
timestamp_url = "http://timestamp.digicert.com"
directorio = r"C:\Users\alber\Desktop\TFG\source"

def firmar_archivo(ruta_archivo):
    if not os.path.exists(signtool_path):
        print("❌ No se encontró signtool.exe en la ruta especificada.")
        return

    comando = [
        signtool_path,
        "sign",
        "/n", nombre_certificado,
        "/fd", "SHA256",
        "/td", "SHA256",
        "/tr", timestamp_url,
        ruta_archivo
    ]

    print(f"🔧 Firmando: {os.path.basename(ruta_archivo)}")
    resultado = subprocess.run(comando, capture_output=True, text=True)

    if resultado.returncode == 0:
        print(f"✅ Firmado correctamente: {ruta_archivo}")
    else:
        print(f"❌ Error al firmar: {ruta_archivo}")
        print(resultado.stdout)
        print(resultado.stderr)

def firmar_directorio(carpeta):
    for archivo in os.listdir(carpeta):
        if archivo.lower().endswith(".xlsm") and not archivo.startswith("~$"):
            ruta_completa = os.path.join(carpeta, archivo)
            firmar_archivo(ruta_completa)

if __name__ == "__main__":
    if not os.path.exists(directorio):
        print(f"❌ Carpeta no encontrada: {directorio}")
    else:
        firmar_directorio(directorio)
