import os
import subprocess

# === CONFIGURACIÓN ===
offsign_path = r"C:\Program Files (x86)\Windows Kits\10\bin\10.0.22621.0\x86\offsign.bat"  # Ruta al offsign.bat
signtool_dir = r"C:\Program Files (x86)\Windows Kits\10\bin\10.0.22621.0\x86\\"  # Directorio (no el exe)
cert_path = r"C:\Certificados\certificado.pfx"
cert_pass = "1234"  # Contraseña del .pfx
timestamp_url = "http://timestamp.digicert.com"
directorio = r"C:\Users\alber\Desktop\TFG\source"

# === FUNCIONES ===
def firmar_con_offsign(archivo_xlsm):
    cert_subject = "Cert1"
    sign_cmd = f'sign /n "{cert_subject}" /fd SHA256 /tr {timestamp_url} /td SHA256'
    verify_cmd = "verify /pa"

    comando = [
        offsign_path,
        signtool_dir,
        sign_cmd,
        verify_cmd,
        archivo_xlsm          # ✅ Correcto
    ]


    print(f"🔐 Firmando con certificado instalado: {os.path.basename(archivo_xlsm)}")
    print(comando)
    resultado = subprocess.run(comando, capture_output=True, text=True)
    if resultado.returncode == 0:
        print(f"✅ Firmado correctamente: {archivo_xlsm}")
    else:
        print(f"❌ Error al firmar: {archivo_xlsm}")
        print(resultado.stdout)
        print(resultado.stderr)


def firmar_directorio_con_offsign(carpeta):
    for archivo in os.listdir(carpeta):
        if archivo.lower().endswith(".xlsm"):
            ruta = os.path.join(carpeta, archivo)
            firmar_con_offsign(ruta)

# === EJECUCIÓN PRINCIPAL ===
if __name__ == "__main__":
    firmar_directorio_con_offsign(directorio)
