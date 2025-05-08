import os
import zipfile

def proyecto_vba_esta_firmado(ruta_xlsm):
    try:
        with zipfile.ZipFile(ruta_xlsm, 'r') as zip_ref:
            entradas = zip_ref.namelist()
            firmas = [
                "xl/vbaProjectSignature.bin",
                "xl/vbaProjectSignatureAgile.bin"
            ]
            return any(f in entradas for f in firmas)
    except Exception as e:
        print(f"⚠️ Error en {os.path.basename(ruta_xlsm)}: {e}")
        return False

def verificar_firmas_en_directorio(carpeta):
    firmados = []
    no_firmados = []

    for archivo in os.listdir(carpeta):
        if archivo.lower().endswith('.xlsm'):
            ruta = os.path.join(carpeta, archivo)
            if proyecto_vba_esta_firmado(ruta):
                firmados.append(archivo)
            else:
                no_firmados.append(archivo)

    print("\n✅ Firmados:")
    for f in firmados:
        print("  -", f)

    print("\n❌ No firmados:")
    for nf in no_firmados:
        print("  -", nf)

    return firmados, no_firmados

# === USO ===
if __name__ == "__main__":
    carpeta = r"C:\Users\alber\Desktop\TFG\source"  # ← Cambia esto si es necesario
    verificar_firmas_en_directorio(carpeta)
