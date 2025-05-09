import os
import subprocess

# === CONFIGURACIÓN ===
offclearsig_path = r"C:\Program Files (x86)\Windows Kits\10\bin\10.0.22621.0\x86\offclearsig.exe"  # Ruta al ejecutable
directorio = r"C:\Users\alber\Desktop\TFG\source"  # Carpeta con archivos .xlsm

# === FUNCIONES ===
def desfirmar_archivo(archivo):
    archivo_comillas = f'"{archivo}"'
    comando = [offclearsig_path, archivo]
    print(f"🧹 Desfirmando: {os.path.basename(archivo)}")

    resultado = subprocess.run(" ".join(comando), capture_output=True, text=True)

    if resultado.returncode == 0:
        print(f"✅ Firma eliminada de: {archivo}")
    else:
        print(f"❌ Error al eliminar firma de: {archivo}")
        print(resultado.stdout)
        print(resultado.stderr)


def desfirmar_directorio(carpeta):
    for archivo in os.listdir(carpeta):
        if archivo.lower().endswith(".xlsm"):
            ruta = os.path.join(carpeta, archivo)
            desfirmar_archivo(ruta)

# === EJECUCIÓN PRINCIPAL ===
if __name__ == "__main__":
    desfirmar_directorio(directorio)
