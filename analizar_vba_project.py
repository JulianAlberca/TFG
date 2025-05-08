import zipfile
import olefile
import os
''' 
def extraer_y_analizar_vbaproject(ruta_xlsm, carpeta_salida="vbaproject_temp"):
    if not os.path.exists(carpeta_salida):
        os.makedirs(carpeta_salida)

    vbapath = os.path.join(carpeta_salida, "vbaProject.bin")

    # Extraer el archivo vbaProject.bin del .xlsm
    with zipfile.ZipFile(ruta_xlsm, 'r') as zip_ref:
        try:
            zip_ref.extract("xl/vbaProject.bin", path=carpeta_salida)
            print("✅ vbaProject.bin extraído correctamente.")
        except KeyError:
            print("❌ El archivo no contiene macros (vbaProject.bin no encontrado).")
            return

    full_path = os.path.join(carpeta_salida, "xl", "vbaProject.bin")

    # Verificar si es un archivo OLE válido
    if not olefile.isOleFile(full_path):
        print("❌ El archivo extraído no es un contenedor OLE válido.")
        return

    # Leer el contenedor OLE y mostrar los streams internos
    ole = olefile.OleFileIO(full_path)
    print("\n📁 Streams encontrados en vbaProject.bin:")
    for stream in ole.listdir():
        print(" -", "/".join(stream))
    
    ole.close()

    return full_path
'''
import zipfile
import os

def listar_contenido_zip(ruta_xlsm):
    if not os.path.exists(ruta_xlsm):
        print(f"❌ El archivo no existe: {ruta_xlsm}")
        return

    try:
        with zipfile.ZipFile(ruta_xlsm, 'r') as zip_ref:
            print(f"\n📄 Contenido de {os.path.basename(ruta_xlsm)}:")
            for name in zip_ref.namelist():
                print(" -", name)
    except Exception as e:
        print("❌ Error al abrir el archivo:", e)

# === USO ===
if __name__ == "__main__":
    ruta_documento = r"C:\Users\alber\Desktop\TFG\source\prueba.xlsm"  # ← Cambia esta ruta si es necesario
    listar_contenido_zip(ruta_documento)

    #extraer_y_analizar_vbaproject(ruta_xlsm)
