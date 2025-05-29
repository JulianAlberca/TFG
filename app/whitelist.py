import hashlib
import zipfile
import olefile
import os
import tempfile
from config import whitelist


def calcular_hash_vba_code(ruta_archivo):
    try:
        # Extraer vbaProject.bin del .xlsm
        with zipfile.ZipFile(ruta_archivo, 'r') as zip_ref:
            candidatos = [name for name in zip_ref.namelist(
            ) if name.endswith("vbaProject.bin")]

            if not candidatos:
                return None
            vba_data = zip_ref.read(candidatos[0])

        # Guardar vbaProject.bin temporalmente
        with tempfile.NamedTemporaryFile(delete=False, suffix=".bin") as tmp:
            tmp.write(vba_data)
            temp_path = tmp.name

        # Leer con olefile
        ole = olefile.OleFileIO(temp_path)
        code_streams = [
            s for s in ole.listdir()
            if s[0] == "VBA" and not s[1].startswith("__") and s[1] not in ("_VBA_PROJECT", "dir")
        ]

        contenido = b''
        for stream in sorted(code_streams):  # Orden alfabético para consistencia
            contenido += ole.openstream(stream).read()

        ole.close()
        os.remove(temp_path)

        return hashlib.sha256(contenido).hexdigest()

    except Exception as e:
        print(f"⚠️ Error al calcular hash del código VBA: {e}")
        return None


def cargar_hashes_autorizados(ruta=whitelist):
    if not os.path.exists(ruta):
        return set()
    with open(ruta, "r", encoding="utf-8") as f:
        return set(line.strip() for line in f if line.strip())
