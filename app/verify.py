import subprocess
import re
from datetime import datetime


def verificar_firma(signtool_path, archivo):

    comando = [signtool_path, "verify", "/pa", "/v", archivo]
    resultado = subprocess.run(comando, capture_output=True, text=True)
    salida = resultado.stdout + resultado.stderr

    firmado = "Signing Certificate Chain:" in salida
    expiracion = None
    expirada = None
    timestamp = None
    emitido_a = None

    if firmado:
        # CN del certificado
        match_cn = re.search(r"Issued to:\s*(.+)", salida)
        if match_cn:
            emitido_a = match_cn.group(1).strip()

        # Fecha de expiración
        match_exp = re.search(r"Expires:\s+([^\r\n]+)", salida)
        if match_exp:
            raw = match_exp.group(1).strip()
            expiracion = raw
            for fmt in ("%a %b %d %H:%M:%S %Y", "%A %B %d %H:%M:%S %Y"):
                try:
                    fecha = datetime.strptime(raw, fmt)
                    expirada = fecha < datetime.now()
                    break
                except ValueError:
                    continue

        # Fecha de firma
        match_ts = re.search(
            r"The signature is timestamped:\s+([^\r\n]+)", salida)
        if match_ts:
            timestamp = match_ts.group(1).strip()

    return firmado, expiracion, expirada, timestamp, emitido_a
