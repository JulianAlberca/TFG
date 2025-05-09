import subprocess


def get_cert_names():
    comando = ['powershell', '-Command',
               'Get-ChildItem Cert:\\CurrentUser\\My | Select-Object -ExpandProperty Subject']
    resultado = subprocess.run(comando, capture_output=True, text=True)
    nombres = []
    for linea in resultado.stdout.splitlines():
        if "CN=" in linea:
            nombres.append(linea.split("CN=")[1].split(",")[0])
    return nombres
