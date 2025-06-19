import subprocess
import sys
from tkinter import messagebox

frecuencias = {
    "DIARIAMENTE": "DAILY",
    "SEMANALMENTE": "WEEKLY"
}


def formatear_hora(hora_str):
    try:
        partes = hora_str.strip().split(":")
        h = int(partes[0])
        m = int(partes[1])
        return f"{h:02}:{m:02}"
    except:
        return "08:00"  # fallback seguro


def crear_tarea_programada(hora, script_path, frecuencia, cert):

    frecuencia_real = frecuencias.get(frecuencia, "DAILY")
    hora_formateada = formatear_hora(hora)

    comando = (
        f'schtasks /Create /F /SC {frecuencia_real} '
        f'/TN "FirmaAutomaticaMacros" '
        f'/TR "\\"{sys.executable}\\" \\"{script_path}\\" \\"{cert}\\"" '
        f'/ST {hora_formateada}'
    )

    try:
        resultado = subprocess.run(
            comando, shell=True, capture_output=True, text=True)
        if resultado.returncode == 0:
            messagebox.showinfo("Tarea programada",
                                "La firma fue programada correctamente.")
        else:
            messagebox.showerror("Error al programar", resultado.stderr)
    except Exception as e:
        messagebox.showerror("Error", str(e))
