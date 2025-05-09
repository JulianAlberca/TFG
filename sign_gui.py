import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client

# === FUNCIONES DE UTILIDAD ===
def listar_certificados_personales():
    """Obtiene los certificados del almacén personal del usuario"""
    certificados = []
    store = win32com.client.Dispatch("CAPICOM.Store")
    store.Open(2, "My", 0)  # CAPICOM_CURRENT_USER_STORE, "My", CAPICOM_STORE_OPEN_READ_ONLY
    for cert in store.Certificates:
        certificados.append(cert.SubjectName)
    return certificados

def listar_archivos_xlsm(directorio):
    return [f for f in os.listdir(directorio) if f.lower().endswith(".xlsm") and not f.startswith("~$")]

def firmar_archivo(ruta_archivo, certificado_nombre, signtool_path, timestamp_url):
    comando = [
        signtool_path,
        "sign",
        "/n", certificado_nombre,
        "/fd", "SHA256",
        "/td", "SHA256",
        "/tr", timestamp_url,
        ruta_archivo
    ]
    resultado = subprocess.run(comando, capture_output=True, text=True)
    return resultado.returncode == 0, resultado.stdout + resultado.stderr

# === INTERFAZ GRÁFICA ===
class AplicacionFirma:
    def __init__(self, root):
        self.root = root
        self.root.title("Firmador de Documentos VBA")

        # Configuración inicial
        self.directorio = tk.StringVar()
        self.certificados = listar_certificados_personales()
        self.cert_seleccionado = tk.StringVar()
        self.cert_seleccionado.set(self.certificados[0] if self.certificados else "")
        self.timestamp_url = tk.StringVar(value="http://timestamp.digicert.com")
        self.signtool_path = tk.StringVar(value=r"C:\\Program Files (x86)\\Windows Kits\\10\\bin\\10.0.22621.0\\x86\\signtool.exe")

        self.crear_interfaz()

    def crear_interfaz(self):
        frm = ttk.Frame(self.root, padding=10)
        frm.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frm, text="Directorio de documentos:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.directorio, width=50).grid(row=1, column=0, sticky="ew")
        ttk.Button(frm, text="Seleccionar...", command=self.seleccionar_directorio).grid(row=1, column=1, sticky="w")

        ttk.Label(frm, text="Certificado:").grid(row=2, column=0, sticky="w", pady=(10,0))
        self.combo_cert = ttk.Combobox(frm, textvariable=self.cert_seleccionado, values=self.certificados, state="readonly")
        self.combo_cert.grid(row=3, column=0, columnspan=2, sticky="ew")

        ttk.Label(frm, text="Archivos a firmar:").grid(row=4, column=0, sticky="w", pady=(10,0))
        self.lista_archivos = tk.Listbox(frm, height=10, width=50)
        self.lista_archivos.grid(row=5, column=0, columnspan=2, sticky="nsew")

        ttk.Button(frm, text="Firmar documentos", command=self.firmar_documentos).grid(row=6, column=0, columnspan=2, pady=(10,0))

        self.mensaje = tk.StringVar()
        ttk.Label(frm, textvariable=self.mensaje, foreground="green").grid(row=7, column=0, columnspan=2, sticky="w")

        frm.columnconfigure(0, weight=1)

    def seleccionar_directorio(self):
        carpeta = filedialog.askdirectory()
        if carpeta:
            self.directorio.set(carpeta)
            archivos = listar_archivos_xlsm(carpeta)
            self.lista_archivos.delete(0, tk.END)
            for archivo in archivos:
                self.lista_archivos.insert(tk.END, archivo)

    def firmar_documentos(self):
        carpeta = self.directorio.get()
        cert = self.cert_seleccionado.get()
        errores = []

        for i in range(self.lista_archivos.size()):
            archivo = self.lista_archivos.get(i)
            ruta = os.path.join(carpeta, archivo)
            ok, salida = firmar_archivo(ruta, cert, self.signtool_path.get(), self.timestamp_url.get())
            if not ok:
                errores.append(f"{archivo}:\n{salida}")

        if errores:
            messagebox.showerror("Errores al firmar", "\n\n".join(errores))
        else:
            messagebox.showinfo("Firmado", "Todos los documentos fueron firmados correctamente.")

# === EJECUCIÓN ===
if __name__ == "__main__":
    root = tk.Tk()
    app = AplicacionFirma(root)
    root.mainloop()
