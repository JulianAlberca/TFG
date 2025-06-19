import tkinter as tk
from tkinter import ttk, messagebox
import hashlib


def hash_contrasena(texto):
    return hashlib.sha256(texto.encode()).hexdigest()


def cargar_hash():
    try:
        with open("password.hash", "r") as f:
            return f.read().strip()
    except FileNotFoundError:
        return None


def lanzar_login(callback_aceptado):
    hash_correcto = cargar_hash()
    if not hash_correcto:
        messagebox.showerror("Error", "No se ha encontrado 'password.hash'.")
        return

    root = tk.Tk()

    root.withdraw()  # Oculta la ventana principal

    ventana = tk.Toplevel()
    ventana.title("Login")
    ventana.geometry("300x150")
    ventana.resizable(False, False)
    # Cerrar todo si cierra el login
    ventana.protocol("WM_DELETE_WINDOW", root.quit)
    ventana.grab_set()  # Hace que sea modal (bloquea el resto)

    ttk.Label(ventana, text="Introduce la contraseña:").pack(pady=(20, 5))
    entrada = ttk.Entry(ventana, show="*")
    entrada.pack(pady=5)
    entrada.focus()

    ventana.bind("<Return>", lambda event: verificar())

    def verificar():
        entrada_hash = hash_contrasena(entrada.get())
        if entrada_hash == hash_correcto:
            ventana.destroy()
            root.destroy()
            callback_aceptado()
        else:
            messagebox.showerror("Acceso denegado", "Contraseña incorrecta.")
            entrada.delete(0, tk.END)

    ttk.Button(ventana, text="Entrar", command=verificar).pack(pady=10)

    ventana.mainloop()
