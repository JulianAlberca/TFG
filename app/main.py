from login import lanzar_login
from gui import FirmaApp
import tkinter as tk


def lanzar_aplicacion():
    root = tk.Tk()
    app = FirmaApp(root)
    root.mainloop()


if __name__ == "__main__":
    lanzar_login(lanzar_aplicacion)
