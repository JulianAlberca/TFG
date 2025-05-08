import os
import time
import threading
import pyautogui
from pywinauto import Application, Desktop
import win32com.client

def cerrar_advertencia_excel_en_segundo_plano(timeout=20):
    def worker():
        inicio = time.time()
        while time.time() - inicio < timeout:
            try:
                ventana = Desktop(backend="win32").window(title_re=".*Microsoft Excel.*")
                if ventana.exists() and 'Aceptar' in ventana.children_texts():
                    print("la ventana existe")
                    ventana.set_focus()
                    ventana['Aceptar'].click_input()
                    print("Advertencia cerrada automáticamente.")
                    break
                else:
                    print("la ventana NO existe")
            except:
                pass
            time.sleep(0.5)
    threading.Thread(target=worker, daemon=True).start()

def abrir_excel_y_documento(ruta_documento):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    wb = excel.Workbooks.Open(ruta_documento)
    return excel, wb

def activar_excel():
    app = Application(backend="uia").connect(title_re=".*Excel")
    excel_win = app.window(title_re=".*Excel")
    excel_win.set_focus()
    return excel_win

def abrir_editor_vba(excel_win):
    excel_win.type_keys('%{F11}', pause=0.5)
    time.sleep(2)
    vba = Desktop(backend="uia").window(title_re=".*Microsoft Visual Basic")
    vba.set_focus()
    return vba

def firmar_macro_vba_con_gui(vba_win):
    vba_win.type_keys('%h')  # Menú Herramientas
    time.sleep(0.3)
    vba_win.type_keys('d')   # Firma digital...
    time.sleep(2)

    # Interacción con ventana no accesible
    #pyautogui.press('tab')       # Mover al botón "Seleccionar..."
    pyautogui.press('enter')     # Pulsar "Seleccionar..."
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.press('enter')     # Elegir primer certificado
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.press('enter')     # Confirmar firma

    vba_win.type_keys('%{F4}')   # Cerrar el editor
    time.sleep(1)
from pywinauto import Desktop

def cerrar_ventana_advertencia_excel():
    try:
        ventana_alerta = Desktop(backend="win32").window(title_re=".*Microsoft Excel.*")
        if ventana_alerta.exists(timeout=2):
            print("la ventana existe")
            ventana_alerta.set_focus()
            if ventana_alerta['Aceptar'].exists():
                ventana_alerta['Aceptar'].click_input()
                time.sleep(0.5)
        else:
            print("la ventana NO existe")
    except Exception as e:
        print("No se pudo cerrar la advertencia:", e)


def firmar_archivo(ruta_documento):
    try:
        print(f"Procesando: {os.path.basename(ruta_documento)}")
        excel, wb = abrir_excel_y_documento(ruta_documento)
        time.sleep(2)
        excel_win = activar_excel()
        vba_win = abrir_editor_vba(excel_win)
        firmar_macro_vba_con_gui(vba_win)

        cerrar_advertencia_excel_en_segundo_plano(timeout=10)
        print("Guardando...")
        time.sleep(3)
        wb.Save()
        cerrar_ventana_advertencia_excel()
        print("Cerrando...")
        time.sleep(3)
        wb.Close()
        print("Quitando...")
        time.sleep(3)
        excel.Quit()
        return f"Firmado: {os.path.basename(ruta_documento)}"
    except Exception as e:
        return f"Error en {os.path.basename(ruta_documento)}: {e}"

def firmar_todos_en_carpeta(ruta_carpeta):
    resultados = []
    for archivo in os.listdir(ruta_carpeta):
        if archivo.lower().endswith('.xlsm'):
            ruta_completa = os.path.join(ruta_carpeta, archivo)
            resultado = firmar_archivo(ruta_completa)
            print(resultado)
            resultados.append(resultado)
    return resultados

# === USO PRINCIPAL ===
if __name__ == "__main__":
    carpeta = r"C:\Users\alber\Desktop\TFG\source"
    resultados = firmar_todos_en_carpeta(carpeta)
    print("\nResumen:")
    for r in resultados:
        print(r)