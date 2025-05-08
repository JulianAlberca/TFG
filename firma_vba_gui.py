# firma_vba_gui.py
import pyautogui
import time
import win32com.client
import pygetwindow as gw

def abrir_excel_con_documento(ruta_documento):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    workbook = excel.Workbooks.Open(ruta_documento)
    return excel, workbook

def activar_ventana_excel():
    time.sleep(2)
    ventanas = gw.getWindowsWithTitle("Excel")
    for ventana in ventanas:
        if ventana.isActive == False and ventana.isMinimized == False:
            ventana.activate()
            time.sleep(1)
            break

def firmar_vba_con_gui(ruta_documento):
    print("Abriendo documento en Excel...")
    excel, workbook = abrir_excel_con_documento(ruta_documento)

    time.sleep(3)
    activar_ventana_excel()

    print("Abriendo el editor de VBA...")
    pyautogui.hotkey('alt', 'f11')
    time.sleep(2)

    print("Activando ventana del editor VBA...")
    # Activar el editor de VBA (asume que la ventana contiene 'Microsoft Visual Basic')
    vba_ventanas = gw.getWindowsWithTitle("Microsoft Visual Basic")
    if vba_ventanas:
        vba_ventanas[0].activate()
        time.sleep(1)

    print("Abriendo el menú 'Herramientas > Firma digital'...")
    pyautogui.hotkey('alt', 'h')  # Menú Herramientas
    time.sleep(1)
    pyautogui.press('d')          # Firma digital
    time.sleep(2)

    print("Seleccionando certificado...")
    pyautogui.press('s')          # Botón 'Seleccionar'
    time.sleep(1)
    pyautogui.press('enter')      # Elegir certificado (asume primero)
    time.sleep(1)
    pyautogui.press('tab')      # Elegir certificado (asume primero)
    time.sleep(1)
    pyautogui.press('enter')      # Confirmar firma
    time.sleep(1)
    pyautogui.press('tab')      # Elegir certificado (asume primero)
    time.sleep(1)
    pyautogui.press('enter') 

    print("Cerrando editor VBA y guardando Excel...")
    pyautogui.hotkey('alt', 'f4')  # Cerrar editor
    time.sleep(1)

    activar_ventana_excel()
    workbook.Save()
    workbook.Close()
    print("Cerrando worbook...")
    time.sleep(2)
    print("Cerrando excel...")
    excel.Quit()

    print("Documento firmado automáticamente.")

# Uso
if __name__ == "__main__":
    documento = r"C:\Users\alber\Desktop\TFG\source\prueba.xlsm"
    firmar_vba_con_gui(documento)
