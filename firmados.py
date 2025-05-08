import win32com.client
import winreg
import os

def comprobar_acceso_vba_registro():
    try:
        key_path = r"Software\Microsoft\Office\16.0\Excel\Security"
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path)
        valor, _ = winreg.QueryValueEx(key, "AccessVBOM")
        winreg.CloseKey(key)
        return valor == 1
    except FileNotFoundError:
        return False
    except Exception as e:
        print(f"⚠️ No se pudo leer la configuración del registro: {e}")
        return False

def macro_esta_firmada(ruta_documento):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(ruta_documento, ReadOnly=True)
        vbproject = wb.VBProject

        if vbproject.DigitalSignature and vbproject.DigitalSignature.Signed:
            resultado = True
        else:
            resultado = False

        wb.Close(SaveChanges=False)
        return resultado

    except Exception as e:
        if not comprobar_acceso_vba_registro():
            print(f"❌ No se puede acceder a las macros de {os.path.basename(ruta_documento)}.")
            print("   → Asegúrate de activar:")
            print("     - 'Confiar en el acceso al modelo de objetos del proyecto VBA'")
            print("     - 'Habilitar todas las macros' en el Centro de confianza de Excel")
        else:
            print(f"⚠️ Error inesperado con {os.path.basename(ruta_documento)}: {e}")
        return False
    finally:
        excel.Quit()


def listar_documentos_firmados(carpeta):
    firmados = []
    no_firmados = []

    for archivo in os.listdir(carpeta):
        if archivo.lower().endswith('.xlsm'):
            ruta = os.path.join(carpeta, archivo)
            if macro_esta_firmada(ruta):
                firmados.append(archivo)
            else:
                no_firmados.append(archivo)

    return firmados, no_firmados


carpeta = r"C:\Users\alber\Desktop\TFG\source"
firmados, no_firmados = listar_documentos_firmados(carpeta)

print("✅ Firmados:")
for f in firmados:
    print("  -", f)

print("\n❌ No firmados:")
for nf in no_firmados:
    print("  -", nf)
