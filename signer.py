# signer.py
import os
import win32com.client
import pythoncom
import traceback
from win32com.client import gencache


def firmar_proyecto_vba(ruta_documento, nombre_certificado):
    """
    Firma el proyecto VBA de un documento de Office usando un certificado instalado.
    :param ruta_documento: Ruta completa al archivo .xlsm, .docm, .pptm.
    :param nombre_certificado: Nombre visible del certificado digital.
    """
    ext = os.path.splitext(ruta_documento)[1].lower()
    pythoncom.CoInitialize()
    app = None
    doc = None
    try:
        if ext == '.xlsm':
            app = gencache.EnsureDispatch("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            doc = app.Workbooks.Open(ruta_documento)
            vba_project = doc.VBProject

        elif ext == '.docm':
            app = win32com.client.Dispatch("Word.Application")
            app.Visible = False
            app.DisplayAlerts = False
            doc = app.Documents.Open(ruta_documento)
            vba_project = doc.VBProject

        elif ext == '.pptm':
            app = win32com.client.Dispatch("PowerPoint.Application")
            app.Visible = False
            doc = app.Presentations.Open(ruta_documento, WithWindow=False)
            vba_project = doc.VBProject

        else:
            raise ValueError(f"Extensión no compatible: {ext}")

        # Buscar el certificado por nombre
        certificados = app.DigitalSignatureCertificates
        certificado = None

        for i in range(1, certificados.Count + 1):
            if certificados.Item(i).Subject == nombre_certificado or nombre_certificado in certificados.Item(i).Subject:
                certificado = certificados.Item(i)
                break

        if not certificado:
            raise Exception(f"No se encontró el certificado con nombre: {nombre_certificado}")

        vba_project.DigitalSignature.Sign(certificado)
        doc.Save()
        print(f"Documento firmado correctamente: {ruta_documento}")

    except Exception as e:
        print(f"Error al firmar {ruta_documento}:\n{traceback.format_exc()}")

    finally:
        try:
            doc.Close(SaveChanges=True)
        except:
            pass
        app.Quit()
        pythoncom.CoUninitialize()
