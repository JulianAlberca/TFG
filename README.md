
# 🛡️ Herramienta de Firma de Macros VBA

Esta aplicación permite **firmar automáticamente macros VBA** contenidas en documentos de Office 365 (Excel, Word, PowerPoint), utilizando un certificado digital. Incorpora una interfaz gráfica sencilla, soporte para whitelist y posibilidad de programación periódica de firma.

---

## ⚙️ Requisitos previos

### ✅ Sistema
- Windows 10 o superior.
- Python 3.9+.
- Office 365 o Microsoft Office con soporte para macros VBA.

### 📦 Dependencias de Python

Instálalas ejecutando:

```bash
pip install -r requirements.txt
```

`requirements.txt` debería incluir:
```txt
pywin32
tk
pyautogui
pywinauto
```

---

## 🧱 Instalación y configuración de herramientas necesarias

### 1. Instalar el Windows SDK

Descargar e instalar desde:  
👉 https://developer.microsoft.com/en-us/windows/downloads/windows-sdk/

Durante la instalación, **asegúrate de marcar la opción**:
- ✅ "Windows SDK Signing Tools for Desktop Apps"

---

### 2. Registrar las DLLs necesarias para firmar macros VBA

Abre una terminal de administrador (cmd o PowerShell) y ejecuta:

```cmd
regsvr32 "C:\Program Files (x86)\Microsoft SDKs\ClickOnce\SignTool\msosip.dll"
regsvr32 "C:\Program Files (x86)\Microsoft SDKs\ClickOnce\SignTool\msosipx.dll"
```

> ⚠️ Ajusta las rutas si tus DLLs están en otra ubicación.

---

### 3. Verificar que `signtool.exe` está disponible

La herramienta lo buscará automáticamente en rutas estándar del SDK. Si quieres comprobarlo manualmente:

```cmd
where signtool
```

Si no aparece, puedes buscarlo manualmente en una ruta como:

```
C:\Program Files (x86)\Windows Kits\10\bin\<versión>\x64\signtool.exe
```

---

## 🚀 Uso básico

### 1. Iniciar la aplicación

```bash
python main.py
```

Se abrirá la interfaz gráfica (la contraseña es admin123).

---

### 2. Interfaz gráfica

- 📁 Navega por carpetas y selecciona documentos `.xlsm`, `.docm`, `.pptm`.
- 🔒 Elige el certificado digital para firmar.
- ✅ Marca los archivos que deseas firmar.
- 📝 Usa filtros por estado: Firmado, No firmado, Caducado.
- ⏱️ Accede a la pestaña "Programar firma" para automatizar el proceso.

---

## 📝 Firma programada

Puedes programar una firma automática mediante el panel de “Programar firma”.  
Esto utiliza `schtasks` de Windows. El script que se ejecutará debe estar preparado (como `firma_programada.py`).

> Asegúrate de que el certificado seleccionado está disponible en el sistema cuando se ejecute.

---

## 🔐 Certificados digitales

Puedes usar:
- Certificados autofirmados (`selfcert.exe`).
- Certificados emitidos por una CA.
- Almacén personal de Windows ("Mis certificados").

Para listar certificados desde consola:

```powershell
Get-ChildItem -Path Cert:\CurrentUser\My
```

---

## 🧪 Verificar firmas

El sistema permite comprobar automáticamente:
- Si el documento está firmado.
- Fecha de firma.
- Fecha de expiración del certificado.
- Nombre del firmante.

---

## 📋 Whitelist (Macros de confianza)

La aplicación utiliza un sistema de **hashes** para permitir solo la firma de macros previamente autorizadas.  
Puedes añadir macros a la whitelist desde la interfaz.

---

## 🛠 Estructura del proyecto

```
TFG/
├── app/
│   ├── main.py
│   ├── gui.py
│   ├── signer.py
│   ├── verify.py
│   ├── whitelist.py
│   └── ...
├── documentos/
│   └── (Archivos Office con macros)
├── logs/
│   └── firma_log.csv (log para las firmas manuales)
│   └── (Logs generados por el script programado)
├── passsword.hash
├── requirements.txt
└── README.md
```
