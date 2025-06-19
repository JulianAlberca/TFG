from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
import os
import csv

from files import list_office_files_recursively
from certs import get_cert_names
from signer import sign_file
import config
from verify import verificar_firma
from whitelist import calcular_hash_vba_code, cargar_hashes_autorizados
from programar import crear_tarea_programada


class FirmaApp:
    def __init__(self, master):
        self.master = master
        master.title("Firmador de macros VBA")

        self.folder = config.default_directory
        self.cert_name = tk.StringVar()
        self.filtro_estado = tk.StringVar(value="Todos")

        self.cert_list = get_cert_names()

        self.paned = ttk.PanedWindow(master, orient="horizontal")

        self.paned.pack(fill="both", expand=True)

        self.frame_izq = ttk.Frame(self.paned, width=200, padding=5)
        self.folder_tree = ttk.Treeview(
            self.frame_izq, show="tree", selectmode="browse")

        self.folder_tree.pack(fill="both", expand=True)

        self.folder_root = config.default_directory
        self.populate_folder_tree(self.folder_root)

        self.folder_tree.bind("<<TreeviewSelect>>", self.on_folder_select)

        self.notebook = ttk.Notebook(self.paned)
        self.frame_der = ttk.Frame(self.notebook, padding=5)
        self.frame_prog = ttk.Frame(self.notebook, padding=5)

        self.notebook.add(self.frame_der, text="Gestor de macros")
        self.notebook.add(self.frame_prog, text="Programar firma")
        self.paned.add(self.frame_izq, weight=1)
        self.paned.add(self.notebook, weight=4)

        ttk.Label(self.frame_der, text="Certificado:").pack()
        self.combo_cert = ttk.Combobox(
            self.frame_der, textvariable=self.cert_name, values=self.cert_list, width=50)
        self.combo_cert.pack(pady=5)
        if self.cert_list:
            self.cert_name.set(self.cert_list[0])
        self.todos_seleccionados = False
        ttk.Button(self.frame_der, text="✓ Seleccionar todos",
                   command=self.toggle_seleccionar_todos).pack(pady=(0, 5))

        self.seleccionados = {}  # clave: path absoluto, valor: True/False
        ttk.Label(self.frame_der, text="Filtrar por estado:").pack(anchor="w")

        self.combo_filtro = ttk.Combobox(
            self.frame_der,
            textvariable=self.filtro_estado,
            values=["Todos", "No firmados", "Firmados", "Firmados caducados"],
            state="readonly",
            width=30
        )
        self.combo_filtro.pack(anchor="w", pady=(0, 5))
        self.combo_filtro.bind("<<ComboboxSelected>>",
                               lambda e: self.cargar_archivos())

        self.tree = ttk.Treeview(
            self.frame_der,
            columns=("checkbox", "path", "status", "expira", "firma", "cert"),
            show="headings",
            selectmode="none",
            height=15
        )

        # Agregamos una columna fantasma (#0) para el "checkbox"
        self.tree["columns"] = (
            "checkbox", "path", "status", "expira", "firma", "cert")
        self.tree.column("#0", width=0, stretch=False)  # Oculta columna real
        self.tree.column("checkbox", width=30, anchor="center")
        self.tree.heading("checkbox", text="✓")

        # Columnas existentes
        self.tree.heading("path", text="Archivo")
        self.tree.heading("status", text="Estado")
        self.tree.heading("expira", text="Expira")
        self.tree.heading("firma", text="Firmado el")
        self.tree.heading("cert", text="Certificado")

        self.tree.column("path", width=400)
        self.tree.column("status", width=120, anchor="center")
        self.tree.column("expira", width=150, anchor="center")
        self.tree.column("firma", width=150, anchor="center")
        self.tree.column("cert", width=160, anchor="w")

        # Estilos visuales por estado de firma
        style = ttk.Style()
        style.configure("Treeview", rowheight=25)

        # Verde claro → firmado válido
        self.tree.tag_configure("firmado_ok", background="#d0f0c0")
        # Rojo claro → firmado pero expirado
        self.tree.tag_configure("firmado_exp", background="#f8d7da")
        # Gris claro → no firmado
        self.tree.tag_configure("no_firmado", background="#e2e3e5")

        self.tree.pack(pady=5, fill="both", expand=True)

        self.progress = ttk.Progressbar(
            self.frame_der, orient="horizontal", mode="determinate", maximum=100)

        self.tree.bind("<Button-1>", self.toggle_checkbox)

        ttk.Button(self.frame_der, text="Firmar archivos",
                   command=self.firmar_todo).pack(pady=10)
        ttk.Button(self.frame_der, text="Refrescar documentos",
                   command=self.cargar_archivos).pack(pady=(5, 0))
        ttk.Button(self.frame_der, text="➕ Añadir a whitelist",
                   command=self.anadir_a_whitelist).pack(pady=(0, 5))

        self.cargar_archivos()
        self.configurar_pestana_programar()

    def configurar_pestana_programar(self):
        ttk.Label(self.frame_prog, text="Hora de ejecución (HH:MM):").grid(
            row=0, column=0, sticky="w", padx=5, pady=5)
        self.hora = tk.StringVar(value="10:00")
        ttk.Entry(self.frame_prog, textvariable=self.hora,
                  width=10).grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(self.frame_prog, text="Frecuencia:").grid(
            row=1, column=0, sticky="w", padx=5, pady=5)
        self.frecuencia = tk.StringVar()
        combo_frec = ttk.Combobox(self.frame_prog, textvariable=self.frecuencia, values=[
            "DIARIAMENTE", "SEMANALMENTE"], state="readonly", width=20)
        combo_frec.grid(row=1, column=1, padx=5, pady=5)
        combo_frec.set("DIARIAMENTE")

        ttk.Label(self.frame_prog, text="Script a ejecutar:").grid(
            row=2, column=0, sticky="w", padx=5, pady=5)
        self.script = tk.StringVar(value="firma_programada.py")
        ttk.Entry(self.frame_prog, textvariable=self.script,
                  width=40).grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(self.frame_prog, text="Certificado:").grid(
            row=3, column=0, sticky="w", padx=5, pady=5)
        self.cert_programado = tk.StringVar()
        combo_cert = ttk.Combobox(
            self.frame_prog, textvariable=self.cert_programado, values=self.cert_list, width=50)
        combo_cert.grid(row=3, column=1, padx=5, pady=5)
        if self.cert_list:
            self.cert_programado.set(self.cert_list[0])

        ttk.Button(
            self.frame_prog,
            text="📅 Programar firma",
            command=lambda: crear_tarea_programada(
                self.hora.get(),
                os.path.abspath(self.script.get()),
                self.frecuencia.get().upper(),
                self.cert_programado.get()
            )
        ).grid(row=4, column=0, columnspan=2, pady=10)

    def cargar_archivos(self):
        self.tree.delete(*self.tree.get_children())
        self.archivos = list_office_files_recursively(self.folder)
        hashes_validos = cargar_hashes_autorizados()

        for full_path, relative_path in self.archivos:
            hash_actual = calcular_hash_vba_code(full_path)

            if not hash_actual or hash_actual not in hashes_validos:
                continue  # Ignorar si no está en la whitelist
            firmado, expiracion, expirada, timestamp, cert_name = verificar_firma(
                config.signtool_path, full_path)

            if firmado:
                if expirada:
                    estado = "Expirado"
                    tag = "firmado_exp"
                else:
                    estado = "Firmado"
                    tag = "firmado_ok"

            else:
                estado = "No firmado"
                tag = "no_firmado"

            expira_str = expiracion if expiracion else "-"
            firma_str = timestamp if timestamp else "-"
            cert_str = cert_name if cert_name else "-"
            # Aplicar filtro seleccionado
            mostrar = True
            if self.filtro_estado.get() == "No firmados" and firmado:
                mostrar = False
            elif self.filtro_estado.get() == "Firmados" and not (firmado and not expirada):
                mostrar = False
            elif self.filtro_estado.get() == "Firmados caducados" and not (firmado and expirada):
                mostrar = False

            if not mostrar:
                continue

            self.seleccionados[full_path] = False  # inicia desmarcado
            self.tree.insert(
                "",
                "end",
                iid=full_path,
                values=("☐", relative_path, estado,
                        expira_str, firma_str, cert_str),
                tags=(tag,)
            )

    def toggle_checkbox(self, event):
        region = self.tree.identify_region(event.x, event.y)
        column = self.tree.identify_column(event.x)
        row = self.tree.identify_row(event.y)

        if region == "cell" and column == "#1":  # checkbox está en la 1ª columna visible
            estado_actual = self.seleccionados.get(row, False)
            nuevo_estado = not estado_actual
            self.seleccionados[row] = nuevo_estado
            simbolo = "✓" if nuevo_estado else "☐"
            valores = list(self.tree.item(row, "values"))
            valores[0] = simbolo
            self.tree.item(row, values=valores)

    def firmar_todo(self):
        if not self.cert_name.get():
            messagebox.showerror("Error", "Selecciona un certificado.")
            return

        seleccionados_a_firmar = [
            (full_path, relative_path)
            for full_path, relative_path in self.archivos
            if self.seleccionados.get(full_path, False)
            and (self.tree.set(full_path, "status") in ("No firmado", "Expirado"))
        ]

        total = len(seleccionados_a_firmar)
        if total == 0:
            messagebox.showinfo("Nada que firmar",
                                "No hay archivos seleccionados o pendientes.")
            return

        self.progress["value"] = 0
        self.progress.pack(fill="x", pady=(5, 10))
        self.master.update_idletasks()

        log_path = os.path.join(config.logs_directory, "firma_log.csv")
        archivo_nuevo = not os.path.exists(log_path)
        firmados_ok = 0  # Contador de archivos firmados correctamente

        with open(log_path, mode="a", newline="", encoding="utf-8-sig") as log_file:
            writer = csv.writer(log_file, delimiter=';',
                                quoting=csv.QUOTE_MINIMAL)

            if archivo_nuevo:
                writer.writerow(["Archivo", "Fecha", "Resultado",
                                "Certificado", "Mensaje"])

            for i, (full_path, relative_path) in enumerate(seleccionados_a_firmar, 1):
                try:
                    self.tree.set(full_path, column="status",
                                  value="Firmando...")
                    self.master.update_idletasks()

                    ok, salida = sign_file(
                        config.signtool_path,
                        self.cert_name.get(),
                        config.timestamp_url,
                        full_path
                    )

                    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    if ok:
                        estado = "Firmado"
                        mensaje = "OK"
                        firmados_ok += 1
                    else:
                        estado = "Error"
                        mensaje = salida.strip()[
                            :250] if salida else "Fallo desconocido"

                except Exception as ex:
                    estado = "Error"
                    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    mensaje = f"Excepción: {str(ex)}"
                    ok = False

                self.tree.set(full_path, column="status", value=estado)
                writer.writerow([relative_path, fecha, estado,
                                self.cert_name.get(), mensaje])

                if not ok:
                    print(f"❌ {relative_path}\n{mensaje}")

                self.progress["value"] = int(i / total * 100)
                self.master.update_idletasks()

        self.progress["value"] = 100
        self.progress.pack_forget()
        messagebox.showinfo("Proceso completado",
                            f"Firma completada: {firmados_ok} de {total} archivos firmados correctamente.\n\nLog guardado en:\n{log_path}")
        self.cargar_archivos()

    def toggle_seleccionar_todos(self):
        self.todos_seleccionados = not self.todos_seleccionados
        nuevo_simbolo = "✓" if self.todos_seleccionados else "☐"

        for full_path, _ in self.archivos:
            self.seleccionados[full_path] = self.todos_seleccionados
            valores = list(self.tree.item(full_path, "values"))
            valores[0] = nuevo_simbolo
            self.tree.item(full_path, values=valores)

    def populate_folder_tree(self, root_path):
        self.folder_tree.delete(*self.folder_tree.get_children())

        def insert_node(parent, path):
            for nombre in sorted(os.listdir(path)):
                ruta = os.path.join(path, nombre)
                if os.path.isdir(ruta):
                    node = self.folder_tree.insert(
                        parent, "end", text=nombre, open=False, values=[ruta])
                    insert_node(node, ruta)

        root_node = self.folder_tree.insert(
            "", "end", text=os.path.basename(root_path), open=True, values=[root_path])
        insert_node(root_node, root_path)

    def on_folder_select(self, event):
        selected_item = self.folder_tree.selection()
        if selected_item:
            ruta = self.folder_tree.item(selected_item[0], "values")[0]
            self.folder = ruta
            self.cargar_archivos()

    def anadir_a_whitelist(self):

        ventana = tk.Toplevel(self.master)
        ventana.title("Añadir macros a la whitelist")
        ventana.geometry("700x400")
        ventana.transient(self.master)

        ttk.Label(ventana, text="Macros disponibles (no están en whitelist):").pack(
            pady=5)

        frame_lista = ttk.Frame(ventana)
        frame_lista.pack(fill="both", expand=True, padx=10)

        tree = ttk.Treeview(frame_lista, columns=("archivo",),
                            show="headings", selectmode="extended")
        tree.heading("archivo", text="Archivo")
        tree.column("archivo", anchor="w", width=600)
        tree.pack(fill="both", expand=True)

        scrollbar = ttk.Scrollbar(
            frame_lista, orient="vertical", command=tree.yview)
        scrollbar.pack(side="right", fill="y")
        tree.configure(yscrollcommand=scrollbar.set)

        hashes_actuales = cargar_hashes_autorizados()
        archivos_no_whitelist = []

        for full_path, relative_path in self.archivos:
            hash_actual = calcular_hash_vba_code(full_path)
            if not hash_actual or hash_actual in hashes_actuales:
                continue
            archivos_no_whitelist.append((full_path, relative_path))
            tree.insert("", "end", iid=full_path, values=(relative_path,))

        def confirmar_agregado():
            seleccionados = tree.selection()
            if not seleccionados:
                messagebox.showinfo(
                    "Sin selección", "No has seleccionado ninguna macro.")
                return

            añadidos = 0
            with open(config.whitelist, "a", encoding="utf-8") as f:
                for full_path in seleccionados:
                    hash_macro = calcular_hash_vba_code(full_path)
                    if hash_macro and hash_macro not in hashes_actuales:
                        f.write(hash_macro + "\n")
                        añadidos += 1

            messagebox.showinfo("Whitelist actualizada",
                                f"{añadidos} macro(s) añadida(s).")
            ventana.destroy()
            self.cargar_archivos()

        ttk.Button(ventana, text="➕ Añadir seleccionadas",
                   command=confirmar_agregado).pack(pady=10)
