import tkinter as tk
from tkinter import ttk, messagebox
import os

from files import list_office_files_recursively
from certs import get_cert_names
from signer import sign_file
import config
from verify import verificar_firma


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

        self.frame_der = ttk.Frame(self.paned, padding=5)

        self.paned.add(self.frame_izq, weight=1)
        self.paned.add(self.frame_der, weight=4)

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

        self.cargar_archivos()

    def cargar_archivos(self):
        self.tree.delete(*self.tree.get_children())
        self.archivos = list_office_files_recursively(self.folder)

        for full_path, relative_path in self.archivos:
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
        # muestra la barra de progreso
        self.progress.pack(fill="x", pady=(5, 10))
        self.master.update_idletasks()

        for i, (full_path, relative_path) in enumerate(seleccionados_a_firmar, 1):
            self.tree.set(full_path, column="status", value="Firmando...")
            self.master.update_idletasks()

            ok, salida = sign_file(
                config.signtool_path, self.cert_name.get(), config.timestamp_url, full_path)
            if not ok:
                self.tree.set(full_path, column="status", value="Error")
                print(f"❌ {relative_path}\n{salida}")

            # Actualizar barra de progreso
            self.progress["value"] = int(i / total * 100)
            self.master.update_idletasks()

        self.progress["value"] = 100
        self.progress.pack_forget()  # <-- Oculta la barra
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
