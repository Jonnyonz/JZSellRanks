"""
Ranking de Ventas - App de escritorio (Tkinter): ventana nativa, sin navegador ni servidor local.
La lógica está en core.py; acá solo viven las pestañas, los hilos y los cuadros de diálogo.
"""
import os
import sys
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

import pandas as pd

import core

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Los datos del usuario (perfiles de reglas, historial) viven en AppData, NO junto al .exe:
# así sobreviven si se reconstruye o se mueve el ejecutable.
DATOS_DIR = os.path.join(os.environ.get('APPDATA', BASE_DIR), 'RankingVentas')
os.makedirs(DATOS_DIR, exist_ok=True)

_CANDIDATOS_REGLAS_LEGACY = [
    os.path.join(BASE_DIR, "reglas_parametrizacion.csv"),
    os.path.join(os.path.dirname(BASE_DIR), "reglas_parametrizacion.csv"),
]
ARCHIVO_REGLAS_LEGACY = next((p for p in _CANDIDATOS_REGLAS_LEGACY if os.path.exists(p)), _CANDIDATOS_REGLAS_LEGACY[0])
ARCHIVO_ACUMULADO = os.path.join(DATOS_DIR, "datos_acumulados.csv")
ARCHIVO_EXCLUSIONES = os.path.join(DATOS_DIR, "facturas_excluidas.csv")


class FilaRegla(ttk.Frame):
    """Una fila editable de la tabla de reglas, con autocompletado desde el Excel cargado."""

    def __init__(self, master, tab_reglas, on_delete, valores=None):
        super().__init__(master)
        self.tab_reglas = tab_reglas
        valores = valores or {}

        self.var_origen = tk.StringVar(value=valores.get('Sucursal Origen (Opcional)', ''))
        self.var_columna = tk.StringVar(value=valores.get('Columna a Buscar', core.COLUMNAS_REGLA[0]))
        self.var_texto = tk.StringVar(value=valores.get('Contiene el Texto', ''))
        self.var_nueva_suc = tk.StringVar(value=valores.get('Asignar a Sucursal', ''))
        self.var_nueva_rs = tk.StringVar(value=valores.get('Asignar a Razón Social', ''))

        self.combo_origen = ttk.Combobox(self, textvariable=self.var_origen, width=16)
        self.combo_origen['postcommand'] = lambda: self.combo_origen.configure(
            values=self.tab_reglas.valores_detectados.get('Sucursal', []))
        self.combo_origen.grid(row=0, column=0, padx=2, pady=1, sticky="ew")

        ttk.Combobox(self, textvariable=self.var_columna, values=core.COLUMNAS_REGLA, width=28, state="readonly").grid(
            row=0, column=1, padx=2, pady=1, sticky="ew")

        self.combo_texto = ttk.Combobox(self, textvariable=self.var_texto, width=22)
        self.combo_texto['postcommand'] = lambda: self.combo_texto.configure(
            values=self.tab_reglas.valores_detectados.get(self.var_columna.get(), []))
        self.combo_texto.grid(row=0, column=2, padx=2, pady=1, sticky="ew")

        self.combo_nueva_suc = ttk.Combobox(self, textvariable=self.var_nueva_suc, width=18)
        self.combo_nueva_suc['postcommand'] = lambda: self.combo_nueva_suc.configure(
            values=self.tab_reglas.opciones_sucursal_destino())
        self.combo_nueva_suc.grid(row=0, column=3, padx=2, pady=1, sticky="ew")

        self.combo_nueva_rs = ttk.Combobox(self, textvariable=self.var_nueva_rs, width=18)
        self.combo_nueva_rs['postcommand'] = lambda: self.combo_nueva_rs.configure(
            values=self.tab_reglas.opciones_razon_social_destino())
        self.combo_nueva_rs.grid(row=0, column=4, padx=2, pady=1, sticky="ew")

        ttk.Button(self, text="X", width=3, command=lambda: on_delete(self)).grid(row=0, column=5, padx=2, pady=1)

    def a_dict(self):
        return {
            'Sucursal Origen (Opcional)': self.var_origen.get().strip(),
            'Columna a Buscar': self.var_columna.get().strip(),
            'Contiene el Texto': self.var_texto.get().strip(),
            'Asignar a Sucursal': self.var_nueva_suc.get().strip(),
            'Asignar a Razón Social': self.var_nueva_rs.get().strip(),
        }


class TabReglas(ttk.Frame):
    def __init__(self, master, perfiles_dir, archivo_legacy):
        super().__init__(master)
        self.perfiles_dir = perfiles_dir
        self.archivo_legacy = archivo_legacy
        self.filas = []
        self.valores_detectados = {}
        self.perfil_actual = None

        # --- BARRA DE PERFILES ---
        barra_perfil = ttk.Frame(self)
        barra_perfil.pack(fill="x", padx=10, pady=(10, 0))
        ttk.Label(barra_perfil, text="Perfil de reglas:", font=("Segoe UI", 9, "bold")).pack(side="left")
        self.var_perfil = tk.StringVar()
        self.combo_perfil = ttk.Combobox(barra_perfil, textvariable=self.var_perfil, state="readonly", width=25)
        self.combo_perfil.pack(side="left", padx=6)
        self.combo_perfil.bind("<<ComboboxSelected>>", self.al_cambiar_perfil)
        ttk.Button(barra_perfil, text="Nuevo", command=self.nuevo_perfil).pack(side="left", padx=2)
        ttk.Button(barra_perfil, text="Guardar como...", command=self.guardar_como).pack(side="left", padx=2)
        ttk.Button(barra_perfil, text="Eliminar perfil", command=self.eliminar_perfil_actual).pack(side="left", padx=2)

        ttk.Label(
            self,
            text="Agrega a qué Sucursal / Razón Social pertenece cada registro para el reporte agrupado. "
                 "Los campos con flecha sugieren valores detectados en el último Excel que cargaste.",
            wraplength=850
        ).pack(anchor="w", padx=10, pady=(8, 4))

        header = ttk.Frame(self)
        header.pack(fill="x", padx=10)
        titulos = ["Sucursal Origen", "Buscar en", "Contiene el texto", "Nueva Sucursal", "Nueva Razón Social", ""]
        anchos = [16, 28, 22, 18, 18, 3]
        for i, (t, a) in enumerate(zip(titulos, anchos)):
            ttk.Label(header, text=t, width=a, font=("Segoe UI", 9, "bold")).grid(row=0, column=i, padx=2)

        contenedor = ttk.Frame(self)
        contenedor.pack(fill="both", expand=True, padx=10, pady=4)

        canvas = tk.Canvas(contenedor, highlightthickness=0)
        scrollbar = ttk.Scrollbar(contenedor, orient="vertical", command=canvas.yview)
        self.frame_filas = ttk.Frame(canvas)

        self.frame_filas.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.frame_filas, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        botones = ttk.Frame(self)
        botones.pack(fill="x", padx=10, pady=6)
        ttk.Button(botones, text="+ Agregar fila", command=self.agregar_fila_vacia).pack(side="left")
        ttk.Button(botones, text="Guardar cambios en este perfil", command=self.guardar).pack(side="left", padx=8)
        self.lbl_estado = ttk.Label(botones, text="")
        self.lbl_estado.pack(side="left", padx=8)

        self.cargar_lista_perfiles(seleccionar=None)

    # --- MANEJO DE FILAS ---

    def agregar_fila(self, valores=None):
        fila = FilaRegla(self.frame_filas, self, self.eliminar_fila, valores)
        fila.grid(row=len(self.filas), column=0, sticky="ew", pady=1)
        self.filas.append(fila)

    def agregar_fila_vacia(self):
        self.agregar_fila()

    def eliminar_fila(self, fila):
        fila.destroy()
        self.filas.remove(fila)
        for i, f in enumerate(self.filas):
            f.grid(row=i, column=0, sticky="ew", pady=1)

    def limpiar_filas(self):
        for f in self.filas:
            f.destroy()
        self.filas = []

    def obtener_dataframe(self):
        import pandas as pd
        filas = [f.a_dict() for f in self.filas]
        if not filas:
            return pd.DataFrame(columns=core.COLUMNAS_CSV_REGLAS)
        return pd.DataFrame(filas)[core.COLUMNAS_CSV_REGLAS]

    # --- AUTOCOMPLETADO ---

    def set_valores_detectados(self, valores):
        self.valores_detectados = valores or {}

    def _valores_usados(self, campo):
        atributo = {'Asignar a Sucursal': 'var_nueva_suc', 'Asignar a Razón Social': 'var_nueva_rs'}[campo]
        return {getattr(f, atributo).get().strip() for f in self.filas if getattr(f, atributo).get().strip()}

    def opciones_sucursal_destino(self):
        detectados = self.valores_detectados.get('Sucursal', [])
        usados = self._valores_usados('Asignar a Sucursal')
        return sorted(set(detectados) | usados | {"EXCLUIR"})

    def opciones_razon_social_destino(self):
        detectados = self.valores_detectados.get('Razón Social', [])
        usados = self._valores_usados('Asignar a Razón Social')
        return sorted(set(detectados) | usados)

    # --- PERFILES ---

    def cargar_lista_perfiles(self, seleccionar=None):
        nombres = core.listar_perfiles(self.perfiles_dir, self.archivo_legacy)
        self.combo_perfil['values'] = nombres
        objetivo = seleccionar or self.perfil_actual or nombres[0]
        if objetivo not in nombres:
            objetivo = nombres[0]
        self.var_perfil.set(objetivo)
        self.perfil_actual = objetivo
        self.cargar_perfil_actual()

    def cargar_perfil_actual(self):
        self.limpiar_filas()
        df = core.cargar_perfil(self.perfiles_dir, self.perfil_actual)
        for _, row in df.iterrows():
            self.agregar_fila(row.to_dict())

    def al_cambiar_perfil(self, event=None):
        nuevo = self.var_perfil.get()
        if nuevo == self.perfil_actual:
            return
        if not messagebox.askyesno(
            "Cambiar de perfil",
            f"¿Cambiar al perfil '{nuevo}'?\nLos cambios sin guardar en '{self.perfil_actual}' se van a perder."
        ):
            self.var_perfil.set(self.perfil_actual)
            return
        self.perfil_actual = nuevo
        self.cargar_perfil_actual()

    def nuevo_perfil(self):
        nombre = simpledialog.askstring("Nuevo perfil", "Nombre del nuevo perfil:", parent=self)
        if nombre is None:
            return
        nombre = nombre.strip()
        if not core.nombre_perfil_valido(nombre):
            messagebox.showerror("Nombre inválido", "Usa un nombre sin \\ / : * ? \" < > |")
            return
        if os.path.exists(core.ruta_perfil(self.perfiles_dir, nombre)):
            messagebox.showerror("Ya existe", f"Ya existe un perfil llamado '{nombre}'.")
            return
        import pandas as pd
        core.guardar_perfil(self.perfiles_dir, nombre, pd.DataFrame(columns=core.COLUMNAS_CSV_REGLAS))
        self.cargar_lista_perfiles(seleccionar=nombre)
        self.lbl_estado.config(text=f"Perfil '{nombre}' creado")
        self.after(3000, lambda: self.lbl_estado.config(text=""))

    def guardar_como(self):
        nombre = simpledialog.askstring("Guardar como", "Nombre del perfil:", initialvalue=self.perfil_actual, parent=self)
        if nombre is None:
            return
        nombre = nombre.strip()
        if not core.nombre_perfil_valido(nombre):
            messagebox.showerror("Nombre inválido", "Usa un nombre sin \\ / : * ? \" < > |")
            return
        if os.path.exists(core.ruta_perfil(self.perfiles_dir, nombre)) and not messagebox.askyesno(
            "Sobrescribir", f"Ya existe un perfil '{nombre}'. ¿Sobrescribirlo?"
        ):
            return
        core.guardar_perfil(self.perfiles_dir, nombre, self.obtener_dataframe())
        self.cargar_lista_perfiles(seleccionar=nombre)
        self.lbl_estado.config(text=f"Guardado como '{nombre}'")
        self.after(3000, lambda: self.lbl_estado.config(text=""))

    def eliminar_perfil_actual(self):
        nombres = core.listar_perfiles(self.perfiles_dir, self.archivo_legacy)
        if len(nombres) <= 1:
            messagebox.showwarning("No se puede eliminar", "Tiene que quedar al menos un perfil.")
            return
        if not messagebox.askyesno("Eliminar perfil", f"¿Eliminar el perfil '{self.perfil_actual}'? No se puede deshacer."):
            return
        core.eliminar_perfil(self.perfiles_dir, self.perfil_actual)
        self.perfil_actual = None
        self.cargar_lista_perfiles()

    def guardar(self):
        df = self.obtener_dataframe()
        core.guardar_perfil(self.perfiles_dir, self.perfil_actual, df)
        self.lbl_estado.config(text=f"Cambios guardados en '{self.perfil_actual}'")
        self.after(3000, lambda: self.lbl_estado.config(text=""))


class TabExclusiones(ttk.Frame):
    """Buscar comprobantes del Excel cargado y marcarlos como excluidos del reporte (persisten en
    facturas_excluidas.csv). No lista nada hasta que se busca."""

    COLS_RESULTADO = ("Tipo", "Documento", "Fecha", "Cliente", "Sucursal", "Total")
    COLS_EXCLUIDAS = ("Tipo", "Documento", "Fecha", "Cliente", "Sucursal", "Total", "Motivo", "Excluida el")
    ANCHOS = {"Tipo": 50, "Documento": 130, "Fecha": 90, "Cliente": 260, "Sucursal": 150, "Total": 110,
              "Motivo": 220, "Excluida el": 90}

    def __init__(self, master, archivo_exclusiones):
        super().__init__(master)
        self.archivo_exclusiones = archivo_exclusiones
        self.docs = None
        self.exclusiones = core.cargar_exclusiones(archivo_exclusiones)

        ttk.Label(self, text="Excluir facturas del reporte", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=10, pady=(10, 4))
        ttk.Label(
            self,
            text="Buscá los comprobantes del Excel cargado por cliente (razón social), punto de venta, número o "
                 "tipo, seleccioná los que no van y apretá Excluir. Las exclusiones quedan guardadas y se aplican "
                 "cada vez que generás el reporte, aunque vuelvas a procesar el mismo mes.",
            wraplength=900
        ).pack(anchor="w", padx=10)

        # --- FILTROS ---
        filtros = ttk.Frame(self)
        filtros.pack(fill="x", padx=10, pady=(8, 2))
        self.var_cliente, self.var_punto, self.var_numero, self.var_tipo = (tk.StringVar() for _ in range(4))
        ttk.Label(filtros, text="Cliente (razón social)").grid(row=0, column=0, sticky="w", padx=2)
        self.combo_cliente = ttk.Combobox(filtros, textvariable=self.var_cliente, width=40)
        self.combo_cliente.grid(row=1, column=0, padx=2, sticky="ew")
        ttk.Label(filtros, text="Punto de venta").grid(row=0, column=1, sticky="w", padx=2)
        ttk.Entry(filtros, textvariable=self.var_punto, width=10).grid(row=1, column=1, padx=2)
        ttk.Label(filtros, text="Número").grid(row=0, column=2, sticky="w", padx=2)
        ttk.Entry(filtros, textvariable=self.var_numero, width=12).grid(row=1, column=2, padx=2)
        ttk.Label(filtros, text="Tipo").grid(row=0, column=3, sticky="w", padx=2)
        self.combo_tipo = ttk.Combobox(filtros, textvariable=self.var_tipo, width=8, state="readonly", values=[""])
        self.combo_tipo.grid(row=1, column=3, padx=2)
        self.btn_buscar = ttk.Button(filtros, text="Buscar", command=self.buscar, state="disabled")
        self.btn_buscar.grid(row=1, column=4, padx=(8, 2))
        ttk.Button(filtros, text="Limpiar", command=self.limpiar_filtros).grid(row=1, column=5, padx=2)
        self.combo_cliente.bind("<Return>", lambda e: self.buscar())
        self.lbl_estado = ttk.Label(self, text="Primero elegí el Excel del ERP en la pestaña Ejecutar Reporte.")
        self.lbl_estado.pack(anchor="w", padx=10, pady=(2, 0))

        # --- RESULTADOS ---
        self.tabla = self._crear_tabla(self, self.COLS_RESULTADO, alto=9)
        acciones = ttk.Frame(self)
        acciones.pack(fill="x", padx=10, pady=(2, 8))
        ttk.Button(acciones, text="Excluir seleccionadas", command=lambda: self.excluir(solo_seleccion=True)).pack(side="left")
        ttk.Button(acciones, text="Excluir todas las listadas", command=lambda: self.excluir(solo_seleccion=False)).pack(side="left", padx=6)

        # --- EXCLUIDAS ---
        barra_ex = ttk.Frame(self)
        barra_ex.pack(fill="x", padx=10)
        ttk.Label(barra_ex, text="Facturas excluidas", font=("Segoe UI", 10, "bold")).pack(side="left")
        self.lbl_cant_excluidas = ttk.Label(barra_ex, text="")
        self.lbl_cant_excluidas.pack(side="left", padx=8)
        ttk.Label(barra_ex, text="Filtrar:").pack(side="left", padx=(16, 2))
        self.var_filtro_ex = tk.StringVar()
        ent = ttk.Entry(barra_ex, textvariable=self.var_filtro_ex, width=30)
        ent.pack(side="left")
        ent.bind("<KeyRelease>", lambda e: self.refrescar_excluidas())
        ttk.Button(barra_ex, text="Restaurar seleccionadas", command=self.restaurar).pack(side="right")
        self.tabla_ex = self._crear_tabla(self, self.COLS_EXCLUIDAS, alto=7)
        self.refrescar_excluidas()

    def _crear_tabla(self, master, columnas, alto):
        marco = ttk.Frame(master)
        marco.pack(fill="both", expand=True, padx=10, pady=(4, 0))
        tabla = ttk.Treeview(marco, columns=columnas, show="headings", height=alto, selectmode="extended")
        for c in columnas:
            tabla.heading(c, text=c)
            tabla.column(c, width=self.ANCHOS[c], anchor="e" if c == "Total" else "w", stretch=(c in ("Cliente", "Motivo")))
        barra = ttk.Scrollbar(marco, orient="vertical", command=tabla.yview)
        tabla.configure(yscrollcommand=barra.set)
        tabla.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        return tabla

    # --- DATOS DEL EXCEL ---

    def set_documentos(self, docs):
        self.docs = docs
        clientes = sorted(c for c in docs['Cliente'].unique() if c)
        self.combo_cliente['values'] = clientes
        self.combo_tipo['values'] = [""] + sorted(t for t in docs['Tipo'].unique() if t)
        self.btn_buscar.config(state="normal")
        self.tabla.delete(*self.tabla.get_children())
        self.lbl_estado.config(text=f"Excel cargado: {len(docs)} comprobantes de {len(clientes)} clientes. Buscá para listar.")

    def limpiar_filtros(self):
        for v in (self.var_cliente, self.var_punto, self.var_numero, self.var_tipo):
            v.set("")
        self.tabla.delete(*self.tabla.get_children())

    def buscar(self):
        if self.docs is None:
            return
        if not any(v.get().strip() for v in (self.var_cliente, self.var_punto, self.var_numero, self.var_tipo)):
            self.lbl_estado.config(text="Completá al menos un filtro (cliente, punto, número o tipo).")
            return
        res = core.buscar_documentos(self.docs, self.var_cliente.get(), self.var_punto.get(), self.var_numero.get(), self.var_tipo.get())
        res = res[~res['Clave'].isin(set(self.exclusiones['Clave']))]
        self.tabla.delete(*self.tabla.get_children())
        filas = core.filas_a_exclusiones(res)
        for _, r in filas.iterrows():
            self.tabla.insert("", "end", iid=r['Clave'], values=self._valores_fila(r, self.COLS_RESULTADO))
        self.lbl_estado.config(text=f"{len(filas)} comprobantes encontrados (sin contar los ya excluidos).")

    @staticmethod
    def _valores_fila(r, columnas):
        out = []
        for c in columnas:
            v = r[c]
            if c == "Total":
                try:
                    v = f"{float(v):,.2f}"
                except (TypeError, ValueError):
                    pass
            out.append(v)
        return out

    # --- EXCLUIR / RESTAURAR ---

    def excluir(self, solo_seleccion):
        claves = list(self.tabla.selection()) if solo_seleccion else list(self.tabla.get_children())
        if not claves:
            messagebox.showinfo("Excluir facturas", "No hay comprobantes seleccionados.")
            return
        motivo = simpledialog.askstring("Motivo", f"Motivo de la exclusión ({len(claves)} comprobantes, opcional):", parent=self)
        if motivo is None:
            return
        nuevas = core.filas_a_exclusiones(self.docs[self.docs['Clave'].isin(claves)], motivo.strip())
        nuevas = nuevas[~nuevas['Clave'].isin(set(self.exclusiones['Clave']))]
        self.exclusiones = pd.concat([self.exclusiones, nuevas], ignore_index=True)
        core.guardar_exclusiones(self.exclusiones, self.archivo_exclusiones)
        for clave in claves:
            self.tabla.delete(clave)
        self.refrescar_excluidas()
        self.lbl_estado.config(text=f"{len(nuevas)} comprobantes excluidos. Se aplican al generar el reporte.")

    def restaurar(self):
        claves = list(self.tabla_ex.selection())
        if not claves:
            messagebox.showinfo("Restaurar", "Seleccioná en la lista de excluidas las facturas a restaurar.")
            return
        if not messagebox.askyesno("Restaurar", f"¿Volver a incluir {len(claves)} comprobantes en el reporte?"):
            return
        self.exclusiones = self.exclusiones[~self.exclusiones['Clave'].isin(claves)].reset_index(drop=True)
        core.guardar_exclusiones(self.exclusiones, self.archivo_exclusiones)
        self.refrescar_excluidas()

    def refrescar_excluidas(self):
        filtro = self.var_filtro_ex.get().strip().lower()
        self.tabla_ex.delete(*self.tabla_ex.get_children())
        for _, r in self.exclusiones.iterrows():
            if filtro and filtro not in f"{r['Cliente']} {r['Documento']} {r['Motivo']} {r['Sucursal']}".lower():
                continue
            self.tabla_ex.insert("", "end", iid=r['Clave'], values=self._valores_fila(r, self.COLS_EXCLUIDAS))
        self.lbl_cant_excluidas.config(text=f"({len(self.exclusiones)} en total)")


class TabReporte(ttk.Frame):
    def __init__(self, master, tab_reglas: TabReglas, tab_exclusiones: TabExclusiones):
        super().__init__(master)
        self.tab_reglas = tab_reglas
        self.tab_exclusiones = tab_exclusiones
        self.ruta_excel = None

        ttk.Label(self, text="Generador de Reportes Ejecutivos", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=10, pady=(10, 4))

        ttk.Label(
            self,
            text="Cada Excel que subís se suma al histórico acumulado: si hoy subís agosto y el mes que viene "
                 "septiembre, el reporte final va a traer las dos solapas (y todas las que vengan después).",
            wraplength=850
        ).pack(anchor="w", padx=10)

        frame_sel = ttk.Frame(self)
        frame_sel.pack(fill="x", padx=10, pady=6)
        ttk.Button(frame_sel, text="Seleccionar Excel del ERP (.xlsx)", command=self.seleccionar_archivo).pack(side="left")
        self.lbl_archivo = ttk.Label(frame_sel, text="Ningún archivo seleccionado")
        self.lbl_archivo.pack(side="left", padx=10)
        ttk.Button(frame_sel, text="Reiniciar histórico acumulado", command=self.reiniciar_acumulado).pack(side="right")

        self.lbl_deteccion = ttk.Label(self, text="")
        self.lbl_deteccion.pack(anchor="w", padx=10)

        self.btn_generar = ttk.Button(self, text="Generar reporte", command=self.generar, state="disabled")
        self.btn_generar.pack(anchor="w", padx=10, pady=6)

        self.progreso = ttk.Progressbar(self, mode="indeterminate")
        self.progreso.pack(fill="x", padx=10, pady=(0, 6))

        ttk.Label(self, text="Registro:").pack(anchor="w", padx=10)
        self.txt_log = tk.Text(self, height=14, state="disabled", wrap="word")
        self.txt_log.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def log(self, mensaje):
        def _escribir():
            self.txt_log.config(state="normal")
            self.txt_log.insert("end", str(mensaje) + "\n")
            self.txt_log.see("end")
            self.txt_log.config(state="disabled")
        self.after(0, _escribir)

    def seleccionar_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Selecciona el Excel del ERP",
            filetypes=[("Archivos Excel", "*.xlsx")]
        )
        if ruta:
            self.ruta_excel = ruta
            self.lbl_archivo.config(text=os.path.basename(ruta))
            self.btn_generar.config(state="normal")
            self.lbl_deteccion.config(text="Detectando valores para autocompletar reglas...")
            threading.Thread(target=self._detectar_en_hilo, args=(ruta,), daemon=True).start()

    def _detectar_en_hilo(self, ruta):
        try:
            docs = core.leer_erp(ruta)  # una sola lectura: sirve para las reglas y para excluir facturas
            valores = core.detectar_valores(docs)
        except Exception as e:
            # `e` deja de existir al salir del except: el lambda corre después, así que se copia el texto.
            msg = f"No se pudieron autodetectar valores: {e}"
            self.after(0, lambda: self.lbl_deteccion.config(text=msg))
            return
        self.after(0, lambda: self._aplicar_deteccion(valores, docs))

    def _aplicar_deteccion(self, valores, docs):
        self.tab_reglas.set_valores_detectados(valores)
        self.tab_exclusiones.set_documentos(docs)
        resumen = ", ".join(f"{k}: {len(v)}" for k, v in valores.items())
        self.lbl_deteccion.config(text=f"Autocompletado listo ({resumen})")

    def generar(self):
        if not self.ruta_excel:
            return
        self.btn_generar.config(state="disabled")
        self.progreso.start(12)
        self.txt_log.config(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.config(state="disabled")
        # Las variables Tk se leen acá, en el hilo principal; el hilo de trabajo recibe el DataFrame.
        df_reglas = self.tab_reglas.obtener_dataframe()
        threading.Thread(target=self._generar_en_hilo, args=(df_reglas,), daemon=True).start()

    def _generar_en_hilo(self, df_reglas):
        try:
            wb, nombre_sugerido, acumulado = core.generar_reporte(
                self.ruta_excel, df_reglas, archivo_acumulado=ARCHIVO_ACUMULADO, log=self.log,
                archivo_exclusiones=ARCHIVO_EXCLUSIONES
            )
            self.after(0, lambda: self._guardar_resultado(wb, nombre_sugerido, acumulado))
        except Exception as e:
            detalle = traceback.format_exc()
            msg = str(e)  # `e` deja de existir al salir del except; el lambda corre después
            self.log(f"Error: {msg}")
            self.log(detalle)
            self.after(0, self._finalizar)
            self.after(0, lambda: messagebox.showerror("Error procesando el archivo", msg))

    def _guardar_resultado(self, wb, nombre_sugerido, acumulado):
        self._finalizar()
        ruta_salida = filedialog.asksaveasfilename(
            title="Guardar reporte como",
            initialfile=nombre_sugerido,
            defaultextension=".xlsx",
            filetypes=[("Archivo Excel", "*.xlsx")]
        )
        if not ruta_salida:
            self.log("Guardado cancelado por el usuario. El histórico acumulado no se modificó.")
            return
        try:
            wb.save(ruta_salida)
        except OSError as e:
            # Típico: el archivo está abierto en Excel (PermissionError). El exe no tiene consola,
            # así que sin este aviso el usuario no se enteraba de nada.
            self.log(f"No se pudo guardar el reporte: {e}")
            messagebox.showerror(
                "No se pudo guardar",
                f"No se pudo escribir el archivo:\n{ruta_salida}\n\n"
                "Si está abierto en Excel, cerralo y volvé a generar el reporte.\n"
                "El histórico acumulado no se modificó."
            )
            return
        core.guardar_acumulado(acumulado, ARCHIVO_ACUMULADO)
        self.log(f"Reporte guardado en: {ruta_salida}")
        messagebox.showinfo("Listo", f"Reporte generado correctamente:\n{ruta_salida}")

    def _finalizar(self):
        self.progreso.stop()
        self.btn_generar.config(state="normal")

    def reiniciar_acumulado(self):
        if not messagebox.askyesno(
            "Reiniciar histórico acumulado",
            "Esto borra TODOS los meses acumulados hasta ahora (ej. para arrancar un año nuevo).\n"
            "No se puede deshacer. ¿Continuar?"
        ):
            return
        core.reiniciar_acumulado(ARCHIVO_ACUMULADO)
        messagebox.showinfo("Listo", "Histórico acumulado reiniciado.")


def main():
    root = tk.Tk()
    root.title("Ranking de Ventas")
    root.geometry("950x680")
    # Ícono de la ventana y la barra de tareas. En el exe de PyInstaller va dentro de _MEIPASS.
    ruta_icono = os.path.join(getattr(sys, '_MEIPASS', BASE_DIR), 'installer', 'icono.png')
    if os.path.exists(ruta_icono):
        root.iconphoto(True, tk.PhotoImage(file=ruta_icono))

    # Tema nativo de Windows si existe; en otros sistemas queda el tema por defecto de ttk.
    style = ttk.Style()
    if "vista" in style.theme_names():
        style.theme_use("vista")

    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True)

    tab_reglas = TabReglas(notebook, DATOS_DIR, ARCHIVO_REGLAS_LEGACY)
    tab_exclusiones = TabExclusiones(notebook, ARCHIVO_EXCLUSIONES)
    tab_reporte = TabReporte(notebook, tab_reglas, tab_exclusiones)

    notebook.add(tab_reporte, text="Ejecutar Reporte")
    notebook.add(tab_reglas, text="Reglas y Razones Sociales")
    notebook.add(tab_exclusiones, text="Excluir facturas")

    root.mainloop()


if __name__ == "__main__":
    main()
