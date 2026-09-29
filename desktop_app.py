"""
Ranking de Ventas - App de escritorio (Tkinter)
Reemplaza la versión web (Streamlit) por una ventana nativa, sin navegador ni servidor local.
"""
import os
import sys
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

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

        ttk.Combobox(self, textvariable=self.var_columna, values=core.COLUMNAS_REGLA, width=16, state="readonly").grid(
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
        titulos = ["Sucursal Origen", "Buscar en", "Texto Exacto", "Nueva Sucursal", "Nueva Razón Social", ""]
        anchos = [16, 16, 22, 18, 18, 3]
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


class TabReporte(ttk.Frame):
    def __init__(self, master, tab_reglas: TabReglas):
        super().__init__(master)
        self.tab_reglas = tab_reglas
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

        self.btn_generar = ttk.Button(self, text="▶ Generar Reporte", command=self.generar, state="disabled")
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
            valores = core.detectar_valores(ruta)
        except Exception as e:
            self.after(0, lambda: self.lbl_deteccion.config(text=f"No se pudieron autodetectar valores: {e}"))
            return
        self.after(0, lambda: self._aplicar_deteccion(valores))

    def _aplicar_deteccion(self, valores):
        self.tab_reglas.set_valores_detectados(valores)
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
        threading.Thread(target=self._generar_en_hilo, daemon=True).start()

    def _generar_en_hilo(self):
        try:
            df_reglas = self.tab_reglas.obtener_dataframe()
            wb, nombre_sugerido = core.generar_reporte(
                self.ruta_excel, df_reglas, archivo_acumulado=ARCHIVO_ACUMULADO, log=self.log
            )
            self.after(0, lambda: self._guardar_resultado(wb, nombre_sugerido))
        except Exception as e:
            detalle = traceback.format_exc()
            self.log(f"Error: {e}")
            self.log(detalle)
            self.after(0, self._finalizar)
            self.after(0, lambda: messagebox.showerror("Error procesando el archivo", str(e)))

    def _guardar_resultado(self, wb, nombre_sugerido):
        self._finalizar()
        ruta_salida = filedialog.asksaveasfilename(
            title="Guardar reporte como",
            initialfile=nombre_sugerido,
            defaultextension=".xlsx",
            filetypes=[("Archivo Excel", "*.xlsx")]
        )
        if ruta_salida:
            wb.save(ruta_salida)
            self.log(f"Reporte guardado en: {ruta_salida}")
            messagebox.showinfo("Listo", f"Reporte generado correctamente:\n{ruta_salida}")
        else:
            self.log("Guardado cancelado por el usuario.")

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

    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass

    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True)

    tab_reglas = TabReglas(notebook, DATOS_DIR, ARCHIVO_REGLAS_LEGACY)
    tab_reporte = TabReporte(notebook, tab_reglas)

    notebook.add(tab_reporte, text="Ejecutar Reporte")
    notebook.add(tab_reglas, text="Reglas y Razones Sociales")

    root.mainloop()


if __name__ == "__main__":
    main()
