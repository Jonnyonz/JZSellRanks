# JZSellRanks (Ranking de Ventas)

App de escritorio para Windows que arma, mes a mes, el ranking de ventas por sucursal y por
razón social a partir del Excel exportado del ERP. Funciona sin internet ni navegador. Parte de
**JZTech Suite**.

- Aplica **reglas configurables** para reasignar ventas a otra sucursal o razón social (por
  vendedor, usuario creador, localidad o sucursal), guardadas en **perfiles**.
- Autocompleta las reglas con los valores reales que detecta en el Excel cargado.
- Mantiene un **histórico acumulado**: cada mes procesado se suma a los anteriores.
- Genera un Excel con ranking de locales y de empresas, variación de posición contra el mes
  anterior, gráficos de evolución y de torta.

- [Instalación (usuarios)](#instalación-usuarios)
- [Cómo se usa](#cómo-se-usa)
- [Detalle técnico](#detalle-técnico)
- [Configuración y datos](#configuración-y-datos)
- [Desarrollo: correr desde el código](#desarrollo-correr-desde-el-código)
- [Construir el ejecutable y el instalador](#construir-el-ejecutable-y-el-instalador)
- [Limitaciones conocidas](#limitaciones-conocidas)
- [Contribuir y licencia](#contribuir-y-licencia)

---

## Instalación (usuarios)

1. Descargar `RankingVentas_Setup_x.x.x.exe` desde
   [Releases](https://github.com/Jonnyonz/JZSellRanks/releases/latest).
2. Ejecutarlo y seguir el asistente. **No requiere permisos de administrador** (instala para el
   usuario actual; el asistente permite elegir instalar para todos).
3. Queda en el menú Inicio y, opcionalmente, en el Escritorio.

No hace falta instalar Python: el ejecutable trae todo. Es una app de escritorio (no usa
servidor ni Docker) para Windows de 64 bits.

---

## Cómo se usa

### 1. El Excel de entrada

El archivo exportado del ERP tiene que ser `.xlsx` con:

- una hoja llamada **`Sheet2`**,
- los encabezados en la **fila 2**,
- estas columnas (los nombres tienen que coincidir; se ignoran espacios en los extremos y
  saltos de línea):

| Columna | Obligatoria | Uso |
|---|---|---|
| `Sucursal` | Sí | Agrupación principal. Las filas sin sucursal se descartan. |
| `Código Fecha Documento` | Sí | Mes de la venta. Las fechas inválidas se descartan. |
| `Total Total` | Sí | Total. |
| `Neto Gravado Neto Gravado` | Sí | Subtotal. |
| `IVA IVA` | Sí | IVA. |
| `Razon Social Empresa - Division` | Sí | Razón social (sin tilde en "Razon"). |
| `Vendedor`, `Usuario Creador`, `Localidad Documento` | No | Solo para las reglas. |

Los montos que no son números se toman como 0.

### 2. Las reglas (pestaña "Reglas y Razones Sociales")

Cada regla dice: *si la columna X **contiene** el texto T (sin distinguir mayúsculas),
opcionalmente solo cuando la sucursal actual es S, asignar la sucursal A y, si se indica, la
razón social R*.

- Se aplican **en orden**, de arriba hacia abajo: una regla ve el resultado de las anteriores.
- Asignar la sucursal **`EXCLUIR`** (en mayúsculas) saca esas ventas del reporte.
- Las reglas se guardan en **perfiles** (Nuevo, Guardar como, Eliminar). Siempre queda al
  menos uno.
- Al elegir un Excel en la otra pestaña, la app lee sus valores reales (vendedores, usuarios,
  localidades, sucursales, razones sociales) y los ofrece en los desplegables de las reglas.

### 3. Generar (pestaña "Ejecutar Reporte")

Elegir el Excel, generar y elegir dónde guardar. El nombre sugerido es
`Reporte_<Mes>_<Año>.xlsx`, o `Reporte_Acumulado_<Mes1>_a_<MesN>.xlsx` si hay varios meses.

El reporte incluye **todos los meses del histórico**. Si el Excel nuevo trae un mes que ya
estaba, ese mes se reemplaza por los datos nuevos. El botón de reinicio borra el histórico.

### 4. Qué contiene el Excel generado

- Hoja **Análisis**: tres gráficos de barras (total de la empresa, por razón social con línea
  de tendencia, y por sucursal apilado).
- Una hoja **por mes** ("Agosto 2026", etc.):
  - **TOP POR LOCALES:** posición, variación contra el mes anterior (`+N`, `-N` o `=`),
    posición anterior, razón social, local, subtotal, IVA y total, con escala de color.
  - **TOP POR EMPRESAS:** lo mismo agrupado por razón social.
  - Dos gráficos de torta (locales y razones sociales).
- Una hoja oculta `Data_Oculta` con los datos de los gráficos.

---

## Detalle técnico

| Componente | Detalle |
|---|---|
| Lenguaje | Python 3.11+ |
| Interfaz | Tkinter / ttk (biblioteca estándar) |
| Datos | `pandas==3.0.6` (lectura y cálculo), `openpyxl==3.1.5` (Excel de salida y gráficos) |
| Empaquetado | PyInstaller (un solo `.exe`, sin consola) + Inno Setup 6 (instalador) |

```
JZSellRanks/
├── core.py               # Lógica: lectura del Excel, reglas, perfiles, histórico, reporte
├── desktop_app.py        # Interfaz Tkinter (pestañas de reporte y de reglas)
├── requirements.txt      # pandas y openpyxl, fijados
├── RankingVentas.spec    # Configuración de PyInstaller
└── installer/setup.iss   # Script de Inno Setup
```

`core.py` no depende de la interfaz: `generar_reporte(ruta_excel, reglas, archivo_acumulado,
log)` devuelve el libro de Excel y el nombre sugerido. La interfaz corre la lectura y la
generación en hilos para no congelar la ventana.

**Seguridad del Excel de salida:** todo texto que viene del archivo de entrada y empieza con
`=`, `+`, `-`, `@`, tabulación o salto de línea se guarda con un apóstrofo delante, para que
Excel no lo interprete como fórmula (inyección de fórmulas / CSV injection).

---

## Configuración y datos

No hay archivo de configuración: todo se maneja desde la interfaz. Los datos del usuario se
guardan en **`%APPDATA%\RankingVentas\`**, sin importar dónde esté instalado el programa:

| Archivo | Contenido |
|---|---|
| `datos_acumulados.csv` | Histórico mensual por sucursal (lo que alimenta el reporte acumulado) |
| `perfiles_reglas\<perfil>.csv` | Reglas de cada perfil. Columnas: `Sucursal Origen (Opcional)`, `Columna a Buscar`, `Contiene el Texto`, `Asignar a Sucursal`, `Asignar a Razón Social` |

En el primer uso se crea el perfil `Default` (a partir de un `reglas_parametrizacion.csv`
viejo si lo encuentra junto al programa, o con reglas de ejemplo). Para hacer una copia de
seguridad o pasar la configuración a otra PC, copiar esa carpeta. El desinstalador no la borra.

---

## Desarrollo: correr desde el código

```powershell
git clone https://github.com/Jonnyonz/JZSellRanks.git
cd JZSellRanks
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python desktop_app.py
```

---

## Construir el ejecutable y el instalador

```powershell
pip install pyinstaller
# Usar el .spec versionado (no regenerarlo con opciones sueltas, que lo pisan):
pyinstaller RankingVentas.spec --distpath dist_desktop --workpath build_desktop
```

Queda `dist_desktop\RankingVentas.exe`. Después, con
[Inno Setup 6](https://jrsoftware.org/isinfo.php):

```powershell
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\setup.iss
```

Queda `dist_installer\RankingVentas_Setup_<versión>.exe`. La versión se define en
`installer\setup.iss` (`MyAppVersion`): subirla antes de cada release. Publicar:

```powershell
gh release create vX.Y.Z dist_installer\RankingVentas_Setup_X.Y.Z.exe --title "vX.Y.Z" --notes "Cambios"
```

---

## Limitaciones conocidas

- La lectura del Excel espera el formato exacto de arriba (hoja `Sheet2`, encabezados en la
  fila 2, nombres de columna fijos). Si falta algo, el reporte falla.
- Los reemplazos de algunas razones sociales y los nombres de columna están fijos en el código.
- Cuando una sucursal vende con varias razones sociales en el mismo mes, todo su total se
  atribuye a la más frecuente.
- El histórico se actualiza al generar, aunque después se cancele el guardado del Excel.
- Si ocurre un error durante la generación, el detalle queda en el registro de la pestaña pero
  hoy no aparece el cuadro de aviso.
- No hay tests automatizados.

---

## Contribuir y licencia

Las contribuciones son bienvenidas. Cada commit tiene que llevar `Signed-off-by`
(`git commit -s`, Developer Certificate of Origin) y ser un único cambio probado. Sin emojis en
la interfaz.

Licencia: ver el archivo `LICENSE`.
