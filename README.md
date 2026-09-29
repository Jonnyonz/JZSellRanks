# Ranking de Ventas

App de escritorio para armar, mes a mes, el ranking de ventas por sucursal y por razón social a partir del Excel exportado del ERP. No depende de internet ni de un navegador: es un programa nativo de Windows.

- Aplica reglas configurables para reasignar sucursales/razón social (por vendedor, usuario creador, localidad, etc.), con **perfiles guardables** para distintos casos de uso.
- Autocompleta esas reglas con los valores reales detectados en el Excel que cargues.
- Mantiene un **histórico acumulado**: cada mes que cargás se suma a los anteriores, así el reporte final siempre tiene todas las solapas de todos los meses procesados hasta la fecha.
- Genera un Excel ejecutivo con ranking, variación de posiciones mes a mes, gráficos de torta y de evolución.

## Instalación (para usuarios)

1. Andá a [Releases](https://github.com/Jonnyonz/JZSellRanks/releases/latest) y descargá `RankingVentas_Setup_x.x.x.exe`.
2. Ejecutalo y seguí el asistente (no requiere permisos de administrador).
3. Al terminar, se abre la app y queda un acceso directo en el Escritorio (opcional, lo podés destildar en el instalador) y en el menú de Inicio.

No hace falta instalar Python ni ninguna otra dependencia: el instalador ya trae todo lo necesario.

### Primer uso

- En la pestaña **Reglas y Razones Sociales** armá o ajustá tus reglas (podés tener varios "perfiles" según el tipo de reporte).
- En la pestaña **Ejecutar Reporte** seleccioná el Excel del ERP (`.xlsx`) y generá el reporte.
- Tus perfiles de reglas y el histórico acumulado se guardan en `%APPDATA%\RankingVentas`, independientemente de dónde esté instalado el programa.

## Para desarrolladores

### Correr desde el código fuente

Requiere Python 3.11+.

```powershell
git clone https://github.com/Jonnyonz/JZSellRanks.git
cd JZSellRanks
pip install -r requirements.txt
python desktop_app.py
```

### Reconstruir el ejecutable (.exe)

```powershell
pip install pyinstaller
python -m PyInstaller --onefile --windowed --name RankingVentas --collect-all numpy --collect-all pandas --distpath dist_desktop --workpath build_desktop --specpath . desktop_app.py
```

El resultado queda en `dist_desktop\RankingVentas.exe`.

### Reconstruir el instalador

Requiere [Inno Setup 6](https://jrsoftware.org/isinfo.php).

```powershell
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\setup.iss
```

El instalador queda en `dist_installer\RankingVentas_Setup_<version>.exe`. Para subir una nueva versión a GitHub Releases:

```powershell
gh release create vX.X.X dist_installer\RankingVentas_Setup_X.X.X.exe --title "vX.X.X" --notes "Descripción de los cambios"
```

## Estructura del proyecto

- `core.py` — lógica de negocio (lectura del Excel, reglas, histórico acumulado, generación del Excel de salida).
- `desktop_app.py` — interfaz de escritorio (Tkinter).
- `requirements.txt` — dependencias de ejecución (pandas, openpyxl).
- `installer/setup.iss` — script de Inno Setup para generar el instalador.

## Licencia

GPLv3. Ver [LICENSE](LICENSE).
