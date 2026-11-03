# Cambios

Formato: una entrada por versión publicada, la más nueva arriba. La versión es la de
`installer/setup.iss` (`MyAppVersion`) y del tag de la release.

## 1.1.2 - 2026-11-03

### Nuevo
- Ícono propio (barras de ranking sobre azul) en la ventana, la barra de tareas, el `.exe`,
  el instalador y el desinstalador. Se genera con `installer/hacer_icono.py` (solo biblioteca
  estándar).

### Arreglos
- El botón dice "Generar reporte", sin el glifo.
- Los nombres de perfil reservados de Windows (`CON`, `NUL`, `COM1`, `LPT1`..., con o sin
  extensión) y los que terminan en punto o espacio se rechazan al crear o guardar un perfil.
- El tema de la interfaz se aplica sin `except/pass`; docstrings sin referencias a la versión
  web vieja.

## 1.1.1 - 2026-11-03

### Arreglos
- Las reglas se leen en el hilo principal antes de lanzar la generación (antes el hilo de
  trabajo tocaba variables de la interfaz).
- Las reglas que se crean en el primer uso (perfil `Default`) son ejemplos genéricos que no
  coinciden con datos reales; muestran los tres casos típicos. Los perfiles existentes no
  cambian.
- Si una sucursal se llama igual que una razón social, el reporte ya no falla al armar los
  datos de los gráficos.

## 1.1.0 - 2026-11-03

### Nuevo
- Las reglas pueden buscar en `Razon Social Empresa - Division` (la razón social tal como
  viene del ERP, con la división). Sirve para separar las ventas "Empresa" de las "Prueba" de
  una misma sucursal: por ejemplo, *contiene "Prueba" → sucursal "CENTRO PRUEBA"*. El
  autocompletado ofrece los valores reales con división.
- Una regla con "Nueva Sucursal" vacía deja la sucursal como está y solo reasigna la razón
  social (antes la dejaba vacía).
- Pruebas automáticas de `core.py` (`tests/test_core.py`, `unittest`, Excel sintético).

### Arreglos
- El cuadro de error al generar el reporte nunca aparecía (`NameError`: la excepción ya no
  existía cuando corría el aviso). Lo mismo pasaba con la etiqueta de autodetección.
- El histórico acumulado se guarda recién cuando el Excel quedó escrito. Si la generación
  falla o se cancela el guardado, `datos_acumulados.csv` no cambia.
- Si no se puede guardar el reporte (típico: el archivo está abierto en Excel), aparece un
  aviso con la ruta y el histórico no se toca. Antes el error se perdía.
- Si todas las ventas quedan excluidas (reglas `EXCLUIR`, fechas inválidas, sin sucursal), se
  avisa con un mensaje claro en vez de fallar.
- Una razón social vacía ya no hace fallar el resumen del mes ni la hoja del reporte
  (se muestra `N/A`).
- La columna de la tabla de reglas dice "Contiene el texto" (antes "Texto Exacto", que era
  engañoso: la comparación siempre fue "contiene").

## 1.0.1 - 2026-09-29

- Neutraliza la inyección de fórmulas en el Excel de salida (`celda_segura`).
- Instalador Inno Setup con `SHA256SUMS.txt` en la release.

## 1.0.0 - 2026-09-29

- Primera versión de escritorio (Tkinter): reglas por perfil, histórico acumulado mes a mes,
  reporte con rankings y gráficos.
