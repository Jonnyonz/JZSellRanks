"""Pruebas de core.py con un Excel sintético (sin datos reales del negocio).

Correr desde la raíz del repo:  python -m unittest discover -s tests -v
"""
import os
import sys
import tempfile
import unittest

import openpyxl
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import core  # noqa: E402

ENCABEZADOS = [
    'Sucursal', 'Código Fecha Documento', 'Total Total', 'Neto Gravado Neto Gravado', 'IVA IVA',
    'Razon Social Empresa - Division', 'Vendedor', 'Usuario Creador', 'Localidad Documento',
]


def escribir_excel(ruta, filas):
    """Arma un .xlsx con el formato del ERP: hoja Sheet2, encabezados en la fila 2."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Sheet2')
    ws.append(['Reporte de ventas'])
    ws.append(ENCABEZADOS)
    for f in filas:
        ws.append(f)
    wb.save(ruta)


def fila(sucursal, rs, total=121.0, fecha='2026-08-15', vendedor='Vendedor A', usuario='usuario.a', localidad='00001-Loc'):
    return [sucursal, fecha, total, round(total / 1.21, 2), round(total - total / 1.21, 2), rs, vendedor, usuario, localidad]


def reglas(*filas):
    return pd.DataFrame(list(filas), columns=core.COLUMNAS_CSV_REGLAS)


class TestCore(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.excel = os.path.join(self.tmp.name, 'ventas.xlsx')
        self.acumulado = os.path.join(self.tmp.name, 'datos_acumulados.csv')

    def tearDown(self):
        self.tmp.cleanup()

    def generar(self, filas, df_reglas=None):
        escribir_excel(self.excel, filas)
        if df_reglas is None:
            df_reglas = reglas()
        return core.generar_reporte(self.excel, df_reglas, archivo_acumulado=self.acumulado, log=lambda *_: None)

    def test_historico_no_se_escribe_hasta_guardar_acumulado(self):
        wb, nombre, acumulado = self.generar([fila('CENTRO', 'Alfa SRL - Empresa')])
        self.assertEqual(nombre, 'Reporte_Agosto_2026.xlsx')
        self.assertFalse(os.path.exists(self.acumulado), 'generar_reporte no debe tocar el histórico')
        core.guardar_acumulado(acumulado, self.acumulado)
        self.assertTrue(os.path.exists(self.acumulado))
        self.assertEqual(list(pd.read_csv(self.acumulado).columns), core.COLUMNAS_ACUMULADO)

    def test_regla_por_division_prueba_distingue_la_misma_sucursal(self):
        filas = [
            fila('CENTRO', 'Alfa SRL - Empresa', total=1000),
            fila('CENTRO', 'Alfa SRL - Prueba', total=100),
        ]
        df_reglas = reglas(['', core.COL_RAZON_SOCIAL_ERP, 'Prueba', 'CENTRO PRUEBA', 'Beta SA'])
        _, _, acumulado = self.generar(filas, df_reglas)
        por_sucursal = acumulado.set_index('Sucursal')
        self.assertEqual(set(por_sucursal.index), {'CENTRO', 'CENTRO PRUEBA'})
        self.assertAlmostEqual(por_sucursal.loc['CENTRO', core.COL_TOTAL], 1000)
        self.assertAlmostEqual(por_sucursal.loc['CENTRO PRUEBA', core.COL_TOTAL], 100)
        self.assertEqual(por_sucursal.loc['CENTRO PRUEBA', 'Razón Social'], 'Beta SA')
        self.assertEqual(por_sucursal.loc['CENTRO', 'Razón Social'], 'Alfa SRL')

    def test_regla_sin_sucursal_destino_conserva_la_sucursal(self):
        df_reglas = reglas(['', core.COL_RAZON_SOCIAL_ERP, 'Prueba', '', 'Beta SA'])
        _, _, acumulado = self.generar([fila('CENTRO', 'Alfa SRL - Prueba')], df_reglas)
        self.assertEqual(acumulado['Sucursal'].tolist(), ['CENTRO'])
        self.assertEqual(acumulado['Razón Social'].tolist(), ['Beta SA'])

    def test_razon_social_vacia_no_revienta(self):
        _, _, acumulado = self.generar([fila('CENTRO', None)])
        self.assertEqual(len(acumulado), 1)

    def test_detectar_valores_incluye_division(self):
        escribir_excel(self.excel, [fila('CENTRO', 'Alfa SRL - Empresa'), fila('CENTRO', 'Alfa SRL - Prueba')])
        valores = core.detectar_valores(self.excel)
        self.assertEqual(valores[core.COL_RAZON_SOCIAL_ERP], ['Alfa SRL - Empresa', 'Alfa SRL - Prueba'])
        self.assertEqual(valores['Razón Social'], ['Alfa SRL'])

    def test_excluir_y_celda_segura(self):
        # openpyxl guarda '=...' como fórmula, así que el texto peligroso de prueba arranca con '+'.
        df_reglas = reglas(['', 'Vendedor', 'Vendedor A', 'EXCLUIR', ''])
        filas = [fila('CENTRO', 'Alfa SRL - Empresa'), fila('+CMD|calc', 'Alfa SRL - Empresa', vendedor='Otro')]
        wb, _, acumulado = self.generar(filas, df_reglas)
        self.assertEqual(acumulado['Sucursal'].tolist(), ['+CMD|calc'])
        celdas = [c.value for row in wb['Agosto 2026'].iter_rows() for c in row if isinstance(c.value, str)]
        self.assertIn("'+CMD|calc", celdas)
        self.assertNotIn('+CMD|calc', celdas)

    def test_sucursal_con_el_mismo_nombre_que_una_razon_social(self):
        # Antes el join de los datos de gráficos fallaba por columnas repetidas.
        filas = [fila('Alfa SRL', 'Alfa SRL - Empresa'), fila('CENTRO', 'Alfa SRL - Empresa')]
        wb, _, acumulado = self.generar(filas)
        self.assertEqual(len(acumulado), 2)
        encabezados = [c.value for c in wb['Data_Oculta'][1]]
        self.assertEqual(len(encabezados), len(set(encabezados)), 'encabezados repetidos en Data_Oculta')

    def test_reglas_por_defecto_son_inertes(self):
        _, _, acumulado = self.generar([fila('CENTRO', 'Alfa SRL - Prueba')], core.reglas_por_defecto())
        self.assertEqual(acumulado['Sucursal'].tolist(), ['CENTRO'])
        self.assertEqual(acumulado['Razón Social'].tolist(), ['Alfa SRL'])

    def test_nombre_perfil_valido(self):
        for ok in ['Default', 'Mayoristas 2026', 'con.sucursales', 'nul_pass']:
            self.assertTrue(core.nombre_perfil_valido(ok), ok)
        for malo in ['', '  ', 'a/b', 'a:b', 'CON', 'con', 'Nul.csv', 'COM1', 'lpt9', 'final.', 'final ']:
            self.assertFalse(core.nombre_perfil_valido(malo), malo)

    def test_todo_excluido_avisa_en_vez_de_reventar(self):
        df_reglas = reglas(['', 'Vendedor', 'Vendedor A', 'EXCLUIR', ''])
        with self.assertRaises(ValueError):
            self.generar([fila('CENTRO', 'Alfa SRL - Empresa')], df_reglas)


if __name__ == '__main__':
    unittest.main()
