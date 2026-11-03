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
    'Dirección Tipo', 'CUIT',
]


def cliente(nombre, cuit=20123456789, codigo='CLI00001'):
    """Fila de cliente del export del ERP: sin sucursal; razón social, código y CUIT en las columnas
    compartidas (Razon Social / Código / CUIT). Los comprobantes que siguen son de este cliente."""
    return [None, codigo, 1000.0, 826.45, 173.55, ' ' + nombre, None, None, 'Ciudad', 'Calle 123', cuit]


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


_contador = [0]


def fila(sucursal, rs, total=121.0, fecha='2026-08-15', vendedor='Vendedor A', usuario='usuario.a',
         documento=None, tipo='FCV'):
    """Fila de comprobante. 'Localidad Documento' lleva el número (letra + punto - número) y
    'Dirección Tipo' el tipo (FCV factura, NCV nota de crédito)."""
    if documento is None:
        _contador[0] += 1
        documento = f'A00001-{_contador[0]:08d}'
    return [sucursal, fecha, total, round(total / 1.21, 2), round(total - total / 1.21, 2), rs, vendedor, usuario,
            documento, tipo, None]


def reglas(*filas):
    return pd.DataFrame(list(filas), columns=core.COLUMNAS_CSV_REGLAS)


class TestCore(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.excel = os.path.join(self.tmp.name, 'ventas.xlsx')
        self.acumulado = os.path.join(self.tmp.name, 'datos_acumulados.csv')

    def tearDown(self):
        self.tmp.cleanup()

    def generar(self, filas, df_reglas=None, archivo_exclusiones=None):
        escribir_excel(self.excel, filas)
        if df_reglas is None:
            df_reglas = reglas()
        return core.generar_reporte(self.excel, df_reglas, archivo_acumulado=self.acumulado, log=lambda *_: None,
                                    archivo_exclusiones=archivo_exclusiones)

    def test_leer_erp_copia_el_cliente_y_desarma_el_documento(self):
        escribir_excel(self.excel, [
            cliente('Cliente Uno SA', cuit=20111111111),
            fila('CENTRO', 'Alfa SRL - Empresa', documento='A00005-00012345'),
            fila('CENTRO', 'Alfa SRL - Empresa', documento='B00005-00000007', tipo='NCV', total=-50),
            cliente('Cliente Dos SRL', cuit=27222222222),
            fila('NORTE', 'Alfa SRL - Empresa', documento='A00012-00000001'),
        ])
        docs = core.leer_erp(self.excel)
        self.assertEqual(len(docs), 3, 'solo filas de comprobante')
        self.assertEqual(docs['Cliente'].tolist(), ['Cliente Uno SA', 'Cliente Uno SA', 'Cliente Dos SRL'])
        self.assertEqual(docs['CUIT Cliente'].tolist(), ['20111111111', '20111111111', '27222222222'])
        self.assertEqual(docs['Tipo'].tolist(), ['FCV', 'NCV', 'FCV'])
        self.assertEqual(docs['Punto'].tolist(), ['00005', '00005', '00012'])
        self.assertEqual(docs['Número'].tolist(), ['00012345', '00000007', '00000001'])
        self.assertEqual(docs['Clave'].iloc[0], 'FCV|A00005-00012345')

    def test_buscar_documentos(self):
        escribir_excel(self.excel, [
            cliente('Cliente Uno SA'), fila('CENTRO', 'Alfa SRL - Empresa', documento='A00005-00012345'),
            fila('CENTRO', 'Alfa SRL - Empresa', documento='B00005-00000007', tipo='NCV'),
            cliente('Otro Cliente SRL'), fila('NORTE', 'Alfa SRL - Empresa', documento='A00012-00000001'),
        ])
        docs = core.leer_erp(self.excel)
        self.assertEqual(len(core.buscar_documentos(docs, cliente='cliente uno')), 2)
        self.assertEqual(len(core.buscar_documentos(docs, cliente='cliente')), 3, 'contiene, sin mayúsculas')
        self.assertEqual(core.buscar_documentos(docs, punto='5', numero='7')['Documento'].tolist(), ['B00005-00000007'])
        self.assertEqual(core.buscar_documentos(docs, punto='00012')['Documento'].tolist(), ['A00012-00000001'])
        self.assertEqual(len(core.buscar_documentos(docs, tipo='NCV')), 1)
        self.assertEqual(len(core.buscar_documentos(docs)), 3, 'sin filtros devuelve todo')

    def test_exclusiones_se_aplican_y_quedan_en_hoja_excluidas(self):
        archivo_ex = os.path.join(self.tmp.name, 'facturas_excluidas.csv')
        filas = [
            cliente('Cliente Uno SA'),
            fila('CENTRO', 'Alfa SRL - Empresa', documento='A00005-00012345', total=1000),
            fila('CENTRO', 'Alfa SRL - Empresa', documento='A00005-00012346', total=100),
        ]
        escribir_excel(self.excel, filas)
        docs = core.leer_erp(self.excel)
        excl = core.filas_a_exclusiones(core.buscar_documentos(docs, numero='12345'), motivo='Cliente en disputa')
        self.assertEqual(excl['Clave'].tolist(), ['FCV|A00005-00012345'])
        core.guardar_exclusiones(excl, archivo_ex)
        self.assertEqual(core.cargar_exclusiones(archivo_ex)['Motivo'].tolist(), ['Cliente en disputa'])

        wb, _, acumulado = self.generar(filas, archivo_exclusiones=archivo_ex)
        self.assertAlmostEqual(acumulado[core.COL_TOTAL].sum(), 100, msg='la factura excluida no suma')
        self.assertIn('Excluidas', wb.sheetnames)
        ws = wb['Excluidas']
        self.assertEqual([c.value for c in ws[2]][:2], ['Tipo', 'Documento'])
        self.assertEqual(ws.cell(row=3, column=2).value, 'A00005-00012345')
        self.assertEqual(ws.cell(row=3, column=7).value, 'Cliente en disputa')

        # Sin exclusiones no hay hoja ni filtro.
        wb2, _, acumulado2 = self.generar(filas, archivo_exclusiones=os.path.join(self.tmp.name, 'no_existe.csv'))
        self.assertAlmostEqual(acumulado2[core.COL_TOTAL].sum(), 1100)
        self.assertNotIn('Excluidas', wb2.sheetnames)

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
        for ok in ['Default', 'Mayoristas 2026', 'consucursales', 'nul_pass', 'COM10']:
            self.assertTrue(core.nombre_perfil_valido(ok), ok)
        # 'con.sucursales' también es reservado: Windows ignora la extensión para CON, NUL, etc.
        for malo in ['', '  ', 'a/b', 'a:b', 'CON', 'con', 'con.sucursales', 'Nul.csv', 'COM1', 'lpt9', 'final.', 'final ']:
            self.assertFalse(core.nombre_perfil_valido(malo), malo)

    def test_todo_excluido_avisa_en_vez_de_reventar(self):
        df_reglas = reglas(['', 'Vendedor', 'Vendedor A', 'EXCLUIR', ''])
        with self.assertRaises(ValueError):
            self.generar([fila('CENTRO', 'Alfa SRL - Empresa')], df_reglas)


if __name__ == '__main__':
    unittest.main()
