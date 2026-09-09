import streamlit as st
import pandas as pd
import os
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.chart import PieChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.trendline import Trendline

st.set_page_config(page_title="Dashboard de Ventas", layout="wide")

ARCHIVO_REGLAS = "reglas_parametrizacion.csv"
ARCHIVO_HISTORIAL = "historial_rankings.csv"

# --- UTILIDAD: FORMATO MES EN ESPAÑOL ---
def formatear_mes_es(fecha_str):
    if pd.isna(fecha_str) or str(fecha_str).lower() == 'nat':
        return "Mes_Desconocido"
    meses = {'01':'Enero', '02':'Febrero', '03':'Marzo', '04':'Abril', '05':'Mayo', '06':'Junio', 
             '07':'Julio', '08':'Agosto', '09':'Septiembre', '10':'Octubre', '11':'Noviembre', '12':'Diciembre'}
    try:
        y, m = str(fecha_str).split('-')
        return f"{meses[m]} {y}"
    except:
        return str(fecha_str)

# --- 1. INICIALIZAR MEMORIA LOCAL ---
if 'df_reglas' not in st.session_state:
    if os.path.exists(ARCHIVO_REGLAS):
        st.session_state.df_reglas = pd.read_csv(ARCHIVO_REGLAS).fillna('')
    else:
        st.session_state.df_reglas = pd.DataFrame({
            'Sucursal Origen (Opcional)': ['PAVON', 'PAVON', 'PAVON', 'PAVON', 'PAVON', ''],
            'Columna a Buscar': ['Vendedor', 'Usuario Creador', 'Localidad Documento', 'Usuario Creador', 'Localidad Documento', 'Usuario Creador'],
            'Contiene el Texto': ['Victor Bascuñan', 'Caniuqueo Veronica', '00014-', 'Roman Ayrton', '00020-', 'Aguilar Mónica'],
            'Asignar a Sucursal': ['Geo Mayorista', 'Dunia Mayorista', 'Mercado Libre', 'Mercado Libre', 'Mercado Libre', 'EXCLUIR'],
            'Asignar a Razón Social': ['Mnos Import SRL', 'Dunia Import SRL', 'Bekir 1 SRL', 'Bekir 1 SRL', 'Bekir 1 SRL', '']
        })

# --- PESTAÑAS APP ---
tab_dash, tab_conf = st.tabs(["📊 Ejecutar Reporte", "⚙️ Reglas y Razones Sociales"])

with tab_conf:
    st.header("⚙️ Configuración de Reglas")
    st.write("Agrega a qué Razón Social pertenece cada sucursal para el reporte agrupado.")
    df_editado = st.data_editor(
        st.session_state.df_reglas, num_rows="dynamic", use_container_width=True,
        column_config={
            "Sucursal Origen (Opcional)": st.column_config.TextColumn("Sucursal Origen"),
            "Columna a Buscar": st.column_config.SelectboxColumn("Buscar en", options=['Vendedor', 'Usuario Creador', 'Localidad Documento', 'Sucursal']),
            "Contiene el Texto": st.column_config.TextColumn("Texto Exacto"),
            "Asignar a Sucursal": st.column_config.TextColumn("Nueva Sucursal"),
            "Asignar a Razón Social": st.column_config.TextColumn("Nueva Razón Social", help="Deja en blanco para mantener la del ERP.")
        }
    )
    if st.button("💾 Guardar Configuración"):
        df_editado = df_editado.fillna('')
        df_editado.to_csv(ARCHIVO_REGLAS, index=False)
        st.session_state.df_reglas = df_editado
        st.success("Reglas guardadas.")

with tab_dash:
    st.title("📤 Generador de Reportes Ejecutivos")
    uploaded_file = st.file_uploader("Sube el archivo Excel de tu ERP (.xlsx)", type=["xlsx"])

    if uploaded_file is not None:
        try:
            # LECTURA Y FILTRADO INICIAL
            df = pd.read_excel(uploaded_file, sheet_name='Sheet2', header=1)
            df.columns = df.columns.str.strip().str.replace('\n', ' ')
            df = df.dropna(subset=['Sucursal']).copy()

            COL_TOTAL, COL_SUBTOTAL, COL_IVA = 'Total Total', 'Neto Gravado Neto Gravado', 'IVA IVA'
            for col in [COL_TOTAL, COL_SUBTOTAL, COL_IVA]:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            # FECHAS
            fechas_limpias = pd.to_datetime(df['Código Fecha Documento'], errors='coerce')
            df = df.loc[fechas_limpias.notna()].copy()
            df['Mes_Año'] = pd.to_datetime(df['Código Fecha Documento']).dt.to_period('M').astype(str)
            df['Mes_Format'] = df['Mes_Año'].apply(formatear_mes_es)
            
            # LIMPIEZA INTELIGENTE DE RAZÓN SOCIAL
            df['Razón Social'] = df['Razon Social Empresa - Division'].astype(str)
            df['Razón Social'] = df['Razón Social'].str.replace(r'\s*-\s*Empresa', '', regex=True, case=False)
            df['Razón Social'] = df['Razón Social'].str.replace(r'\s*-\s*Prueba', '', regex=True, case=False)
            df['Razón Social'] = df['Razón Social'].str.strip()
            df['Razón Social'] = df['Razón Social'].replace({
                'Bekir I SRL': 'Bekir 1 SRL',
                'Dunia Importer SRL': 'Dunia Import SRL'
            })

            # APLICAR REGLAS
            for _, row in st.session_state.df_reglas.iterrows():
                suc_ori, col, texto, nueva_suc, nueva_rs = map(lambda x: str(x).strip(), [row.get('Sucursal Origen (Opcional)',''), row['Columna a Buscar'], row['Contiene el Texto'], row['Asignar a Sucursal'], row.get('Asignar a Razón Social', '')])
                if col in df.columns and texto:
                    mask = df[col].astype(str).str.contains(texto, case=False, na=False, regex=False)
                    if suc_ori: mask &= (df['Sucursal'].astype(str).str.upper() == suc_ori.upper())
                    df.loc[mask, 'Sucursal'] = nueva_suc
                    if nueva_rs: df.loc[mask, 'Razón Social'] = nueva_rs

            # EXCLUIR Y REPASO FINAL DE DUNIA IMPORTER (Por si la regla metió un Importer)
            df = df[df['Sucursal'] != 'EXCLUIR']
            df['Razón Social'] = df['Razón Social'].str.replace('Dunia Importer SRL', 'Dunia Import SRL', regex=False)

            st.success("¡Datos procesados! Generando Excel...")

            # --- GENERACIÓN DE EXCEL ---
            output = io.BytesIO()
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            
            # --- HOJA 1: ANÁLISIS ---
            ws_graficos = wb.create_sheet("📈 Análisis")
            ws_graficos.sheet_view.showGridLines = False
            ws_data = wb.create_sheet("Data_Oculta"); ws_data.sheet_state = 'hidden'
            
            # Pivot Tables para Gráficos
            df_tot_mes = df.groupby(['Mes_Año', 'Mes_Format'])[COL_TOTAL].sum().reset_index().rename(columns={COL_TOTAL: 'Total Empresa'})
            df_rs_mes = df.groupby(['Mes_Format', 'Razón Social'])[COL_TOTAL].sum().unstack().fillna(0)
            df_loc_mes = df.groupby(['Mes_Format', 'Sucursal'])[COL_TOTAL].sum().unstack().fillna(0)
            
            df_charts = df_tot_mes.set_index('Mes_Format').drop(columns=['Mes_Año']).join(df_rs_mes).join(df_loc_mes).reset_index()
            ws_data.append(df_charts.columns.tolist())
            for _, r in df_charts.iterrows(): ws_data.append(r.tolist())
            
            # Función maestra para crear gráficos ejecutivos
            def dibujar_grafico(titulo, min_col, max_col, dest_cell, tipo="clustered", tendencia_unica=False):
                if max_col < min_col: return 
                c = BarChart()
                c.type = "col" 
                c.grouping = tipo # "clustered" (lado a lado) o "stacked" (apilado)
                if tipo == "stacked": c.overlap = 100
                
                c.title = titulo
                c.style = 13 # Estilo moderno
                c.width, c.height = 30, 14 # Más anchos para acomodar leyenda inferior
                c.y_axis.title, c.y_axis.number_format, c.x_axis.title = 'Ventas ($)', '#,##0', 'Mes'
                c.legend.position = 'b' # LEYENDA ABAJO
                
                data = Reference(ws_data, min_col=min_col, min_row=1, max_col=max_col, max_row=ws_data.max_row)
                cats = Reference(ws_data, min_col=1, min_row=2, max_row=ws_data.max_row)
                c.add_data(data, titles_from_data=True)
                c.set_categories(cats)
                
                # Paleta de colores forzada
                colores = ["1F4E78", "ED7D31", "7030A0", "C00000", "00B0F0", "FFC000", "92D050", "002060", "FF00FF"]
                for i, s in enumerate(c.series):
                    s.graphicalProperties.solidFill = colores[i % len(colores)]
                
                # Línea de tendencia única en el gráfico de empresas
                if tendencia_unica and len(c.series) > 0:
                    c.series[0].trendline = Trendline(trendlineType='movingAvg', period=2)
                        
                ws_graficos.add_chart(c, dest_cell)

            ws_graficos['B2'] = "ANÁLISIS DE CRECIMIENTO TEMPORAL"; ws_graficos['B2'].font = Font(size=20, bold=True, color="1F4E78")
            
            # Gráfico 1: Empresa Total
            dibujar_grafico("1. Crecimiento Total Empresa", 2, 2, "B4", tipo="clustered")
            # Gráfico 2: Razones Sociales (Agrupadas, Colores sólidos y UNA Sola línea de Promedio Móvil)
            dibujar_grafico("2. Crecimiento por Razón Social (Colores separados)", 3, 2 + len(df_rs_mes.columns), "Q4", tipo="clustered", tendencia_unica=True)
            # Gráfico 3: Todas las sucursales (BARRAS APILADAS - Stacked para no ver minibarras)
            dibujar_grafico("3. Crecimiento de TODAS las Sucursales y Canales (Composición por Mes)", 3 + len(df_rs_mes.columns), ws_data.max_column, "B32", tipo="stacked")

            # --- HOJAS MENSUALES ---
            historial = pd.read_csv(ARCHIVO_HISTORIAL) if os.path.exists(ARCHIVO_HISTORIAL) else pd.DataFrame(columns=['Mes', 'Sucursal', 'Rank'])
            meses_ordenados = sorted(df['Mes_Año'].unique())
            
            for mes in meses_ordenados:
                df_mes = df[df['Mes_Año'] == mes]
                nombre_solapa = df_mes['Mes_Format'].iloc[0]
                ws_mes = wb.create_sheet(nombre_solapa)
                ws_mes.sheet_view.showGridLines = False

                rs_map = df_mes.groupby('Sucursal')['Razón Social'].agg(lambda x: x.value_counts().index[0] if len(x)>0 else '')
                rk_locales = df_mes.groupby('Sucursal')[[COL_SUBTOTAL, COL_IVA, COL_TOTAL]].sum().reset_index()
                rk_locales['Razón Social'] = rk_locales['Sucursal'].map(rs_map)
                
                rk_locales = rk_locales.sort_values(by=COL_TOTAL, ascending=False).reset_index(drop=True)
                rk_locales.index += 1; rk_locales = rk_locales.reset_index().rename(columns={'index': 'Rank Actual'})

                rk_empresas = df_mes.groupby('Razón Social')[[COL_SUBTOTAL, COL_IVA, COL_TOTAL]].sum().reset_index()
                rk_empresas = rk_empresas.sort_values(by=COL_TOTAL, ascending=False).reset_index(drop=True)

                mes_dt = pd.to_datetime(mes + '-01')
                mes_ant_str = (mes_dt - pd.DateOffset(months=1)).strftime('%Y-%m')
                
                hist_ant = historial[historial['Mes'] == mes_ant_str]
                dict_rank_ant = dict(zip(hist_ant['Sucursal'], hist_ant['Rank']))
                rk_locales['Rank Anterior'] = rk_locales['Sucursal'].map(dict_rank_ant).fillna(rk_locales['Rank Actual'])
                rk_locales['Diff Num'] = rk_locales['Rank Anterior'] - rk_locales['Rank Actual']
                rk_locales['Diff Rank'] = rk_locales['Diff Num'].apply(lambda x: f"+{int(x)}" if x > 0 else (f"{int(x)}" if x < 0 else "="))

                n_hist = rk_locales[['Sucursal', 'Rank Actual']].copy().rename(columns={'Rank Actual': 'Rank'})
                n_hist['Mes'] = mes
                historial = pd.concat([historial[historial['Mes'] != mes], n_hist])

                # DISEÑO 
                F_HEAD, BG_DARK = Font(color="FFFFFF", bold=True), PatternFill(start_color="0B3D59", end_color="0B3D59", fill_type="solid")
                C_AL, L_AL, R_AL = Alignment(horizontal="center", vertical="center"), Alignment(horizontal="left", vertical="center"), Alignment(horizontal="right", vertical="center")
                THIN_B = Border(left=Side(style='thin', color="000000"), right=Side(style='thin', color="000000"), top=Side(style='thin', color="000000"), bottom=Side(style='thin', color="000000"))
                
                ws_mes.merge_cells('A2:K2')
                ws_mes['A2'] = f"TOP POR LOCALES - {nombre_solapa.upper()}"; ws_mes['A2'].fill, ws_mes['A2'].font, ws_mes['A2'].alignment = BG_DARK, Font(color="FFFFFF", bold=True, size=12), C_AL

                headers = ['Diff Rank', 'Rank', 'Anterior', 'Razón Social', 'Local', 'ARS', 'Subtotal', 'ARS', 'IVA', 'ARS', 'Total']
                for c_i, h in enumerate(headers, 1):
                    c = ws_mes.cell(row=3, column=c_i, value=h)
                    c.fill, c.font, c.alignment, c.border = BG_DARK, F_HEAD, C_AL, THIN_B

                rn = 4
                for _, r in rk_locales.iterrows():
                    c_d = ws_mes.cell(row=rn, column=1, value=r['Diff Rank'])
                    c_d.alignment = C_AL
                    c_d.font = Font(color="00B050", bold=True) if '+' in str(r['Diff Rank']) else (Font(color="C00000", bold=True) if '-' in str(r['Diff Rank']) else Font(bold=True))
                    
                    ws_mes.cell(row=rn, column=2, value=r['Rank Actual']).alignment = C_AL
                    ws_mes.cell(row=rn, column=3, value=r['Rank Anterior']).alignment = C_AL
                    ws_mes.cell(row=rn, column=4, value=str(r['Razón Social']).split()[0] if str(r['Razón Social']) != 'nan' else 'N/A').alignment = C_AL
                    ws_mes.cell(row=rn, column=5, value=r['Sucursal']).alignment = L_AL
                    
                    for i, v in enumerate([r[COL_SUBTOTAL], r[COL_IVA], r[COL_TOTAL]]):
                        ws_mes.cell(row=rn, column=6+(i*2), value="ARS").alignment = L_AL
                        cx = ws_mes.cell(row=rn, column=7+(i*2), value=v)
                        cx.number_format, cx.alignment = '#,##0.00', R_AL
                    for col in range(1, 12): ws_mes.cell(row=rn, column=col).border = THIN_B
                    rn += 1

                ws_mes.merge_cells(start_row=rn, start_column=1, end_row=rn, end_column=5)
                ws_mes.cell(row=rn, column=1, value="TOTAL").alignment, ws_mes.cell(row=rn, column=1).font = C_AL, Font(bold=True)
                for i, v in enumerate([rk_locales[COL_SUBTOTAL].sum(), rk_locales[COL_IVA].sum(), rk_locales[COL_TOTAL].sum()]):
                    ws_mes.cell(row=rn, column=6+(i*2), value="ARS").alignment = L_AL
                    cx = ws_mes.cell(row=rn, column=7+(i*2), value=v)
                    cx.number_format, cx.font = '#,##0.00', Font(bold=True)
                for col in range(1, 12): ws_mes.cell(row=rn, column=col).border = THIN_B

                color_scale = ColorScaleRule(start_type='min', start_color='F8696B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='63BE7B')
                for c_l in ['G', 'I', 'K']: ws_mes.conditional_formatting.add(f"{c_l}4:{c_l}{rn-1}", color_scale)

                # TOP EMPRESAS
                rn += 3
                ws_mes.merge_cells(start_row=rn, start_column=4, end_row=rn, end_column=11)
                ws_mes.cell(row=rn, column=4, value="TOP POR EMPRESAS").fill, ws_mes.cell(row=rn, column=4).font, ws_mes.cell(row=rn, column=4).alignment = BG_DARK, F_HEAD, C_AL
                rn += 1
                ws_mes.merge_cells(start_row=rn, start_column=4, end_row=rn, end_column=5)
                ws_mes.cell(row=rn, column=4, value="RAZON SOCIAL").fill, ws_mes.cell(row=rn, column=4).font, ws_mes.cell(row=rn, column=4).alignment = BG_DARK, F_HEAD, C_AL
                for idx, text in enumerate(["SUBTOTAL", "IVA", "TOTAL"]):
                    ws_mes.merge_cells(start_row=rn, start_column=6+(idx*2), end_row=rn, end_column=7+(idx*2))
                    c = ws_mes.cell(row=rn, column=6+(idx*2), value=text)
                    c.fill, c.font, c.alignment = BG_DARK, F_HEAD, C_AL

                rn += 1
                for _, r in rk_empresas.iterrows():
                    ws_mes.merge_cells(start_row=rn, start_column=4, end_row=rn, end_column=5)
                    ws_mes.cell(row=rn, column=4, value=r['Razón Social']).alignment = C_AL
                    for i, v in enumerate([r[COL_SUBTOTAL], r[COL_IVA], r[COL_TOTAL]]):
                        ws_mes.cell(row=rn, column=6+(i*2), value="ARS").alignment = L_AL
                        cx = ws_mes.cell(row=rn, column=7+(i*2), value=v)
                        cx.number_format = '#,##0.00'
                    for col in range(4, 12): ws_mes.cell(row=rn, column=col).border = THIN_B
                    rn += 1
                
                ws_mes.merge_cells(start_row=rn, start_column=4, end_row=rn, end_column=5)
                ws_mes.cell(row=rn, column=4, value="TOTAL").alignment, ws_mes.cell(row=rn, column=4).font = C_AL, Font(bold=True)
                for i, v in enumerate([rk_empresas[COL_SUBTOTAL].sum(), rk_empresas[COL_IVA].sum(), rk_empresas[COL_TOTAL].sum()]):
                    ws_mes.cell(row=rn, column=6+(i*2), value="ARS").alignment = L_AL
                    cx = ws_mes.cell(row=rn, column=7+(i*2), value=v)
                    cx.number_format, cx.font = '#,##0.00', Font(bold=True)
                for col in range(4, 12): ws_mes.cell(row=rn, column=col).border = THIN_B

                ws_mes.column_dimensions['A'].width, ws_mes.column_dimensions['B'].width, ws_mes.column_dimensions['C'].width = 8, 8, 8
                ws_mes.column_dimensions['D'].width, ws_mes.column_dimensions['E'].width = 15, 25
                for c in ['F', 'H', 'J']: ws_mes.column_dimensions[c].width = 5
                for c in ['G', 'I', 'K']: ws_mes.column_dimensions[c].width = 17

                # TORTAS
                pie_loc = PieChart()
                pie_loc.add_data(Reference(ws_mes, min_col=11, min_row=3, max_row=3+len(rk_locales)), titles_from_data=True)
                pie_loc.set_categories(Reference(ws_mes, min_col=5, min_row=4, max_row=3+len(rk_locales)))
                pie_loc.title = "LOCALES"; pie_loc.dataLabels = DataLabelList(); pie_loc.dataLabels.showPercent, pie_loc.dataLabels.showCatName = True, True
                pie_loc.width, pie_loc.height = 15, 12
                ws_mes.add_chart(pie_loc, "M3")

                pie_emp = PieChart()
                start_emp_row = rn - len(rk_empresas)
                pie_emp.add_data(Reference(ws_mes, min_col=11, min_row=start_emp_row-1, max_row=rn-1), titles_from_data=True)
                pie_emp.set_categories(Reference(ws_mes, min_col=4, min_row=start_emp_row, max_row=rn-1))
                pie_emp.title = "RAZON SOCIAL"; pie_emp.dataLabels = DataLabelList(); pie_emp.dataLabels.showPercent, pie_emp.dataLabels.showCatName = True, True
                pie_emp.width, pie_emp.height = 15, 12
                ws_mes.add_chart(pie_emp, "U3")

            historial.to_csv(ARCHIVO_HISTORIAL, index=False)
            wb.save(output)
            
            nom_arch = f"Reporte_{formatear_mes_es(meses_ordenados[0]).replace(' ', '_')}.xlsx" if len(meses_ordenados)==1 else f"Reporte_Anual_{formatear_mes_es(meses_ordenados[0]).replace(' ', '_')}_a_{formatear_mes_es(meses_ordenados[-1]).replace(' ', '_')}.xlsx"
            st.download_button("📥 Descargar Excel Ejecutivo", data=output.getvalue(), file_name=nom_arch)
            
        except Exception as e:
            st.error(f"Error procesando el archivo: {e}")