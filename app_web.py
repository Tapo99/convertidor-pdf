import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Convertidor Planilla Pro", page_icon="📊")

st.title("Convertidor de Planillas Profesional")
st.markdown("CONVERTIDOR DE PLANILLA POR CENTROS DE COSTOS EN EXCEL")

archivo_subido = st.file_uploader("Sube tu planilla PDF", type="pdf")

if archivo_subido:
    if st.button("🚀 Procesar y Dar Formato Contable"):
        datos_completos = []
        
        with pdfplumber.open(archivo_subido) as pdf:
            for pagina in pdf.pages:
                # Ajuste de sensibilidad para el formato de Caja de Crédito
                tabla = pagina.extract_table({
                    "vertical_strategy": "text", 
                    "horizontal_strategy": "text",
                    "snap_tolerance": 5,
                    "join_tolerance": 3,
                })
                
                if tabla:
                    for fila in tabla:
                        fila_limpia = [str(c).replace('\n', ' ').strip() if c else "" for c in fila]
                        texto_completo = " ".join(fila_limpia).upper()
                        
                        # FILTROS: Saltamos encabezados y totales
                        if not any(fila_limpia) or "PLANILLA" in texto_completo or "CORR." in texto_completo:
                            continue
                        if "TOTALES" in texto_completo or "CENTRO DE" in texto_completo:
                            continue
                        if len([c for c in fila_limpia if c]) < 5: # Ignora filas con poca info
                            continue
                            
                        datos_completos.append(fila_limpia)

        if datos_completos:
            # Columnas exactas de tu Excel de ejemplo
            columnas_finales = [
                "N°", "CODIGO", "EMPLEADO", "Días laborados", "Salario", 
                "Salario quincenal", "Horas extras", "Comisiones", "Otr", 
                "SALARIO DEVENGADO", "AFP", "ISSS", "RENTA", 
                "INSTITUCIONES FINANCIERAS", "Préstamos", "Otros", 
                "Total DE DESCUENTOS", "Líquido a recibir"
            ]
            
            df = pd.DataFrame(datos_completos)
            
            # Ajustamos el ancho si el PDF trajo columnas de más o de menos
            if df.shape[1] > len(columnas_finales):
                df = df.iloc[:, :len(columnas_finales)]
            df.columns = columnas_finales[:df.shape[1]]

            # --- LIMPIEZA DE DATOS ---
            def limpiar_monto(val, es_entero=False):
                if not val: return 0
                limpio = str(val).replace('$', '').replace(',', '').strip()
                try:
                    num = float(limpio)
                    return int(num) if es_entero else num
                except: return 0

            # Días laborados a entero
            if "Días laborados" in df.columns:
                df["Días laborados"] = df["Días laborados"].apply(lambda x: limpiar_monto(x, True))
            
            # Dinero a float (columnas desde Salario hasta Líquido)
            cols_dinero = df.columns[4:]
            for col in cols_dinero:
                df[col] = df[col].apply(lambda x: limpiar_monto(x, False))

            # --- GENERACIÓN DE EXCEL CON FORMATO PROFESIONAL ---
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Planilla_Limpia')
                
                workbook  = writer.book
                worksheet = writer.sheets['Planilla_Limpia']

                # Estilos
                format_contable = workbook.add_format({'num_format': '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', 'font_name': 'Arial', 'font_size': 10})
                format_entero = workbook.add_format({'num_format': '0', 'align': 'center', 'font_name': 'Arial', 'font_size': 10})
                format_texto = workbook.add_format({'font_name': 'Arial', 'font_size': 10})
                format_header = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1})

                # Aplicar estilo al encabezado
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, format_header)

                # Configurar anchos y formatos de columnas
                worksheet.set_column('A:B', 10, format_texto)   # N° y Código
                worksheet.set_column('C:C', 45, format_texto)   # Empleado
                worksheet.set_column('D:D', 12, format_entero)  # Días (Sin decimales)
                worksheet.set_column('E:R', 16, format_contable)# Columnas de dinero

            st.success("✅ ¡Conversión completada con éxito!")
            st.download_button(
                label="📥 Descargar Excel Formateado",
                data=output.getvalue(),
                file_name="Planilla_Excel_Contable.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No se detectaron datos de empleados. Verifica que el PDF no sea una imagen escaneada.")

