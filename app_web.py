import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Convertidor Pro", page_icon="mi_logo.png")
st.image("mi_logo.png", width=200) 
st.title("Convertidor de Planillas")

archivo_subido = st.file_uploader("Sube tu planilla PDF", type="pdf")

def limpiar_y_separar(texto):
    """Separa el código del empleado del nombre cuando vienen pegados"""
    if not texto: return "", ""
    # Busca un patrón de código al final del texto (ej: JAPP1 o D46V11U)
    match = re.search(r'([A-Z0-9]+)\s*$', str(texto).strip())
    if match:
        codigo = match.group(1)
        nombre = str(texto).strip()[:match.start()].strip()
        return codigo, nombre
    return "", str(texto).strip()

if archivo_subido:
    if st.button("🚀 Generar Excel Profesional"):
        try:
            todas_las_filas = []
            with pdfplumber.open(archivo_subido) as pdf:
                for pagina in pdf.pages:
                    tabla = pagina.extract_table({
                        "vertical_strategy": "text", 
                        "horizontal_strategy": "lines"
                    })
                    if tabla:
                        for fila in tabla:
                            # Filtramos filas de "AGENCIA", "TOTALES" o vacías
                            texto_fila = " ".join([str(x) for x in fila if x])
                            if any(palabra in texto_fila.upper() for palabra in ["AGENCIA", "TOTALES", "CUENTA", "FECHA"]):
                                continue
                            if len([x for x in fila if x]) < 5: # Filtra ruidos
                                continue
                            todas_las_filas.append(fila)

            if todas_las_filas:
                df = pd.DataFrame(todas_las_filas)
                
                # Procesar Columna A y B (Separar Código y Nombre)
                # Basado en la estructura del PDF donde el nombre y código vienen en la col 0
                datos_empleados = df[0].apply(limpiar_y_separar)
                df.insert(0, 'Codigo Emp', [x[0] for x in datos_empleados])
                df[0] = [x[1] for x in datos_empleados]
                
                # Reorganizar y Nombrar Columnas según tu requerimiento
                columnas_finales = [
                    'Corr.', 'Codigo Emp', 'Nombre Empleado', 'Días Laborados', 
                    'Salario Mensual', 'Salario Quincenal', 'Horas Extra', 'Festivo', 
                    'Comisiones', 'Vacaciones', 'Otros Ingresos', 'Salario Devengado', 
                    'AFP', 'ISSS', 'Renta', 'Inst. Financieras', 'Préstamos', 
                    'Otros Desc.', 'Total Desc.', 'Líquido a Recibir'
                ]
                
                # Ajustamos el número de columnas dinámicamente
                df = df.iloc[:, :len(columnas_finales)]
                df.columns = columnas_finales

                # Convertir a números para que sean "contables"
                cols_moneda = df.columns[4:] # Desde Salario Mensual en adelante
                for col in cols_moneda:
                    df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[^\d.]', '', regex=True), errors='coerce').fillna(0)

                # Agregar Fila de Totales al final
                suma_totales = df[cols_moneda].sum()
                fila_total = pd.Series([''] * len(df.columns), index=df.columns)
                fila_total['Nombre Empleado'] = 'TOTAL GENERAL'
                for col in cols_moneda:
                    fila_total[col] = suma_totales[col]
                
                df = pd.concat([df, pd.DataFrame([fila_total])], ignore_index=True)

                # Crear el Excel con formato profesional
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Planilla')
                    workbook  = writer.book
                    worksheet = writer.sheets['Planilla']
                    
                    # Formato contable ($ #,##0.00)
                    formato_moneda = workbook.add_format({'num_format': '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'})
                    
                    # Aplicar formato a las columnas numéricas
                    for i, col in enumerate(df.columns):
                        if i >= 4:
                            worksheet.set_column(i, i, 15, formato_moneda)
                        else:
                            worksheet.set_column(i, i, 12)

                st.success("¡Excel generado con éxito y totales calculados!")
                st.download_button(
                    label="📥 Descargar Excel Contable",
                    data=output.getvalue(),
                    file_name="planilla_contable_final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Error detallado: {e}")
