import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Convertidor Contable Pro", page_icon="mi_logo.png", layout="wide")
st.image("mi_logo.png", width=180) 
st.title("Convertidor de Planillas Profesional")

archivo_subido = st.file_uploader("Sube tu planilla PDF", type="pdf")

def limpiar_valor(valor):
    if not valor: return 0.0
    # Limpia símbolos de dólar, comas y espacios para dejar solo el número
    limpio = re.sub(r'[^\d.]', '', str(valor).replace(',', ''))
    try:
        return float(limpio)
    except:
        return 0.0

if archivo_subido:
    if st.button("🚀 Generar Excel Ordenado"):
        try:
            datos_finales = []
            
            with pdfplumber.open(archivo_subido) as pdf:
                for pagina in pdf.pages:
                    # Extraemos las palabras con sus coordenadas para no mezclarlas
                    words = pagina.extract_words()
                    
                    # Extraer tabla con ajustes de precisión para evitar mezclas
                    tabla = pagina.extract_table({
                        "vertical_strategy": "text", 
                        "horizontal_strategy": "lines",
                        "snap_tolerance": 3,
                    })
                    
                    if not tabla: continue
                    
                    for fila in tabla:
                        # Unimos la fila para analizar si es basura
                        linea_texto = " ".join([str(x) for x in fila if x]).upper()
                        
                        # Filtros solicitados: quitar agencias, encabezados repetidos y totales intermedios
                        if any(x in linea_texto for x in ["AGENCIA", "TOTALES", "CORR.", "NOMBRE EMPLEADO", "CUENTA"]):
                            continue
                        
                        # Si la fila tiene muy pocos datos, la saltamos (ruido)
                        if len([x for x in fila if x]) < 8:
                            continue

                        # Procesamos la primera celda que suele traer Código + Nombre pegado
                        primera_celda = str(fila[0]).strip() if fila[0] else ""
                        
                        # Separación inteligente: El código suele ser la primera palabra corta (letras y números)
                        partes = primera_celda.split(maxsplit=1)
                        codigo = partes[0] if len(partes) > 1 else ""
                        nombre = partes[1] if len(partes) > 1 else primera_celda

                        # Construimos la fila limpia alineada a tus columnas
                        fila_limpia = {
                            'Corr.': fila[1] if len(fila) > 1 else "",
                            'Codigo Emp': codigo,
                            'Nombre Empleado': nombre,
                            'Días Laborados': limpiar_valor(fila[2]) if len(fila) > 2 else 0,
                            'Salario Mensual': limpiar_valor(fila[3]) if len(fila) > 3 else 0,
                            'Salario Quincenal': limpiar_valor(fila[4]) if len(fila) > 4 else 0,
                            'Horas Extra': limpiar_valor(fila[5]) if len(fila) > 5 else 0,
                            'Festivo': limpiar_valor(fila[6]) if len(fila) > 6 else 0,
                            'Comisiones': limpiar_valor(fila[7]) if len(fila) > 7 else 0,
                            'Vacaciones': limpiar_valor(fila[8]) if len(fila) > 8 else 0,
                            'Otros Ingresos': limpiar_valor(fila[9]) if len(fila) > 9 else 0,
                            'Salario Devengado': limpiar_valor(fila[10]) if len(fila) > 10 else 0,
                            'AFP': limpiar_valor(fila[11]) if len(fila) > 11 else 0,
                            'ISSS': limpiar_valor(fila[12]) if len(fila) > 12 else 0,
                            'Renta': limpiar_valor(fila[13]) if len(fila) > 13 else 0,
                            'Inst. Financieras': limpiar_valor(fila[14]) if len(fila) > 14 else 0,
                            'Préstamos': limpiar_valor(fila[15]) if len(fila) > 15 else 0,
                            'Otros Desc.': limpiar_valor(fila[16]) if len(fila) > 16 else 0,
                            'Total Desc.': limpiar_valor(fila[17]) if len(fila) > 17 else 0,
                            'Líquido a Recibir': limpiar_valor(fila[18]) if len(fila) > 18 else 0
                        }
                        datos_finales.append(fila_limpia)

            if datos_finales:
                df = pd.DataFrame(datos_finales)
                
                # TOTALIZACIÓN FINAL
                columnas_numericas = df.columns[3:] # De Días Laborados en adelante
                totales = df[columnas_numericas].sum()
                
                fila_totales = {col: "" for col in df.columns}
                fila_totales['Nombre Empleado'] = "TOTAL GENERAL"
                for col in columnas_numericas:
                    fila_totales[col] = totales[col]
                
                df = pd.concat([df, pd.DataFrame([fila_totales])], ignore_index=True)

                # Exportación a Excel con Formato
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Planilla_Limpia')
                    
                    workbook  = writer.book
                    worksheet = writer.sheets['Planilla_Limpia']
                    
                    # Formato contable para las columnas de dinero
                    formato_contable = workbook.add_format({'num_format': '#,##0.00', 'align': 'right'})
                    
                    for i, col in enumerate(df.columns):
                        width = max(len(str(col)), 15)
                        if i >= 3: # Columnas numéricas
                            worksheet.set_column(i, i, 15, formato_contable)
                        else:
                            worksheet.set_column(i, i, 25)

                st.success("¡Planilla procesada sin mezclas y con totales!")
                st.download_button(
                    label="📥 Descargar Excel Final",
                    data=output.getvalue(),
                    file_name="planilla_corregida.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("No se pudo extraer información clara del PDF.")
        except Exception as e:
            st.error(f"Error en el proceso: {e}")
