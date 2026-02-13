import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Convertidor Contable Pro", page_icon="mi_logo.png", layout="wide")
st.image("mi_logo.png", width=180) 
st.title("Convertidor de Planillas Profesional")

archivo_subido = st.file_uploader("Sube tu planilla PDF", type="pdf")

def limpiar_monto(texto):
    if not texto: return 0.0
    # Deja solo números, puntos y el signo menos
    limpio = re.sub(r'[^\d.-]', '', str(texto).replace(',', ''))
    try:
        return float(limpio)
    except:
        return 0.0

if archivo_subido:
    if st.button("🚀 Generar Excel Sin Desplazamientos"):
        try:
            filas_finales = []
            
            with pdfplumber.open(archivo_subido) as pdf:
                for pagina in pdf.pages:
                    # Extraemos el texto crudo por líneas (más confiable que tablas en este caso)
                    texto_pagina = pagina.extract_text()
                    if not texto_pagina: continue
                    
                    lineas = texto_pagina.split('\n')
                    
                    for linea in lineas:
                        l_up = linea.upper()
                        # FILTROS: Ignoramos basura y totales intermedios
                        if any(x in l_up for x in ["AGENCIA", "TOTALES", "CUENTA", "FECHA", "CORR.", "SALARIO", "NOMBRE", "CAJA"]):
                            continue
                        
                        # Buscamos todos los números en la línea (que tengan formato de moneda o días)
                        # Esta expresión busca números con decimales o enteros aislados
                        numeros_encontrados = re.findall(r'[\d,]+\.\d+|(?<=\s)\d+(?=\s)|(?<=\s)\d+$', linea)
                        
                        if len(numeros_encontrados) >= 10:
                            # 1. El nombre y código es todo lo que está ANTES del primer número de la línea
                            primer_numero = numeros_encontrados[0]
                            indice_corte = linea.find(primer_numero)
                            identidad_empleado = linea[:indice_corte].strip()
                            
                            # 2. Limpiamos los números encontrados
                            datos_num = [limpiar_monto(n) for n in numeros_encontrados]
                            
                            # 3. Alineación: Si detectamos que el correlativo se coló al inicio, lo quitamos
                            # Normalmente el primer número "pequeño" es el correlativo, lo saltamos.
                            if len(datos_num) > 17:
                                datos_num = datos_num[1:] # Eliminamos el correlativo del inicio

                            # Rellenamos de derecha a izquierda para que el Líquido siempre cuadre
                            while len(datos_num) < 17:
                                datos_num.insert(0, 0.0)
                            
                            datos_num = datos_num[-17:] # Nos quedamos con las 17 columnas contables

                            fila_dict = {
                                'Código y Nombre del Empleado': identidad_empleado,
                                'Días Laborados': datos_num[0],
                                'Salario Mensual': datos_num[1],
                                'Salario Quincenal': datos_num[2],
                                'Horas Extra': datos_num[3],
                                'Festivo': datos_num[4],
                                'Comisiones': datos_num[5],
                                'Vacaciones': datos_num[6],
                                'Otros Ingresos': datos_num[7],
                                'Salario Devengado': datos_num[8],
                                'AFP': datos_num[9],
                                'ISSS': datos_num[10],
                                'Renta': datos_num[11],
                                'Inst. Financieras': datos_num[12],
                                'Préstamos': datos_num[13],
                                'Otros Desc.': datos_num[14],
                                'Total Desc.': datos_num[15],
                                'Líquido a Recibir': datos_num[16]
                            }
                            filas_finales.append(fila_dict)

            if filas_finales:
                df = pd.DataFrame(filas_finales)
                
                # TOTAL GENERAL
                cols_n = df.columns[1:]
                sumas = df[cols_n].sum()
                fila_t = {c: "" for c in df.columns}
                fila_t['Código y Nombre del Empleado'] = "TOTAL GENERAL PLANILLA"
                for c in cols_n: fila_t[c] = sumas[c]
                
                df = pd.concat([df, pd.DataFrame([fila_t])], ignore_index=True)

                # EXCEL
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Planilla')
                    wb = writer.book
                    ws = writer.sheets['Planilla']
                    fmt_mon = wb.add_format({'num_format': '#,##0.00', 'border': 1})
                    fmt_txt = wb.add_format({'border': 1})
                    
                    ws.set_column(0, 0, 65, fmt_txt) # Columna de Identidad
                    ws.set_column(1, 18, 15, fmt_mon) # Columnas de Dinero

                st.success("¡Excel generado! Se eliminó el correlativo y se unificó nombre y código.")
                st.download_button("📥 Descargar Excel Final", output.getvalue(), "planilla_perfecta.xlsx")
            else:
                st.error("No se detectaron empleados en el PDF.")
        except Exception as e:
            st.error(f"Error: {e}")
