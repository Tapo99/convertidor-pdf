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
    # Mantiene solo números, puntos y signo menos
    limpio = re.sub(r'[^\d.-]', '', str(texto).replace(',', ''))
    try:
        return float(limpio)
    except:
        return 0.0

if archivo_subido:
    if st.button("🚀 Generar Excel con Datos Unificados"):
        try:
            filas_finales = []
            
            with pdfplumber.open(archivo_subido) as pdf:
                for pagina in pdf.pages:
                    # Usamos una estrategia de extracción de palabras para identificar montos
                    tabla = pagina.extract_table({
                        "vertical_strategy": "text",
                        "horizontal_strategy": "text",
                        "snap_tolerance": 8,
                    })
                    
                    if not tabla: continue
                    
                    for fila in tabla:
                        # Limpiamos y unimos la fila para filtrar basura
                        f = [str(x).strip() if x else "" for x in fila]
                        texto_fila = " ".join(f).upper()
                        
                        # FILTROS: Saltamos encabezados de página y totales de centros de costo
                        if any(x in texto_fila for x in ["AGENCIA", "TOTALES", "CUENTA", "FECHA", "CORR.", "SALARIO", "NOMBRE", "CAJA"]):
                            continue
                        
                        # Extraemos todos los números de la fila (Días + 17 montos financieros)
                        todos_los_numeros = []
                        texto_acumulado = []
                        
                        for celda in f:
                            # Si la celda es puramente texto o código (tiene letras), la guardamos para el nombre
                            if re.search(r'[a-zA-Z]', celda):
                                texto_acumulado.append(celda)
                            else:
                                # Si parece número, extraemos los montos
                                encontrados = re.findall(r'[\d,]+\.\d+|(?<=\s)\d+(?=\s)|^\d+$', celda)
                                if encontrados:
                                    for n in encontrados:
                                        todos_los_numeros.append(limpiar_monto(n))
                                elif celda != "":
                                    # Caso de números sin decimales que quedaron solos
                                    todos_los_numeros.append(limpiar_monto(celda))

                        # Procesamos solo si detectamos la estructura de un empleado (Nombre + varios montos)
                        if len(todos_los_numeros) >= 10:
                            # FUSIONAMOS: Correlativo, Código y Nombre en un solo campo
                            identidad_empleado = " ".join(texto_acumulado).strip()
                            
                            # Si el primer número es pequeño (ej: 1, 2, 3), es el correlativo que se coló
                            # Lo movemos al texto de identidad para no desordenar los saldos
                            if len(todos_los_numeros) > 18:
                                identidad_empleado += f" (Corr: {int(todos_los_numeros[0])})"
                                datos_finanzas = todos_los_numeros[1:]
                            else:
                                datos_finanzas = todos_los_numeros

                            # Aseguramos exactamente 18 columnas numéricas (Días + 17 rubros)
                            # Rellenamos de derecha a izquierda para que el "Líquido a recibir" siempre cuadre
                            while len(datos_finanzas) < 18:
                                datos_finanzas.insert(0, 0.0)
                            
                            datos_finanzas = datos_finanzas[-18:]

                            fila_dict = {
                                'Datos del Empleado (Corr - Código - Nombre)': identidad_empleado,
                                'Días Laborados': datos_finanzas[0],
                                'Salario Mensual': datos_finanzas[1],
                                'Salario Quincenal': datos_finanzas[2],
                                'Horas Extra': datos_finanzas[3],
                                'Festivo': datos_finanzas[4],
                                'Comisiones': datos_finanzas[5],
                                'Vacaciones': datos_finanzas[6],
                                'Otros Ingresos': datos_finanzas[7],
                                'Salario Devengado': datos_finanzas[8],
                                'AFP': datos_finanzas[9],
                                'ISSS': datos_finanzas[10],
                                'Renta': datos_finanzas[11],
                                'Inst. Financieras': datos_finanzas[12],
                                'Préstamos': datos_finanzas[13],
                                'Otros Desc.': datos_finanzas[14],
                                'Total Desc.': datos_finanzas[15],
                                'Líquido a Recibir': datos_finanzas[16]
                            }
                            filas_finales.append(fila_dict)

            if filas_finales:
                df = pd.DataFrame(filas_finales)
                
                # TOTAL GENERAL AL FINAL
                cols_num = df.columns[1:]
                sumas = df[cols_num].sum()
                fila_t = {c: "" for c in df.columns}
                fila_t['Datos del Empleado (Corr - Código - Nombre)'] = "TOTAL GENERAL DE PLANILLA"
                for c in cols_num: fila_t[c] = sumas[c]
                df = pd.concat([df, pd.DataFrame([fila_t])], ignore_index=True)

                # EXCEL PROFESIONAL
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Planilla_Unificada')
                    wb = writer.book
                    ws = writer.sheets['Planilla_Unificada']
                    fmt_mon = wb.add_format({'num_format': '#,##0.00', 'border': 1})
                    fmt_txt = wb.add_format({'border': 1})
                    
                    ws.set_column(0, 0, 70, fmt_txt) # Columna de nombre muy ancha para que quepa todo
                    ws.set_column(1, 18, 15, fmt_mon)

                st.success("¡Hecho! Se unificó la identidad del empleado para proteger la alineación de los saldos.")
                st.download_button("📥 Descargar Excel Unificado", output.getvalue(), "planilla_unificada.xlsx")
        except Exception as e:
            st.error(f"Error técnico: {e}")
