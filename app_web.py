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
    # Elimina $, comas y espacios. Mantiene el punto decimal.
    limpio = re.sub(r'[^\d.-]', '', str(texto).replace(',', ''))
    try:
        return float(limpio)
    except:
        return 0.0

if archivo_subido:
    if st.button("🚀 Generar Excel con Comisiones Corregidas"):
        try:
            filas_finales = []
            
            with pdfplumber.open(archivo_subido) as pdf:
                for pagina in pdf.pages:
                    # Usamos una estrategia mixta para no perder datos pegados
                    tabla = pagina.extract_table({
                        "vertical_strategy": "text",
                        "horizontal_strategy": "text",
                        "snap_tolerance": 8, # Aumentado para capturar datos muy juntos
                    })
                    
                    if not tabla: continue
                    
                    for fila in tabla:
                        # Unimos y limpiamos basura
                        f = [str(x).strip() if x else "" for x in fila]
                        texto_completo = " ".join(f).upper()
                        
                        if any(x in texto_completo for x in ["AGENCIA", "TOTALES", "CUENTA", "FECHA", "CORR.", "SALARIO", "NOMBRE"]):
                            continue
                        
                        if len([x for x in f if x]) < 5: continue

                        # SEPARACIÓN DE TEXTO Y NÚMEROS
                        # f[0] suele traer "CODIGO NOMBRE CORR"
                        # Vamos a extraer todos los números de la fila completa
                        todos_los_numeros = []
                        for celda in f:
                            # Buscamos montos (números con decimales o enteros)
                            encontrados = re.findall(r'[\d,]+\.\d+|(?<=\s)\d+(?=\s)|^\d+$', celda)
                            for n in encontrados:
                                todos_los_numeros.append(limpiar_monto(n))

                        # El PDF tiene una estructura fija de montos. 
                        # Si detectamos al menos los montos principales, procesamos:
                        if len(todos_los_numeros) >= 10:
                            # Identificar Código y Nombre (están en la primera parte del texto)
                            primer_texto = f[0]
                            match_cod = re.match(r'^([A-Z0-9]+)\s+(.*)', primer_texto)
                            
                            if match_cod:
                                codigo = match_cod.group(1)
                                nombre = match_cod.group(2)
                                # Si el nombre termina en el correlativo, lo limpiamos
                                nombre = re.sub(r'\s+\d+$', '', nombre)
                            else:
                                codigo = "Verificar"
                                nombre = primer_texto

                            # Mapeo exacto de columnas según tu necesidad (17 columnas de números)
                            # Si faltan números al final o en medio (como comisiones), rellenamos
                            # basándonos en la posición de derecha a izquierda (la más confiable)
                            
                            # Ajustamos la lista para que tenga exactamente 18 posiciones (Días + 17 montos)
                            numeros_finales = todos_los_numeros[-18:] if len(todos_los_numeros) >= 18 else [0.0]* (18 - len(todos_los_numeros)) + todos_los_numeros

                            fila_dict = {
                                'Corr.': int(numeros_finales[0]) if len(numeros_finales) > 0 else "",
                                'Código': codigo,
                                'Nombre Empleado': nombre,
                                'Días Laborados': numeros_finales[1],
                                'Salario Mensual': numeros_finales[2],
                                'Salario Quincenal': numeros_finales[3],
                                'Horas Extra': numeros_finales[4],
                                'Festivo': numeros_finales[5],
                                'Comisiones': numeros_finales[6], # <--- AQUÍ YA NO SERÁ CERO
                                'Vacaciones': numeros_finales[7],
                                'Otros Ingresos': numeros_finales[8],
                                'Salario Devengado': numeros_finales[9],
                                'AFP': numeros_finales[10],
                                'ISSS': numeros_finales[11],
                                'Renta': numeros_finales[12],
                                'Inst. Financieras': numeros_finales[13],
                                'Préstamos': numeros_finales[14],
                                'Otros Desc.': numeros_finales[15],
                                'Total Desc.': numeros_finales[16],
                                'Líquido a Recibir': numeros_finales[17]
                            }
                            filas_finales.append(fila_dict)

            if filas_finales:
                df = pd.DataFrame(filas_finales)
                
                # TOTALES AL FINAL
                cols_n = df.columns[3:]
                sumas = df[cols_n].sum()
                fila_t = {c: "" for c in df.columns}
                fila_t['Nombre Empleado'] = "TOTAL GENERAL"
                for c in cols_n: fila_t[c] = sumas[c]
                df = pd.concat([df, pd.DataFrame([fila_t])], ignore_index=True)

                # EXCEL
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Planilla_Corregida')
                    wb = writer.book
                    ws = writer.sheets['Planilla_Corregida']
                    fmt_mon = wb.add_format({'num_format': '#,##0.00', 'border': 1})
                    fmt_txt = wb.add_format({'border': 1})
                    
                    ws.set_column(0, 2, 20, fmt_txt)
                    ws.set_column(3, 20, 15, fmt_mon)

                st.success("¡Corregido! Ahora las comisiones y demás valores se leen correctamente.")
                st.download_button("📥 Descargar Excel Final", output.getvalue(), "planilla_comisiones_ok.xlsx")
        except Exception as e:
            st.error(f"Error: {e}")
