import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Convertidor Contable Pro", page_icon="mi_logo.png", layout="wide")
st.image("mi_logo.png", width=180) 
st.title("Convertidor de Planillas Profesional")

archivo_subido = st.file_uploader("Sube tu planilla PDF", type="pdf")

def limpiar_num(txt):
    if not txt: return 0.0
    # Quita $, comas y espacios. Maneja paréntesis de saldos negativos si los hay.
    s = str(txt).replace('$', '').replace(',', '').strip()
    if '(' in s and ')' in s:
        s = '-' + s.replace('(', '').replace(')', '')
    try:
        return float(s)
    except:
        return 0.0

if archivo_subido:
    if st.button("🚀 Generar Excel Perfecto"):
        try:
            filas_limpias = []
            
            with pdfplumber.open(archivo_subido) as pdf:
                for pagina in pdf.pages:
                    # Extraemos la tabla definiendo "puntos de corte" visuales (columnas)
                    # Esto evita que el nombre se pase a la columna de Días
                    tabla = pagina.extract_table({
                        "vertical_strategy": "text", 
                        "horizontal_strategy": "text",
                        "intersection_x_tolerance": 15
                    })
                    
                    if not tabla: continue
                    
                    for fila in tabla:
                        # Unimos la fila para ver si es basura (Encabezados o Agencias)
                        txt_fila = " ".join([str(x) for x in fila if x]).upper()
                        
                        if any(x in txt_fila for x in ["AGENCIA", "TOTALES", "CORR.", "SALARIO", "NOMBRE"]):
                            continue
                        
                        # Limpiamos nulos
                        f = [x.strip() if x else "" for x in fila]
                        
                        # Si la fila no tiene al menos un nombre o código, saltar
                        if len(f) < 5 or f[0] == "": continue

                        # SEPARACIÓN DE CÓDIGO Y NOMBRE (Evita que se mezclen)
                        # El código suele ser la primera palabra (ej: D46V11U)
                        primer_bloque = f[0].split()
                        cod = primer_bloque[0] if len(primer_bloque) > 0 else ""
                        nom = " ".join(primer_bloque[1:]) if len(primer_bloque) > 1 else ""

                        # Mapeo exacto según tu imagen de encabezado
                        fila_dict = {
                            'Corr.': f[1] if len(f) > 1 else "",
                            'Codigo Emp': cod,
                            'Nombre Empleado': nom,
                            'Días Laborados': limpiar_num(f[2]) if len(f) > 2 else 0,
                            'Salario Mensual': limpiar_num(f[3]) if len(f) > 3 else 0,
                            'Salario Quincenal': limpiar_num(f[4]) if len(f) > 4 else 0,
                            'Horas Extra': limpiar_num(f[5]) if len(f) > 5 else 0,
                            'Festivo': limpiar_num(f[6]) if len(f) > 6 else 0,
                            'Comisiones': limpiar_num(f[7]) if len(f) > 7 else 0,
                            'Vacaciones': limpiar_num(f[8]) if len(f) > 8 else 0,
                            'Otros Ingresos': limpiar_num(f[9]) if len(f) > 9 else 0,
                            'Salario Devengado': limpiar_num(f[10]) if len(f) > 10 else 0,
                            'AFP': limpiar_num(f[11]) if len(f) > 11 else 0,
                            'ISSS': limpiar_num(f[12]) if len(f) > 12 else 0,
                            'Renta': limpiar_num(f[13]) if len(f) > 13 else 0,
                            'Inst. Financieras': limpiar_num(f[14]) if len(f) > 14 else 0,
                            'Préstamos': limpiar_num(f[15]) if len(f) > 15 else 0,
                            'Otros Desc.': limpiar_num(f[16]) if len(f) > 16 else 0,
                            'Total Desc.': limpiar_num(f[17]) if len(f) > 17 else 0,
                            'Líquido a Recibir': limpiar_num(f[18]) if len(f) > 18 else 0
                        }
                        filas_limpias.append(fila_dict)

            if filas_limpias:
                df = pd.DataFrame(filas_limpias)
                
                # Crear Fila de Totales
                cols_num = df.columns[3:]
                sumas = df[cols_num].sum()
                fila_tot = {c: "" for c in df.columns}
                fila_tot['Nombre Empleado'] = "TOTAL GENERAL"
                for c in cols_num: fila_tot[c] = sumas[c]
                
                df = pd.concat([df, pd.DataFrame([fila_tot])], ignore_index=True)

                # Formato Excel
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Planilla')
                    wb = writer.book
                    ws = writer.sheets['Planilla']
                    
                    fmt_num = wb.add_format({'num_format': '#,##0.00', 'border': 1})
                    fmt_txt = wb.add_format({'border': 1})
                    
                    for i, col in enumerate(df.columns):
                        ws.set_column(i, i, 18, fmt_num if i >= 3 else fmt_txt)

                st.success("¡Planilla procesada correctamente!")
                st.download_button("📥 Descargar Excel", output.getvalue(), "planilla_final.xlsx")
            else:
                st.error("No se detectaron datos válidos. Revisa el formato del PDF.")
        except Exception as e:
            st.error(f"Error: {e}")
