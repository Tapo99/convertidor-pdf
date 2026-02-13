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
    # Deja solo números, puntos y signo menos
    limpio = re.sub(r'[^\d.-]', '', str(texto).replace(',', ''))
    try:
        return float(limpio)
    except:
        return 0.0

if archivo_subido:
    if st.button("🚀 Generar Excel Ordenado"):
        try:
            filas_finales = []
            
            with pdfplumber.open(archivo_subido) as pdf:
                for pagina in pdf.pages:
                    # Usamos la estrategia de texto para que no se pierdan los datos pegados
                    tabla = pagina.extract_table({
                        "vertical_strategy": "text",
                        "horizontal_strategy": "text",
                        "snap_tolerance": 5,
                    })
                    
                    if not tabla: continue
                    
                    for fila in tabla:
                        # Unimos la fila para detectar si es basura o totales
                        contenido = " ".join([str(x) for x in fila if x]).upper()
                        
                        # FILTROS: Quitamos todo lo que no sea un empleado
                        if any(x in contenido for x in ["AGENCIA", "TOTALES", "CUENTA", "FECHA", "CORR.", "SALARIO", "NOMBRE"]):
                            continue
                        
                        # Limpiamos celdas vacías
                        f = [str(x).strip() for x in fila if x is not None and str(x).strip() != ""]
                        
                        # Una fila real debe tener el nombre y varios montos (mínimo 10 datos)
                        if len(f) < 10:
                            continue

                        # UNIMOS CÓDIGO Y NOMBRE (Todo lo que no sea número al principio)
                        # Buscamos dónde empiezan los números (los días laborados)
                        idx_num = -1
                        for i, val in enumerate(f):
                            if re.match(r'^\d+(\.\d+)?$', val.replace(',', '')):
                                idx_num = i
                                break
                        
                        if idx_num != -1:
                            nombre_y_codigo = " ".join(f[:idx_num])
                            datos_numericos = f[idx_num:]
                            
                            # Limpiamos los números
                            nums = [limpiar_monto(n) for n in datos_numericos]
                            
                            # Rellenamos con ceros si faltan columnas
                            while len(nums) < 17:
                                nums.append(0.0)

                            fila_dict = {
                                'Empleado (Código y Nombre)': nombre_y_codigo,
                                'Días Laborados': nums[0],
                                'Salario Mensual': nums[1],
                                'Salario Quincenal': nums[2],
                                'Horas Extra': nums[3],
                                'Festivo': nums[4],
                                'Comisiones': nums[5],
                                'Vacaciones': nums[6],
                                'Otros Ingresos': nums[7],
                                'Salario Devengado': nums[8],
                                'AFP': nums[9],
                                'ISSS': nums[10],
                                'Renta': nums[11],
                                'Inst. Financieras': nums[12],
                                'Préstamos': nums[13],
                                'Otros Desc.': nums[14],
                                'Total Desc.': nums[15],
                                'Líquido a Recibir': nums[16]
                            }
                            filas_finales.append(fila_dict)

            if filas_finales:
                df = pd.DataFrame(filas_finales)
                
                # SUMA TOTAL AL FINAL
                cols_n = df.columns[1:]
                totales = df[cols_n].sum()
                fila_tot = {c: "" for c in df.columns}
                fila_tot['Empleado (Código y Nombre)'] = "TOTAL GENERAL PLANILLA"
                for c in cols_n:
                    fila_tot[c] = totales[c]
                
                df = pd.concat([df, pd.DataFrame([fila_tot])], ignore_index=True)

                # EXPORTAR A EXCEL
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Planilla_Limpia')
                    wb = writer.book
                    ws = writer.sheets['Planilla_Limpia']
                    
                    fmt_contable = wb.add_format({'num_format': '#,##0.00', 'border': 1})
                    fmt_texto = wb.add_format({'border': 1})
                    
                    ws.set_column(0, 0, 60, fmt_texto) # Columna de Nombre mucho más ancha
                    ws.set_column(1, 18, 15, fmt_contable) # Columnas de dinero alineadas

                st.success("¡Excel generado! Sin correlativos y con nombres unificados para mayor orden.")
                st.download_button("📥 Descargar Excel Final", output.getvalue(), "planilla_ordenada.xlsx")
            else:
                st.error("No se detectaron datos. Revisa que el PDF no sea una imagen.")
        except Exception as e:
            st.error(f"Error: {e}")
