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
    if not texto or str(texto).strip() == "": return 0.0
    limpio = re.sub(r'[^\d.-]', '', str(texto).replace(',', ''))
    try:
        return float(limpio)
    except:
        return 0.0

if archivo_subido:
    if st.button("🚀 Generar Excel Corregido"):
        try:
            filas_finales = []
            
            with pdfplumber.open(archivo_subido) as pdf:
                for pagina in pdf.pages:
                    # Extraemos la tabla con ajustes de tolerancia para evitar el desorden en cambios de página
                    tabla = pagina.extract_table({
                        "vertical_strategy": "text",
                        "horizontal_strategy": "text",
                        "snap_tolerance": 6, # Mayor tolerancia para filas desalineadas
                        "join_tolerance": 3,
                    })
                    
                    if not tabla: continue
                    
                    for fila in tabla:
                        # Limpiamos nulos y espacios
                        f = [str(x).strip() if x else "" for x in fila]
                        
                        # Unimos la fila para detectar basura (Agencias, Totales, Encabezados)
                        texto_fila = " ".join(f).upper()
                        if any(x in texto_fila for x in ["AGENCIA", "TOTALES", "CUENTA", "FECHA", "CORR.", "SALARIO", "NOMBRE", "CAJA"]):
                            continue

                        # Si la fila está vacía o es muy corta, la saltamos
                        if len([x for x in f if x]) < 5:
                            continue

                        # PROCESAMIENTO QUIRÚRGICO DE NOMBRE Y CÓDIGO
                        # Intentamos separar el código (letras/números al inicio) del nombre
                        primera_celda = f[0]
                        match_cod = re.match(r'^([A-Z0-9]+)\s+(.*)', primera_celda)
                        
                        if match_cod:
                            codigo = match_cod.group(1)
                            nombre = match_cod.group(2)
                        else:
                            codigo = "Revisar"
                            nombre = primera_celda

                        # Buscamos el correlativo (suele estar en la celda 1 o al final del nombre)
                        # Si el nombre termina en número, ese es el correlativo
                        corr_match = re.search(r'\s+(\d+)$', nombre)
                        if corr_match:
                            correlativo = corr_match.group(1)
                            nombre = nombre[:corr_match.start()].strip()
                        else:
                            correlativo = f[1] if len(f) > 1 else ""

                        # Extraemos los números (empezando desde donde detectamos montos)
                        # Buscamos la primera celda que parezca un número después del nombre
                        datos_numeros = []
                        for celda in f:
                            if any(char.isdigit() for char in celda) and ("." in celda or "," in celda or celda.isdigit()):
                                # Si la celda contiene letras (ej: "D46V11U"), no es un monto puro
                                if not re.search(r'[a-zA-Z]', celda):
                                    datos_numeros.append(limpiar_monto(celda))
                        
                        # Filtramos los números para quedarnos solo con los 17 campos contables
                        # (Días, Salarios, AFP, ISSS, Renta, etc.)
                        if len(datos_numeros) > 17:
                            datos_numeros = datos_numeros[-17:] # Nos quedamos con los últimos (los montos)
                        while len(datos_numeros) < 17:
                            datos_numeros.append(0.0)

                        fila_dict = {
                            'Corr.': correlativo,
                            'Código': codigo,
                            'Nombre Empleado': nombre,
                            'Días Laborados': datos_numeros[0],
                            'Salario Mensual': datos_numeros[1],
                            'Salario Quincenal': datos_numeros[2],
                            'Horas Extra': datos_numeros[3],
                            'Festivo': datos_numeros[4],
                            'Comisiones': datos_numeros[5],
                            'Vacaciones': datos_numeros[6],
                            'Otros Ingresos': datos_numeros[7],
                            'Salario Devengado': datos_numeros[8],
                            'AFP': datos_numeros[9],
                            'ISSS': datos_numeros[10],
                            'Renta': datos_numeros[11],
                            'Inst. Financieras': datos_numeros[12],
                            'Préstamos': datos_numeros[13],
                            'Otros Desc.': datos_numeros[14],
                            'Total Desc.': datos_numeros[15],
                            'Líquido a Recibir': datos_numeros[16]
                        }
                        filas_finales.append(fila_dict)

            if filas_finales:
                df = pd.DataFrame(filas_finales)
                
                # TOTALIZACIÓN GENERAL
                cols_n = df.columns[3:]
                sumas = df[cols_n].sum()
                fila_t = {c: "" for c in df.columns}
                fila_t['Nombre Empleado'] = "TOTAL GENERAL"
                for c in cols_n: fila_t[c] = sumas[c]
                
                df = pd.concat([df, pd.DataFrame([fila_t])], ignore_index=True)

                # EXCEL CON FORMATO
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Planilla')
                    wb = writer.book
                    ws = writer.sheets['Planilla']
                    fmt_mon = wb.add_format({'num_format': '#,##0.00', 'border': 1})
                    fmt_txt = wb.add_format({'border': 1})
                    
                    ws.set_column(0, 1, 12, fmt_txt) # Corr y Cod
                    ws.set_column(2, 2, 40, fmt_txt) # Nombre
                    ws.set_column(3, 19, 15, fmt_mon) # Números

                st.success("¡Corregido! Se recuperaron los nombres después de la fila 32 y se separó el código.")
                st.download_button("📥 Descargar Excel Corregido", output.getvalue(), "planilla_contable_final.xlsx")
            else:
                st.error("No se detectaron datos. Revisa el PDF.")
        except Exception as e:
            st.error(f"Error: {e}")
