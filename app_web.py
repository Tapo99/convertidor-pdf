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
    # Solo deja números, puntos y el signo menos
    limpio = re.sub(r'[^\d.-]', '', str(texto).replace(',', ''))
    try:
        return float(limpio)
    except:
        return 0.0

if archivo_subido:
    if st.button("🚀 Generar Excel Final Ordenado"):
        try:
            lista_empleados = []
            
            with pdfplumber.open(archivo_subido) as pdf:
                for pagina in pdf.pages:
                    # Extraemos las tablas con una configuración de "tolerancia" alta
                    # para que no rompa las filas en pedazos
                    tabla = pagina.extract_table({
                        "vertical_strategy": "text",
                        "horizontal_strategy": "text",
                        "snap_tolerance": 4,
                    })
                    
                    if not tabla: continue
                    
                    for fila in tabla:
                        # Unimos toda la fila para analizar si es basura
                        contenido = " ".join([str(x) for x in fila if x]).upper()
                        
                        # FILTROS: Saltamos encabezados, agencias y totales de centros de costo
                        if any(x in contenido for x in ["AGENCIA", "TOTALES", "CUENTA", "FECHA", "CORR.", "SALARIO", "NOMBRE", "EMPLEADO"]):
                            continue
                        
                        # Limpiamos la fila de valores nulos
                        f = [str(x).strip() if x else "" for x in fila]
                        f = [x for x in f if x != ""] # Quitamos espacios vacíos internos
                        
                        # Una fila válida de empleado debe tener al menos el nombre y los montos (aprox 10+ elementos)
                        if len(f) < 10:
                            continue

                        # ESTRATEGIA: Unimos Código + Nombre para evitar que los números se desplacen
                        # El PDF suele traer: [COD+NOMBRE] [CORR] [DIAS] [SALARIO]...
                        nombre_completo = f[0]
                        correlativo = f[1]
                        
                        # Los números empiezan usualmente después del correlativo
                        numeros_raw = f[2:]
                        numeros = [limpiar_monto(n) for n in numeros_raw]

                        # Aseguramos que tengamos exactamente las columnas necesarias rellenando con 0
                        while len(numeros) < 17:
                            numeros.append(0.0)

                        fila_dict = {
                            'Codigo y Nombre': nombre_completo,
                            'Corr.': correlativo,
                            'Días Laborados': numeros[0],
                            'Salario Mensual': numeros[1],
                            'Salario Quincenal': numeros[2],
                            'Horas Extra': numeros[3],
                            'Festivo': numeros[4],
                            'Comisiones': numeros[5],
                            'Vacaciones': numeros[6],
                            'Otros Ingresos': numeros[7],
                            'Salario Devengado': numeros[8],
                            'AFP': numeros[9],
                            'ISSS': numeros[10],
                            'Renta': numeros[11],
                            'Inst. Financieras': numeros[12],
                            'Préstamos': numeros[13],
                            'Otros Desc.': numeros[14],
                            'Total Desc.': numeros[15],
                            'Líquido a Recibir': numeros[16]
                        }
                        lista_empleados.append(fila_dict)

            if lista_empleados:
                df = pd.DataFrame(lista_empleados)
                
                # TOTALIZACIÓN ÚNICA (Solo al final de todo el archivo)
                cols_num = df.columns[2:]
                suma_total = df[cols_num].sum()
                
                fila_total = {c: "" for c in df.columns}
                fila_total['Codigo y Nombre'] = "TOTAL GENERAL DE PLANILLA"
                for c in cols_num:
                    fila_total[c] = suma_total[c]
                
                df = pd.concat([df, pd.DataFrame([fila_total])], ignore_index=True)

                # FORMATO EXCEL
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Planilla')
                    wb = writer.book
                    ws = writer.sheets['Planilla']
                    
                    # Estilos
                    fmt_moneda = wb.add_format({'num_format': '#,##0.00', 'border': 1})
                    fmt_texto = wb.add_format({'border': 1})
                    
                    # Anchos de columna
                    ws.set_column(0, 0, 50, fmt_texto) # Nombre y Código
                    ws.set_column(1, 1, 10, fmt_texto) # Corr
                    ws.set_column(2, 18, 15, fmt_moneda) # Números

                st.success("¡Excel generado! Se eliminaron los totales por agencia y se alinearon las columnas.")
                st.download_button("📥 Descargar Planilla Limpia", output.getvalue(), "planilla_final.xlsx")
            else:
                st.error("No se encontraron filas de empleados. Verifica el formato del PDF.")
        except Exception as e:
            st.error(f"Error: {e}")
