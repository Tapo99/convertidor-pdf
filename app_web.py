import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Convertidor Pro", page_icon="📊")

st.title("Convertidor de Planillas")

archivo_subido = st.file_uploader("Sube tu planilla PDF", type="pdf")

if archivo_subido:
    if st.button("🚀 Procesar Planilla"):
        try:
            datos_completos = []
            with pdfplumber.open(archivo_subido) as pdf:
                for pagina in pdf.pages:
                    tabla = pagina.extract_table({
                        "vertical_strategy": "text", 
                        "horizontal_strategy": "text",
                        "snap_tolerance": 4,
                    })
                    if tabla:
                        for fila in tabla:
                            fila_limpia = [str(c).replace('\n', ' ').strip() if c else "" for c in fila]
                            # Filtro básico para filas vacías o encabezados
                            if len([c for c in fila_limpia if c]) > 5:
                                datos_completos.append(fila_limpia)

            if datos_completos:
                df = pd.DataFrame(datos_completos)
                
                # Nombres de columnas deseados
                cols_deseadas = ["N°", "CODIGO", "EMPLEADO", "Días laborados", "Salario", "Quincenal", "Extras", "Comisiones", "Otr", "DEVENGADO", "AFP", "ISSS", "RENTA", "Bancos", "Préstamos", "Otros", "Total Desc", "Líquido"]
                
                # Ajustamos el DataFrame para que no falle por tamaño
                num_cols = df.shape[1]
                df.columns = cols_deseadas[:num_cols]

                # --- Limpieza Numérica ---
                def limpiar(v, entero=False):
                    try:
                        n = float(str(v).replace('$','').replace(',','').strip())
                        return int(n) if entero else n
                    except: return 0

                if "Días laborados" in df.columns:
                    df["Días laborados"] = df["Días laborados"].apply(lambda x: limpiar(x, True))
                
                # Columnas de dinero (intentar convertir todas las posibles)
                for col in df.columns:
                    if col not in ["N°", "CODIGO", "EMPLEADO", "Días laborados"]:
                        df[col] = df[col].apply(lambda x: limpiar(x, False))

                # --- Crear Excel ---
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Datos')
                    workbook  = writer.book
                    worksheet = writer.sheets['Datos']
                    
                    # Formato Contable
                    f_money = workbook.add_format({'num_format': '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'})
                    f_int = workbook.add_format({'align': 'center'})
                    
                    worksheet.set_column('C:C', 40) # Nombre
                    worksheet.set_column('D:D', 12, f_int) # Días
                    worksheet.set_column('E:R', 15, f_money) # Dinero

                st.success("✅ ¡Proceso exitoso!")
                st.download_button("📥 Descargar Excel", output.getvalue(), "Planilla.xlsx")
            else:
                st.warning("No se encontraron datos en el PDF.")
        
        except Exception as e:
            st.error(f"Ocurrió un error técnico: {e}")
            st.info("Asegúrate de tener instalada la librería xlsxwriter: pip install xlsxwriter")
