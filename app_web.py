import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

# Configuración de la página
st.set_page_config(page_title="Convertidor Pro", page_icon="mi_logo.png")

# Logo y Título
st.image("mi_logo.png", width=200) 
st.title("Convertidor de Planillas")

archivo_subido = st.file_uploader("Sube tu planilla PDF", type="pdf")

if archivo_subido:
    if st.button("🚀 Procesar Planilla"):
        try:
            datos_completos = []
            with pdfplumber.open(archivo_subido) as pdf:
                for pagina in pdf.pages:
                    tabla = pagina.extract_table()
                    if tabla:
                        df = pd.DataFrame(tabla)
                        datos_completos.append(df)
            
            if datos_completos:
                # UNIÓN Y PREPARACIÓN DEL EXCEL
                df_final = pd.concat(datos_completos, ignore_index=True)
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False)
                
                excel_data = output.getvalue()

                st.success("¡Planilla procesada con éxito!")
                
                # BOTÓN DE DESCARGA
                st.download_button(
                    label="📥 Descargar Excel",
                    data=excel_data,
                    file_name="planilla_convertida.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("No se encontraron datos en el PDF.")
        except Exception as e:
            st.error(f"Error al procesar: {e}")
