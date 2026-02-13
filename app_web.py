import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

# 1. Configuración de la pestaña (ahora usa tu logo como icono)
st.set_page_config(page_title="Convertidor Pro", page_icon="mi_logo.png")

# 2. Tu Logo y Título
st.image("mi_logo.png", width=200) 
st.title("Convertidor de Planillas")

# 3. Subida de archivos
archivo_subido = st.file_uploader("Sube tu planilla PDF", type="pdf")

if archivo_subido:
    if st.button("🚀 Procesar Planilla"):
        try:
            datos_completos = []
            with pdfplumber.open(archivo_subido) as pdf:
                for pagina in pdf.pages:
                    # Usamos una estrategia más precisa para tablas de planillas
                    tabla = pagina.extract_table()
                    if tabla:
                        df = pd.DataFrame(tabla)
                        datos_completos.append(df)
            
            if datos_completos:
                # Aquí continúa tu lógica de limpieza...
                st.success("¡Planilla procesada con éxito!")
                if datos_completos:
                # Unimos todas las páginas en un solo Excel
                df_final = pd.concat(datos_completos, ignore_index=True)
                
                # Convertimos el Excel a datos que el navegador pueda descargar
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Planilla')
                
                excel_data = output.getvalue()

                st.success("¡Planilla procesada con éxito!") # Esto ya lo tienes
                
                # ESTO ES LO QUE FALTA: El botón de descarga
                st.download_button(
                    label="📥 Descargar Excel",
                    data=excel_data,
                    file_name="planilla_convertida.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            else:
                st.error("No se detectaron tablas en el PDF.")
        except Exception as e:
            st.error(f"Hubo un error: {e}")


