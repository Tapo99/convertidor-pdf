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
            else:
                st.error("No se detectaron tablas en el PDF.")
        except Exception as e:
            st.error(f"Hubo un error: {e}")

