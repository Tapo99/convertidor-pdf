import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Convertidor Inteligente", page_icon="📊")

st.title("📊 Convertidor de Planillas Profesional")
st.markdown("Este sistema convierte en excel los pdf de las planilla por centros de costos.")

archivo_subido = st.file_uploader("Sube tu planilla PDF", type="pdf")

if archivo_subido:
    if st.button("🚀 Procesar y Limpiar"):
        datos_completos = []
        
        with pdfplumber.open(archivo_subido) as pdf:
            for pagina in pdf.pages:
                # Usamos una estrategia más precisa para tablas de planillas
                tabla = pagina.extract_table({
                    "vertical_strategy": "text", 
                    "horizontal_strategy": "text",
                    "snap_tolerance": 3,
                })
                if tabla:
                    for fila in tabla:
                        # Limpiamos espacios y saltos de línea molestos
                        fila_limpia = [str(celda).replace('\n', ' ').strip() if celda else "" for celda in fila]
                        # Ignoramos filas vacías o que son títulos de la empresa
                        if not any(fila_limpia) or "Planilla por Centro" in str(fila_limpia):
                            continue
                        datos_completos.append(fila_limpia)

        if datos_completos:
            df = pd.DataFrame(datos_completos)
            
            # --- LÓGICA DE LIMPIEZA DE ENCABEZADOS ---
            # Buscamos la fila que tiene los títulos reales
            for i, fila in df.iterrows():
                if "Código" in str(fila.values) or "INGRESO" in str(fila.values):
                    df.columns = df.iloc[i] # Ponemos esa fila como título
                    df = df.iloc[i+1:] # Borramos todo lo que hay arriba
                    break
            
            # Quitamos filas que repiten la palabra "Código" en cada página
            if not df.empty:
                df = df[df.iloc[:, 0] != df.columns[0]]
            
            # Crear Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            
            st.success("¡Planilla procesada con éxito!")
            st.download_button(
                label="📥 Descargar Excel Limpio",
                data=output.getvalue(),
                file_name="Planilla_Limpia.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:

            st.error("No se pudo extraer información clara del PDF.")
