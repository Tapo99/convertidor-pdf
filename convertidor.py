import pdfplumber
import pandas as pd

def extraer_y_limpiar_planilla(pdf_path, excel_path):
    datos_completos = []
    
    print("Leyendo y procesando planilla...")
    
    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            tabla = pagina.extract_table({
                "vertical_strategy": "text", 
                "horizontal_strategy": "text",
                "snap_tolerance": 3,
            })
            
            if tabla:
                for fila in tabla:
                    # Limpiamos saltos de línea y espacios extra
                    fila_limpia = [str(celda).replace('\n', ' ').strip() if celda else "" for celda in fila]
                    
                    # REGLAS DE LIMPIEZA:
                    # 1. Saltamos filas vacías
                    if not any(fila_limpia):
                        continue
                    # 2. Saltamos filas que contienen el título del reporte o el logo (ajusta según necesites)
                    if "Planilla por Centro Costo" in fila_limpia[0] or "Página" in str(fila_limpia):
                        continue
                        
                    datos_completos.append(fila_limpia)

    if datos_completos:
        df = pd.DataFrame(datos_completos)

        # --- LÓGICA PARA DEJAR SOLO UN ENCABEZADO ---
        # Asumimos que la primera fila que contiene "Código" es nuestro encabezado real
        encabezado_idx = None
        for i, fila in df.iterrows():
            if "Código" in str(fila.values):
                encabezado_idx = i
                break
        
        if encabezado_idx is not None:
            # Definimos los nombres de las columnas
            df.columns = df.iloc[encabezado_idx]
            # Filtramos: eliminamos todas las filas que repiten la palabra "Código" 
            # y también las filas que estaban arriba del primer encabezado
            df = df.iloc[encabezado_idx + 1:] 
            df = df[df.iloc[:, 0] != "Código"] 
        
        # Eliminar columnas completamente vacías (opcional)
        df = df.dropna(axis=1, how='all')

        # Guardar resultado final
        df.to_excel(excel_path, index=False)
        print(f"\n--- ¡Limpieza completada! ---")
        print(f"Los datos limpios están en: {excel_path}")
    else:
        print("No se encontró información procesable.")

# Ejecución
extraer_y_limpiar_planilla("Planilla por Centro Costo (1).pdf", "Planilla_Limpia.xlsx")