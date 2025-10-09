import streamlit as st
import pandas as pd
import os
import glob

# Configuración de la página
st.set_page_config(page_title="Buscar Código en Proveedores", page_icon="🔍", layout="wide")

# Título de la aplicación
st.title("Buscar Código en Proveedores")

# Ruta de la carpeta principal con los proveedores
PROVEEDORES_DIR = r"C:\Users\PC-DEPO\Dropbox\ADMINISTRACION\CONTROL\PENDIENTES"

# Función para obtener la lista de proveedores
def get_proveedores():
    if not os.path.exists(PROVEEDORES_DIR):
        return []
    return [name for name in os.listdir(PROVEEDORES_DIR) 
            if os.path.isdir(os.path.join(PROVEEDORES_DIR, name))]

# Función para buscar el código en los archivos Excel de un proveedor
def buscar_codigo(proveedor, codigo):
    resultados = []
    proveedor_path = os.path.join(PROVEEDORES_DIR, proveedor)
    # Buscar todos los archivos Excel en la carpeta del proveedor
    excel_files = glob.glob(os.path.join(proveedor_path, "*.xlsx")) + \
                  glob.glob(os.path.join(proveedor_path, "*.xls"))
    
    for file in excel_files:
        try:
            # Leer el archivo Excel
            df = pd.read_excel(file)
            # Buscar el código en todas las columnas
            for column in df.columns:
                if df[column].astype(str).str.contains(codigo, case=False, na=False).any():
                    # Filtrar las filas que contienen el código
                    matches = df[df[column].astype(str).str.contains(codigo, case=False, na=False)]
                    for _, row in matches.iterrows():
                        resultados.append({
                            "Archivo": os.path.basename(file),
                            **row.to_dict()
                        })
        except Exception as e:
            resultados.append({
                "Archivo": os.path.basename(file),
                "Error": f"Error al procesar el archivo: {str(e)}"
            })
    
    return resultados

# Formulario para seleccionar proveedor y código
with st.form(key="busqueda_form"):
    proveedores = get_proveedores()
    proveedor = st.selectbox("Seleccionar Proveedor", ["-- Seleccione un proveedor --"] + proveedores)
    codigo = st.text_input("Código a buscar")
    submit_button = st.form_submit_button(label="Buscar")

# Lógica para procesar la búsqueda
if submit_button and proveedor != "-- Seleccione un proveedor --" and codigo:
    st.subheader(f"Resultados de la búsqueda para '{codigo}'")
    resultados = buscar_codigo(proveedor, codigo)
    
    if resultados:
        # Convertir resultados a DataFrame para mostrar en tabla
        df_resultados = pd.DataFrame(resultados)
        st.dataframe(df_resultados, use_container_width=True)
    else:
        st.warning(f"No se encontraron resultados para el código '{codigo}' en los archivos de {proveedor}.")
elif submit_button:
    st.error("Por favor, seleccione un proveedor e ingrese un código.")

# Mensaje inicial si no se ha realizado ninguna búsqueda
if not submit_button:
    st.info("Seleccione un proveedor, ingrese un código y haga clic en 'Buscar' para ver los resultados.")