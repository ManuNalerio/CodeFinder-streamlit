import streamlit as st
import pandas as pd
import os
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
from typing import List, Dict, Tuple, Optional

# Extensiones válidas para archivos Excel
EXCEL_EXTENSIONS = [".xlsx", ".xls"]
# Ruta donde se encuentran las carpetas de proveedores
PROVEEDORES_DIR = Path(r"C:\Users\PC-DEPO\Dropbox\ADMINISTRACION\CONTROL\PENDIENTES")

@st.cache_data
def read_excel_file(filepath: Path) -> Optional[pd.DataFrame]:
    """Leer archivo Excel con manejo de errores y limpieza básica."""
    try:
        df = pd.read_excel(filepath, header=None)
        # Intentar detectar el encabezado en las primeras filas
        for i in range(min(10, len(df))):
            if all(isinstance(x, str) for x in df.iloc[i].dropna()):
                df.columns = df.iloc[i]  # Asignar encabezado
                df = df[(i+1):]          # Eliminar filas de encabezado
                break
        df = df.dropna(axis=1, how='all')  # Quitar columnas completamente vacías
        return df.reset_index(drop=True)
    except Exception:
        return None  # Si hay error, retornar None

def search_code_in_df(df: pd.DataFrame, code: str) -> List[Dict]:
    """Buscar código en todo el DataFrame."""
    results = []
    if df is None or df.empty:
        return results
    # Buscar el código en cada columna
    for col in df.columns:
        try:
            matches = df[col].astype(str).str.contains(code, case=False, na=False)
            for idx in df[matches].index:
                results.append({"Fila": idx + 1, "Columna": col, "Valor": df.at[idx, col]})
        except Exception:
            continue  # Si hay error en una columna, continuar con la siguiente
    return results

def get_proveedores() -> List[str]:
    """Obtener lista de proveedores (carpetas) disponibles."""
    if not PROVEEDORES_DIR.exists():
        return []
    return [p.name for p in PROVEEDORES_DIR.iterdir() if p.is_dir()]

def get_excel_files(proveedor: str) -> List[Path]:
    """Obtener lista de archivos Excel para un proveedor."""
    path = PROVEEDORES_DIR / proveedor
    files = []
    for ext in EXCEL_EXTENSIONS:
        files.extend(path.glob(f"*{ext}"))
    return files

def search_code_in_files(proveedor: str, code: str) -> Tuple[List[Dict], int, int]:
    """Buscar el código en todos los archivos Excel de un proveedor."""
    excel_files = get_excel_files(proveedor)
    results = []
    errors = 0

    def process_file(file):
        df = read_excel_file(file)
        if df is None:
            return None, file.name  # Retornar nombre de archivo si hubo error
        res = search_code_in_df(df, code)
        return res, None

    # Procesar archivos en paralelo para mayor velocidad
    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {executor.submit(process_file, file): file for file in excel_files}
        for future in futures:
            res, err_file = future.result()
            if err_file:
                errors += 1
            elif res:
                for r in res:
                    r["Archivo"] = futures[future].name  # Agregar nombre de archivo al resultado
                results.extend(res)

    return results, len(excel_files), errors

def main():
    """Función principal de la aplicación Streamlit."""
    st.title("Buscar Código en Proveedores Excel")
    proveedores = get_proveedores()
    proveedor = st.selectbox("Seleccionar proveedor", ["--"] + proveedores)
    code = st.text_input("Código a buscar")
    buscar = st.button("Buscar")

    if buscar:
        if proveedor == "--" or not code:
            st.error("Seleccione un proveedor y escriba un código válido")
            return

        with st.spinner("Buscando..."):
            resultados, total, errores = search_code_in_files(proveedor, code)

        st.write(f"Archivos procesados: {total}")
        st.write(f"Errores: {errores}")

        if resultados:
            df_res = pd.DataFrame(resultados)
            st.dataframe(df_res)
            csv = df_res.to_csv(index=False).encode("utf-8")
            st.download_button("Descargar resultados CSV", data=csv, file_name=f"resultados_{code}_{proveedor}.csv")
        else:
            st.warning("No se encontraron resultados")

if __name__ == "__main__":
    main()