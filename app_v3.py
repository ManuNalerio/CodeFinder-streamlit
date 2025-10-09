"""
Buscador de Códigos en Proveedores
==================================

Aplicación Streamlit para buscar códigos específicos en archivos Excel 
de diferentes proveedores de manera eficiente.

Autor: Tu Nombre
Versión: 2.0.0
"""

import streamlit as st
import pandas as pd
import numpy as np
import os
import glob
from typing import List, Tuple, Dict, Any, Optional
from pathlib import Path


# ============================================================================
# CONFIGURACIÓN GLOBAL
# ============================================================================

# Configuración de la página
st.set_page_config(
    page_title="Buscar Código en Proveedores",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Constantes
PROVEEDORES_DIR = Path(r"C:\Users\PC-DEPO\Dropbox\ADMINISTRACION\CONTROL\PENDIENTES")
EXCEL_EXTENSIONS = ["*.xlsx", "*.xls"]
HEADER_KEYWORDS = [
    'codigo', 'descripcion', 'precio', 'stock', 'nombre', 'id', 'item',
    'producto', 'categoria', 'marca', 'modelo', 'sku', 'referencia'
]
PROGRESS_UPDATE_THRESHOLD = 0.1  # Actualizar progreso cada 10%


# ============================================================================
# CLASES Y ESTRUCTURAS DE DATOS
# ============================================================================

class ExcelProcessor:
    """Clase para manejar el procesamiento de archivos Excel."""
    
    def __init__(self):
        self.processed_files = 0
        self.error_files = 0
        
    @staticmethod
    def _calculate_header_score(row: pd.Series) -> int:
        """Calcula la puntuación para determinar si una fila es un header válido."""
        score = 0
        non_empty = row.notna().sum()
        
        if non_empty == 0:
            return 0
            
        # Puntuación base por cantidad de datos
        score += non_empty * 2
        
        # Bonus por contenido de texto vs números
        text_count = sum(
            1 for val in row 
            if pd.notna(val) and str(val).strip() and 
            not str(val).strip().replace('.', '').replace(',', '').replace('-', '').isdigit()
        )
        score += text_count * 3
        
        # Bonus por palabras clave de headers
        for val in row:
            if pd.notna(val) and any(keyword in str(val).lower().strip() for keyword in HEADER_KEYWORDS):
                score += 5
                break
                
        return score
    
    def detect_header_row(self, df: pd.DataFrame) -> int:
        """Detecta automáticamente qué fila contiene los headers reales."""
        best_score = 0
        best_row = 0
        
        for i in range(min(15, len(df))):
            row = df.iloc[i]
            score = self._calculate_header_score(row)
            
            if score > best_score:
                best_score = score
                best_row = i
                
        return best_row
    
    @staticmethod
    def clean_column_names(df: pd.DataFrame) -> pd.DataFrame:
        """Limpia y mejora los nombres de las columnas."""
        new_columns = []
        
        for i, col in enumerate(df.columns):
            if pd.isna(col) or not str(col).strip() or 'Unnamed' in str(col):
                # Intentar inferir nombre desde los datos
                column_data = df.iloc[:, i].dropna().head(5)
                inferred_name = f'Campo_{i+1}'
                
                if not column_data.empty:
                    first_val = str(column_data.iloc[0]).strip()
                    if first_val and any(char.isdigit() for char in first_val) and len(first_val) >= 3:
                        inferred_name = f'Datos_{i+1}'
                        
                new_columns.append(inferred_name)
            else:
                # Limpiar nombre existente
                clean_name = ' '.join(str(col).strip().split())
                clean_name = clean_name.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
                
                if len(clean_name) <= 2 or clean_name.lower() in ['col', 'column', 'field', 'data']:
                    clean_name = f'Campo_{i+1}'
                    
                new_columns.append(clean_name)
        
        df.columns = new_columns
        return df
    
    def try_multiple_read_strategies(self, file_path: Path) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
        """Intenta múltiples estrategias para leer el archivo Excel."""
        strategies = [
            {'header': 0},
            {'header': 1},
            {'header': 2},
            {'header': None}
        ]
        
        best_df = None
        best_score = -1
        
        for strategy in strategies:
            try:
                df_test = pd.read_excel(file_path, **strategy)
                
                if df_test.empty:
                    continue
                
                # Calcular puntuación de calidad
                score = self._calculate_dataframe_quality(df_test)
                
                if score > best_score:
                    best_df = df_test.copy()
                    best_score = score
                    
            except Exception:
                continue
        
        return best_df, None if best_df is not None else "No se pudo leer el archivo"
    
    @staticmethod
    def _calculate_dataframe_quality(df: pd.DataFrame) -> float:
        """Calcula la calidad de un DataFrame para seleccionar la mejor estrategia de lectura."""
        score = 0.0
        
        # Puntos por headers con texto
        text_headers = sum(
            1 for col in df.columns 
            if isinstance(col, str) and len(str(col).strip()) > 0 and 'Unnamed' not in str(col)
        )
        score += text_headers * 3
        
        # Puntos por variedad de datos
        for col in df.columns:
            if df[col].notna().sum() > 0:
                unique_ratio = df[col].nunique() / len(df[col].dropna())
                score += unique_ratio * 2
        
        # Penalizar columnas vacías
        empty_cols = sum(1 for col in df.columns if df[col].notna().sum() == 0)
        score -= empty_cols * 5
        
        return score
    
    def read_excel_optimized(self, file_path: Path) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
        """Lee un archivo Excel con detección optimizada de headers."""
        try:
            # Intentar múltiples estrategias
            df, error = self.try_multiple_read_strategies(file_path)
            
            if error or df is None:
                return None, error or "Error desconocido"
            
            # Si el header es problemático, detectar manualmente
            if any('Unnamed' in str(col) for col in df.columns):
                df_raw = pd.read_excel(file_path, header=None)
                header_row = self.detect_header_row(df_raw)
                
                if header_row > 0:
                    df = pd.read_excel(file_path, header=header_row)
            
            # Limpiar y optimizar
            df = self.clean_column_names(df)
            df = df.dropna(how='all').dropna(axis=1, how='all')
            
            return df, None
            
        except Exception as e:
            return None, str(e)


class SearchEngine:
    """Motor de búsqueda para códigos en archivos Excel."""
    
    def __init__(self, proveedores_dir: Path):
        self.proveedores_dir = proveedores_dir
        self.processor = ExcelProcessor()
    
    def get_proveedores(self) -> List[str]:
        """Obtiene la lista de proveedores disponibles."""
        if not self.proveedores_dir.exists():
            return []
            
        return [
            item.name for item in self.proveedores_dir.iterdir()
            if item.is_dir()
        ]
    
    def _get_excel_files(self, proveedor_path: Path) -> List[Path]:
        """Obtiene todos los archivos Excel de un proveedor."""
        excel_files = []
        for extension in EXCEL_EXTENSIONS:
            excel_files.extend(proveedor_path.glob(extension))
        return excel_files
    
    def _search_in_dataframe(self, df: pd.DataFrame, codigo: str, archivo: str) -> List[Dict[str, Any]]:
        """Busca un código en un DataFrame y retorna los resultados."""
        resultados = []
        
        for column in df.columns:
            try:
                # Búsqueda optimizada usando pandas
                col_str = df[column].astype(str).str.strip()
                matches_mask = col_str.str.contains(codigo, case=False, na=False, regex=False)
                
                if matches_mask.any():
                    matches = df[matches_mask].copy()
                    
                    for _, row in matches.iterrows():
                        resultado = {
                            "Archivo": archivo,
                            "Columna_encontrada": column
                        }
                        
                        # Agregar solo datos relevantes
                        for col_name, value in row.items():
                            if pd.notna(value) and str(value).strip():
                                resultado[col_name] = value
                        
                        resultados.append(resultado)
                        
            except Exception:
                continue
                
        return resultados
    
    def search_codigo(self, proveedor: str, codigo: str) -> Tuple[List[Dict[str, Any]], Dict[str, int]]:
        """Busca un código en todos los archivos Excel de un proveedor."""
        resultados = []
        stats = {"procesados": 0, "errores": 0}
        
        proveedor_path = self.proveedores_dir / proveedor
        excel_files = self._get_excel_files(proveedor_path)
        
        if not excel_files:
            return resultados, stats
        
        # Configurar barra de progreso
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, file_path in enumerate(excel_files):
            stats["procesados"] += 1
            status_text.text(f'Procesando {stats["procesados"]}/{len(excel_files)}: {file_path.name}')
            
            # Leer archivo
            df, error = self.processor.read_excel_optimized(file_path)
            
            if error or df is None or df.empty:
                stats["errores"] += 1
                if error:
                    resultados.append({
                        "Archivo": file_path.name,
                        "Error": f"Error: {error}"
                    })
                continue
            
            # Buscar en el DataFrame
            file_results = self._search_in_dataframe(df, codigo, file_path.name)
            resultados.extend(file_results)
            
            # Actualizar progreso
            progress_bar.progress((idx + 1) / len(excel_files))
        
        # Limpiar UI
        progress_bar.empty()
        status_text.empty()
        
        return resultados, stats


class ResultsProcessor:
    """Procesador de resultados de búsqueda."""
    
    @staticmethod
    def filter_relevant_columns(df: pd.DataFrame, min_data_ratio: float = 0.3) -> pd.DataFrame:
        """Filtra columnas relevantes basado en la cantidad de datos."""
        essential_columns = ['Archivo', 'Columna_encontrada']
        relevant_columns = essential_columns.copy()
        
        for col in df.columns:
            if col not in essential_columns:
                non_empty_ratio = df[col].notna().sum() / len(df)
                if non_empty_ratio > min_data_ratio:
                    relevant_columns.append(col)
        
        return df[relevant_columns]
    
    @staticmethod
    def separate_results_and_errors(resultados: List[Dict[str, Any]]) -> Tuple[List[Dict], List[Dict]]:
        """Separa resultados válidos de errores."""
        valid_results = [r for r in resultados if 'Error' not in r]
        errors = [r for r in resultados if 'Error' in r]
        return valid_results, errors
    
    @staticmethod
    def create_download_csv(df: pd.DataFrame) -> str:
        """Crea CSV para descarga con encoding correcto."""
        return df.to_csv(index=False, encoding='utf-8-sig')


# ============================================================================
# INTERFAZ DE USUARIO
# ============================================================================

class StreamlitUI:
    """Clase para manejar la interfaz de usuario de Streamlit."""
    
    def __init__(self):
        self.search_engine = SearchEngine(PROVEEDORES_DIR)
        self.results_processor = ResultsProcessor()
    
    def render_header(self):
        """Renderiza el header de la aplicación."""
        st.title("Buscar Código en Proveedores")
    
    def render_sidebar(self) -> Tuple[bool, bool]:
        """Renderiza la barra lateral con opciones."""
        st.sidebar.header("Opciones de Búsqueda")
        busqueda_exacta = st.sidebar.checkbox("Búsqueda exacta", value=False)
        mostrar_solo_relevantes = st.sidebar.checkbox("Mostrar solo columnas relevantes", value=True)
        return busqueda_exacta, mostrar_solo_relevantes
    
    def render_search_form(self) -> Tuple[str, str, bool]:
        """Renderiza el formulario de búsqueda."""
        with st.form(key="busqueda_form"):
            proveedores = self.search_engine.get_proveedores()
            
            col1, col2 = st.columns(2)
            with col1:
                proveedor = st.selectbox(
                    "Seleccionar Proveedor",
                    ["-- Seleccione un proveedor --"] + proveedores
                )
            with col2:
                codigo = st.text_input(
                    "Código a buscar",
                    placeholder="Ingrese el código a buscar..."
                )
            
            submit_button = st.form_submit_button(
                label="🔍 Buscar",
                use_container_width=True
            )
        
        return proveedor, codigo, submit_button
    
    def render_results(self, resultados: List[Dict], stats: Dict[str, int], 
                      codigo: str, proveedor: str, mostrar_solo_relevantes: bool):
        """Renderiza los resultados de la búsqueda."""
        st.subheader(f"Resultados de la búsqueda para: **'{codigo}'**")
        
        # Mostrar estadísticas
        st.info(f"Procesados: {stats['procesados']} archivos. Errores: {stats['errores']}")
        
        if not resultados:
            st.warning(f"❌ No se encontraron resultados para el código '{codigo}' en {proveedor}.")
            return
        
        # Separar resultados y errores
        valid_results, errors = self.results_processor.separate_results_and_errors(resultados)
        
        if valid_results:
            self._render_valid_results(valid_results, mostrar_solo_relevantes, codigo, proveedor)
        
        if errors:
            self._render_errors(errors)
        
        if not valid_results:
            st.warning(f"❌ No se encontraron resultados válidos para '{codigo}' en {proveedor}.")
    
    def _render_valid_results(self, valid_results: List[Dict], mostrar_solo_relevantes: bool,
                             codigo: str, proveedor: str):
        """Renderiza los resultados válidos."""
        df_resultados = pd.DataFrame(valid_results)
        
        # Estadísticas de éxito
        st.success(
            f"✅ Se encontraron {len(valid_results)} coincidencias "
            f"en {df_resultados['Archivo'].nunique()} archivos diferentes"
        )
        
        # Filtrar columnas si es necesario
        if mostrar_solo_relevantes:
            df_mostrar = self.results_processor.filter_relevant_columns(df_resultados)
        else:
            df_mostrar = df_resultados
        
        # Mostrar tabla
        st.dataframe(df_mostrar, use_container_width=True, height=400)
        
        # Botón de descarga
        csv = self.results_processor.create_download_csv(df_resultados)
        st.download_button(
            label="📥 Descargar resultados como CSV",
            data=csv,
            file_name=f"busqueda_{codigo}_{proveedor}.csv",
            mime="text/csv"
        )
    
    def _render_errors(self, errors: List[Dict]):
        """Renderiza los errores encontrados."""
        st.warning(f"⚠️ Se encontraron {len(errors)} archivos con errores:")
        df_errores = pd.DataFrame(errors)
        st.dataframe(df_errores, use_container_width=True)
    
    def render_help_section(self):
        """Renderiza la sección de ayuda."""
        st.info("👆 Seleccione un proveedor, ingrese un código y haga clic en 'Buscar' para ver los resultados.")
        
        with st.expander("ℹ️ Ayuda y características"):
            st.markdown("""
            **Características de esta aplicación:**
            
            - 🔍 **Búsqueda inteligente**: Detecta automáticamente los headers correctos en archivos Excel
            - 🧹 **Limpieza de datos**: Elimina columnas y filas vacías automáticamente  
            - 📊 **Resultados organizados**: Muestra solo las columnas relevantes por defecto
            - 📈 **Progreso en tiempo real**: Muestra el progreso del procesamiento
            - 📥 **Exportación**: Descarga los resultados en formato CSV
            - ⚠️ **Manejo de errores**: Reporta archivos problemáticos sin interrumpir la búsqueda
            
            **Consejos:**
            - Los códigos se buscan en todas las columnas de todos los archivos Excel
            - La búsqueda no distingue entre mayúsculas y minúsculas
            - Use la opción "Mostrar solo columnas relevantes" para ver resultados más limpios
            """)
    
    def run(self):
        """Ejecuta la aplicación principal."""
        self.render_header()
        busqueda_exacta, mostrar_solo_relevantes = self.render_sidebar()
        proveedor, codigo, submit_button = self.render_search_form()
        
        # Validar entrada y procesar búsqueda
        if submit_button:
            if proveedor == "-- Seleccione un proveedor --" or not codigo:
                st.error("⚠️ Por favor, seleccione un proveedor e ingrese un código.")
            else:
                with st.spinner('Buscando en los archivos...'):
                    resultados, stats = self.search_engine.search_codigo(proveedor, codigo)
                
                self.render_results(resultados, stats, codigo, proveedor, mostrar_solo_relevantes)
        else:
            self.render_help_section()


# ============================================================================
# PUNTO DE ENTRADA PRINCIPAL
# ============================================================================

def main():
    """Función principal de la aplicación."""
    app = StreamlitUI()
    app.run()


if __name__ == "__main__":
    main()