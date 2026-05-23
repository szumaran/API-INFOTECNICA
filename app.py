import streamlit as st
import asyncio
import os
import re

# Importación del módulo de Infotécnica independiente
from infotecnica import exportar_datos_async

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Extractor Infotécnica - Disgilent", page_icon="⚡", layout="wide")

st.title("⚡ Extractor Masivo Infotécnica (Input desde Disgilent)")
st.write("Pega tu lista de IDs directamente para descargar los parámetros desde el Coordinador Eléctrico")
st.markdown("---")

# 1. Cuadro de texto grande para pegar la lista de IDs masivamente
input_ids = st.text_area(
    "Pega aquí la lista de IDs (Puedes pegarlos separados por comas, espacios o uno por fila copiado de Excel/Disgilent):",
    placeholder="Ejemplo:\n3789\n4521\n8912",
    height=200
)

# 2. Selector de la forma de búsqueda (Ambas habilitadas por listas de IDs)
tipo_busqueda = st.radio(
    "Selecciona cómo deseas procesar esta lista de IDs:",
    [
        "🔍 Modo Directo: Los IDs corresponden directamente a los Equipos (Interruptores, Transformadores, etc.)",
        "🛤️ Modo Tramo: Los IDs corresponden a Tramos (Buscar todos los equipos asociados a estos tramos)"
    ]
)

st.markdown("---")
boton_extraer = st.button("🚀 Iniciar Extracción Masiva Asíncrona")

if boton_extraer:
    if not input_ids.strip():
        st.warning("⚠️ Por favor, pega al menos un ID para iniciar la extracción.")
    else:
        try:
            # Limpieza ultra flexible de IDs usando expresiones regulares.
            # Esto detecta cualquier número entero sin importar si están separados por comas, espacios o saltos de línea (Excel)
            list_ids = [int(x) for x in re.findall(r'\b\d+\b', input_ids)]
            
            if not list_ids:
                raise ValueError("No se encontraron números válidos en el cuadro de texto.")
                
            # Determinar el flag 'por_secciones' según la opción seleccionada del radio button
            es_modo_tramo = "Modo Tramo" in tipo_busqueda

            st.info(f"📋 Se identificaron {len(list_ids)} IDs únicos para procesar en el backend.")

            with st.spinner("Llamando a las APIs del Coordinador Eléctrico de forma asíncrona..."):
                loop = asyncio.get_event_loop()
                
                # Pasamos la lista limpia de IDs al backend original
                excel_path = loop.run_until_complete(
                    exportar_datos_async(ids_st=list_ids, por_secciones=es_modo_tramo)
                )
                
            if os.path.exists(excel_path):
                st.success("🎉 ¡Extracción masiva finalizada con éxito!")
                
                with open(excel_path, "rb") as file:
                    st.download_button(
                        label="📥 Descargar Parámetros en Excel (.xlsx)",
                        data=file,
                        file_name="parametros_infotecnica.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
        except ValueError:
            st.error("❌ Error: Asegúrate de que la lista contenga solo números enteros (IDs).")
        except Exception as e:
            st.error(f"❌ Error en el proceso de Infotécnica: {e}")

st.markdown("---")
st.caption("© 2026 Plataforma Eléctrica - Conector de Parámetros Masivos")
