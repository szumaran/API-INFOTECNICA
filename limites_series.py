import streamlit as st
import asyncio
import os
import re

# Importación de los módulos independientes de forma limpia
from infotecnica import ejecutar_extraccion_motor
from limites_series import buscar_limites_series_motor

# --- CONFIGURACIÓN DE LA PLATAFORMA ---
st.set_page_config(page_title="Sistema Integral Infotécnica", page_icon="⚡", layout="wide")

# Barra lateral para conmutar entre las dos herramientas homólogas
st.sidebar.title("🧭 Panel de Control")
aplicacion_activa = st.sidebar.radio(
    "Selecciona la herramienta de trabajo:",
    [
        "📊 1. Extractor Masivo Infotécnica",
        "⛓️ 2. Extractor de Elementos en Serie"
    ]
)

st.markdown("---")

# ==============================================================================
# ENTORNO 1: EXTRACTOR MASIVO INFOTÉCNICA
# ==============================================================================
if aplicacion_activa == "📊 1. Extractor Masivo Infotécnica":
    st.title("📊 Extractor Masivo Infotécnica (Input desde DIgSILENT)")
    st.write("Descarga los parámetros generales organizados en pestañas independientes (IM, T2D, Secciones Tramos)")
    
    input_ids = st.text_area(
        "Pega aquí la lista de IDs (Separados por comas, espacios o saltos de línea):",
        placeholder="Ejemplo:\n3789\n4521\n8912",
        height=200,
        key="txt_extractor_general"
    )

    tipo_busqueda = st.radio(
        "Selecciona cómo deseas procesar esta lista de IDs:",
        [
            "🔍 Modo Directo: Los IDs corresponden directamente a los Equipos (Interruptores, Transformadores)",
            "🛤️ Modo Tramo: Los IDs corresponden a Tramos (Buscar todos los equipos asociados)"
        ],
        key="radio_extractor_general"
    )

    st.markdown("---")
    boton_extraer = st.button("🚀 Iniciar Extracción Masiva Asíncrona", key="btn_extractor_general")

    if boton_extraer:
        if not input_ids.strip():
            st.warning("⚠️ Por favor, pega al menos un ID para iniciar.")
        else:
            try:
                list_ids = [int(x) for x in re.findall(r'\b\d+\b', input_ids)]
                if not list_ids: raise ValueError("No se encontraron números válidos.")
                
                es_modo_tramo = "Modo Tramo" in tipo_busqueda

                st.info(f"📋 Se identificaron {len(list_ids)} IDs únicos para el Extractor Masivo.")
                with St.spinner("Llamando a las APIs de Infotécnica..."):
                    loop = asyncio.new_event_loop()
                    asyncio.set_event_loop(loop)
                    try:
                        excel_path = loop.run_until_complete(
                            ejecutar_extraccion_motor(list_ids=list_ids, es_modo_tramo=es_modo_tramo)
                        )
                    finally:
                        loop.close()
                    
                if os.path.exists(excel_path):
                    st.success("🎉 ¡Extracción masiva finalizada con éxito!")
                    with open(excel_path, "rb") as file:
                        st.download_button(
                            label="📥 Descargar Parámetros en Excel (.xlsx)",
                            data=file,
                            file_name="parametros_infotecnica.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            except Exception as e:
                st.error(f"❌ Error: {e}")

# ==============================================================================
# ENTORNO 2: EXTRACTOR DE ELEMENTOS EN SERIE
# ==============================================================================
elif aplicacion_activa == "⛓️ 2. Extractor de Elementos en Serie":
    st.title("⛓️ Extractor de Elementos en Serie (Eslabón Más Débil)")
    st.write("Calcula el límite térmico global y capacidad de ruptura analizando el paño completo de forma vertical")
    
    input_ids_series = st.text_area(
        "Pega aquí la lista de IDs (Separados por comas, espacios o saltos de línea de Excel):",
        placeholder="Ejemplo:\n3789\n4521\n8912",
        height=200,
        key="txt_extractor_series"
    )

    tipo_busqueda_series = st.radio(
        "Selecciona el tipo de entrada para la búsqueda en serie:",
        [
            "🔍 Modo Directo: Los IDs corresponden a Equipos (Interruptores, Transformadores 2D/3D)",
            "🛤️ Modo Tramo: Los IDs corresponden a Secciones de Tramo (Buscar los paños de sus extremos)"
        ],
        key="radio_extractor_series"
    )

    st.markdown("---")
    boton_extraer_series = st.button("🚀 Iniciar Extracción de Objetos en Serie", key="btn_extractor_series")

    if boton_extraer_series:
        if not input_ids_series.strip():
            st.warning("⚠️ Por favor, pega al menos un ID para iniciar el análisis.")
        else:
            try:
                list_ids = [int(x) for x in re.findall(r'\b\d+\b', input_ids_series)]
                if not list_ids: raise ValueError("No se encontraron números válidos.")
                
                es_modo_tramo = "Modo Tramo" in tipo_busqueda_series

                st.info(f"📋 Se identificaron {len(list_ids)} IDs únicos para la búsqueda en serie.")
                
                status_text = st.empty()
                status_text.text("⚡ Iniciando conexión asíncrona con el Coordinador...")
                
                with st.spinner("Analizando elementos en serie del paño..."):
                    loop = asyncio.new_event_loop()
                    asyncio.set_event_loop(loop)
                    try:
                        status_text.text("📡 Cosechando fichas técnicas de los paños... Por favor espera.")
                        excel_path = loop.run_until_complete(
                            buscar_limites_series_motor(list_ids=list_ids, es_modo_tramo=es_modo_tramo)
                        )
                    finally:
                        loop.close()
                
                status_text.empty()
                    
                if os.path.exists(excel_path):
                    st.success("🎉 ¡Análisis de elementos en serie finalizado con éxito!")
                    with open(excel_path, "rb") as file:
                        st.download_button(
                            label="📥 Descargar Reporte de Serie (.xlsx)",
                            data=file,
                            file_name="limites_corriente_series.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            except Exception as e:
                st.error(f"❌ Error: {e}")

st.markdown("---")
st.caption("© 2026 Plataforma Eléctrica Integral - Módulos Modulares Homólogos")
