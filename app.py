import streamlit as st
import asyncio
import os
import re

# Importamos los dos motores aislados
from infotecnica import ejecutar_extraccion_motor
from limites_series import buscar_limites_series_motor

st.set_page_config(page_title="Extractor Infotécnica - Disgilent", page_icon="⚡", layout="wide")

st.title("⚡ Extractor Masivo Infotécnica (Input desde Disgilent)")
st.write("Pega tu lista de IDs directamente para descargar los parámetros desde el Coordinador Eléctrico")
st.markdown("---")

input_ids = st.text_area(
    "Pega aquí la lista de IDs (Puedes pegarlos separados por comas, espacios o uno por fila copiado de Excel/Disgilent):",
    placeholder="Ejemplo:\n3789\n4521\n8912",
    height=200
)

tipo_busqueda = st.radio(
    "Selecciona cómo deseas procesar esta lista de IDs:",
    [
        "🔍 Modo Directo: Los IDs corresponden directamente a los Equipos (Interruptores, Transformadores, etc.)",
        "🛤️ Modo Tramo: Los IDs corresponden a Tramos (Buscar todos los equipos asociados a estos tramos)",
        "⛓️ Modo Objetos en Serie: Evaluar el paño completo del elemento y calcular el Límite de Corriente global (Eslabón más débil)"
    ]
)

st.markdown("---")
boton_extraer = st.button("🚀 Iniciar Extracción Masiva Asíncrona")

if boton_extraer:
    if not input_ids.strip():
        st.warning("⚠️ Por favor, pega al menos un ID para iniciar la extracción.")
    else:
        try:
            list_ids = [int(x) for x in re.findall(r'\b\d+\b', input_ids)]
            if not list_ids:
                raise ValueError("No se encontraron números válidos en el cuadro de texto.")

            st.info(f"📋 Se identificaron {len(list_ids)} IDs únicos para procesar.")

            with st.spinner("Llamando a las APIs del Coordinador Eléctrico..."):
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                
                try:
                    # Enrutamiento modular e independiente
                    if "Modo Objetos en Serie" in tipo_busqueda:
                        excel_path = loop.run_until_complete(buscar_limites_series_motor(list_ids=list_ids))
                    else:
                        es_modo_tramo = "Modo Tramo" in tipo_busqueda
                        excel_path = loop.run_until_complete(ejecutar_extraccion_motor(list_ids=list_ids, es_modo_tramo=es_modo_tramo))
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
                    
        except ValueError as ve:
            st.error(f"❌ Error de entrada: {ve}")
        except Exception as e:
            st.error(f"❌ Error en el proceso de Infotécnica: {e}")

st.markdown("---")
st.caption("© 2026 Plataforma Eléctrica - Conector de Parámetros Masivos")
