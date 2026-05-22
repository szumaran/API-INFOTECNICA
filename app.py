import streamlit as st
import google.generativeai as genai
import asyncio
import os

# --- IMPORTACIÓN DEL CÓDIGO SEPARADO ---
# Importamos la función encargada de conectarse a las APIs desde infotecnica.py
from infotecnica import exportar_datos_async

# --- 1. CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Sistema Integral Psico-IA & Infotécnica", page_icon="🧠", layout="wide")

# --- 2. CONFIGURACIÓN DE LA IA ---
if "GEMINI_API_KEY" in st.secrets:
    try:
        genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
        model = genai.GenerativeModel('gemini-3-flash-preview')
    except Exception as e:
        st.error(f"Error al configurar la IA: {e}")
else:
    st.error("❌ Falta GEMINI_API_KEY en los Secrets.")

# --- 3. LOGIN ---
if "autenticado" not in st.session_state:
    st.title("🔐 Acceso al Sistema")
    with st.form("login"):
        u = st.text_input("Usuario")
        p = st.text_input("Contraseña", type="password")
        if st.form_submit_button("Entrar"):
            if u == "admin" and p == "centro2026":
                st.session_state["autenticado"] = True
                st.rerun()
            else:
                st.error("❌ Credenciales incorrectas")
    st.stop()

# --- 4. CONTROL DE NAVEGACIÓN EN SIDEBAR ---
st.sidebar.title("🧭 Navegación")
opcion_modulo = st.sidebar.radio(
    "Selecciona el Módulo de trabajo:",
    ["🧠 Generador de Informes (IA)", "⚡ Extractor Infotécnica (API)"]
)

st.sidebar.markdown("---")
if st.sidebar.button("🔒 Cerrar Sesión"):
    del st.session_state["autenticado"]
    st.rerun()


# ==============================================================================
# MÓDULO 1: GENERADOR DE INFORMES (TU CÓDIGO ORIGINAL)
# ==============================================================================
if opcion_modulo == "🧠 Generador de Informes (IA)":
    st.title("🧠 Generador Inteligente de Informes")
    st.write("Potenciado por Gemini 3 Flash Preview")
    st.markdown("---")

    col1, col2 = st.columns([1, 1])

    with col1:
        st.header("📥 Datos del Alumno")
        nombre = st.text_input("Nombre completo")
        curso = st.text_input("Curso")
        obs = st.text_area("Notas de la sesión", height=300, placeholder="Escribe tus notas aquí...")
        boton_generar = st.button("✨ Generar Informe")

    with col2:
        st.header("📄 Informe Sugerido")
        if boton_generar:
            if nombre and obs:
                with st.spinner("Redactando informe..."):
                    try:
                        prompt = f"Actúa como psicopedagogo clínico. Redacta un informe para {nombre} del curso {curso}. Notas: {obs}. Estructura con Identificación, Análisis y Recomendaciones."
                        response = model.generate_content(prompt)
                        
                        if response.text:
                            st.success("¡Informe generado con éxito!")
                            st.markdown(response.text)
                            st.download_button(
                                label="📥 Descargar TXT",
                                data=response.text,
                                file_name=f"Informe_{nombre}.txt",
                                mime="text/plain"
                            )
                    except Exception as e:
                        st.error(f"Error técnico: {e}")
            else:
                st.warning("⚠️ Rellena el nombre y las notas.")


# ==============================================================================
# MÓDULO 2: ASOCIACIÓN DEL EXTRACTOR DE INFOTÉCNICA
# ==============================================================================
elif opcion_modulo == "⚡ Extractor Infotécnica (API)":
    st.title("⚡ Extractor de Fichas Técnicas del Coordinador Eléctrico")
    st.write("Consulta directa asíncrona a la plataforma Infotécnica")
    st.markdown("---")
    
    st.info("Ingresa los IDs correspondientes. Si dejas los campos en blanco, se ejecutará el ID de Sección de Tramo 3789 por defecto.")
    
    # Inputs de la interfaz gráfica
    col_st, col_sub = st.columns(2)
    with col_st:
        input_st = st.text_input("IDs de Secciones de Tramos (Separados por coma)", placeholder="Ej: 3789")
    with col_sub:
        input_sub = st.text_input("IDs de Subestaciones (Separados por coma)", placeholder="Ej: 124, 85")
        
    por_tramos = st.checkbox("Buscar todos los equipos asociados a las Secciones de Tramos", value=False)
    
    st.markdown("---")
    boton_extraer = st.button("🚀 Iniciar Extracción Asíncrona")
    
    if boton_extraer:
        # Procesamiento de las entradas de texto a listas de enteros nativas
        try:
            list_st = [int(i.strip()) for i in input_st.split(",") if i.strip()] if input_st else []
            list_sub = [int(i.strip()) for i in input_sub.split(",") if i.strip()] if input_sub else []
            
            # Comportamiento por defecto si no se digita nada
            if not list_st and not list_sub:
                list_st = [3789]
                st.write("ℹ️ Campos vacíos. Usando ID de prueba `3789`")
            
            # Ejecución controlada mediante un Spinner de Streamlit
            with st.spinner("Conectando con la API y estructurando el archivo Excel..."):
                # Capturamos el bucle de eventos asíncronos activo y ejecutamos de forma segura en Streamlit
                loop = asyncio.get_event_loop()
                excel_path = loop.run_until_complete(
                    exportar_datos_async(ids_st=list_st, ids_sub=list_sub, por_secciones=por_tramos)
                )
                
            if os.path.exists(excel_path):
                st.success("🎉 ¡Extracción completada con éxito!")
                
                # Leemos el archivo en binario para habilitar la descarga directa en el navegador
                with open(excel_path, "rb") as file:
                    st.download_button(
                        label="📥 Descargar Libro de Excel (.xlsx)",
                        data=file,
                        file_name="equipos_datos.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
        except ValueError:
            st.error("❌ Error: Asegúrate de ingresar únicamente números enteros separados por comas.")
        except Exception as e:
            st.error(f"❌ Ocurrió un problema en la API: {e}")

st.markdown("---")
st.caption("© 2026 Centro Psicopedagógico & Plataforma Eléctrica")
