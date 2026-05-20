import streamlit as st

# Configuración
st.set_page_config(page_title="Test Estructura", layout="wide")

# 1. BARRA LATERAL (AGENTE)
with st.sidebar:
    st.title("🤖 Agente de IA")
    st.info("Aquí está el chat del agente listo para el desarrollo que involucra IA.")
    st.chat_input("Escribe al agente...")

# 2. CUERPO PRINCIPAL
st.title("⚡ Visualización de Plataforma Infotécnica")
st.markdown("---")

# 3. LOS 3 BOTONES EN COLUMNAS
col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("🆔 XID")
    st.text_input("Parámetro XID", "XID-001")
    st.button("Ejecutar XID", type="primary", use_container_width=True)

with col2:
    st.subheader("📉 Lineales")
    st.text_input("ID Línea", "L-100")
    st.button("Ejecutar Lineales", type="primary", use_container_width=True)

with col3:
    st.subheader("🏢 Subestación")
    st.text_input("Nombre SE", "Norte")
    st.button("Ejecutar Subestación", type="primary", use_container_width=True)

st.markdown("---")
st.write("✅ **Si ves esto, la estructura es correcta.** Los botones dispararán tus otros códigos de Python cuando los pegues en el archivo `motores.py`.")
