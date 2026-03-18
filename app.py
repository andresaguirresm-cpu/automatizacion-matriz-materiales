"""
app.py — Streamlit UI para la Automatización de la Matriz de Materiales Digitales.
"""
import io
import re
import streamlit as st
import pandas as pd

from src.flow_parser import parse_flow
from src.base_loader import load_base
from src.matrix_builder import build_matrix
from src.excel_writer import write_excel


st.set_page_config(
    page_title="Matriz de Materiales Digitales",
    page_icon="📊",
    layout="wide",
)

st.title("Automatización — Matriz de Materiales Digitales")
st.markdown(
    "Sube el **FLOW de medios** aprobado (.xlsx) y descarga la **Matriz de Materiales** "
    "lista para el equipo de creatividad."
)

# ── Nombre de campaña
campaign_name = st.text_input(
    "Nombre de campaña",
    value="CAMPAÑA",
    help="Se usará en el título del Excel generado.",
)

# ── File uploader
uploaded_file = st.file_uploader(
    "Subir FLOW de medios (.xlsx)",
    type=["xlsx"],
    help="Solo subir el FLOW. La MATRIZ BASE se carga automáticamente.",
)

if uploaded_file is not None:
    with st.spinner("Procesando FLOW..."):
        try:
            # 1. Parsear el FLOW
            flow_df = parse_flow(uploaded_file)
        except Exception as e:
            st.error(f"Error al parsear el FLOW: {e}")
            st.stop()

        try:
            # 2. Cargar la BASE
            specs = load_base()
        except Exception as e:
            st.error(f"Error al cargar la MATRIZ BASE: {e}")
            st.stop()

        # 3. Construir la matriz
        try:
            rows, warnings = build_matrix(flow_df, specs)
        except Exception as e:
            st.error(f"Error al construir la matriz: {e}")
            st.stop()

    # ── Preview de combinaciones extraídas
    st.subheader("Combinaciones únicas extraídas del FLOW")
    st.dataframe(
        flow_df.rename(columns={
            "etapa": "ETAPA",
            "sub_etapa": "SUB ETAPA",
            "medio": "MEDIO",
            "formato": "FORMATO",
            "mensaje": "MENSAJE",
        }),
        use_container_width=True,
        hide_index=True,
    )
    st.caption(f"Total: {len(flow_df)} combinaciones únicas (ETAPA + SUB ETAPA + MEDIO + FORMATO)")

    # ── Warnings de matching
    missing_warnings = [w for w in warnings if "no encontrado" in w.lower()]
    fuzzy_warnings = [w for w in warnings if "fuzzy" in w.lower()]

    can_download = True

    if missing_warnings:
        st.warning(
            f"⚠️ {len(missing_warnings)} formato(s) no encontrado(s) en la BASE. "
            "Las columnas de specs quedarán vacías para esos formatos."
        )
        with st.expander("Ver formatos no encontrados"):
            for w in missing_warnings:
                st.write(f"- {w}")
        can_download = st.checkbox(
            "Continuar de todos modos (algunos specs quedarán vacíos)",
            value=False,
        )

    if fuzzy_warnings:
        with st.expander(f"ℹ️ {len(fuzzy_warnings)} match(es) aproximado(s) — verificar"):
            for w in fuzzy_warnings:
                st.write(f"- {w}")

    # ── Botón de descarga
    if can_download or not missing_warnings:
        try:
            excel_bytes = write_excel(rows, campaign_name=campaign_name)
        except Exception as e:
            st.error(f"Error al generar el Excel: {e}")
            st.stop()

        # Nombre limpio del archivo
        safe_name = re.sub(r"[^\w\s-]", "", campaign_name).strip().replace(" ", "_").upper()
        filename = f"MATRIZ_DIGITAL_{safe_name}.xlsx"

        st.download_button(
            label="⬇️ Descargar Matriz de Materiales",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
