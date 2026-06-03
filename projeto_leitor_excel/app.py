import importlib
import json
import sys

import streamlit as st

if "exportar_escoragem" in sys.modules:
    importlib.reload(sys.modules["exportar_escoragem"])

from exportar_escoragem import (
    build_classificacao_html,
    build_preview_html,
    export_to_bytes,
)

st.set_page_config(
    page_title="Motor de Escoragem",
    page_icon=None,
    layout="wide",
)

st.title("Motor de Escoragem — Visualizador e Exportador")

uploaded = st.file_uploader("Selecione o arquivo JSON de escoragem", type=["json"])

if uploaded:
    json_bytes = uploaded.read()

    try:
        data = json.loads(json_bytes)
    except json.JSONDecodeError as e:
        st.error(f"JSON invalido: {e}")
        st.stop()

    model_name = data.get("detalhes", uploaded.name)
    st.subheader(f"Modelo: {model_name}")

    tab1, tab2 = st.tabs(["Variaveis e Regras", "Classificacao de Risco"])

    with tab1:
        try:
            st.html(build_preview_html(data))
        except Exception as e:
            st.error(f"Erro ao gerar preview: {e}")
            st.exception(e)

    with tab2:
        try:
            st.html(build_classificacao_html(data))
        except Exception as e:
            st.error(f"Erro ao gerar classificacao: {e}")
            st.exception(e)

    st.divider()

    try:
        filename, excel_bytes = export_to_bytes(json_bytes)
        st.download_button(
            label="Baixar Excel (.xlsx)",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as e:
        st.error(f"Erro ao gerar Excel: {e}")
        st.exception(e)
