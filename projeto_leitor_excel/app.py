import importlib
import json
import sys

import streamlit as st

if "exportar_escoragem" in sys.modules:
    importlib.reload(sys.modules["exportar_escoragem"])

from exportar_escoragem import (
    _parse_classificacao,
    _parse_grupos,
    build_classificacao_html,
    build_preview_html,
    export_to_bytes,
)

st.set_page_config(
    page_title="Visualizador e Exportador",
    page_icon=None,
    layout="wide",
)

st.title("Visualizador e Exportador")

input_mode = st.radio(
    "Fonte dos dados",
    ["Upload de arquivo", "Colar JSON"],
    horizontal=True,
)

json_bytes: bytes | None = None

if input_mode == "Upload de arquivo":
    uploaded = st.file_uploader("Selecione o arquivo JSON de escoragem", type=["json"])
    if uploaded:
        json_bytes = uploaded.read()
else:
    pasted = st.text_area(
        "Cole o conteúdo JSON aqui",
        height=200,
        placeholder='{"detalhes": "Nome do modelo", "grupos": [...], "classificacao": [...]}',
    )
    if pasted and pasted.strip():
        json_bytes = pasted.encode("utf-8")

if json_bytes:
    try:
        data = json.loads(json_bytes)
    except json.JSONDecodeError as e:
        st.error(f"JSON inválido: {e}")
        st.stop()

    model_name = data.get("detalhes", "—")
    st.subheader(f"Modelo: {model_name}")

    grupos = _parse_grupos(data)
    classificacoes = _parse_classificacao(data)
    total_vars = sum(len(g["variaveis"]) for g in grupos)

    m1, m2, m3 = st.columns(3)
    m1.metric("Grupos", len(grupos))
    m2.metric("Variáveis", total_vars)
    m3.metric("Classificações", len(classificacoes))

    st.divider()

    tab1, tab2 = st.tabs(["Variáveis e Regras", "Classificação de Risco"])

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
            st.error(f"Erro ao gerar classificação: {e}")
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
