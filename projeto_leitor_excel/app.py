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

st.markdown("""
<style>
/* Esconde botões Share, estrela, lápis e menu hamburguer */
header button { display: none !important; }
#MainMenu { display: none !important; }
footer { display: none !important; }

/* Ícone GitHub fixo no canto superior direito */
.github-link {
    position: fixed;
    top: 12px;
    right: 16px;
    z-index: 9999;
    opacity: 0.7;
    transition: opacity 0.2s;
}
.github-link:hover { opacity: 1; }
.github-link svg { width: 26px; height: 26px; fill: #ffffff; }
</style>

<a class="github-link" href="https://github.com/eoCoRe/projeto-leitor-excel" target="_blank" title="Ver no GitHub">
  <svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
    <path d="M12 0C5.37 0 0 5.37 0 12c0 5.3 3.438 9.8 8.205 11.385.6.113.82-.258.82-.577
    0-.285-.01-1.04-.015-2.04-3.338.724-4.042-1.61-4.042-1.61-.546-1.387-1.333-1.756-1.333-1.756
    -1.09-.745.083-.729.083-.729 1.205.085 1.84 1.237 1.84 1.237 1.07 1.834 2.807 1.304
    3.492.997.108-.775.418-1.305.762-1.605-2.665-.3-5.466-1.332-5.466-5.93
    0-1.31.465-2.38 1.235-3.22-.135-.303-.54-1.523.105-3.176
    0 0 1.005-.322 3.3 1.23.96-.267 1.98-.399 3-.405
    1.02.006 2.04.138 3 .405 2.28-1.552 3.285-1.23 3.285-1.23
    .645 1.653.24 2.873.12 3.176.765.84 1.23 1.91 1.23 3.22
    0 4.61-2.805 5.625-5.475 5.92.42.36.81 1.096.81 2.22
    0 1.605-.015 2.896-.015 3.286 0 .315.21.69.825.57
    C20.565 21.795 24 17.295 24 12c0-6.63-5.37-12-12-12z"/>
  </svg>
</a>
""", unsafe_allow_html=True)

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
