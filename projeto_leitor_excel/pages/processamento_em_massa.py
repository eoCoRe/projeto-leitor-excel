"""
Processamento em Massa
Versao integrada da ferramenta local "envio-massa-api.html": anexa uma
planilha, monta uma requisicao por linha a partir de um template com
placeholders {{Coluna}} e envia para uma API (ex.: disparo de esteiras de
credito), com token Bearer, intervalo controlado entre envios e modo de
teste (dry-run) que monta as requisicoes sem enviar de verdade. Rodar o
envio no servidor (em vez de fetch no navegador) evita bloqueios de CORS.
"""

import json
import re
import time

import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="Processamento em Massa", layout="wide")

_PLACEHOLDER_RE = re.compile(r"\{\{\s*([^}]+?)\s*\}\}")

DEFAULT_URL = "https://app-api.neocredit.com.br/empresa-esteira-solicitacao/3142/integracao"
DEFAULT_HEADERS = '{"accept": "application/json"}'
DEFAULT_BODY = '{"documento": "{{documento}}"}'


def _cell_to_str(val) -> str:
    """Converte uma celula da planilha para string, evitando o problema
    classico de numero inteiro (ex.: CNPJ) virar "123.0" quando a coluna
    e lida como float pelo pandas/openpyxl."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()


def _fill_template(node, row: pd.Series):
    if isinstance(node, str):
        return _PLACEHOLDER_RE.sub(lambda m: _cell_to_str(row.get(m.group(1))), node)
    if isinstance(node, list):
        return [_fill_template(v, row) for v in node]
    if isinstance(node, dict):
        return {k: _fill_template(v, row) for k, v in node.items()}
    return node


st.title("Processamento em Massa")
st.caption(
    "Envia uma planilha linha a linha para uma API (ex.: disparar uma esteira "
    "de crédito para cada CNPJ). Monta o corpo da requisição a partir de um "
    "template com placeholders `{{Coluna}}`, com intervalo controlado entre "
    "envios e um modo de teste (dry-run) que monta as requisições sem enviar "
    "de verdade. Nada é enviado até você clicar em \"Iniciar envio\"."
)

st.subheader("1. Requisição")
col_url, col_method = st.columns([3, 1])
with col_url:
    url = st.text_input("URL do endpoint", value=DEFAULT_URL)
with col_method:
    metodo = st.selectbox("Método", ["POST", "GET", "PUT", "PATCH", "DELETE"])

col_delay, col_token = st.columns([1, 2])
with col_delay:
    delay_ms = st.number_input("Intervalo entre envios (ms)", min_value=0, step=100, value=2000)
with col_token:
    token = st.text_input(
        "Bearer Token",
        type="password",
        placeholder="Cole o token aqui (sem a palavra 'Bearer')",
        help="Fica só nesta sessão do navegador — não é salvo em nenhum arquivo.",
    )

headers_text = st.text_area(
    "Headers extras (JSON)",
    value=DEFAULT_HEADERS,
    height=70,
    help="Content-Type e Authorization são adicionados automaticamente.",
)

if "pm_body" not in st.session_state:
    st.session_state.pm_body = DEFAULT_BODY
st.text_area(
    "Template do corpo (JSON)",
    key="pm_body",
    height=90,
    help="Use {{NomeDaColuna}} para inserir valores da planilha.",
)

st.subheader("2. Planilha")
uploaded = st.file_uploader("Selecione a planilha", type=["xlsx", "xls", "csv"])

if not uploaded:
    st.info("Envie uma planilha (.xlsx, .xls ou .csv) para continuar.")
    st.stop()

if uploaded.name.lower().endswith(".csv"):
    df_raw = pd.read_csv(uploaded, sep=None, engine="python")
    aba = ""
else:
    xls = pd.ExcelFile(uploaded)
    if len(xls.sheet_names) > 1:
        aba = st.selectbox("Aba", xls.sheet_names)
    else:
        aba = xls.sheet_names[0]
    df_raw = pd.read_excel(xls, sheet_name=aba)

if df_raw.empty:
    st.warning("A planilha/aba selecionada está vazia.")
    st.stop()

df_raw = df_raw.reset_index(drop=True)
data_key = f"{uploaded.name}-{uploaded.size}-{aba}"

if st.session_state.get("pm_data_key") != data_key:
    st.session_state.pm_data_key = data_key
    st.session_state.pm_selected = [True] * len(df_raw)
    st.session_state.pm_nonce = 0
    st.session_state.pop("pm_results", None)

col_insert, col_btn = st.columns([3, 1])
with col_insert:
    placeholder_col = st.selectbox("Inserir coluna no template do corpo", df_raw.columns)
with col_btn:
    st.write("")
    if st.button("Inserir placeholder", use_container_width=True):
        st.session_state.pm_body += "{{" + placeholder_col + "}}"
        st.rerun()

st.caption(f"{len(df_raw)} linha(s) carregada(s) de \"{uploaded.name}\".")

st.subheader("3. Linhas a enviar")
b1, b2, b3, b4 = st.columns([1, 1, 2, 1])
with b1:
    if st.button("Selecionar todas", use_container_width=True):
        st.session_state.pm_selected = [True] * len(df_raw)
        st.session_state.pm_nonce += 1
        st.rerun()
with b2:
    if st.button("Nenhuma", use_container_width=True):
        st.session_state.pm_selected = [False] * len(df_raw)
        st.session_state.pm_nonce += 1
        st.rerun()
with b3:
    n_primeiras = st.number_input("Primeiras N", min_value=1, value=3, label_visibility="collapsed")
with b4:
    if st.button("Selecionar N", use_container_width=True):
        n = int(n_primeiras)
        st.session_state.pm_selected = [i < n for i in range(len(df_raw))]
        st.session_state.pm_nonce += 1
        st.rerun()

colunas = df_raw.columns.tolist()
default_filter_col = 0
default_filter_val = ""
for candidato, valor in (("CNPJ_ENVIO_FINAL", ""), ("TIPO_RECEITA_FEDERAL", "MATRIZ")):
    if candidato in colunas:
        default_filter_col = colunas.index(candidato)
        default_filter_val = valor
        break

f1, f2, f3 = st.columns([2, 2, 1])
with f1:
    filtro_col = st.selectbox("Filtro de validação — coluna", colunas, index=default_filter_col)
with f2:
    filtro_val = st.text_input(
        "Valor (vazio = só linhas preenchidas)",
        value=default_filter_val,
        placeholder="ex: MATRIZ",
    )
with f3:
    st.write("")
    if st.button("Aplicar filtro", use_container_width=True):
        val = filtro_val.strip().upper()
        if val == "":
            nova_selecao = [_cell_to_str(x) != "" for x in df_raw[filtro_col]]
        else:
            nova_selecao = [_cell_to_str(x).upper() == val for x in df_raw[filtro_col]]
        st.session_state.pm_selected = nova_selecao
        st.session_state.pm_nonce += 1
        st.rerun()

view_df = df_raw.copy()
view_df.insert(0, "Enviar", st.session_state.pm_selected)

edited_df = st.data_editor(
    view_df,
    column_config={"Enviar": st.column_config.CheckboxColumn()},
    disabled=colunas,
    hide_index=True,
    use_container_width=True,
    height=420,
    key=f"pm_editor_{data_key}_{st.session_state.pm_nonce}",
)
st.session_state.pm_selected = edited_df["Enviar"].tolist()
selecionadas = df_raw.loc[edited_df["Enviar"]]

st.caption(f"{len(selecionadas)} linha(s) selecionada(s) de {len(df_raw)}.")

st.subheader("4. Enviar")
dry_run = st.checkbox(
    "Modo teste (dry-run) — monta as requisições mas NÃO envia de verdade",
    value=True,
)
stop_on_error = st.checkbox("Parar no primeiro erro", value=False)

confirmar_envio_real = True
if not dry_run:
    confirmar_envio_real = st.checkbox(
        f"Confirmo que quero enviar {len(selecionadas)} requisição(ões) REAIS para {url or '(defina a URL acima)'} "
        "— isso não é reversível do lado do servidor de destino."
    )

iniciar = st.button(
    "Iniciar envio",
    type="primary",
    disabled=len(selecionadas) == 0,
)

if iniciar:
    try:
        template = json.loads(st.session_state.pm_body)
    except json.JSONDecodeError as e:
        st.error(f"O template do corpo não é um JSON válido: {e}")
        st.stop()

    try:
        extra_headers = json.loads(headers_text) if headers_text.strip() else {}
    except json.JSONDecodeError as e:
        st.error(f"Headers extras não são um JSON válido: {e}")
        st.stop()

    if not url.strip():
        st.error("Informe a URL do endpoint.")
        st.stop()
    if not dry_run and not token.strip():
        st.error("Informe o Bearer Token para enviar de verdade (ou marque o modo teste).")
        st.stop()
    if not dry_run and not confirmar_envio_real:
        st.error("Marque a confirmação de envio real antes de iniciar.")
        st.stop()

    progress = st.progress(0.0)
    live_area = st.empty()
    resultados = []
    total = len(selecionadas)

    for i, (idx, row) in enumerate(selecionadas.iterrows()):
        body = _fill_template(template, row)
        headers = {"accept": "application/json", **extra_headers}
        if token.strip():
            headers["Authorization"] = f"Bearer {token.strip()}"

        if dry_run:
            resultado = {
                "linha": idx + 1,
                "status": "dry-run",
                "http": "",
                "detalhe": json.dumps(body, ensure_ascii=False),
            }
        else:
            req_kwargs = {"headers": headers, "timeout": 30}
            if metodo not in ("GET", "HEAD"):
                req_kwargs["json"] = body
            try:
                resp = requests.request(metodo, url, **req_kwargs)
                detalhe = resp.text[:300]
                resultado = {
                    "linha": idx + 1,
                    "status": "sucesso" if resp.ok else "erro",
                    "http": resp.status_code,
                    "detalhe": detalhe,
                }
            except requests.RequestException as e:
                resultado = {"linha": idx + 1, "status": "erro", "http": "", "detalhe": str(e)}

        resultados.append(resultado)
        progress.progress((i + 1) / total)
        live_area.dataframe(pd.DataFrame(resultados), use_container_width=True, hide_index=True)

        if resultado["status"] == "erro" and stop_on_error:
            break
        if i < total - 1 and delay_ms > 0:
            time.sleep(delay_ms / 1000)

    live_area.empty()
    progress.empty()
    st.session_state.pm_results = resultados

if st.session_state.get("pm_results"):
    resultados = st.session_state.pm_results
    df_res = pd.DataFrame(resultados)
    ok = int((df_res["status"] != "erro").sum())
    err = int((df_res["status"] == "erro").sum())

    c1, c2, c3 = st.columns(3)
    c1.metric("Enviadas", len(df_res))
    c2.metric("Sucesso", ok)
    c3.metric("Erro", err)

    if err:
        st.warning(f'{err} requisição(ões) falharam. Revise a coluna "detalhe" antes de reenviar.')

    st.dataframe(df_res, use_container_width=True, hide_index=True)

    st.download_button(
        "⬇️ Baixar resultado (CSV)",
        data=df_res.to_csv(index=False, sep=";").encode("utf-8-sig"),
        file_name="resultado_envio.csv",
        mime="text/csv",
        use_container_width=True,
    )
