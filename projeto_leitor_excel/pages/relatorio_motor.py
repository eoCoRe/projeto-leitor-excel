"""
Relatório do Motor de Crédito
Analisa o JSON completo de um motor (estrutura_motor.componentes) e exibe um
resumo estruturado: fontes de dados, decisões, condicionais, escoragem e variáveis.
"""

import html as _html
import json

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Relatório do Motor", layout="wide")


_HTML_ENT = {"&gt;=": ">=", "&lt;=": "<=", "&gt;": ">", "&lt;": "<", "&amp;": "&"}


def _decode(text) -> str:
    if not isinstance(text, str):
        return str(text) if text is not None else ""
    for k, v in _HTML_ENT.items():
        text = text.replace(k, v)
    return text


def _tiptap_text(doc) -> str:
    """Extrai texto legível de uma string simples ou de um documento tiptap."""
    if isinstance(doc, str):
        return _decode(doc)
    if not isinstance(doc, dict) or doc.get("type") != "doc":
        return str(doc) if doc else ""
    parts: list[str] = []
    for block in doc.get("content", []):
        for inline in block.get("content", []):
            t = inline.get("type")
            if t == "text":
                txt = inline.get("text", "")
                if txt:
                    parts.append(txt)
            elif t == "variableComponent":
                label = inline.get("attrs", {}).get("label", "")
                if label:
                    parts.append(f"[{label}]")
            elif t == "mathComponent":
                for child in inline.get("content", []):
                    if child.get("type") == "variableComponent":
                        label = child.get("attrs", {}).get("label", "")
                        if label:
                            parts.append(f"[{label}]")
                    elif child.get("type") == "text":
                        txt = child.get("text", "")
                        if txt:
                            parts.append(txt)
    return " ".join(parts).strip()


def _categorize(componentes: list) -> dict:
    buckets: dict[str, list] = {
        "start": [],
        "etl": [],
        "status": [],
        "conditional": [],
        "scoring": [],
        "variables": [],
        "other": [],
    }
    for comp in componentes:
        t = comp.get("type", "")
        if t == "startNode":
            buckets["start"].append(comp)
        elif t in ("etlNode", "etlNodeV2"):
            buckets["etl"].append(comp)
        elif t in ("statusNode", "statusNodeV2"):
            buckets["status"].append(comp)
        elif "condicional" in t.lower():
            buckets["conditional"].append(comp)
        elif "escoragem" in t.lower():
            buckets["scoring"].append(comp)
        elif t == "variaveisNode":
            buckets["variables"].append(comp)
        else:
            buckets["other"].append(comp)
    return buckets


_STATUS_EMOJI = {"reprovar": "🔴", "derivar": "🟡", "aprovar": "🟢"}
_STATUS_LABEL = {"reprovar": "Reprovar", "derivar": "Derivar", "aprovar": "Aprovar"}

_STATUS_COLORS = {
    "reprovar": ("#FFEEEE", "#C00000"),
    "derivar":  ("#FFFBE6", "#9B6A00"),
    "aprovar":  ("#EEFFEE", "#006400"),
}


def _status_card(comp: dict) -> None:
    d = comp.get("data", {})
    status = d.get("status", "")
    final = d.get("final", False)
    msg_raw = d.get("message", "")
    msg = _tiptap_text(msg_raw) if isinstance(msg_raw, dict) else str(msg_raw)

    bg, fg = _STATUS_COLORS.get(status, ("#F5F5F5", "#333333"))
    emoji = _STATUS_EMOJI.get(status, "⚪")
    badge = "Final" if final else "Intermediário"

    st.markdown(
        f"""<div style="background:{bg};border-left:4px solid {fg};
            border-radius:6px;padding:10px 14px;margin-bottom:6px;">
            <span style="font-weight:bold;color:{fg}">{emoji} {badge}</span>
            <span style="font-size:12px;color:#555;margin-left:8px;">
                {_html.escape(msg)}
            </span></div>""",
        unsafe_allow_html=True,
    )


st.title("Relatório do Motor de Crédito")

input_mode = st.radio(
    "Origem do JSON",
    ["Colar JSON", "Upload de arquivo"],
    horizontal=True,
)

json_text: str | None = None

if input_mode == "Colar JSON":
    json_text = st.text_area(
        "Cole o JSON completo do motor aqui",
        height=220,
        placeholder='{"configuracao": {}, "estrutura_motor": {"componentes": [...]}}',
    )
else:
    uploaded = st.file_uploader("Selecione o arquivo JSON do motor", type=["json"])
    if uploaded:
        json_text = uploaded.read().decode("utf-8")

if not json_text or not json_text.strip():
    st.info("Cole ou faça upload do JSON do motor para gerar o relatório.")
    st.stop()

try:
    data = json.loads(json_text)
except json.JSONDecodeError as e:
    st.error(f"JSON inválido: {e}")
    st.stop()

componentes: list = data.get("estrutura_motor", {}).get("componentes", [])
if not componentes:
    st.warning("Nenhum componente encontrado em `estrutura_motor.componentes`.")
    st.stop()

buckets = _categorize(componentes)


c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Nós totais", len(componentes))
c2.metric("Fontes de dados", len(buckets["etl"]))
c3.metric("Decisões", len(buckets["status"]))
c4.metric("Condicionais", len(buckets["conditional"]))
c5.metric("Escoragem", len(buckets["scoring"]))

st.divider()


tab_sources, tab_status, tab_cond, tab_score, tab_vars, tab_tech = st.tabs([
    f"Fontes de Dados ({len(buckets['etl'])})",
    f"Decisões ({len(buckets['status'])})",
    f"Condicionais ({len(buckets['conditional'])})",
    f"Escoragem ({len(buckets['scoring'])})",
    f"Variáveis ({len(buckets['variables'])})",
    "Resumo Técnico",
])

with tab_sources:
    st.subheader("Fontes de Dados Consultadas")
    if not buckets["etl"]:
        st.info("Nenhum nó ETL encontrado.")
    else:
        rows = []
        for comp in buckets["etl"]:
            d = comp.get("data", {})
            label = d.get("label", "—")
            comp_id = d.get("componentId", "—")
            version = "v2" if comp.get("type") == "etlNodeV2" else "v1"

            doc_raw = d.get("documento", "")
            if isinstance(doc_raw, list):
                doc_str = ", ".join(str(x) for x in doc_raw)
            elif isinstance(doc_raw, dict):
                doc_str = _tiptap_text(doc_raw)
            else:
                doc_str = str(doc_raw)
            doc_str = doc_str.replace("{{input_inicial.documento}}", "Documento de entrada")

            rows.append({
                "Nome": label,
                "Componente ID": comp_id,
                "Versão API": version,
                "Documento consultado": doc_str,
            })

        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

with tab_status:
    st.subheader("Decisões do Motor")

    def _group_status(nodes: list, group_label: str) -> None:
        if not nodes:
            return
        st.markdown(f"**{group_label}** — {len(nodes)} ocorrências")
        for comp in nodes:
            _status_card(comp)
        st.write("")

    reprovar_final   = [c for c in buckets["status"]
                        if c["data"].get("status") == "reprovar" and c["data"].get("final")]
    reprovar_mid     = [c for c in buckets["status"]
                        if c["data"].get("status") == "reprovar" and not c["data"].get("final")]
    derivar_nodes    = [c for c in buckets["status"] if c["data"].get("status") == "derivar"]
    aprovar_nodes    = [c for c in buckets["status"] if c["data"].get("status") == "aprovar"]

    _group_status(reprovar_final, "🔴 Reprovações Finais")
    _group_status(derivar_nodes,  "🟡 Derivações (análise manual)")
    _group_status(aprovar_nodes,  "🟢 Aprovações")
    _group_status(reprovar_mid,   "🔴 Reprovações Intermediárias")

with tab_cond:
    st.subheader("Regras Condicionais")
    if not buckets["conditional"]:
        st.info("Nenhum nó condicional encontrado.")
    else:
        rows = []
        for comp in buckets["conditional"]:
            d = comp.get("data", {})
            detalhes = d.get("detalhes", "—")
            operador = _decode(d.get("condicao", ""))
            tipo_text  = _tiptap_text(d.get("tipo", ""))
            valor_text = _tiptap_text(d.get("valor", ""))

            tipo_clean = tipo_text.replace("[", "").replace("]", "")
            valor_clean = valor_text.replace("[", "").replace("]", "")

            rows.append({
                "Pergunta / Detalhe": detalhes,
                "Variável verificada": tipo_clean,
                "Operador": operador,
                "Valor comparado": valor_clean,
            })

        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

with tab_score:
    st.subheader("Nós de Escoragem")
    if not buckets["scoring"]:
        st.info("Nenhum nó de escoragem encontrado.")
    else:
        for idx, comp in enumerate(buckets["scoring"]):
            d = comp.get("data", {})
            label    = d.get("label", "Escoragem")
            detalhes = d.get("detalhes", "")
            title    = label + (f" — {detalhes}" if detalhes else "")

            with st.expander(title, expanded=(idx == 0)):
                esc = d.get("escoragem", {})
                grupos = esc.get("grupos", [])
                for grupo in grupos:
                    st.markdown(f"**Grupo: {grupo.get('nome', '')}**")
                    variaveis = grupo.get("variaveis", [])
                    var_rows = []
                    for var in variaveis:
                        nome = var.get("nome", "")
                        peso_raw = var.get("peso", "")
                        peso_fmt = (
                            f"{float(peso_raw) * 100:.0f}%"
                            if isinstance(peso_raw, float) and 0 < peso_raw <= 1
                            else str(peso_raw)
                        )
                        for p in var.get("pontuacao", []):
                            var_rows.append({
                                "Variável": nome,
                                "Peso": peso_fmt,
                                "Condição": _decode(str(p.get("condicao_1", ""))),
                                "Pontuação base": p.get("pontuacao_1", ""),
                                "Score faixa": p.get("faixa_pontuacao", ""),
                            })
                    if var_rows:
                        st.dataframe(
                            pd.DataFrame(var_rows),
                            use_container_width=True,
                            hide_index=True,
                        )

                classificacao = esc.get("classificacao", [])
                if classificacao:
                    st.markdown("**Classificação de risco:**")
                    cls_rows = []
                    for cls in classificacao:
                        cls_rows.append({
                            "Classe / Nome": cls.get("nome", cls.get("classe", "")),
                            "Risco": cls.get("risco", ""),
                            "Condição": _decode(str(cls.get("condicao_1", ""))),
                            "Pontuação": cls.get("pontuacao_1", ""),
                            "Fator": cls.get("faixa_pontuacao", ""),
                            "Validade (dias)": cls.get("validade_limite", ""),
                        })
                    st.dataframe(
                        pd.DataFrame(cls_rows),
                        use_container_width=True,
                        hide_index=True,
                    )

with tab_vars:
    st.subheader("Nós de Variáveis Customizadas")
    if not buckets["variables"]:
        st.info("Nenhum nó de variáveis encontrado.")
    else:
        for idx, comp in enumerate(buckets["variables"]):
            d = comp.get("data", {})
            label    = d.get("label", "Variáveis")
            detalhes = d.get("detalhes", "")
            title    = label + (f" — {detalhes}" if detalhes else "")

            with st.expander(title, expanded=(idx == 0)):
                variaveis = d.get("variaveis", [])
                rows = []
                for var in variaveis:
                    val_raw = var.get("valor", "")
                    val_text = (
                        _tiptap_text(val_raw) if isinstance(val_raw, dict) else str(val_raw)
                    )
                    rows.append({
                        "Nome interno": var.get("nome_variavel", var.get("nome", "")),
                        "Label": var.get("label", ""),
                        "Tipo": var.get("tipo", ""),
                        "Valor / Fórmula": val_text,
                        "Ocultar": "Sim" if var.get("ocultar") else "Não",
                    })
                if rows:
                    st.dataframe(
                        pd.DataFrame(rows),
                        use_container_width=True,
                        hide_index=True,
                    )

with tab_tech:
    st.subheader("Resumo Técnico")

    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown("**Distribuição por tipo de nó:**")
        type_counts: dict[str, int] = {}
        for comp in componentes:
            t = comp.get("type", "unknown")
            type_counts[t] = type_counts.get(t, 0) + 1
        type_rows = [
            {"Tipo de nó": t, "Quantidade": cnt}
            for t, cnt in sorted(type_counts.items(), key=lambda x: -x[1])
        ]
        st.dataframe(pd.DataFrame(type_rows), use_container_width=True, hide_index=True)

    with col_right:
        cfg = data.get("configuracao", {})
        st.markdown("**Configuração do motor:**")
        url_hook = cfg.get("urlExternalHook", "")
        st.markdown(f"- **URL External Hook:** `{url_hook or '(não definida)'}`")
        st.markdown(f"- **Headers:** {len(cfg.get('headers', []))}")
        st.markdown(f"- **Body params:** {len(cfg.get('body', []))}")

        st.markdown("")
        st.markdown("**Desfechos finais possíveis:**")
        final_nodes = [
            c for c in buckets["status"]
            if c.get("data", {}).get("final", False)
        ]
        if final_nodes:
            for comp in final_nodes:
                d = comp.get("data", {})
                status = d.get("status", "").upper()
                msg_raw = d.get("message", "")
                msg = _tiptap_text(msg_raw) if isinstance(msg_raw, dict) else str(msg_raw)
                if len(msg) > 110:
                    msg = msg[:107] + "…"
                emoji = _STATUS_EMOJI.get(d.get("status", ""), "⚪")
                st.markdown(f"- {emoji} **{status}**: {msg}")
        else:
            st.info("Nenhum nó de status final encontrado.")
