"""
Conversor de Escoragem: formato antigo -> formato novo
Converte nos `escoragemNode` (formato antigo) de um motor de credito completo
para `escoragemNodeV2` (formato novo), preenchendo a metadata de cada
`variavel` (label, nodeWorkflowId, nodeWorkflowType, variableId, variableType)
a partir de referencias {{...}} identicas ja presentes em outras partes do
proprio motor (nos condicionais, outros nos ja migrados etc).
"""

import json
import re

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Atualizar Escoragens", layout="wide")

_VAR_RE = re.compile(r"\{\{.*?\}\}")
_PLACEHOLDER_LABEL = "[REVISAR] campo nao encontrado no motor"


def _walk_variable_components(node, index: dict) -> None:
    """Percorre recursivamente o motor procurando nos tiptap
    'variableComponent' e indexa sua metadata pelo conteudo {{...}} que
    carregam, para reaproveitar em referencias identicas."""
    if isinstance(node, dict):
        if node.get("type") == "variableComponent":
            attrs = node.get("attrs", {})
            content = attrs.get("content")
            if content and content not in index:
                index[content] = {
                    "label": attrs.get("label", ""),
                    "nodeWorkflowId": attrs.get("nodeWorkflowId", ""),
                    "nodeWorkflowType": attrs.get("nodeWorkflowType", ""),
                    "variableId": attrs.get("variableId", ""),
                    "variableType": attrs.get("variableType", "text"),
                }
        for v in node.values():
            _walk_variable_components(v, index)
    elif isinstance(node, list):
        for item in node:
            _walk_variable_components(item, index)


def _build_reference_index(componentes: list) -> dict:
    index: dict = {}
    _walk_variable_components(componentes, index)
    return index


def _make_variable_doc(expr: str, meta: dict) -> dict:
    attrs = {
        "content": expr,
        "label": meta.get("label") or _PLACEHOLDER_LABEL,
        "id": "",
        "nodeWorkflowId": meta.get("nodeWorkflowId", ""),
        "nodeWorkflowType": meta.get("nodeWorkflowType", ""),
        "variableId": meta.get("variableId", ""),
        "variableType": meta.get("variableType", "text"),
    }
    return {
        "content": [
            {
                "content": [{"attrs": attrs, "type": "variableComponent"}],
                "type": "paragraph",
            }
        ],
        "mathEnabled": False,
        "type": "doc",
    }


def _convert_variavel(variavel, index: dict):
    """Converte o campo `variavel` de uma variavel de escoragem antiga.
    Referencias simples (uma unica {{...}}) sao promovidas ao formato tiptap
    rico. Formulas compostas (varios {{...}} + operadores) sao mantidas no
    formato de lista original -- o motor ja aceita esse formato dentro de
    escoragemNodeV2, como visto em nos ja migrados manualmente."""
    if not isinstance(variavel, list):
        return variavel, False

    if (
        len(variavel) == 1
        and isinstance(variavel[0], str)
        and _VAR_RE.fullmatch(variavel[0].strip())
    ):
        expr = variavel[0].strip()
        meta = index.get(expr, {})
        faltando = expr not in index
        return _make_variable_doc(expr, meta), faltando

    faltando = any(
        isinstance(item, str) and _VAR_RE.fullmatch(item.strip()) and item.strip() not in index
        for item in variavel
    )
    return variavel, faltando


def _convert_escoragem_node(comp: dict, index: dict):
    d = comp.get("data", {})
    esc = d.get("escoragem", {})
    detalhes = esc.get("detalhes", d.get("detalhes", ""))

    novos_grupos = []
    faltando_total = 0
    for grupo in esc.get("grupos", []):
        novo_grupo = dict(grupo)
        novas_variaveis = []
        for var in grupo.get("variaveis", []):
            novo_var = dict(var)
            novo_valor, faltando = _convert_variavel(var.get("variavel"), index)
            novo_var["variavel"] = novo_valor
            if faltando:
                faltando_total += 1
            novas_variaveis.append(novo_var)
        novo_grupo["variaveis"] = novas_variaveis
        novos_grupos.append(novo_grupo)

    novo_comp = dict(comp)
    novo_comp["type"] = "escoragemNodeV2"
    novo_data = dict(d)
    novo_data["detalhes"] = detalhes
    novo_escoragem = dict(esc)
    novo_escoragem.pop("detalhes", None)
    novo_escoragem["grupos"] = novos_grupos
    novo_data["escoragem"] = novo_escoragem
    novo_comp["data"] = novo_data
    return novo_comp, faltando_total


st.title("Atualizar Escoragens: Antigo → Novo")
st.caption(
    "Converte nos `escoragemNode` (formato antigo) de um motor de crédito "
    "completo para `escoragemNodeV2` (formato novo), preenchendo a metadata "
    "de cada variável a partir de referências idênticas já presentes no "
    "restante do motor."
)

input_mode = st.radio("Origem do JSON", ["Colar JSON", "Upload de arquivo"], horizontal=True)

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
    st.info("Cole ou faça upload do JSON do motor para converter.")
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

index = _build_reference_index(componentes)

alvo = [c for c in componentes if c.get("type") == "escoragemNode"]
if not alvo:
    st.info("Nenhum nó `escoragemNode` (formato antigo) encontrado neste motor.")
    st.stop()

st.metric("Nós de escoragem no formato antigo encontrados", len(alvo))

novos_componentes = []
relatorio_faltantes = []
for comp in componentes:
    if comp.get("type") == "escoragemNode":
        novo_comp, faltando = _convert_escoragem_node(comp, index)
        novos_componentes.append(novo_comp)
        if faltando:
            relatorio_faltantes.append(
                {
                    "Nó": comp.get("data", {}).get("label", comp.get("id")),
                    "Variáveis sem metadata encontrada": faltando,
                }
            )
    else:
        novos_componentes.append(comp)

nova_data = dict(data)
nova_estrutura = dict(data.get("estrutura_motor", {}))
nova_estrutura["componentes"] = novos_componentes
nova_data["estrutura_motor"] = nova_estrutura

st.divider()

if relatorio_faltantes:
    st.warning(
        f"{len(relatorio_faltantes)} nó(s) de escoragem têm variáveis cuja "
        "metadata (label/nó de origem) não foi encontrada em nenhuma outra "
        "parte do motor — provavelmente referenciam um nó ETL que já foi "
        "removido ou substituído. Essas variáveis ficaram marcadas com "
        f'"{_PLACEHOLDER_LABEL}" no JSON de saída e precisam de revisão manual.'
    )
    st.dataframe(pd.DataFrame(relatorio_faltantes), use_container_width=True, hide_index=True)
else:
    st.success("Todas as variáveis puderam ser mapeadas com sucesso.")

st.download_button(
    label="⬇️ Baixar motor convertido (JSON)",
    data=json.dumps(nova_data, ensure_ascii=False, indent=2).encode("utf-8"),
    file_name="motor_convertido.json",
    mime="application/json",
    use_container_width=True,
)

with st.expander("Ver JSON convertido"):
    st.json(nova_data)
