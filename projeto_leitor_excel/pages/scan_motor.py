"""
Scan do Motor
Valida os nos de escoragem (escoragemNode / escoragemNodeV2) de um motor de
credito completo: soma de pesos, pontuacao maxima, variaveis quebradas,
formato antigo e cobertura da classificacao de risco.
"""

import json

import pandas as pd
import streamlit as st

from exportar_escoragem import _fmt_num, _parse_classificacao, _parse_grupos, _peso_to_float

st.set_page_config(page_title="Scan do Motor", layout="wide")


_HTML_ENT = {"&gt;=": ">=", "&lt;=": "<=", "&gt;": ">", "&lt;": "<", "&amp;": "&"}


def _decode(text) -> str:
    if not isinstance(text, str):
        return str(text) if text is not None else ""
    for k, v in _HTML_ENT.items():
        text = text.replace(k, v)
    return text


def _fix_mojibake(obj):
    """Corrige texto UTF-8 que foi salvo/exportado como Latin-1."""
    if isinstance(obj, str):
        try:
            return obj.encode("latin-1").decode("utf-8")
        except (UnicodeDecodeError, UnicodeEncodeError):
            return obj
    if isinstance(obj, dict):
        return {k: _fix_mojibake(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_fix_mojibake(v) for v in obj]
    return obj


# Mesmo texto usado em atualizar_escoragens.py — precisa ficar em sync.
_PLACEHOLDER_LABEL = "[REVISAR] campo nao encontrado no motor"


def _has_placeholder(node) -> bool:
    """Procura recursivamente por variableComponent com metadata pendente
    de revisao (marcada pelo conversor antigo -> V2)."""
    if isinstance(node, dict):
        if node.get("type") == "variableComponent" and node.get("attrs", {}).get("label") == _PLACEHOLDER_LABEL:
            return True
        return any(_has_placeholder(v) for v in node.values())
    if isinstance(node, list):
        return any(_has_placeholder(item) for item in node)
    return False


def _empty(v) -> bool:
    return v is None or v == ""


_NUM_OPS_LOW = (">=", ">")
_NUM_OPS_HIGH = ("<=", "<")


def _to_number(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _classificacao_interval(cls: dict):
    """Deriva um intervalo numerico [lo, hi] a partir das condicoes brutas
    de uma classe de risco. Retorna None se nao for possivel parsear."""
    lo, hi = None, None
    for cond_key, val_key in (("condicao_1", "pontuacao_1"), ("condicao_2", "pontuacao_2")):
        cond = cls.get(cond_key)
        val = _to_number(cls.get(val_key))
        if not cond or val is None:
            continue
        cond = str(cond).strip()
        if cond in _NUM_OPS_LOW:
            lo = val if lo is None else min(lo, val)
        elif cond in _NUM_OPS_HIGH:
            hi = val if hi is None else max(hi, val)
        elif cond == "==":
            lo = hi = val
    if lo is None and hi is None:
        return None
    return (lo if lo is not None else float("-inf"), hi if hi is not None else float("inf"))


def _scan_classificacao(classificacao_raw: list) -> list:
    """Best-effort: aponta gaps/sobreposicoes entre as faixas de pontuacao
    das classes de risco. Faixas que nao dao pra parsear numericamente sao
    sinalizadas para revisao manual, nao contam como erro."""
    issues = []
    parsed = []
    for cls in classificacao_raw:
        nome = cls.get("nome") or cls.get("classe") or "—"
        interval = _classificacao_interval(cls)
        if interval is None:
            issues.append({
                "nivel": "aviso", "categoria": "Classificação de risco",
                "detalhe": f'Classe "{nome}": faixa não pôde ser validada automaticamente, revisar manualmente.',
            })
            continue
        parsed.append((interval[0], interval[1], nome))

    parsed.sort(key=lambda x: x[0])
    for (lo1, hi1, nome1), (lo2, hi2, nome2) in zip(parsed, parsed[1:]):
        if hi1 == float("inf") or lo2 == float("-inf"):
            continue
        if lo2 > hi1 + 1:
            issues.append({
                "nivel": "aviso", "categoria": "Classificação de risco",
                "detalhe": f'Gap entre "{nome1}" (até {_fmt_num(hi1)}) e "{nome2}" (a partir de {_fmt_num(lo2)}) — faixa não coberta.',
            })
        elif lo2 < hi1:
            issues.append({
                "nivel": "aviso", "categoria": "Classificação de risco",
                "detalhe": f'Faixas sobrepostas entre "{nome1}" e "{nome2}".',
            })
    return issues


def _scan_node(comp: dict) -> dict:
    d = comp.get("data", {})
    esc = d.get("escoragem", {})
    nome = esc.get("detalhes") or d.get("detalhes") or d.get("label") or comp.get("id", "—")
    formato = "V1 (antigo)" if comp.get("type") == "escoragemNode" else "V2"

    grupos_raw = esc.get("grupos", [])
    classificacao_raw = esc.get("classificacao", [])
    grupos_parsed = _parse_grupos({"grupos": grupos_raw})

    issues = []

    if formato.startswith("V1"):
        issues.append({
            "nivel": "aviso", "categoria": "Formato",
            "detalhe": "Nó ainda no formato antigo (escoragemNode) — não migrado para V2.",
        })

    soma_pesos = sum(_peso_to_float(v["peso"]) for g in grupos_parsed for v in g["variaveis"])
    total_pontos = soma_pesos * 1000
    if abs(soma_pesos - 1.0) > 0.001:
        diff_pp = (1.0 - soma_pesos) * 100
        issues.append({
            "nivel": "erro", "categoria": "Peso total",
            "detalhe": (
                f"Soma dos pesos = {soma_pesos * 100:.1f}% (pontuação máxima = "
                f"{_fmt_num(round(total_pontos, 1))} pts) — "
                f"{'faltam' if diff_pp > 0 else 'excedem em'} {abs(diff_pp):.1f} pontos percentuais para 100%."
            ),
        })

    for grupo_raw, grupo_parsed in zip(grupos_raw, grupos_parsed):
        nome_grupo = grupo_parsed["nome"]
        variaveis_raw = grupo_raw.get("variaveis", [])
        if not variaveis_raw:
            issues.append({
                "nivel": "aviso", "categoria": "Grupo vazio",
                "detalhe": f'Grupo "{nome_grupo}" não tem nenhuma variável.',
            })
            continue

        for var_raw, var_parsed in zip(variaveis_raw, grupo_parsed["variaveis"]):
            nome_var = var_parsed["nome"]
            pontuacao_raw = var_raw.get("pontuacao", [])

            if not pontuacao_raw:
                issues.append({
                    "nivel": "erro", "categoria": "Variável quebrada",
                    "detalhe": f'Variável "{nome_var}" (grupo "{nome_grupo}") sem nenhuma regra de pontuação.',
                })
                continue

            if _peso_to_float(var_parsed["peso"]) == 0:
                issues.append({
                    "nivel": "aviso", "categoria": "Variável quebrada",
                    "detalhe": f'Variável "{nome_var}" (grupo "{nome_grupo}") com peso 0% — não contribui para o score.',
                })

            max_score = var_parsed["max_score"]
            if abs(max_score - 100) > 0.01:
                issues.append({
                    "nivel": "aviso", "categoria": "Faixa incompleta",
                    "detalhe": (
                        f'Variável "{nome_var}" (grupo "{nome_grupo}"): maior faixa de pontuação é '
                        f"{_fmt_num(max_score)}, não 100 — pode indicar regra faltando (revisar manualmente)."
                    ),
                })

            for p in pontuacao_raw:
                if _empty(p.get("condicao_1")) or _empty(p.get("faixa_pontuacao")):
                    issues.append({
                        "nivel": "erro", "categoria": "Regra incompleta",
                        "detalhe": f'Variável "{nome_var}" (grupo "{nome_grupo}") tem uma regra sem condição ou sem valor de faixa.',
                    })

            if _has_placeholder(var_raw.get("variavel")):
                issues.append({
                    "nivel": "erro", "categoria": "Metadata pendente",
                    "detalhe": f'Variável "{nome_var}" (grupo "{nome_grupo}") com metadata não resolvida (placeholder de migração) — precisa revisão manual.',
                })

    issues.extend(_scan_classificacao(classificacao_raw))

    return {
        "id": comp.get("id"),
        "nome": _decode(nome),
        "formato": formato,
        "soma_pesos": soma_pesos,
        "total_pontos": total_pontos,
        "issues": issues,
    }


def _node_label(scan: dict) -> str:
    n_err = sum(1 for i in scan["issues"] if i["nivel"] == "erro")
    n_warn = sum(1 for i in scan["issues"] if i["nivel"] == "aviso")
    if not n_err and not n_warn:
        badge = "✅ OK"
    else:
        parts = []
        if n_err:
            parts.append(f"{n_err} erro(s)")
        if n_warn:
            parts.append(f"{n_warn} aviso(s)")
        badge = "⚠️ " + ", ".join(parts)
    return f"{scan['nome']} — {badge}"


st.title("Scan do Motor")
st.caption(
    "Valida os nós de escoragem do motor: soma de pesos, pontuação máxima, "
    "variáveis quebradas, formato antigo e cobertura da classificação de risco."
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
    st.info("Cole ou faça upload do JSON do motor para rodar o scan.")
    st.stop()

try:
    data = json.loads(json_text)
except json.JSONDecodeError as e:
    st.error(f"JSON inválido: {e}")
    st.stop()

data = _fix_mojibake(data)

componentes: list = data.get("estrutura_motor", {}).get("componentes", [])
if not componentes:
    st.warning("Nenhum componente encontrado em `estrutura_motor.componentes`.")
    st.stop()

nos_escoragem = [c for c in componentes if c.get("type") in ("escoragemNode", "escoragemNodeV2")]
if not nos_escoragem:
    st.info("Nenhum nó de escoragem encontrado neste motor.")
    st.stop()

scans = [_scan_node(c) for c in nos_escoragem]

n_antigo = sum(1 for s in scans if s["formato"].startswith("V1"))
n_com_pendencia = sum(1 for s in scans if s["issues"])

c1, c2, c3 = st.columns(3)
c1.metric("Nós de escoragem", len(scans))
c2.metric("Formato antigo", n_antigo)
c3.metric("Com pendências", n_com_pendencia)

st.divider()

st.subheader("Resumo")
resumo_rows = []
for s in scans:
    n_err = sum(1 for i in s["issues"] if i["nivel"] == "erro")
    n_warn = sum(1 for i in s["issues"] if i["nivel"] == "aviso")
    status = "✅ OK" if not s["issues"] else f"⚠️ {n_err} erro(s), {n_warn} aviso(s)"
    resumo_rows.append({
        "Nó": s["nome"],
        "Formato": s["formato"],
        "Soma dos pesos": f"{s['soma_pesos'] * 100:.1f}%",
        "Pontuação máxima": f"{_fmt_num(round(s['total_pontos'], 1))} pts",
        "Status": status,
    })
st.dataframe(pd.DataFrame(resumo_rows), use_container_width=True, hide_index=True)

st.divider()

st.subheader("Detalhe por nó")
options = {_node_label(s): s for s in scans}
escolha = st.selectbox("Escolha o nó de escoragem para revisar", list(options.keys()))
scan = options[escolha]

m1, m2 = st.columns(2)
if abs(scan["soma_pesos"] - 1.0) <= 0.001:
    m1.success(f"Soma dos pesos: {scan['soma_pesos'] * 100:.1f}%")
else:
    m1.error(f"Soma dos pesos: {scan['soma_pesos'] * 100:.1f}%")
m2.metric("Pontuação máxima", f"{_fmt_num(round(scan['total_pontos'], 1))} pts")

if not scan["issues"]:
    st.success("Nenhuma pendência encontrada neste nó.")
else:
    categorias: dict = {}
    for issue in scan["issues"]:
        categorias.setdefault(issue["categoria"], []).append(issue)
    for categoria, issues in categorias.items():
        nivel = "error" if any(i["nivel"] == "erro" for i in issues) else "warning"
        getattr(st, nivel)(f"**{categoria}** — {len(issues)} ocorrência(s)")
        st.dataframe(
            pd.DataFrame([{"Nível": i["nivel"].capitalize(), "Detalhe": i["detalhe"]} for i in issues]),
            use_container_width=True,
            hide_index=True,
        )

st.caption(
    "As checagens de \"faixa não atingir 100%\" e de cobertura da classificação "
    "de risco são heurísticas — revise manualmente antes de confirmar."
)
