#!/usr/bin/env python3
"""
motor_graph.py
Helpers de grafo e resolucao de referencias de variavel, compartilhados
entre as paginas que analisam `estrutura_motor.componentes` (Scan do Motor)
e `estrutura.componentes` de uma esteira (Scan da Esteira).
"""

from collections import defaultdict, deque

_HTML_ENT = {"&gt;=": ">=", "&lt;=": "<=", "&gt;": ">", "&lt;": "<", "&amp;": "&"}

# nodeWorkflowId usados pela plataforma para valores que nao vem de um no
# real do fluxo (opcao escolhida num dropdown, ou saida "default"/agregada
# do motor como status.tipo_status e politica.*).
SENTINEL_WORKFLOW_IDS = {"options", "default"}


def decode(text) -> str:
    if not isinstance(text, str):
        return str(text) if text is not None else ""
    for k, v in _HTML_ENT.items():
        text = text.replace(k, v)
    return text


def fix_mojibake(obj):
    """Corrige texto UTF-8 que foi salvo/exportado como Latin-1."""
    if isinstance(obj, str):
        try:
            return obj.encode("latin-1").decode("utf-8")
        except (UnicodeDecodeError, UnicodeEncodeError):
            return obj
    if isinstance(obj, dict):
        return {k: fix_mojibake(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [fix_mojibake(v) for v in obj]
    return obj


def build_graph(componentes: list) -> tuple:
    """Indexa os nos por id e as ligacoes de saida (fonte de verdade do fluxo
    -- `ligacoes.entrada` no JSON do motor/esteira e auto-referente/
    inconsistente e nao deve ser usado)."""
    nodes_by_id = {c["id"]: c for c in componentes if c.get("id")}
    out_edges: dict = defaultdict(list)
    for c in componentes:
        cid = c.get("id")
        for e in c.get("ligacoes", {}).get("saida", []):
            tgt = e.get("target")
            if tgt:
                out_edges[cid].append((e.get("sourceHandle"), tgt))
    return nodes_by_id, out_edges


def reachable_from(start_ids: list, out_edges: dict) -> set:
    seen = set(start_ids)
    fila = deque(start_ids)
    while fila:
        cur = fila.popleft()
        for _, tgt in out_edges.get(cur, []):
            if tgt not in seen:
                seen.add(tgt)
                fila.append(tgt)
    return seen


def build_in_edges_index(out_edges: dict) -> dict:
    idx: dict = defaultdict(set)
    for src, edges in out_edges.items():
        for _, tgt in edges:
            idx[tgt].add(src)
    return idx


def ancestors_of(node_id: str, in_edges_index: dict) -> set:
    """Todos os nos que possuem algum caminho no fluxo ate `node_id`."""
    seen: set = set()
    fila = deque([node_id])
    while fila:
        cur = fila.popleft()
        for src in in_edges_index.get(cur, []):
            if src not in seen:
                seen.add(src)
                fila.append(src)
    return seen


def walk_variable_refs(node):
    """Percorre recursivamente um bloco `data` de componente e retorna os
    `attrs` de cada `variableComponent` encontrado (referencias a variaveis
    de outros nos, em condicionais, formulas, mensagens etc.)."""
    if isinstance(node, dict):
        if node.get("type") == "variableComponent":
            attrs = node.get("attrs", {})
            if attrs:
                yield attrs
        for v in node.values():
            yield from walk_variable_refs(v)
    elif isinstance(node, list):
        for item in node:
            yield from walk_variable_refs(item)


def declared_vars_of_variaveis_node(comp: dict) -> set:
    return {
        f"variable_{var['nome_variavel']}"
        for var in comp.get("data", {}).get("variaveis", [])
        if var.get("nome_variavel")
    }


def comp_label(comp: dict) -> str:
    d = comp.get("data", {})
    return d.get("detalhes") or d.get("label") or comp.get("id", "—")


def ref_label(attrs: dict) -> str:
    """Nome legivel de uma referencia de variavel (para exibicao), preferindo
    o `label` humano ao token bruto `{{...}}`."""
    label = decode(attrs.get("label") or "").strip()
    if label:
        return label
    content = attrs.get("content")
    if content:
        return decode(content)
    return attrs.get("variableId") or "variável"
