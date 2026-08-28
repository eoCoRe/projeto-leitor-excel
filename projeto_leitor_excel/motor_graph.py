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


def scan_estrutura_motor(componentes: list) -> list:
    """Valida o grafo do motor inteiro (todos os tipos de no, nao so
    escoragem): nos soltos/desconectados, ramos de condicional/switch sem
    ligacao, status intermediario sem continuacao, referencias de variavel
    para nos que nao existem mais, variaveis customizadas nao declaradas no
    no de origem, e variaveis referenciadas antes de terem sido calculadas
    em qualquer caminho possivel do fluxo. Compartilhado entre Scan do Motor
    (motor isolado) e Scan da Esteira (motores enviados junto com a esteira)."""
    issues = []

    def add(nivel, categoria, comp_id, comp_lbl, detalhe):
        issues.append({
            "nivel": nivel, "categoria": categoria,
            "id": comp_id, "no": comp_lbl, "detalhe": detalhe,
        })

    nodes_by_id, out_edges = build_graph(componentes)
    in_edges_index = build_in_edges_index(out_edges)
    start_ids = [c["id"] for c in componentes if c.get("type") == "startNode"]
    reachable = reachable_from(start_ids, out_edges)

    # Nos desconectados do fluxo (nenhum caminho a partir do Inicio chega neles)
    for c in componentes:
        cid = c.get("id")
        if cid in start_ids or cid in reachable:
            continue
        label = comp_label(c)
        add("erro", "Nó desconectado", cid, label,
            f'Nó "{label}" ({c.get("type")}) não é alcançável a partir do Início — '
            "nenhuma ligação de outro nó chega até ele.")

    # Ramos de condicional (se-sim / se-nao) sem ligacao de saida
    for c in componentes:
        if "condicional" not in c.get("type", "").lower():
            continue
        cid = c.get("id")
        label = comp_label(c)
        handles = {h for h, _ in out_edges.get(cid, [])}
        for esperado in ("se-sim", "se-nao"):
            if esperado not in handles:
                add("erro", "Ramo de condicional sem conexão", cid, label,
                    f'Condicional "{label}": ramo "{esperado}" não tem nenhuma ligação de saída.')

    # switchNode: cases sem ligacao, e ausencia de caminho default
    for c in componentes:
        if c.get("type") != "switchNode":
            continue
        cid = c.get("id")
        label = comp_label(c)
        handles = {h for h, _ in out_edges.get(cid, [])}
        for case in c.get("data", {}).get("cases", []):
            case_id = case.get("id")
            if case_id not in handles:
                add("erro", "Case de switch sem conexão", cid, label,
                    f'Switch "{label}": case "{case.get("name", case_id)}" não tem nenhuma ligação de saída.')
        if "default" not in handles:
            add("aviso", "Switch sem caminho default", cid, label,
                f'Switch "{label}": não há ligação para o caso "default" — '
                "se nenhum case corresponder em tempo de execução, o fluxo fica sem destino.")

    # Status intermediario (final != true) sem nenhuma ligacao de saida
    for c in componentes:
        if c.get("type") not in ("statusNode", "statusNodeV2"):
            continue
        cid = c.get("id")
        d = c.get("data", {})
        if d.get("final") or out_edges.get(cid):
            continue
        add("erro", "Status intermediário sem continuação", cid, comp_label(c),
            f'Status (status="{d.get("status")}") não está marcado como final e não tem '
            "nenhuma ligação de saída — o fluxo terminaria implicitamente, sem desfecho formal.")

    # Referencias de variavel: no de origem inexistente / variavel nao declarada / usada antes de calculada
    ancestors_cache: dict = {}

    def ancestors(node_id):
        if node_id not in ancestors_cache:
            ancestors_cache[node_id] = ancestors_of(node_id, in_edges_index)
        return ancestors_cache[node_id]

    # Indexa nos por (tipo, nome) para detectar nos duplicados (mesmo nome,
    # ids diferentes -- ex.: um fluxo copiado para outro ramo/produto) na
    # hora de diagnosticar uma referencia fora de ordem.
    nodes_by_type_label: dict = defaultdict(list)
    for c in componentes:
        nodes_by_type_label[(c.get("type"), comp_label(c))].append(c.get("id"))

    vistos = set()
    usados: set = set()
    for c in componentes:
        cid = c.get("id")
        label = comp_label(c)
        for attrs in walk_variable_refs(c.get("data", {})):
            nwid = attrs.get("nodeWorkflowId")
            nwtype = attrs.get("nodeWorkflowType")
            varid = attrs.get("variableId")
            if not nwid or nwid in SENTINEL_WORKFLOW_IDS:
                continue

            usados.add((nwid, varid))

            chave = (cid, nwid, varid)
            if chave in vistos:
                continue
            vistos.add(chave)

            rlabel = ref_label(attrs)

            producer = nodes_by_id.get(nwid)
            if producer is None:
                add("erro", "Referência a nó inexistente", cid, label,
                    f'Variável "{rlabel}" está quebrada: o nó de onde ela vinha não existe mais '
                    "no motor (removido ou renomeado).")
                continue

            if nwtype == "variaveisNode" and varid not in declared_vars_of_variaveis_node(producer):
                add("erro", "Variável não declarada", cid, label,
                    f'Variável "{rlabel}" está quebrada: o nó de Variáveis "{comp_label(producer)}" '
                    "não declara mais essa variável (foi renomeada ou removida lá).")

            if nwid != cid and cid in reachable and nwid not in ancestors(cid):
                producer_label = comp_label(producer)
                candidatos = [
                    outro_id for outro_id in nodes_by_type_label.get((producer.get("type"), producer_label), [])
                    if outro_id != nwid and outro_id in ancestors(cid)
                ]
                if candidatos:
                    ids_cand = ", ".join(candidatos)
                    add("erro", "Variável referencia nó duplicado errado", cid, label,
                        f'Variável "{rlabel}" referencia o nó "{producer_label}" (id {nwid}), que não é '
                        "alcançado em nenhum caminho do fluxo antes deste ponto — mas existe outro nó com "
                        f"esse mesmo nome que É alcançado antes (id {ids_cand}). Isso indica um nó duplicado "
                        "no motor (por exemplo, copiado para outro ramo/produto) cuja referência de variável "
                        "não foi reapontada para a cópia certa — corrija a referência para usar o nó de "
                        "origem correto.")
                else:
                    add("erro", "Variável usada antes de calculada", cid, label,
                        f'Variável "{rlabel}" está fora de ordem: depende do nó "{producer_label}", '
                        "que não é alcançado em nenhum caminho do fluxo antes deste ponto — a variável "
                        "nunca teria sido calculada quando este nó executa.")

    # Variavel customizada declarada mas nunca referenciada em nenhum outro
    # lugar do motor -- provavel sobra de uma versao antiga. So conta
    # "variavel_customizada" de fato -- nomes reservados (ex: "classificacao",
    # "data_validade_limite") sao consumidos pela esteira/plataforma, nao por
    # referencias internas, entao nao contam como "nao usada".
    for c in componentes:
        if c.get("type") != "variaveisNode":
            continue
        cid = c.get("id")
        for var in c.get("data", {}).get("variaveis", []):
            if var.get("nome") != "variavel_customizada":
                continue
            nv = var.get("nome_variavel")
            if not nv:
                continue
            if (cid, f"variable_{nv}") not in usados:
                add("aviso", "Variável customizada não usada", cid, comp_label(c),
                    f'Variável "{var.get("label") or nv}" é declarada neste nó de Variáveis, mas '
                    "não é referenciada em nenhum outro lugar do motor — provável sobra de uma versão anterior.")

    # Condicional comparando operador numerico com variavel do tipo "options"
    # (categorica) -- combinacao que nao faz sentido e costuma indicar
    # condicao montada errada.
    for c in componentes:
        if "condicional" not in c.get("type", "").lower():
            continue
        cid = c.get("id")
        operador = decode(c.get("data", {}).get("condicao", "")).strip()
        if operador not in (">", ">=", "<", "<="):
            continue
        tipo_refs = list(walk_variable_refs(c.get("data", {}).get("tipo", {})))
        for attrs in tipo_refs:
            if attrs.get("variableType") == "options":
                add("erro", "Condicional com tipos incompatíveis", cid, comp_label(c),
                    f'Condicional usa o operador numérico "{operador}" sobre a variável '
                    f'"{ref_label(attrs)}", que é do tipo categórico (options) — '
                    "operador numérico não faz sentido para um valor categórico.")

    # Nomes de variavel customizada duplicados entre nos de Variaveis diferentes
    donos: dict = defaultdict(list)
    for c in componentes:
        if c.get("type") != "variaveisNode":
            continue
        for var in c.get("data", {}).get("variaveis", []):
            nv = var.get("nome_variavel")
            if nv:
                donos[nv].append((c.get("id"), comp_label(c)))
    for nv, lista in donos.items():
        if len(lista) > 1:
            nomes = ", ".join(f'"{lbl}"' for _, lbl in lista)
            add("aviso", "Nome de variável duplicado", lista[0][0], nv,
                f'A variável customizada "{nv}" é declarada em mais de um nó de Variáveis: {nomes}. '
                "As referências usam o nó de origem para desambiguar, mas isso é um risco de confusão/copy-paste.")

    return issues
