"""
Scan da Esteira
Valida uma esteira de credito e a integracao com os motores que ela chama
(nos `motorNode`): roda a checagem estrutural completa do Scan do Motor
(nos desconectados, ramos sem ligacao, referencias de variavel quebradas ou
fora de ordem) em cada motor enviado, confere se cada motor realmente
produz o que a esteira espera dele (status final, classificacao/risco/
escore/limite/data de validade, ou variaveis customizadas especificas), se
os campos iniciais da esteira referenciados existem, e faz uma checagem
estrutural leve da propria esteira (nos desconectados, ramos de
condicional/switch sem ligacao) -- tudo num relatorio so, sem precisar
rodar o Scan do Motor separadamente para cada motor.
"""

import json
import re
from collections import defaultdict

import pandas as pd
import streamlit as st

from motor_graph import (
    ancestors_of,
    build_graph,
    build_in_edges_index,
    comp_label,
    fix_mojibake,
    reachable_from,
    ref_label,
    scan_estrutura_motor,
    walk_variable_refs,
)

st.set_page_config(page_title="Scan da Esteira", layout="wide")

# Saidas reservadas que a plataforma sempre devolve automaticamente de um
# motorNode quando ha uma escoragem com tabela de classificacao alcancavel
# (risco/escore/limite_sugerido/data_validade/classificacao) ou um status
# final alcancavel (status.tipo_status) -- confirmado com o usuario e com o
# padrao repetido em varias esteiras reais.
_RESERVED_STATUS_KEYS = {"tipo_status"}
_RESERVED_POLITICA_KEYS = {"classificacao", "risco", "escore", "limite_sugerido", "data_validade"}


def _analisar_motor(componentes: list) -> dict:
    """Resume o que um motor e capaz de fornecer: status final alcancavel,
    tabela de classificacao alcancavel, e nomes de variaveis customizadas
    alcancaveis (nome_variavel de nos variaveisNode reachable)."""
    nodes_by_id, out_edges = build_graph(componentes)
    start_ids = [c["id"] for c in componentes if c.get("type") == "startNode"]
    reachable = reachable_from(start_ids, out_edges)

    tem_status_final = any(
        c.get("id") in reachable
        and c.get("type") in ("statusNode", "statusNodeV2")
        and c.get("data", {}).get("final")
        for c in componentes
    )
    tem_classificacao = any(
        c.get("id") in reachable
        and c.get("type") in ("escoragemNode", "escoragemNodeV2")
        and c.get("data", {}).get("escoragem", {}).get("classificacao")
        for c in componentes
    )
    variaveis_customizadas: set = set()
    for c in componentes:
        if c.get("id") not in reachable or c.get("type") != "variaveisNode":
            continue
        for var in c.get("data", {}).get("variaveis", []):
            nv = var.get("nome_variavel")
            if nv:
                variaveis_customizadas.add(nv)

    return {
        "tem_status_final": tem_status_final,
        "tem_classificacao": tem_classificacao,
        "variaveis_customizadas": variaveis_customizadas,
    }


def _scan_esteira_estrutura(componentes: list) -> list:
    """Checagem estrutural leve da propria esteira: nos desconectados e
    ramos de condicional/switch sem ligacao -- os mesmos padroes ja
    validados no Scan do Motor, que se aplicam igual aqui porque a esteira
    usa a mesma convencao de ligacoes/sourceHandle."""
    issues = []

    def add(nivel, categoria, comp_id, comp_lbl, detalhe):
        issues.append({"nivel": nivel, "categoria": categoria, "id": comp_id, "no": comp_lbl, "detalhe": detalhe})

    nodes_by_id, out_edges = build_graph(componentes)
    start_ids = [c["id"] for c in componentes if c.get("type") == "startNode"]
    reachable = reachable_from(start_ids, out_edges)

    for c in componentes:
        cid = c.get("id")
        if cid in start_ids or cid in reachable:
            continue
        lbl = comp_label(c)
        add("erro", "Nó desconectado", cid, lbl,
            f'Nó "{lbl}" ({c.get("type")}) não é alcançável a partir do Início.')

    for c in componentes:
        if "condicional" not in c.get("type", "").lower():
            continue
        cid = c.get("id")
        lbl = comp_label(c)
        handles = {h for h, _ in out_edges.get(cid, [])}
        for esperado in ("se-sim", "se-nao"):
            if esperado not in handles:
                add("erro", "Ramo de condicional sem conexão", cid, lbl,
                    f'Condicional "{lbl}": ramo "{esperado}" não tem nenhuma ligação de saída.')

    for c in componentes:
        if c.get("type") != "switchNode":
            continue
        cid = c.get("id")
        lbl = comp_label(c)
        handles = {h for h, _ in out_edges.get(cid, [])}
        for case in c.get("data", {}).get("cases", []):
            case_id = case.get("id")
            if case_id not in handles:
                add("erro", "Case de switch sem conexão", cid, lbl,
                    f'Switch "{lbl}": case "{case.get("name", case_id)}" não tem nenhuma ligação de saída.')
        if "default" not in handles:
            add("aviso", "Switch sem caminho default", cid, lbl,
                f'Switch "{lbl}": não há ligação para o caso "default".')

    return issues


st.title("Scan da Esteira")
st.caption(
    "Valida a integração entre uma esteira de crédito e os motores que ela chama: roda a checagem "
    "estrutural completa do Scan do Motor em cada motor enviado, confere se cada motor realmente "
    "produz o que a esteira espera dele (status, classificação, risco, escore, limite sugerido, "
    "data de validade, ou variáveis customizadas específicas) — tudo num relatório só."
)

st.subheader("1. Esteira")
input_mode = st.radio("Origem do JSON da esteira", ["Colar JSON", "Upload de arquivo"], horizontal=True)

esteira_text: str | None = None
if input_mode == "Colar JSON":
    esteira_text = st.text_area(
        "Cole o JSON completo da esteira aqui",
        height=200,
        placeholder='{"estrutura": {"componentes": [...]}, "estrutura_inicio_execucao": {...}}',
    )
else:
    uploaded = st.file_uploader("Selecione o arquivo JSON da esteira", type=["json"], key="esteira_upload")
    if uploaded:
        esteira_text = uploaded.read().decode("utf-8")

if not esteira_text or not esteira_text.strip():
    st.info("Cole ou faça upload do JSON da esteira para continuar.")
    st.stop()

try:
    esteira_data = json.loads(esteira_text)
except json.JSONDecodeError as e:
    st.error(f"JSON da esteira inválido: {e}")
    st.stop()

esteira_data = fix_mojibake(esteira_data)
esteira_componentes: list = esteira_data.get("estrutura", {}).get("componentes", [])
if not esteira_componentes:
    st.warning("Nenhum componente encontrado em `estrutura.componentes` da esteira.")
    st.stop()

motor_nodes = [c for c in esteira_componentes if c.get("type") == "motorNode"]

st.subheader("2. Motores chamados por essa esteira")
if not motor_nodes:
    st.info("Essa esteira não tem nenhum nó `motorNode` — nada para cruzar.")
else:
    nomes_motornode = [{"Nó da esteira": comp_label(c), "motor_nome": c["data"].get("motor_nome"),
                         "motor_id": c["data"].get("motor_id")} for c in motor_nodes]
    st.dataframe(pd.DataFrame(nomes_motornode), use_container_width=True, hide_index=True)

motor_files = st.file_uploader(
    "Envie o(s) JSON(s) dos motores chamados por essa esteira (pode selecionar vários)",
    type=["json"], accept_multiple_files=True, key="motores_upload",
)

motores_uploaded: dict = {}
if motor_files:
    for f in motor_files:
        try:
            motor_data = json.loads(f.read().decode("utf-8"))
        except json.JSONDecodeError as e:
            st.error(f'Arquivo "{f.name}": JSON inválido ({e}).')
            continue
        motor_data = fix_mojibake(motor_data)
        comps = motor_data.get("estrutura_motor", {}).get("componentes", [])
        if not comps:
            st.warning(f'Arquivo "{f.name}" não tem `estrutura_motor.componentes` — não parece um motor válido, ignorado.')
            continue
        motores_uploaded[f.name] = comps

if motor_nodes and not motor_files:
    st.info("Envie os motores acima para rodar a validação cruzada — sem eles, só a checagem estrutural da esteira roda.")


def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


# O motor exportado nao guarda o proprio nome (configuracao.nome fica vazio
# na pratica) -- entao a associacao com o motor_nome do motorNode e por
# nome de arquivo, com override manual pra garantir (o usuario confere/troca
# se o palpite automatico errar).
motor_para_arquivo: dict = {}
if motor_nodes and motores_uploaded:
    st.markdown("Associe cada `motorNode` da esteira ao arquivo de motor correspondente:")
    for mn in motor_nodes:
        motor_nome = (mn.get("data", {}).get("motor_nome") or "").strip()
        opcoes = ["— não enviado —"] + list(motores_uploaded.keys())
        alvo_norm = _norm(motor_nome)
        default_idx = 0
        for i, fname in enumerate(motores_uploaded.keys(), start=1):
            base = re.sub(r"^motor-?", "", fname, flags=re.IGNORECASE)
            base = re.sub(r"\.json$", "", base, flags=re.IGNORECASE)
            base_norm = _norm(base)
            if base_norm and alvo_norm and (base_norm == alvo_norm or alvo_norm in base_norm or base_norm in alvo_norm):
                default_idx = i
                break
        escolha = st.selectbox(
            f'Motor esperado: "{motor_nome}" (nó "{comp_label(mn)}" na esteira)',
            opcoes, index=default_idx, key=f"match_{mn.get('id')}",
        )
        motor_para_arquivo[mn.get("id")] = None if escolha == "— não enviado —" else escolha

st.divider()

if motor_nodes:
    cross_issues = []
    nao_validados = []
    estrutura_motor_issues: dict = {}
    motores_usados_por: dict = defaultdict(list)

    def add_cross(nivel, categoria, comp_id, comp_lbl, detalhe):
        cross_issues.append({"nivel": nivel, "categoria": categoria, "id": comp_id, "no": comp_lbl, "detalhe": detalhe})

    _, esteira_out_edges = build_graph(esteira_componentes)
    esteira_in_edges_index = build_in_edges_index(esteira_out_edges)
    esteira_start_ids = [c["id"] for c in esteira_componentes if c.get("type") == "startNode"]
    esteira_reachable = reachable_from(esteira_start_ids, esteira_out_edges)
    ancestors_cache: dict = {}

    def ancestors(node_id):
        if node_id not in ancestors_cache:
            ancestors_cache[node_id] = ancestors_of(node_id, esteira_in_edges_index)
        return ancestors_cache[node_id]

    for mn in motor_nodes:
        mn_id = mn.get("id")
        mn_label = comp_label(mn)
        motor_nome_esteira = (mn.get("data", {}).get("motor_nome") or "").strip()
        arquivo = motor_para_arquivo.get(mn_id)

        if not arquivo:
            nao_validados.append({"Nó motorNode": mn_label, "motor_nome esperado": motor_nome_esteira})
            continue

        nome_real = arquivo
        motor_comps = motores_uploaded[arquivo]
        motores_usados_por[arquivo].append(mn_label)
        if arquivo not in estrutura_motor_issues:
            estrutura_motor_issues[arquivo] = scan_estrutura_motor(motor_comps)
        analise = _analisar_motor(motor_comps)

        for c in esteira_componentes:
            cid = c.get("id")
            lbl = comp_label(c)
            for attrs in walk_variable_refs(c.get("data", {})):
                if attrs.get("nodeWorkflowId") != mn_id:
                    continue
                content = attrs.get("content") or ""
                partes = content.strip("{}").split(".")
                if len(partes) < 2:
                    continue
                namespace, chave = partes[0], partes[-1]
                rlabel = ref_label(attrs)

                if mn_id != cid and cid in esteira_reachable and mn_id not in ancestors(cid):
                    add_cross("erro", "Motor referenciado antes de executar", cid, lbl,
                        f'"{lbl}" lê "{rlabel}" do motor "{nome_real}", mas esse motorNode não está em '
                        "nenhum caminho do fluxo da esteira que leve até aqui — o motor nunca teria rodado "
                        "quando este nó executa.")

                if namespace == "status" and chave in _RESERVED_STATUS_KEYS:
                    if not analise["tem_status_final"]:
                        add_cross("erro", "Saída não disponível no motor", cid, lbl,
                            f'"{lbl}" espera "{rlabel}" (status final) do motor "{nome_real}", mas nenhum status '
                            "final está em nenhum caminho alcançável desse motor.")
                elif namespace == "politica" and chave in _RESERVED_POLITICA_KEYS:
                    if not analise["tem_classificacao"]:
                        add_cross("erro", "Saída não disponível no motor", cid, lbl,
                            f'"{lbl}" espera "{rlabel}" do motor "{nome_real}", mas nenhuma escoragem com tabela '
                            "de classificação de risco está em nenhum caminho alcançável desse motor.")
                elif namespace == "politica":
                    if chave not in analise["variaveis_customizadas"]:
                        add_cross("erro", "Variável customizada não encontrada no motor", cid, lbl,
                            f'"{lbl}" espera a variável customizada "{rlabel}" do motor "{nome_real}", mas nenhum '
                            "nó de Variáveis alcançável nesse motor declara essa variável (renomeada, removida, "
                            "ou o motor enviado não é a versão certa).")

    st.subheader("3. Estrutura interna dos motores enviados")
    st.caption(
        "Mesma checagem estrutural completa do Scan do Motor (nós desconectados, ramos de "
        "condicional/switch sem ligação, referências de variável quebradas ou usadas fora de "
        "ordem) — rodada automaticamente em cada motor enviado acima, sem precisar abrir o Scan "
        "do Motor à parte."
    )
    if not estrutura_motor_issues:
        st.info("Nenhum motor associado ainda — envie e associe os motores na seção 2 acima.")
    else:
        n_err_estrutura = sum(1 for issues in estrutura_motor_issues.values() for i in issues if i["nivel"] == "erro")
        n_warn_estrutura = sum(1 for issues in estrutura_motor_issues.values() for i in issues if i["nivel"] == "aviso")
        me1, me2 = st.columns(2)
        me1.metric("Erros estruturais (todos os motores)", n_err_estrutura)
        me2.metric("Avisos estruturais (todos os motores)", n_warn_estrutura)
        for arquivo, issues in estrutura_motor_issues.items():
            usado_por = ", ".join(f'"{lbl}"' for lbl in motores_usados_por[arquivo])
            n_err = sum(1 for i in issues if i["nivel"] == "erro")
            n_warn = sum(1 for i in issues if i["nivel"] == "aviso")
            titulo = f'Motor "{arquivo}" (nó(s) da esteira: {usado_por})'
            if not issues:
                st.success(f"{titulo} — nenhuma pendência estrutural.")
                continue
            with st.expander(f"{titulo} — ⚠️ {n_err} erro(s), {n_warn} aviso(s)", expanded=n_err > 0):
                st.dataframe(
                    pd.DataFrame([
                        {"Nível": i["nivel"].capitalize(), "Nó": i["no"], "Detalhe": i["detalhe"]}
                        for i in issues
                    ]),
                    use_container_width=True, hide_index=True,
                )

    st.divider()

    st.subheader("4. Validação cruzada: esteira × motor")

    if nao_validados:
        st.warning(f"{len(nao_validados)} motorNode(s) não puderam ser validados — motor correspondente não foi enviado.")
        st.dataframe(pd.DataFrame(nao_validados), use_container_width=True, hide_index=True)

    if any(motor_para_arquivo.values()):
        n_err = sum(1 for i in cross_issues if i["nivel"] == "erro")
        if not cross_issues:
            st.success("Nenhuma inconsistência encontrada entre a esteira e os motores enviados.")
        else:
            st.error(f"{n_err} inconsistência(s) encontrada(s) entre a esteira e os motores enviados.")
            st.dataframe(
                pd.DataFrame([{"Nível": i["nivel"].capitalize(), "Nó da esteira": i["no"], "Detalhe": i["detalhe"]}
                              for i in cross_issues]),
                use_container_width=True, hide_index=True,
            )

    st.caption(
        "As saídas \"status.tipo_status\", \"politica.classificacao/risco/escore/limite_sugerido/data_validade\" "
        "são consideradas disponíveis automaticamente sempre que o motor tem, respectivamente, um status final ou "
        "uma escoragem com tabela de classificação de risco alcançável — não é checado se o valor específico bate, "
        "só se a estrutura que o produz existe no caminho do fluxo."
    )

    st.divider()

# Campos iniciais da esteira referenciados vs declarados
st.subheader("5. Campos iniciais da esteira")
campos_declarados = {
    c.get("nome") for c in esteira_data.get("estrutura_inicio_execucao", {}).get("campos", []) if c.get("nome")
}
campos_referenciados: set = set()
campos_issues = []
for c in esteira_componentes:
    for attrs in walk_variable_refs(c.get("data", {})):
        content = attrs.get("content") or ""
        if content.startswith("{{campos."):
            nome_campo = content.strip("{}").split(".", 1)[-1]
            campos_referenciados.add(nome_campo)
            if nome_campo not in campos_declarados:
                campos_issues.append({
                    "Nó": comp_label(c),
                    "Detalhe": f'Referencia o campo inicial "{nome_campo}", que não está declarado em '
                               "`estrutura_inicio_execucao.campos` — provavelmente renomeado ou removido.",
                })

if campos_issues:
    st.error(f"{len(campos_issues)} referência(s) a campo inicial inexistente.")
    st.dataframe(pd.DataFrame(campos_issues), use_container_width=True, hide_index=True)
else:
    st.success("Todos os campos iniciais referenciados estão declarados em `estrutura_inicio_execucao`.")

st.divider()

st.subheader("6. Validação estrutural da esteira")
estrutura_issues = _scan_esteira_estrutura(esteira_componentes)
if not estrutura_issues:
    st.success("Nenhuma pendência estrutural encontrada na esteira (nós desconectados, ramos sem ligação).")
else:
    por_categoria: dict = {}
    for issue in estrutura_issues:
        por_categoria.setdefault(issue["categoria"], []).append(issue)
    for categoria, issues_cat in por_categoria.items():
        nivel = "error" if any(i["nivel"] == "erro" for i in issues_cat) else "warning"
        getattr(st, nivel)(f"**{categoria}** — {len(issues_cat)} ocorrência(s)")
        st.dataframe(
            pd.DataFrame([{"Nível": i["nivel"].capitalize(), "Nó": i["no"], "Detalhe": i["detalhe"]} for i in issues_cat]),
            use_container_width=True, hide_index=True,
        )
