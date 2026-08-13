"""
Relatório do Motor de Crédito
Analisa o JSON completo de um motor (estrutura_motor.componentes) e exibe um
resumo estruturado: fontes de dados, decisões, condicionais, escoragem e variáveis.
"""

import html as _html
import io
import json
import re
from datetime import datetime

import pandas as pd
import streamlit as st
from reportlab.graphics.charts.barcharts import HorizontalBarChart
from reportlab.graphics.shapes import Drawing
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    HRFlowable,
    KeepTogether,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

st.set_page_config(page_title="Relatório do Motor", layout="wide")


_HTML_ENT = {"&gt;=": ">=", "&lt;=": "<=", "&gt;": ">", "&lt;": "<", "&amp;": "&"}


def _decode(text) -> str:
    if not isinstance(text, str):
        return str(text) if text is not None else ""
    for k, v in _HTML_ENT.items():
        text = text.replace(k, v)
    return text


def _safe_filename(text: str) -> str:
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", text).strip() or "escoragem"


def _fix_mojibake(obj):
    """Corrige texto UTF-8 que foi salvo/exportado como Latin-1 (ex: 'SituaÃ§Ã£o' -> 'Situação')."""
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
                attrs = inline.get("attrs", {})
                if attrs.get("nodeWorkflowType") == "options":
                    val = attrs.get("content", "")
                    if val:
                        parts.append(val)
                else:
                    label = attrs.get("label", "")
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


def _format_documento(doc_raw) -> str:
    """Formata o campo 'documento' de um nÃ³ ETL, reconhecendo tanto a lista simples
    (etlNode) quanto o rich-text com referÃªncia de variÃ¡vel (etlNodeV2)."""
    if isinstance(doc_raw, list):
        text = ", ".join(str(x) for x in doc_raw)
    elif isinstance(doc_raw, dict):
        is_input_doc = any(
            inline.get("type") == "variableComponent"
            and "input_inicial.documento" in inline.get("attrs", {}).get("content", "")
            for block in doc_raw.get("content", [])
            for inline in block.get("content", [])
        )
        text = "{{input_inicial.documento}}" if is_input_doc else _tiptap_text(doc_raw)
    else:
        text = str(doc_raw) if doc_raw else ""
    text = text.replace("{{input_inicial.documento}}", "Documento de entrada")
    return text or "—"


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


_PDF_NAVY  = colors.HexColor("#1F4E79")
_PDF_BLUE  = colors.HexColor("#2E75B6")
_PDF_ALT   = colors.HexColor("#EBF3FA")
_PDF_GRID  = colors.HexColor("#AAAAAA")
_PDF_LIGHT = colors.HexColor("#F0F4FA")

_STATUS_BADGE = {
    "reprovar": colors.HexColor("#C0392B"),
    "derivar":  colors.HexColor("#D4A017"),
    "aprovar":  colors.HexColor("#1E8449"),
}

_RISK_BG = {
    "muito baixo": colors.HexColor("#00B050"),
    "baixo":       colors.HexColor("#92D050"),
    "medio":       colors.HexColor("#FFFF00"),
    "medio -":     colors.HexColor("#FFFF00"),
    "medio +":     colors.HexColor("#FFC000"),
    "alto":        colors.HexColor("#FFC000"),
    "muito alto":  colors.HexColor("#FF4444"),
}
_RISK_DARK_BG = {"muito baixo", "muito alto"}

_pdf_styles = getSampleStyleSheet()
_pdf_styles.add(ParagraphStyle("Cell", parent=_pdf_styles["Normal"], fontSize=8, leading=10))
_pdf_styles.add(ParagraphStyle("CellBold", parent=_pdf_styles["Cell"], fontName="Helvetica-Bold"))
_pdf_styles.add(ParagraphStyle("CellWhite", parent=_pdf_styles["Cell"], textColor=colors.white,
                                fontName="Helvetica-Bold"))
_pdf_styles.add(ParagraphStyle("Subtitle", parent=_pdf_styles["Normal"], textColor=colors.white,
                                fontSize=10, fontName="Helvetica-Oblique"))
_pdf_styles.add(ParagraphStyle("MetricLabel", parent=_pdf_styles["Normal"], textColor=colors.white,
                                fontSize=8, alignment=1))
_pdf_styles.add(ParagraphStyle("MetricValue", parent=_pdf_styles["Normal"], textColor=_PDF_NAVY,
                                fontSize=20, fontName="Helvetica-Bold", alignment=1))
_pdf_styles.add(ParagraphStyle("H2x", parent=_pdf_styles["Heading2"], textColor=_PDF_NAVY,
                                spaceBefore=18, spaceAfter=2))
_pdf_styles.add(ParagraphStyle("H3x", parent=_pdf_styles["Heading3"], textColor=_PDF_BLUE,
                                spaceBefore=10, spaceAfter=4))

_OP_PHRASES = {
    "==": "for igual a",
    "!=": "for diferente de",
    ">":  "for maior que",
    ">=": "for maior ou igual a",
    "<":  "for menor que",
    "<=": "for menor ou igual a",
}


_OP_VERB_PHRASES = {
    "CONTEM": "contém",
    "CONTIDO_EM": "está contido em",
}


def _explain_condition(tipo: str, operador: str, valor: str) -> str:
    tipo_e = _html.escape(tipo or "—")
    op = (operador or "").strip()
    if op.lower() == "vazio":
        return f"Verifica se <b>{tipo_e}</b> está vazio."
    valor_e = _html.escape(str(valor)) if valor not in (None, "") else ""
    verb_phrase = _OP_VERB_PHRASES.get(op.upper())
    if verb_phrase:
        return f"Verifica se <b>{tipo_e}</b> {verb_phrase} <b>{valor_e}</b>."
    phrase = _OP_PHRASES.get(op, f"for \"{_html.escape(op)}\"")
    return f"Verifica se <b>{tipo_e}</b> {phrase} <b>{valor_e}</b>."


def _build_graph_index(componentes: list) -> tuple:
    nodes_by_id = {c.get("id"): c for c in componentes if c.get("id")}
    edges_by_source = {
        c.get("id"): c.get("ligacoes", {}).get("saida", [])
        for c in componentes if c.get("id")
    }
    return nodes_by_id, edges_by_source


def _trace_outcome(start_id: str, handle: str, nodes_by_id: dict, edges_by_source: dict, max_hops: int = 8) -> str:
    """Segue o fluxo (sem narrar passos intermediÃ¡rios) atÃ© achar um status final.
    Se a cadeia bifurcar de novo (outro condicional) ou nÃ£o resolver, retorna um aviso genÃ©rico."""
    edges = [e for e in edges_by_source.get(start_id, []) if e.get("sourceHandle") == handle]
    if not edges:
        return ""
    current_id = edges[0].get("target")
    for _ in range(max_hops):
        comp = nodes_by_id.get(current_id)
        if not comp:
            return "Segue o fluxo"
        t = comp.get("type", "")
        d = comp.get("data", {})
        if "status" in t.lower():
            status = _STATUS_LABEL.get(d.get("status", ""), d.get("status", ""))
            msg_raw = d.get("message", "")
            msg = _tiptap_text(msg_raw) if isinstance(msg_raw, dict) else str(msg_raw)
            return f"<b>{_html.escape(status)}</b>: {_html.escape(msg)}"
        if "condicional" in t.lower():
            return "Segue o fluxo"
        next_edges = edges_by_source.get(current_id, [])
        if len(next_edges) != 1:
            return "Segue o fluxo"
        current_id = next_edges[0].get("target")
    return "Segue o fluxo"


def _section(title: str) -> list:
    return [
        Paragraph(title, _pdf_styles["H2x"]),
        HRFlowable(width="100%", thickness=1.2, color=_PDF_BLUE, spaceBefore=2, spaceAfter=8),
    ]


def _pdf_cell(text, style="Cell") -> Paragraph:
    if isinstance(text, Paragraph):
        return text
    return Paragraph(_html.escape(str(text)) if text not in (None, "") else "—", _pdf_styles[style])


def _pdf_table(rows: list, col_widths=None, extra_style=None) -> Table:
    wrapped = [[_pdf_cell(c) for c in row] for row in rows]
    t = Table(wrapped, colWidths=col_widths, repeatRows=1)
    cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), _PDF_NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, _PDF_GRID),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, _PDF_ALT]),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    if extra_style:
        cmds.extend(extra_style)
    t.setStyle(TableStyle(cmds))
    return t


def _build_cover(model_name: str, generated_at: str) -> list:
    title_html = f"<font color='white' size='20'><b>{_html.escape(model_name)}</b></font>"
    banner = Table(
        [[Paragraph(title_html, _pdf_styles["Normal"])],
         [Paragraph(f"Resumo Técnico do Motor de Decisão &nbsp;•&nbsp; Gerado em {generated_at}",
                    _pdf_styles["Subtitle"])]],
        colWidths=[18.7 * cm],
    )
    banner.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), _PDF_NAVY),
        ("BACKGROUND", (0, 1), (0, 1), _PDF_BLUE),
        ("TOPPADDING", (0, 0), (0, 0), 16),
        ("BOTTOMPADDING", (0, 0), (0, 0), 4),
        ("TOPPADDING", (0, 1), (0, 1), 6),
        ("BOTTOMPADDING", (0, 1), (0, 1), 10),
        ("LEFTPADDING", (0, 0), (-1, -1), 14),
    ]))
    return [banner, Spacer(1, 16)]


def _build_metric_cards(labels_values: list) -> Table:
    label_row = [Paragraph(lbl, _pdf_styles["MetricLabel"]) for lbl, _ in labels_values]
    value_row = [Paragraph(str(val), _pdf_styles["MetricValue"]) for _, val in labels_values]
    n = len(labels_values)
    width = 18.7 * cm / n
    t = Table([label_row, value_row], colWidths=[width] * n)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), _PDF_NAVY),
        ("BACKGROUND", (0, 1), (-1, 1), _PDF_LIGHT),
        ("TOPPADDING", (0, 0), (-1, 0), 6),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
        ("TOPPADDING", (0, 1), (-1, 1), 10),
        ("BOTTOMPADDING", (0, 1), (-1, 1), 12),
        ("LINEBELOW", (0, 0), (-1, 0), 0.5, _PDF_GRID),
        ("BOX", (0, 0), (-1, -1), 0.5, _PDF_GRID),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, _PDF_GRID),
    ]))
    return t


def _type_distribution_chart(type_counts: dict) -> Drawing:
    items = sorted(type_counts.items(), key=lambda x: x[1])[-10:]
    cats = [t for t, _ in items]
    vals = [cnt for _, cnt in items]
    height = max(2.5 * cm, len(items) * 0.9 * cm)

    d = Drawing(15 * cm, height)
    chart = HorizontalBarChart()
    chart.x = 4.2 * cm
    chart.y = 0.3 * cm
    chart.width = 10 * cm
    chart.height = height - 0.6 * cm
    chart.data = [vals]
    chart.categoryAxis.categoryNames = cats
    chart.categoryAxis.labels.fontSize = 7
    chart.valueAxis.valueMin = 0
    chart.bars[0].fillColor = _PDF_BLUE
    chart.barWidth = 8
    chart.barLabels.fontSize = 7
    chart.barLabelFormat = "%d"
    chart.barLabels.nudge = 8
    d.add(chart)
    return d


def _status_row_style(row_idx: int, status: str) -> list:
    bg = _STATUS_BADGE.get(status, _PDF_NAVY)
    return [
        ("BACKGROUND", (0, row_idx), (0, row_idx), bg),
        ("TEXTCOLOR", (0, row_idx), (0, row_idx), colors.white),
        ("FONTNAME", (0, row_idx), (0, row_idx), "Helvetica-Bold"),
        ("ALIGN", (0, row_idx), (0, row_idx), "CENTER"),
    ]


def _risk_row_style(row_idx: int, risco: str) -> list:
    key = risco.lower().strip().replace("é", "e").replace("í", "i")
    bg = _RISK_BG.get(key, colors.white)
    fg = colors.white if key in _RISK_DARK_BG else colors.black
    return [
        ("BACKGROUND", (0, row_idx), (1, row_idx), bg),
        ("TEXTCOLOR", (0, row_idx), (1, row_idx), fg),
        ("FONTNAME", (0, row_idx), (1, row_idx), "Helvetica-Bold"),
    ]


def _pdf_footer(canvas, doc) -> None:
    canvas.saveState()
    canvas.setStrokeColor(_PDF_GRID)
    canvas.line(1.5 * cm, 1.1 * cm, A4[0] - 1.5 * cm, 1.1 * cm)
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(colors.HexColor("#888888"))
    canvas.drawString(1.5 * cm, 0.7 * cm, "Gerado automaticamente pelo Visualizador e Exportador")
    canvas.drawRightString(A4[0] - 1.5 * cm, 0.7 * cm, f"Página {doc.page}")
    canvas.restoreState()


def build_motor_pdf(data: dict, componentes: list, buckets: dict) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        topMargin=1.5 * cm, bottomMargin=1.5 * cm,
        leftMargin=1.5 * cm, rightMargin=1.5 * cm,
    )
    s = _pdf_styles
    story = []

    model_name = data.get("configuracao", {}).get("nome") or "Relatório do Motor de Crédito"
    story.extend(_build_cover(model_name, datetime.now().strftime("%d/%m/%Y %H:%M")))

    story.append(_build_metric_cards([
        ("NÓS TOTAIS", len(componentes)),
        ("FONTES DE DADOS", len(buckets["etl"])),
        ("DECISÕES", len(buckets["status"])),
        ("CONDICIONAIS", len(buckets["conditional"])),
        ("ESCORAGEM", len(buckets["scoring"])),
    ]))

    story.extend(_section("1. Fontes de Dados Consultadas"))
    if buckets["etl"]:
        rows = [["Nome", "Componente ID", "Versão", "Documento consultado"]]
        for comp in buckets["etl"]:
            d = comp.get("data", {})
            doc_str = _format_documento(d.get("documento", ""))
            version = "v2" if comp.get("type") == "etlNodeV2" else "v1"
            rows.append([d.get("label", "—"), d.get("componentId", "—"), version, doc_str])
        story.append(_pdf_table(rows, col_widths=[4.5 * cm, 3.5 * cm, 1.5 * cm, 8 * cm]))
    else:
        story.append(Paragraph("Nenhum nó ETL encontrado.", s["Normal"]))

    story.extend(_section("2. Regras Condicionais"))
    nodes_by_id, edges_by_source = _build_graph_index(componentes)
    if buckets["conditional"]:
        rows = [["Pergunta / Detalhe", "Explicação da regra e consequência"]]
        for comp in buckets["conditional"]:
            d = comp.get("data", {})
            cid = comp.get("id")
            tipo_clean = _tiptap_text(d.get("tipo", "")).replace("[", "").replace("]", "")
            valor_clean = _tiptap_text(d.get("valor", "")).replace("[", "").replace("]", "")
            operador = _decode(d.get("condicao", ""))

            lines = [_explain_condition(tipo_clean, operador, valor_clean)]
            sim_outcome = _trace_outcome(cid, "se-sim", nodes_by_id, edges_by_source)
            nao_outcome = _trace_outcome(cid, "se-nao", nodes_by_id, edges_by_source)
            if sim_outcome:
                lines.append(f"&bull; Se <b>SIM</b>: {sim_outcome}")
            if nao_outcome:
                lines.append(f"&bull; Se <b>NÃO</b>: {nao_outcome}")

            explicacao = Paragraph("<br/>".join(lines), s["Cell"])
            rows.append([d.get("detalhes", "—"), explicacao])
        story.append(_pdf_table(rows, col_widths=[5 * cm, 12.5 * cm]))
    else:
        story.append(Paragraph("Nenhum nó condicional encontrado.", s["Normal"]))

    story.extend(_section("3. Decisões do Motor"))
    if buckets["status"]:
        rows = [["Status", "Tipo", "Mensagem"]]
        badge_style = []
        for i, comp in enumerate(buckets["status"], start=1):
            d = comp.get("data", {})
            raw_status = d.get("status", "")
            status = _STATUS_LABEL.get(raw_status, raw_status)
            tipo = "Final" if d.get("final") else "Intermediário"
            msg_raw = d.get("message", "")
            msg = _tiptap_text(msg_raw) if isinstance(msg_raw, dict) else str(msg_raw)
            rows.append([status, tipo, msg])
            badge_style.extend(_status_row_style(i, raw_status))
        story.append(_pdf_table(rows, col_widths=[2.5 * cm, 3 * cm, 12 * cm], extra_style=badge_style))
    else:
        story.append(Paragraph("Nenhum nó de decisão encontrado.", s["Normal"]))

    story.extend(_section("4. Escoragem"))
    if buckets["scoring"]:
        for comp in buckets["scoring"]:
            d = comp.get("data", {})
            label = d.get("label", "Escoragem")
            detalhes = d.get("detalhes", "")
            story.append(Paragraph(label + (f" — {detalhes}" if detalhes else ""), s["H3x"]))

            esc = d.get("escoragem", {})
            for grupo in esc.get("grupos", []):
                story.append(Paragraph(f"Grupo: {grupo.get('nome', '')}", s["Normal"]))
                rows = [["Variável", "Peso", "Condição", "Pontuação base", "Score faixa"]]
                for var in grupo.get("variaveis", []):
                    peso_raw = var.get("peso", "")
                    peso_fmt = (
                        f"{float(peso_raw) * 100:.0f}%"
                        if isinstance(peso_raw, float) and 0 < peso_raw <= 1
                        else str(peso_raw)
                    )
                    for p in var.get("pontuacao", []):
                        rows.append([
                            var.get("nome", ""), peso_fmt,
                            _decode(str(p.get("condicao_1", ""))),
                            p.get("pontuacao_1", ""), p.get("faixa_pontuacao", ""),
                        ])
                if len(rows) > 1:
                    story.append(_pdf_table(rows, col_widths=[5 * cm, 2 * cm, 5 * cm, 3 * cm, 3 * cm]))
                    story.append(Spacer(1, 6))

            classificacao = esc.get("classificacao", [])
            if classificacao:
                story.append(Paragraph("Classificação de risco:", s["Normal"]))
                rows = [["Classe / Nome", "Risco", "Condição", "Pontuação", "Fator", "Validade (dias)"]]
                risk_style = []
                for i, cls in enumerate(classificacao, start=1):
                    rows.append([
                        cls.get("nome", cls.get("classe", "")), cls.get("risco", ""),
                        _decode(str(cls.get("condicao_1", ""))), cls.get("pontuacao_1", ""),
                        cls.get("faixa_pontuacao", ""), cls.get("validade_limite", ""),
                    ])
                    risk_style.extend(_risk_row_style(i, str(cls.get("risco", ""))))
                story.append(_pdf_table(rows, col_widths=[3 * cm, 2.5 * cm, 4 * cm, 2.5 * cm, 2 * cm, 3 * cm],
                                        extra_style=risk_style))
                story.append(Spacer(1, 6))
    else:
        story.append(Paragraph("Nenhum nó de escoragem encontrado.", s["Normal"]))

    story.extend(_section("5. Variáveis Customizadas"))
    if buckets["variables"]:
        for comp in buckets["variables"]:
            d = comp.get("data", {})
            label = d.get("label", "Variáveis")
            detalhes = d.get("detalhes", "")
            story.append(Paragraph(label + (f" — {detalhes}" if detalhes else ""), s["H3x"]))
            rows = [["Nome interno", "Label", "Tipo", "Valor / Fórmula", "Ocultar"]]
            for var in d.get("variaveis", []):
                val_raw = var.get("valor", "")
                val_text = _tiptap_text(val_raw) if isinstance(val_raw, dict) else str(val_raw)
                rows.append([
                    var.get("nome_variavel", var.get("nome", "")), var.get("label", ""),
                    var.get("tipo", ""), val_text, "Sim" if var.get("ocultar") else "Não",
                ])
            if len(rows) > 1:
                story.append(_pdf_table(rows, col_widths=[3.5 * cm, 3.5 * cm, 2 * cm, 6 * cm, 2 * cm]))
                story.append(Spacer(1, 6))
    else:
        story.append(Paragraph("Nenhum nó de variáveis encontrado.", s["Normal"]))

    story.extend(_section("6. Resumo Técnico"))
    type_counts: dict = {}
    for comp in componentes:
        t = comp.get("type", "unknown")
        type_counts[t] = type_counts.get(t, 0) + 1
    story.append(Paragraph("Distribuição por tipo de nó:", s["Normal"]))
    story.append(Spacer(1, 4))
    story.append(_type_distribution_chart(type_counts))
    story.append(Spacer(1, 10))

    cfg = data.get("configuracao", {})
    story.append(Paragraph(
        f"<b>URL External Hook:</b> {cfg.get('urlExternalHook') or '(não definida)'} &nbsp;|&nbsp; "
        f"<b>Headers:</b> {len(cfg.get('headers', []))} &nbsp;|&nbsp; "
        f"<b>Body params:</b> {len(cfg.get('body', []))}",
        s["Normal"],
    ))
    story.append(Spacer(1, 8))

    final_nodes = [c for c in buckets["status"] if c.get("data", {}).get("final", False)]
    if final_nodes:
        block = [Paragraph("Desfechos finais possíveis:", s["Normal"])]
        for comp in final_nodes:
            d = comp.get("data", {})
            status = _STATUS_LABEL.get(d.get("status", ""), d.get("status", ""))
            msg_raw = d.get("message", "")
            msg = _tiptap_text(msg_raw) if isinstance(msg_raw, dict) else str(msg_raw)
            block.append(Paragraph(f"&bull; <b>{status}</b>: {_html.escape(msg)}", s["Normal"]))
        story.append(KeepTogether(block))

    doc.build(story, onFirstPage=_pdf_footer, onLaterPages=_pdf_footer)
    buf.seek(0)
    return buf.read()


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

data = _fix_mojibake(data)

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

try:
    pdf_bytes = build_motor_pdf(data, componentes, buckets)
    st.download_button(
        label="Baixar resumo em PDF",
        data=pdf_bytes,
        file_name="resumo_motor.pdf",
        mime="application/pdf",
    )
except Exception as e:
    st.error(f"Erro ao gerar PDF: {e}")
    st.exception(e)

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

            doc_str = _format_documento(d.get("documento", ""))

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
                for gidx, grupo in enumerate(grupos):
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

                    grupo_nome = grupo.get("nome", f"Grupo {gidx + 1}")
                    grupo_export = {
                        "detalhes": grupo_nome,
                        "grupos": [grupo],
                        "classificacao": esc.get("classificacao", []),
                        "variavel_especifica": esc.get("variavel_especifica", {}),
                    }
                    st.download_button(
                        label=f"⬇️ Baixar JSON — {grupo_nome}",
                        data=json.dumps(grupo_export, ensure_ascii=False, indent=2).encode("utf-8"),
                        file_name=f"ESCORAGEM-{_safe_filename(grupo_nome)}.json",
                        mime="application/json",
                        key=f"dl_esc_json_{idx}_{gidx}",
                    )
                    st.write("")

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
