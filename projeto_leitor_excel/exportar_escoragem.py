#!/usr/bin/env python3
"""
exportar_escoragem.py
Gera Excel e preview HTML a partir de JSON de motor de escoragem.
"""

import html as _html
import io
import json
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────────────────────────────────────
# Paleta de cores
# ─────────────────────────────────────────────────────────────────────────────

# Cores de score — Excel
_XL_RED    = "FF4444"
_XL_YELLOW = "FFFF00"
_XL_LGREEN = "92D050"
_XL_GREEN  = "00B050"

# Cores de score — CSS
_CSS_COLORS = {
    _XL_RED:    "#FF4444",
    _XL_YELLOW: "#FFFF00",
    _XL_LGREEN: "#92D050",
    _XL_GREEN:  "#00B050",
}

# Fundos claros exigem texto PRETO; fundos escuros exigem texto BRANCO
_LIGHT_BG_XL  = {_XL_YELLOW, _XL_LGREEN, "FFC000", "FFFFFF"}
_LIGHT_BG_CSS = {"#FFFF00", "#92D050", "#FFC000", "#FFFFFF"}

_RISK_COLORS_XL = {
    "muito baixo": _XL_GREEN,
    "baixo":       _XL_LGREEN,
    "médio":       _XL_YELLOW,
    "medio":       _XL_YELLOW,
    "alto":        "FFC000",
    "muito alto":  _XL_RED,
}
_RISK_COLORS_CSS = {
    "muito baixo": "#00B050",
    "baixo":       "#92D050",
    "médio":       "#FFFF00",
    "medio":       "#FFFF00",
    "alto":        "#FFC000",
    "muito alto":  "#FF4444",
}

_THIN   = Side(style="thin",   color="AAAAAA")
_MEDIUM = Side(style="medium", color="1F4E79")
_BTHIN  = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


# ─────────────────────────────────────────────────────────────────────────────
# Helpers de cor
# ─────────────────────────────────────────────────────────────────────────────

def _font_color_xl(bg: str) -> str:
    """Retorna cor de fonte Excel para o fundo dado."""
    return "000000" if bg in _LIGHT_BG_XL else "FFFFFF"


def _font_color_css(bg_css: str) -> str:
    """Retorna cor de fonte CSS para o fundo dado."""
    return "#000000" if bg_css in _LIGHT_BG_CSS else "#FFFFFF"


def _score_color_xl(score, max_score) -> str:
    try:
        s, m = float(score), float(max_score)
    except (TypeError, ValueError):
        return "FFFFFF"
    if m <= 0:
        return "FFFFFF"
    if s == 0:
        return _XL_RED
    pct = s / m
    if pct < 0.6:
        return _XL_YELLOW
    if pct < 1.0:
        return _XL_LGREEN
    return _XL_GREEN


# ─────────────────────────────────────────────────────────────────────────────
# Utilitarios de texto
# ─────────────────────────────────────────────────────────────────────────────

_HTML_ENT = {"&gt;=": ">=", "&lt;=": "<=", "&gt;": ">", "&lt;": "<", "&amp;": "&"}


def _decode(text) -> str:
    if text is None:
        return ""
    if not isinstance(text, str):
        return str(text)
    for k, v in _HTML_ENT.items():
        text = text.replace(k, v)
    return _html.unescape(text)


def _fmt_num(v):
    if v is None or v == "":
        return ""
    if isinstance(v, float) and v == int(v):
        return int(v)
    return v


def _to_str(v) -> str:
    if v is None or v == "":
        return ""
    return str(_fmt_num(v))


def build_condition(cond1=None, val1=None, cond2=None, val2=None) -> str:
    c1 = _decode(cond1)
    c2 = _decode(cond2)
    v1 = _fmt_num(val1)
    v2 = _fmt_num(val2)
    if not c1:
        return ""
    if c1.lower() == "vazio":
        return "vazio"
    if c1 == "==" and (v2 == "" or v2 is None):
        return f"= {v1}"
    if not c2 or v2 == "" or v2 is None:
        return f"{c1} {v1}"
    return f"{c1} {v1} e {c2} {v2}"


def _fmt_peso(peso) -> str:
    if isinstance(peso, float) and 0 < peso <= 1:
        return f"{peso * 100:.0f}%"
    return f"{peso}%" if peso not in ("", None, 0) else ""


def _fmt_validade(val) -> str:
    if val is None or val == "":
        return ""
    if isinstance(val, (int, float)) and val > 0:
        return f"{int(val)} dias"
    return str(val)


def safe_filename(text: str) -> str:
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", text).strip() or "escoragem"


def _clean_var_name(name: str) -> str:
    """Remove sufixos de I.A do nome da variavel para exibicao."""
    name = re.sub(r'\s*[-–]\s*I\.?\s*A\.?\s*$', '', name, flags=re.IGNORECASE)
    return name.strip()


# ─────────────────────────────────────────────────────────────────────────────
# Parsing
# ─────────────────────────────────────────────────────────────────────────────

def _parse_grupos(data: dict) -> List[Dict]:
    result = []
    for grupo in data.get("grupos", []):
        variaveis = []
        for var in grupo.get("variaveis", []):
            pts = var.get("pontuacao", [])
            nums = []
            for p in pts:
                s = p.get("faixa_pontuacao")
                if s is not None and s != "":
                    try:
                        nums.append(float(s))
                    except (ValueError, TypeError):
                        pass
            max_s = max(nums) if nums else 0
            variaveis.append({
                "nome":       _clean_var_name(var.get("nome", "")),
                "aplicar_a":  var.get("aplicar_a", ""),
                "peso":       var.get("peso", 0),
                "max_score":  max_s,
                "pontuacoes": [
                    {
                        "condicao": build_condition(
                            p.get("condicao_1"), p.get("pontuacao_1"),
                            p.get("condicao_2"), p.get("pontuacao_2"),
                        ),
                        "score": _fmt_num(p.get("faixa_pontuacao", 0)),
                    }
                    for p in pts
                ],
            })
        result.append({"nome": grupo.get("nome", ""), "variaveis": variaveis})
    return result


def _parse_classificacao(data: dict) -> List[Dict]:
    result = []
    for cls in data.get("classificacao", []):
        classe   = cls.get("classe") or cls.get("nome", "")
        risco    = cls.get("risco", "")
        condicao = build_condition(
            cls.get("condicao_1"), cls.get("pontuacao_1"),
            cls.get("condicao_2"), cls.get("pontuacao_2"),
        )
        fator = _fmt_num(cls.get("faixa_pontuacao", ""))
        result.append({
            "classe":          classe,
            "risco":           risco,
            "condicao":        condicao,
            "fator":           _to_str(fator),
            "validade":        _fmt_validade(cls.get("validade_limite")),
            "limite_sugerido": _to_str(cls.get("limite_sugerido", "")),
        })
    return result


# ─────────────────────────────────────────────────────────────────────────────
# Helpers de celula Excel
# ─────────────────────────────────────────────────────────────────────────────

def _hdr(cell, text: str, bg: str = "1F4E79", fg: str = "FFFFFF",
         bold: bool = True, rot: int = 0) -> None:
    cell.value     = text
    cell.font      = Font(bold=bold, color=fg, name="Calibri", size=10)
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True, text_rotation=rot)
    cell.border    = _BTHIN


def _dat(cell, value, bg: str = "", bold: bool = False,
         align: str = "center", wrap: bool = False, font_color: str = "") -> None:
    cell.value = value
    fc = font_color or (_font_color_xl(bg) if bg else "000000")
    cell.font      = Font(name="Calibri", size=10, bold=bold, color=fc)
    if bg:
        cell.fill  = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    cell.border    = _BTHIN


# ─────────────────────────────────────────────────────────────────────────────
# Aba EscoreSugerido
# ─────────────────────────────────────────────────────────────────────────────

def _write_escoragem_sheet(ws, grupos: List[Dict], model_name: str) -> None:
    # Colunas: A=Grupo | B=Variavel | C=Faixa de Pontuacao | D=Score | E=Peso
    ws.merge_cells("A1:E1")
    c = ws["A1"]
    c.value     = "REGRA MOTOR ESCORAGEM"
    c.font      = Font(bold=True, size=14, name="Calibri", color="1F4E79")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    hdrs = [
        ("Grupo",              "1F4E79", "FFFFFF"),
        ("Variaveis\nEscore",  "1F4E79", "FFFFFF"),
        ("Faixa de\nPontuacao","1F4E79", "FFFFFF"),
        ("Score\n(0-100)",     "C00000", "FFFFFF"),
        ("Peso\n(0-100)%",     "C00000", "FFFFFF"),
    ]
    for col, (txt, bg, fg) in enumerate(hdrs, 1):
        _hdr(ws.cell(2, col), txt, bg=bg, fg=fg)
    ws.row_dimensions[2].height = 30

    cur = 3
    for grupo in grupos:
        g_start = cur
        g_rows  = sum(len(v["pontuacoes"]) for v in grupo["variaveis"])

        for var in grupo["variaveis"]:
            v_start = cur
            max_s   = var["max_score"]

            for regra in var["pontuacoes"]:
                score = regra["score"]
                try:
                    sf   = float(score) if score != "" else 0.0
                    xcol = _score_color_xl(sf, max_s)
                    sdsp = _fmt_num(score) if score != "" else 0
                except (TypeError, ValueError):
                    xcol, sdsp = "FFFFFF", score

                _dat(ws.cell(cur, 3), regra["condicao"], align="left")
                _dat(ws.cell(cur, 4), sdsp, bg=xcol, bold=True)
                cur += 1

            v_end = cur - 1
            for col in (2, 5):
                if v_end > v_start:
                    ws.merge_cells(
                        f"{get_column_letter(col)}{v_start}:{get_column_letter(col)}{v_end}"
                    )
            _dat(ws.cell(v_start, 2), var["nome"],            align="center", wrap=True)
            _dat(ws.cell(v_start, 5), _fmt_peso(var["peso"]), align="center")

        g_end = cur - 1
        if g_end > g_start:
            ws.merge_cells(f"A{g_start}:A{g_end}")
        gc = ws.cell(g_start, 1)
        gc.value     = grupo["nome"]
        gc.font      = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        gc.fill      = PatternFill("solid", fgColor="2E75B6")
        gc.alignment = Alignment(horizontal="center", vertical="center",
                                  text_rotation=90, wrap_text=False)
        gc.border    = Border(left=_MEDIUM, right=_THIN, top=_MEDIUM, bottom=_MEDIUM)

    for col, w in enumerate([8, 32, 32, 12, 12], 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "C3"


# ─────────────────────────────────────────────────────────────────────────────
# Aba Classificacao de Risco  +  secao Limite Sugerido na parte inferior
# ─────────────────────────────────────────────────────────────────────────────

def _write_classificacao_sheet(ws, classificacao: List[Dict]) -> None:
    # ── Titulo ────────────────────────────────────────────────────────────
    ws.merge_cells("A1:D1")
    c = ws["A1"]
    c.value     = "CLASSIFICACAO DE RISCO"
    c.font      = Font(bold=True, size=14, name="Calibri", color="1F4E79")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    # ── Cabecalhos — apenas faixas ────────────────────────────────────────
    for col, hdr in enumerate(["Classe", "Risco", "Faixa de Pontuacao", "Validade"], 1):
        _hdr(ws.cell(2, col), hdr)
    ws.row_dimensions[2].height = 24

    # ── Linhas de classificacao ───────────────────────────────────────────
    for ri, cls in enumerate(classificacao, 3):
        risco_key = cls["risco"].lower().replace("é", "e")
        bg = _RISK_COLORS_XL.get(risco_key, "FFFFFF")
        _dat(ws.cell(ri, 1), cls["classe"],   bg=bg, bold=True)
        _dat(ws.cell(ri, 2), cls["risco"],    bg=bg)
        _dat(ws.cell(ri, 3), cls["condicao"])
        _dat(ws.cell(ri, 4), cls["validade"])
        ws.row_dimensions[ri].height = 20

    last_cls_row = 2 + len(classificacao)

    # ── Espacador ─────────────────────────────────────────────────────────
    sep_row = last_cls_row + 2

    # ── Titulo da secao inferior ──────────────────────────────────────────
    ws.merge_cells(f"A{sep_row}:D{sep_row}")
    t = ws.cell(sep_row, 1)
    t.value     = "RESUMO — CALCULO DE LIMITE SUGERIDO"
    t.font      = Font(bold=True, size=12, name="Calibri", color="FFFFFF")
    t.fill      = PatternFill("solid", fgColor="1F4E79")
    t.alignment = Alignment(horizontal="center", vertical="center")
    t.border    = _BTHIN
    ws.row_dimensions[sep_row].height = 26

    # ── Cabecalhos da secao inferior ──────────────────────────────────────
    sub_row = sep_row + 1
    for col, hdr in enumerate(["Classe", "Risco", "Fator de Limite", "Limite Sugerido"], 1):
        _hdr(ws.cell(sub_row, col), hdr, bg="2E75B6")
    ws.row_dimensions[sub_row].height = 22

    # ── Linhas da secao inferior ──────────────────────────────────────────
    for i, cls in enumerate(classificacao):
        ri2 = sub_row + 1 + i
        risco_key = cls["risco"].lower().replace("é", "e")
        bg = _RISK_COLORS_XL.get(risco_key, "FFFFFF")
        _dat(ws.cell(ri2, 1), cls["classe"],          bg=bg, bold=True)
        _dat(ws.cell(ri2, 2), cls["risco"],           bg=bg)
        _dat(ws.cell(ri2, 3), cls["fator"],           align="center")
        # Se limite_sugerido estiver no JSON usa; senao mostra formula generica
        limite_txt = cls["limite_sugerido"] if cls["limite_sugerido"] else "Fator x Capacidade de Pagamento"
        _dat(ws.cell(ri2, 4), limite_txt, align="left")
        ws.row_dimensions[ri2].height = 20

    # ── Nota explicativa ──────────────────────────────────────────────────
    nota_row = sub_row + 1 + len(classificacao) + 1
    ws.merge_cells(f"A{nota_row}:D{nota_row + 2}")
    nota = ws.cell(nota_row, 1)
    nota.value = (
        "Como interpretar:\n"
        "O Limite Sugerido e calculado aplicando o Fator de Limite da classe do cliente "
        "sobre a Capacidade de Pagamento Mensal apurada no motor.\n"
        "Exemplo: Classe A — Fator 1 — Limite = 1 x Capacidade de Pagamento."
    )
    nota.font      = Font(name="Calibri", size=9, italic=True, color="444444")
    nota.alignment = Alignment(horizontal="left", vertical="top",
                                wrap_text=True)
    nota.border    = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    nota.fill = PatternFill("solid", fgColor="F5F5F5")
    for r in range(nota_row, nota_row + 3):
        ws.row_dimensions[r].height = 18

    # ── Larguras ──────────────────────────────────────────────────────────
    for col, w in enumerate([16, 16, 28, 36], 1):
        ws.column_dimensions[get_column_letter(col)].width = w


# ─────────────────────────────────────────────────────────────────────────────
# Montar workbook
# ─────────────────────────────────────────────────────────────────────────────

def _build_workbook(data: dict) -> Tuple[str, Workbook]:
    model_name    = data.get("detalhes", "escoragem")
    grupos        = _parse_grupos(data)
    classificacao = _parse_classificacao(data)

    wb = Workbook()
    wb.remove(wb.active)

    _write_escoragem_sheet(wb.create_sheet("EscoreSugerido"), grupos, model_name)
    _write_classificacao_sheet(wb.create_sheet("Classificacao de Risco"), classificacao)

    return safe_filename(model_name) + ".xlsx", wb


def export_to_bytes(json_bytes: bytes) -> Tuple[str, bytes]:
    data      = json.loads(json_bytes)
    fname, wb = _build_workbook(data)
    buf       = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return fname, buf.read()


def export_json_to_excel(json_path: Path, output_dir: Optional[Path] = None) -> Path:
    with json_path.open(encoding="utf-8") as f:
        data = json.load(f)
    fname, wb = _build_workbook(data)
    out_dir   = output_dir or json_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path  = out_dir / fname
    wb.save(out_path)
    print(f"Exportado: {out_path}")
    return out_path


# ─────────────────────────────────────────────────────────────────────────────
# CSS base compartilhado
# ─────────────────────────────────────────────────────────────────────────────

_BASE_CSS = """
<style>
  /* Reset — garante fundo branco e texto preto independente do tema do Streamlit */
  html, body { background: #ffffff; color: #000000; }
  * { box-sizing: border-box; margin: 0; padding: 0; }

  .esc-wrap {
    overflow-x: auto;
    overflow-y: auto;
    max-height: 640px;
    border: 1px solid #D0D7E3;
    border-radius: 8px;
    background: #ffffff;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
  }
  .esc-tbl {
    border-collapse: collapse;
    font-family: Calibri, Arial, sans-serif;
    font-size: 12px;
    width: 100%;
    min-width: 700px;
    background: #ffffff;
  }
  /* Todas as celulas de dados: texto PRETO por padrao */
  .esc-tbl td {
    border: 1px solid #D0D7E3;
    padding: 5px 9px;
    text-align: center;
    vertical-align: middle;
    color: #000000;
    background: #ffffff;
  }
  /* Cabecalhos: texto BRANCO */
  .esc-tbl th {
    border: 1px solid #1A406A;
    padding: 6px 9px;
    text-align: center;
    vertical-align: middle;
    background: #1F4E79;
    color: #ffffff;
    font-weight: bold;
    white-space: nowrap;
  }
  .hdr-score { background: #C00000 !important; color: #ffffff !important; }

  .esc-title {
    background: #F0F4FA;
    color: #1F4E79;
    font-size: 13px;
    font-weight: bold;
    text-align: center;
    padding: 10px;
    letter-spacing: 0.5px;
    border-radius: 8px 8px 0 0;
  }
  .esc-grupo {
    background: #2E75B6;
    color: #ffffff;
    font-weight: bold;
    writing-mode: vertical-lr;
    transform: rotate(180deg);
    white-space: nowrap;
    width: 26px;
    padding: 8px 3px;
  }
  .esc-var {
    text-align: left;
    padding-left: 10px;
    min-width: 150px;
    max-width: 220px;
    white-space: normal;
    word-break: break-word;
    color: #000000;
    background: #ffffff;
  }
  .esc-faixa {
    text-align: left;
    padding-left: 8px;
    min-width: 130px;
    color: #000000;
    background: #ffffff;
  }
  /* Linhas alternadas leves */
  .esc-tbl tbody tr:nth-child(even) td:not([style*="background"]) {
    background: #F7F9FC;
  }
</style>
"""


# ─────────────────────────────────────────────────────────────────────────────
# Preview HTML — Variaveis e Regras
# ─────────────────────────────────────────────────────────────────────────────

def build_preview_html(data: dict) -> str:
    grupos     = _parse_grupos(data)
    model_name = data.get("detalhes", "Escoragem")
    has_aplic  = any(v["aplicar_a"] for g in grupos for v in g["variaveis"])

    n_cols = 7 + (1 if has_aplic else 0)

    th  = "<th>Grupo</th><th>Variaveis Escore</th>"
    th += "<th>Aplicar a</th>" if has_aplic else ""
    th += "<th>Faixa de Pontuacao</th>"
    th += '<th class="hdr-score">Score (0-100)</th>'
    th += '<th class="hdr-score">% do Max</th>'
    th += '<th class="hdr-score">N Variavel</th>'
    th += '<th class="hdr-score">Peso (0-100)%</th>'

    rows = []
    for grupo in grupos:
        g_total = sum(len(v["pontuacoes"]) for v in grupo["variaveis"])
        g_done  = False

        for var in grupo["variaveis"]:
            v_cnt   = len(var["pontuacoes"])
            max_s   = var["max_score"]
            v_first = True

            for regra in var["pontuacoes"]:
                score = regra["score"]
                try:
                    sf    = float(score) if score != "" else 0.0
                    xkey  = _score_color_xl(sf, max_s)
                    bg_c  = _CSS_COLORS.get(xkey, "#FFFFFF")
                    fc    = _font_color_css(bg_c)
                    pct   = sf / max_s * 100 if max_s > 0 else 0.0
                    ptxt  = f"{pct:.1f}%"
                    sdsp  = str(_fmt_num(score)) if score != "" else "0"
                except (TypeError, ValueError):
                    bg_c, fc, ptxt, sdsp = "#FFFFFF", "#000000", "", str(score)

                r = "<tr>"
                if not g_done:
                    r += f'<td class="esc-grupo" rowspan="{g_total}">{_html.escape(grupo["nome"])}</td>'
                    g_done = True

                if v_first:
                    r += f'<td class="esc-var" rowspan="{v_cnt}">{_html.escape(var["nome"])}</td>'
                    if has_aplic:
                        r += f'<td rowspan="{v_cnt}">{_html.escape(var["aplicar_a"])}</td>'

                r += f'<td class="esc-faixa">{_html.escape(regra["condicao"])}</td>'
                r += (f'<td style="background:{bg_c};color:{fc};font-weight:bold">{sdsp}</td>')
                r += (f'<td style="background:{bg_c};color:{fc}">{ptxt}</td>')

                if v_first:
                    nvar = str(_fmt_num(max_s)) if max_s else ""
                    r += f'<td rowspan="{v_cnt}" style="font-weight:bold">{nvar}</td>'
                    r += f'<td rowspan="{v_cnt}">{_fmt_peso(var["peso"])}</td>'
                    v_first = False

                r += "</tr>"
                rows.append(r)

    return f"""
{_BASE_CSS}
<div class="esc-wrap">
  <table class="esc-tbl">
    <thead>
      <tr><td class="esc-title" colspan="{n_cols}">
        REGRA MOTOR ESCORAGEM &mdash; {_html.escape(model_name)}
      </td></tr>
      <tr>{th}</tr>
    </thead>
    <tbody>{"".join(rows)}</tbody>
  </table>
</div>
"""


# ─────────────────────────────────────────────────────────────────────────────
# Preview HTML — Classificacao (so faixas)
# ─────────────────────────────────────────────────────────────────────────────

def build_classificacao_html(data: dict) -> str:
    classificacao = _parse_classificacao(data)
    model_name    = data.get("detalhes", "Escoragem")

    if not classificacao:
        return "<p>Sem dados de classificacao.</p>"

    rows = []
    for cls in classificacao:
        bg  = _RISK_COLORS_CSS.get(cls["risco"].lower(), "#FFFFFF")
        fc  = _font_color_css(bg)
        rows.append(
            f"<tr>"
            f'<td style="background:{bg};color:{fc};font-weight:bold">'
            f'{_html.escape(cls["classe"])}</td>'
            f'<td style="background:{bg};color:{fc}">'
            f'{_html.escape(cls["risco"])}</td>'
            f'<td>{_html.escape(cls["condicao"])}</td>'
            f'<td>{_html.escape(cls["validade"])}</td>'
            f"</tr>"
        )

    return f"""
{_BASE_CSS}
<style>
  .cls-tbl {{
    border-collapse: collapse;
    font-family: Calibri, Arial, sans-serif;
    font-size: 12px;
    width: 100%;
    max-width: 680px;
    background: #ffffff;
  }}
  .cls-tbl td {{
    border: 1px solid #D0D7E3;
    padding: 7px 14px;
    text-align: center;
    vertical-align: middle;
    color: #000000;
    background: #ffffff;
  }}
  .cls-tbl th {{
    border: 1px solid #1A406A;
    padding: 7px 14px;
    text-align: center;
    background: #1F4E79;
    color: #ffffff;
    font-weight: bold;
  }}
</style>
<div style="overflow-x:auto; border:1px solid #D0D7E3; border-radius:8px;
            box-shadow:0 1px 4px rgba(0,0,0,0.08); background:#fff;">
  <table class="cls-tbl">
    <thead>
      <tr>
        <td colspan="4" style="background:#F0F4FA;color:#1F4E79;font-family:Calibri,Arial,sans-serif;
            font-size:13px;font-weight:bold;text-align:center;padding:10px;
            border-bottom:1px solid #D0D7E3;border-radius:8px 8px 0 0;">
          CLASSIFICACAO DE RISCO &mdash; {_html.escape(model_name)}
        </td>
      </tr>
      <tr>
        <th>Classe</th>
        <th>Risco</th>
        <th>Faixa de Pontuacao</th>
        <th>Validade</th>
      </tr>
    </thead>
    <tbody>{"".join(rows)}</tbody>
  </table>
</div>
"""


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    args = sys.argv[1:]
    if not args:
        files = sorted(Path(".").glob("*.json"))
        if not files:
            print("Nenhum .json encontrado.")
            sys.exit(1)
        for f in files:
            try:
                export_json_to_excel(f)
            except Exception as e:
                print(f"Erro em {f.name}: {e}", file=sys.stderr)
        return
    p = Path(args[0])
    if not p.is_file():
        print(f"Arquivo nao encontrado: {p}", file=sys.stderr)
        sys.exit(1)
    export_json_to_excel(p, Path(args[1]) if len(args) > 1 else None)


if __name__ == "__main__":
    main()
