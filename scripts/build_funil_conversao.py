# -*- coding: utf-8 -*-
"""Deck 'Funil lead -> venda por canal' (CV, Liza e Central, jan-ago/26).

Uma pagina de resumo e uma pagina por canal, todas com tabelas nativas do
PowerPoint: cada numero fica na propria celula, entao da para selecionar o
bloco e colar direto no Excel. As taxas sao recalculadas a partir dos volumes
brutos -- nada e transcrito de valores ja arredondados -- e a linha de
conversao final recebe heatmap contra a meta de 0,94%.

Os rotulos e a ordem das linhas seguem a planilha de origem, para o bloco colar
alinhado com ela.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

from deckstyle import cell_border, cell_box, cell_text, set_table_plain, set_row_heights

# ---------------------------------------------------------------- tokens
RED    = RGBColor(0xCC, 0x00, 0x00)   # Bain red
RUBY   = RGBColor(0x99, 0x00, 0x00)   # negativo
FOREST = RGBColor(0x10, 0x4C, 0x3E)   # positivo
GREY1  = RGBColor(0x33, 0x33, 0x33)
GREY2  = RGBColor(0x5C, 0x5C, 0x5C)
GREY3  = RGBColor(0x85, 0x85, 0x85)
GREY4  = RGBColor(0xB4, 0xB4, 0xB4)
GREY5  = RGBColor(0xDC, 0xDC, 0xDC)
BAND   = RGBColor(0xF2, 0xF2, 0xF2)   # linhas de taxa sobre o total de leads
METABG = RGBColor(0xEC, 0xEC, 0xEC)   # coluna de meta
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
FONT   = "Arial"
LANG   = "pt-BR"

# ---------------------------------------------------------------- grid (in)
SW, SH = 13.333, 7.5
MARGIN_L, MARGIN_R = 0.30, 13.03
TITLE_Y, TITLE_H = 0.17, 0.62
RULE_Y = 0.90
READ_Y, READ_H = 1.00, 0.30
TBL_Y = 1.40
NOTE_Y = 6.98

MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago"]
META_CONV = 0.94                      # meta de conversao lead -> venda (%)

# larguras da tabela de detalhe: rotulo + 8 meses + meta
LBL_W, META_W = 2.40, 1.05
MON_W = (MARGIN_R - MARGIN_L - LBL_W - META_W) / 8.0
HEAD_H, ROW_H = 0.36, 0.455

# ---------------------------------------------------------------- dados
# volumes brutos por mes (jan a ago/26); None = canal ainda nao operava
N = None
CANAIS = [
    dict(
        nome="CV (sem Liza e sem Central)",
        meta_leads=400000,
        leads=[454314, 351045, 387905, 426871, 435764, 336078, 324671, 321406],
        cep=[N, N, N, 0.46, 0.48, 0.47, 0.47, 0.39],
        env=[16048, 13734, 15604, 16497, 17662, 19035, 16326, 15215],
        apr=[5711, 4703, 5161, 5203, 5161, 5256, 4648, 4619],
        fat=[1933, 1553, 1803, 1654, 1575, 1646, 1527, 1377],
        ven=[3416, 2674, 3183, 2816, 2742, 2669, 2480, 2144],
        titulo=("CV perde escala e conversão ao mesmo tempo: leads −29% e conversão −11% "
                "(0,75%→0,67%), com o funil interno praticamente estável"),
        leitura=("Queda vem do topo, não do funil: fichas enviadas sobre leads sobe de 3,5% para 4,7%, "
                 "mas aprovação cai de 36% para 30% e o volume de leads recua 29%"),
    ),
    dict(
        nome="Liza",
        meta_leads=50000,
        leads=[N, N, N, 8117, 13735, 14688, 31018, 58349],
        cep=[N, N, N, 0.54, 0.55, 0.53, 0.49, 0.47],
        env=[N, N, N, 584, 890, 1140, 1557, 2684],
        apr=[N, N, N, 141, 236, 285, 466, 809],
        fat=[N, N, N, 50, 81, 85, 168, 248],
        ven=[N, N, N, 82, 157, 166, 297, 389],
        titulo=("Liza escala 7,2x em cinco meses e a conversão cai de 1,01% para 0,67% (−34%), "
                "saindo de acima da meta para abaixo dela"),
        leitura=("Trade-off volume × qualidade: até jul/26 a Liza operou acima ou na meta de 0,94%; "
                 "o salto de 31k para 58k leads em ago/26 derrubou a conversão a 0,67%"),
    ),
    dict(
        nome="Central",
        meta_leads=20000,
        leads=[14502, 12085, 16248, 14333, 13311, 10660, 15577, 9060],
        cep=[N, N, N, 0.64, 0.73, 0.68, 0.54, 0.55],
        env=[134, 137, 167, 175, 175, 188, 719, 397],
        apr=[69, 66, 82, 71, 78, 91, 543, 320],
        fat=[29, 35, 43, 28, 29, 32, 88, 68],
        ven=[116, 101, 136, 98, 101, 83, 210, 107],
        titulo=("Central é o único canal acima da meta: conversão sobe 48% (0,80%→1,18%) com "
                "aprovação de fichas saltando de 51% para 81%"),
        leitura=("Ganho é de qualidade, não de escala: leads caem 38% (14,5k→9,1k), mas o envio de "
                 "fichas sobe de 0,9% para 4,4% dos leads e a aprovação quase dobra"),
    ),
]

# linhas do funil, na ordem e com os rotulos da planilha de origem
LINHAS = [
    ("# de leads",                  "leads",     "num"),
    ("% leads sem info de CEP",     "cep",       "pct0"),
    ("% de fichas enviadas",        "env_leads", "pct1"),
    ("# fichas enviadas",           "env",       "num"),
    ("% fichas aprovadas",          "apr_env",   "pct0"),
    ("# fichas aprovadas",          "apr",       "num"),
    ("% fichas faturadas",          "fat_apr",   "pct0"),
    ("# fichas faturadas",          "fat",       "num"),
    ("% total de fichas faturadas", "fat_leads", "pct2"),
    ("# de vendas totais",          "ven",       "num"),
    ("% de vendas",                 "ven_leads", "pct2"),
]

TITULO_RESUMO = ("Conversão lead→venda cai ~10% no consolidado (0,75%→0,68%) e segue 0,26 p.p. "
                 "abaixo da meta de 0,94%: só a Central avança, Liza perde conversão ao escalar")

NOTA_BASE = ("Nota: CV exclui Liza e Central. Taxas recalculadas sobre os volumes brutos; "
             "% de fichas enviadas, % total de fichas faturadas e % de vendas são sobre o total de leads "
             "do mês, % fichas aprovadas sobre as enviadas e % fichas faturadas sobre as aprovadas. "
             "Meta de conversão de 0,94% aplicada sobre a meta mensal de leads de cada canal "
             "(CV 400k, Liza 50k, Central 20k). % leads sem info de CEP não disponível para jan-mar/26; "
             "Liza inicia em abr/26. Fonte: base de leads, base de fichas e base de vendas faturadas")

LEITURAS = [
    ("O problema é de topo de funil, não do funil interno",
     "Os leads consolidados caíram 17% (469k→389k) e a conversão, 10%. As taxas internas de "
     "aprovação e faturamento seguem no mesmo patamar de janeiro."),
    ("Liza mostra trade-off claro entre volume e qualidade",
     "Operou acima da meta até jul/26, com até 31k leads/mês. Ao saltar para 58k em ago/26, "
     "a conversão caiu a 0,67% — vale checar a origem do volume incremental."),
    ("Central ganha conversão sem precisar de escala",
     "Com 38% menos leads, subiu a conversão a 1,18% elevando o envio de fichas "
     "(0,9%→4,4% dos leads) e a aprovação (51%→81%). Único canal acima da meta."),
]

NOTA_RESUMO = ("Nota: CV exclui Liza e Central. Primeiro mês = jan/26, exceto Liza, que inicia em abr/26. "
               "Conversão = vendas totais / leads do mês. Meta consolidada = soma das metas mensais de "
               "leads dos três canais (470k) à conversão-meta de 0,94%. Detalhe mês a mês de cada canal "
               "nas páginas seguintes. Fonte: base de leads e base de vendas faturadas")


# ---------------------------------------------------------------- formatacao
def fmt(v, kind):
    """numeros em pt-BR, um valor por celula, sem abreviacao: o bloco cola no Excel"""
    if v is None:
        return "–"
    if kind == "num":
        return "{:,}".format(int(round(v))).replace(",", ".")
    return ("%.*f%%" % (int(kind[-1]), v)).replace(".", ",")


def ratio(num, den):
    if num is None or den in (None, 0):
        return None
    return num / float(den) * 100.0


def serie(canal, chave):
    """valores mes a mes de uma linha do funil"""
    L, E, A, F, V = (canal[k] for k in ("leads", "env", "apr", "fat", "ven"))
    if chave in ("leads", "env", "apr", "fat", "ven"):
        return canal[chave]
    if chave == "cep":
        return [None if c is None else c * 100 for c in canal["cep"]]
    pares = {"env_leads": (E, L), "apr_env": (A, E),
             "fat_apr": (F, A), "fat_leads": (F, L), "ven_leads": (V, L)}
    num, den = pares[chave]
    return [ratio(n, d) for n, d in zip(num, den)]


def meta_cell(canal, chave):
    """valor da coluna Meta; None = celula vazia"""
    if chave == "leads":
        return fmt(canal["meta_leads"], "num")
    if chave == "ven":
        return fmt(canal["meta_leads"] * META_CONV / 100.0, "num")
    if chave == "ven_leads":
        return fmt(META_CONV, "pct2")
    return None


def mix(a, b, t):
    t = max(0.0, min(1.0, t))
    return RGBColor(*[int(round(x + (y - x) * t)) for x, y in zip(a, b)])


def heat(v):
    """fundo e cor do texto da linha de conversao, contra a meta de 0,94%"""
    if v is None:
        return None, GREY4, False
    r = v / META_CONV
    if r >= 1.0:
        return mix((0xE1, 0xEE, 0xE8), (0xAF, 0xD2, 0xC2), (r - 1.0) / 0.5), FOREST, True
    return mix((0xFF, 0xFF, 0xFF), (0xEC, 0xAC, 0xAC), 1.0 - r), \
        (RUBY if r < 0.6 else GREY1), r < 0.6


def delta_conv(canal):
    """variacao da conversao entre o primeiro e o ultimo mes com dados"""
    vals = [x for x in serie(canal, "ven_leads") if x is not None]
    return (vals[-1] / vals[0] - 1.0) * 100.0


def primeiro_mes(canal):
    return next(i for i, v in enumerate(canal["leads"]) if v is not None)


# ---------------------------------------------------------------- helpers
def txt(slide, x, y, w, h, parts, size=9, color=GREY1, bold=False,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, line=None):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    if line:
        p.line_spacing = line
    if isinstance(parts, str):
        parts = [(parts, {})]
    for t, o in parts:
        if t == "\n":
            p = tf.add_paragraph()
            p.alignment = align
            if line:
                p.line_spacing = line
            continue
        r = p.add_run()
        r.text = t
        f = r.font
        f.name = FONT
        f.size = Pt(o.get("size", size))
        f.bold = o.get("bold", bold)
        f.color.rgb = o.get("color", color)
        r.font._rPr.set('lang', LANG)
    return tb


def hline(slide, x1, y, x2, color=GREY4, w=0.75):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y),
                                   Inches(x2), Inches(y))
    c.line.color.rgb = color
    c.line.width = Pt(w)
    return c


def nova_pagina(prs, titulo, leitura=None, nota=NOTA_BASE):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    txt(slide, MARGIN_L, TITLE_Y, 10.85, TITLE_H, titulo, size=18, color=GREY1,
        anchor=MSO_ANCHOR.TOP, line=1.06)
    txt(slide, 11.25, TITLE_Y + 0.36, 1.78, 0.22, "/ P R E L I M I N A R", size=9,
        color=GREY3, align=PP_ALIGN.RIGHT)
    hline(slide, MARGIN_L, RULE_Y, MARGIN_R, RED, 2.25)
    if leitura:
        txt(slide, MARGIN_L, READ_Y, MARGIN_R - MARGIN_L, READ_H, leitura,
            size=9.5, color=GREY2, anchor=MSO_ANCHOR.TOP, line=1.12)
    txt(slide, MARGIN_L, NOTE_Y, MARGIN_R - MARGIN_L, 0.48, nota, size=9,
        color=GREY3, anchor=MSO_ANCHOR.TOP, line=1.18)
    return slide


def nova_tabela(slide, x, y, larguras, alturas):
    tbl = slide.shapes.add_table(len(alturas), len(larguras), Inches(x), Inches(y),
                                 Inches(sum(larguras)), Inches(sum(alturas))).table
    set_table_plain(tbl)
    for i, w in enumerate(larguras):
        tbl.columns[i].width = Inches(w)
    set_row_heights(tbl, alturas)
    for r in range(len(alturas)):
        for c in range(len(larguras)):
            cell_box(tbl.cell(r, c), fill=None)
    return tbl


# ---------------------------------------------------------------- pagina de canal
def pagina_canal(prs, canal):
    slide = nova_pagina(prs, canal["titulo"], canal["leitura"])

    d = delta_conv(canal)
    txt(slide, 10.60, READ_Y - 0.02, 2.43, 0.34,
        [("Conversão lead→venda  ", {"size": 9, "color": GREY2}),
         ("%+d%%" % round(d), {"size": 14, "bold": True,
                               "color": FOREST if d > 0 else RUBY})],
        align=PP_ALIGN.RIGHT)

    larg = [LBL_W] + [MON_W] * 8 + [META_W]
    alt = [HEAD_H] + [ROW_H] * len(LINHAS)
    tbl = nova_tabela(slide, MARGIN_L, TBL_Y, larg, alt)

    # cabecalho: nome do canal, meses e coluna de meta
    cell_text(tbl.cell(0, 0), canal["nome"], 10, GREY1, bold=True,
              anchor=MSO_ANCHOR.MIDDLE, mx=0.06)
    for i, mes in enumerate(MESES):
        cell_text(tbl.cell(0, i + 1), "%s/26" % mes, 9.5, GREY1, bold=True,
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    cell_text(tbl.cell(0, 9), "Meta", 9.5, RED, bold=True,
              align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    for c in range(10):
        cell_border(tbl.cell(0, c), 'T', GREY1, 1.25)
        cell_border(tbl.cell(0, c), 'B', GREY1, 1.25)

    for ri, (rotulo, chave, kind) in enumerate(LINHAS, start=1):
        vals = serie(canal, chave)
        # taxas medidas sobre o total de leads ganham fundo, para separar do
        # resto do funil, que usa a etapa anterior como base
        sobre_leads = chave in ("fat_leads", "ven_leads")
        forte = chave == "ven_leads"
        fundo = BAND if sobre_leads and not forte else None
        size = 10 if kind == "num" or sobre_leads else 9.5

        cell_text(tbl.cell(ri, 0), rotulo, size,
                  GREY1 if rotulo.startswith("#") or sobre_leads else GREY2,
                  bold=rotulo.startswith("#") or sobre_leads,
                  anchor=MSO_ANCHOR.MIDDLE, mx=0.06)
        cell_box(tbl.cell(ri, 0), fill=fundo)

        for i, v in enumerate(vals):
            cell = tbl.cell(ri, i + 1)
            cor = GREY1 if kind == "num" or sobre_leads else GREY2
            bold = kind == "num" or sobre_leads
            if forte:
                fill_heat, cor, bold = heat(v)
                cell_box(cell, fill=fill_heat)
            else:
                cell_box(cell, fill=fundo)
            if v is None:
                cor = GREY4
            cell_text(cell, fmt(v, kind), size, cor, bold=bold,
                      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        mv = meta_cell(canal, chave)
        cell_box(tbl.cell(ri, 9), fill=METABG)
        cell_text(tbl.cell(ri, 9), mv or "", 10, RED, bold=True,
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        for c in range(10):
            cell_border(tbl.cell(ri, c), 'T', GREY5, 0.5)
        if forte:
            for c in range(10):
                cell_border(tbl.cell(ri, c), 'T', RED, 1.25)
                cell_border(tbl.cell(ri, c), 'B', RED, 1.25)

    cell_box(tbl.cell(0, 9), fill=METABG,
             borders=(None, None, GREY1, GREY1), lw=1.25)
    cell_text(tbl.cell(0, 9), "Meta", 9.5, RED, bold=True,
              align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return slide


# ---------------------------------------------------------------- pagina de resumo
COLS_RESUMO = ["Canal", "Leads 1º mês", "Leads ago/26", "Δ leads",
               "Vendas 1º mês", "Vendas ago/26", "Conv. 1º mês", "Conv. ago/26",
               "Δ conversão"]


def linha_resumo(nome, leads, vendas, i0):
    c0 = vendas[i0] / leads[i0] * 100.0
    c1 = vendas[-1] / leads[-1] * 100.0
    return [nome, fmt(leads[i0], "num"), fmt(leads[-1], "num"),
            "%+d%%" % round((leads[-1] / leads[i0] - 1) * 100),
            fmt(vendas[i0], "num"), fmt(vendas[-1], "num"),
            fmt(c0, "pct2"), fmt(c1, "pct2"),
            "%+d%%" % round((c1 / c0 - 1) * 100)]


def pagina_resumo(prs):
    slide = nova_pagina(prs, TITULO_RESUMO, nota=NOTA_RESUMO)

    tot_l = [a + b + (c or 0) for a, b, c in
             zip(CANAIS[0]["leads"], CANAIS[2]["leads"], CANAIS[1]["leads"])]
    tot_v = [a + b + (c or 0) for a, b, c in
             zip(CANAIS[0]["ven"], CANAIS[2]["ven"], CANAIS[1]["ven"])]
    meta_l = sum(c["meta_leads"] for c in CANAIS)

    linhas = [linha_resumo(c["nome"], c["leads"], c["ven"], primeiro_mes(c))
              for c in CANAIS]
    linhas.append(linha_resumo("Consolidado", tot_l, tot_v, 0))
    linhas.append(["Meta mensal", fmt(meta_l, "num"), fmt(meta_l, "num"), "–",
                   fmt(meta_l * META_CONV / 100.0, "num"),
                   fmt(meta_l * META_CONV / 100.0, "num"),
                   fmt(META_CONV, "pct2"), fmt(META_CONV, "pct2"), "–"])

    larg = [2.65] + [1.26] * 8
    alt = [0.46] + [0.52] * len(linhas)
    tbl = nova_tabela(slide, MARGIN_L, 1.30, larg, alt)

    for c, titulo in enumerate(COLS_RESUMO):
        cell_text(tbl.cell(0, c), titulo, 10, GREY1, bold=True,
                  align=PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER,
                  anchor=MSO_ANCHOR.MIDDLE, mx=0.07)
        cell_border(tbl.cell(0, c), 'T', GREY1, 1.25)
        cell_border(tbl.cell(0, c), 'B', GREY1, 1.25)

    for ri, vals in enumerate(linhas, start=1):
        meta_row = vals[0] == "Meta mensal"
        cons_row = vals[0] == "Consolidado"
        fundo = METABG if meta_row else (BAND if cons_row else None)
        for c, v in enumerate(vals):
            cell = tbl.cell(ri, c)
            cell_box(cell, fill=fundo)
            cor, bold = GREY1, (cons_row or meta_row or c == 0)
            if meta_row:
                cor = RED
            elif c in (3, 8) and v not in ("–",):
                cor = FOREST if v.startswith("+") else RUBY
                bold = True
            elif c == 7 and not meta_row:
                cor = FOREST if float(v[:-1].replace(",", ".")) >= META_CONV else GREY1
                bold = True
            cell_text(cell, v, 11 if c == 0 else 11, cor, bold=bold,
                      align=PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER,
                      anchor=MSO_ANCHOR.MIDDLE, mx=0.07)
            cell_border(cell, 'T', GREY5, 0.5)
        if cons_row or meta_row:
            for c in range(9):
                cell_border(tbl.cell(ri, c), 'T', GREY1, 1.0)

    txt(slide, MARGIN_L, 4.62, 4.0, 0.26, "Três leituras", size=12, bold=True,
        color=GREY1, anchor=MSO_ANCHOR.TOP)
    col_w = (MARGIN_R - MARGIN_L - 2 * 0.34) / 3.0
    for i, (chapeu, corpo) in enumerate(LEITURAS):
        cx = MARGIN_L + i * (col_w + 0.34)
        hline(slide, cx, 5.04, cx + col_w, RED, 1.5)
        txt(slide, cx, 5.14, 0.26, 0.24, "%d" % (i + 1), size=13, bold=True, color=RED,
            anchor=MSO_ANCHOR.TOP)
        txt(slide, cx + 0.28, 5.14, col_w - 0.28, 1.70,
            [(chapeu, {"size": 10.5, "bold": True, "color": GREY1}), ("\n", {}),
             (corpo, {"size": 10, "color": GREY2})],
            anchor=MSO_ANCHOR.TOP, line=1.24)
    return slide


# ---------------------------------------------------------------- build
prs = Presentation()
prs.slide_width, prs.slide_height = Inches(SW), Inches(SH)
pagina_resumo(prs)
for canal in CANAIS:
    pagina_canal(prs, canal)

OUT = "/home/user/Gerais/output/funil_conversao_canais.pptx"
prs.save(OUT)
print("saved %s (%d paginas)" % (OUT, len(prs.slides)))
