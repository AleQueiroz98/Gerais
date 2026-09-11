# -*- coding: utf-8 -*-
"""Pagina 'Funil lead -> venda por canal' (CV, Liza e Central, jan-ago/26).

Uma unica pagina 16:9 com os tres paineis lado a lado, como no rascunho, numa
tabela nativa do PowerPoint: cada numero fica na propria celula, entao da para
selecionar o bloco e colar direto no Excel. As taxas sao recalculadas a partir
dos volumes brutos -- nada e transcrito de valores ja arredondados -- e a linha
de conversao final recebe heatmap contra a meta de 0,94%.

Os rotulos e a ordem das linhas seguem a planilha de origem, para o bloco colar
alinhado com ela.

As larguras das colunas saem do numero mais largo de cada painel a 8pt (o CV
precisa de mais espaco por causa dos leads na casa dos 454 mil); o conjunto
fecha exatamente os 12,73" uteis da pagina.
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
MINT   = RGBColor(0x7F, 0xD1, 0xAE)   # positivo sobre a barra escura
ROSE   = RGBColor(0xFF, 0x8A, 0x8A)   # negativo sobre a barra escura
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
FONT   = "Arial"
LANG   = "pt-BR"

# ---------------------------------------------------------------- grid (in)
SW, SH = 13.333, 7.5
MARGIN_L, MARGIN_R = 0.30, 13.03
TITLE_Y, TITLE_H = 0.17, 0.62
RULE_Y = 0.90
TBL_Y = 1.00
READ_Y, READ_H = 6.36, 0.52
NOTE_Y = 6.92

MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago"]
META_CONV = 0.94                      # meta de conversao lead -> venda (%)

# a tabela e uma so: rotulo + 3 paineis de 9 colunas (8 meses + meta),
# separados por uma coluna-espacador estreita
PT_DADO = 8.0                         # menor fonte da pagina
LBL_W, GAP_W = 1.52, 0.07
COL_W = {"CV (sem Liza e sem Central)": 0.4515, "Liza": 0.3895, "Central": 0.3895}
PANEL_H, MON_H, ROW_H = 0.38, 0.26, 0.42

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
        leitura=("Queda vem do topo, não do funil: os leads caem 29% (454k→321k) e as taxas "
                 "internas de aprovação e faturamento seguem no patamar de janeiro"),
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
        leitura=("Trade-off volume × qualidade: operou acima da meta até jul/26 e caiu a 0,67% "
                 "ao saltar de 31k para 58k leads em ago/26"),
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
        leitura=("Ganho de qualidade, não de escala: com 38% menos leads, o envio de fichas "
                 "sobe de 0,9% para 4,4% e a aprovação dobra"),
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

TITULO = ("Conversão lead→venda cai ~10% no consolidado (0,75%→0,68%) e segue 0,26 p.p. abaixo "
          "da meta de 0,94%: só a Central avança, Liza perde conversão ao escalar")

NOTA = ("Nota: CV exclui Liza e Central. Taxas recalculadas sobre os volumes brutos; "
        "% de fichas enviadas, % total de fichas faturadas e % de vendas são sobre o total de leads "
        "do mês, % fichas aprovadas sobre as enviadas e % fichas faturadas sobre as aprovadas. "
        "Meta de conversão de 0,94% aplicada sobre a meta mensal de leads de cada canal "
        "(CV 400k, Liza 50k, Central 20k). % leads sem info de CEP não disponível para jan-mar/26; "
        "Liza inicia em abr/26. Fonte: base de leads, base de fichas e base de vendas faturadas")

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


def hcell(tbl, r, c):
    return tbl.cell(r, c)


# ---------------------------------------------------------------- montagem
def colunas():
    """larguras e indice da primeira coluna de dados de cada painel"""
    largs, inicio = [LBL_W], []
    for i, canal in enumerate(CANAIS):
        if i:
            largs.append(GAP_W)
        inicio.append(len(largs))
        largs += [COL_W[canal["nome"]]] * 9
    return largs, inicio


def monta_tabela(slide):
    largs, inicio = colunas()
    alturas = [PANEL_H, MON_H] + [ROW_H] * len(LINHAS)
    tbl = slide.shapes.add_table(len(alturas), len(largs), Inches(MARGIN_L),
                                 Inches(TBL_Y), Inches(sum(largs)),
                                 Inches(sum(alturas))).table
    set_table_plain(tbl)
    for i, w in enumerate(largs):
        tbl.columns[i].width = Inches(w)
    set_row_heights(tbl, alturas)
    for r in range(len(alturas)):
        for c in range(len(largs)):
            cell_box(tbl.cell(r, c), fill=None)
    return tbl, inicio


def cabecalho(tbl, inicio):
    """linha 0: barra preta com o nome do canal e a variacao da conversao;
    linha 1: meses e coluna de meta"""
    cell_text(tbl.cell(0, 0), "Funil lead → venda", 10, GREY1, bold=True,
              anchor=MSO_ANCHOR.BOTTOM, mx=0.04, mb=0.04)
    cell_text(tbl.cell(1, 0), "2026", 8.5, GREY3, bold=True,
              anchor=MSO_ANCHOR.MIDDLE, mx=0.04)
    cell_border(tbl.cell(1, 0), 'B', GREY1, 1.25)

    for canal, c0 in zip(CANAIS, inicio):
        d = delta_conv(canal)
        alvo = tbl.cell(0, c0)
        alvo.merge(tbl.cell(0, c0 + 8))
        cell_box(alvo, fill=GREY1)
        cell_text(alvo, [(canal["nome"] + "     ", False)], 10, WHITE, bold=True,
                  anchor=MSO_ANCHOR.MIDDLE, mx=0.08)
        p = alvo.text_frame.paragraphs[0]
        for trecho, cor in ((" conversão ", GREY4), ("%+d%%" % round(d), None)):
            r = p.add_run()
            r.text = trecho
            r.font.name, r.font.size = FONT, Pt(8 if cor else 10)
            r.font.bold = cor is None
            r.font.color.rgb = cor or (MINT if d > 0 else ROSE)
            r.font._rPr.set('lang', LANG)

        for i, mes in enumerate(MESES):
            cell_text(tbl.cell(1, c0 + i), mes, 8.5, GREY1, bold=True,
                      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, mx=0.01)
        meta = tbl.cell(1, c0 + 8)
        cell_box(meta, fill=METABG)
        cell_text(meta, "Meta", 8.5, RED, bold=True,
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, mx=0.01)
        for i in range(9):
            cell_border(tbl.cell(1, c0 + i), 'B', GREY1, 1.25)


def corpo(tbl, inicio):
    for ri, (rotulo, chave, kind) in enumerate(LINHAS, start=2):
        # taxas medidas sobre o total de leads ganham fundo, para separar do
        # resto do funil, que usa a etapa anterior como base
        sobre_leads = chave in ("fat_leads", "ven_leads")
        forte = chave == "ven_leads"
        fundo = BAND if sobre_leads and not forte else None
        destaque = rotulo.startswith("#") or sobre_leads

        cell_text(tbl.cell(ri, 0), rotulo, PT_DADO, GREY1 if destaque else GREY2,
                  bold=destaque, anchor=MSO_ANCHOR.MIDDLE, mx=0.04)
        cell_box(tbl.cell(ri, 0), fill=fundo)
        cell_border(tbl.cell(ri, 0), 'T', GREY5, 0.5)

        for canal, c0 in zip(CANAIS, inicio):
            vals = serie(canal, chave)
            for i, v in enumerate(vals):
                cell = tbl.cell(ri, c0 + i)
                cor, bold = (GREY1, True) if destaque else (GREY2, False)
                if forte:
                    fill_heat, cor, bold = heat(v)
                    cell_box(cell, fill=fill_heat)
                else:
                    cell_box(cell, fill=fundo)
                cell_text(cell, fmt(v, kind), PT_DADO, GREY4 if v is None else cor,
                          bold=bold, align=PP_ALIGN.CENTER,
                          anchor=MSO_ANCHOR.MIDDLE, mx=0.01)

            meta = tbl.cell(ri, c0 + 8)
            cell_box(meta, fill=METABG)
            cell_text(meta, meta_cell(canal, chave) or "", PT_DADO, RED, bold=True,
                      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, mx=0.01)

            for i in range(9):
                cell_border(tbl.cell(ri, c0 + i), 'T', GREY5, 0.5)
                if forte:
                    cell_border(tbl.cell(ri, c0 + i), 'T', RED, 1.25)
                    cell_border(tbl.cell(ri, c0 + i), 'B', RED, 1.25)


def build():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(SW), Inches(SH)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    txt(slide, MARGIN_L, TITLE_Y, 10.85, TITLE_H, TITULO, size=18, color=GREY1,
        anchor=MSO_ANCHOR.TOP, line=1.06)
    txt(slide, 11.25, TITLE_Y + 0.36, 1.78, 0.22, "/ P R E L I M I N A R", size=9,
        color=GREY3, align=PP_ALIGN.RIGHT)
    hline(slide, MARGIN_L, RULE_Y, MARGIN_R, RED, 2.25)

    tbl, inicio = monta_tabela(slide)
    cabecalho(tbl, inicio)
    corpo(tbl, inicio)

    # leitura de uma linha por painel, alinhada com as colunas do painel
    largs, _ = colunas()
    for canal, c0 in zip(CANAIS, inicio):
        x = MARGIN_L + sum(largs[:c0])
        txt(slide, x, READ_Y, sum(largs[c0:c0 + 9]), READ_H, canal["leitura"],
            size=8.5, color=GREY2, anchor=MSO_ANCHOR.TOP, line=1.16)

    txt(slide, MARGIN_L, NOTE_Y, MARGIN_R - MARGIN_L, 0.48, NOTA, size=8,
        color=GREY3, anchor=MSO_ANCHOR.TOP, line=1.18)

    out = "/home/user/Gerais/output/funil_conversao_canais.pptx"
    prs.save(out)
    print("saved %s (%d pagina)" % (out, len(prs.slides)))


build()
