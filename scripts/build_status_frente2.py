# -*- coding: utf-8 -*-
"""Pagina unica (16:9) do acompanhamento de setembro/26 da frente 2, parte 1/2:
qualificacao de leads (site, Meta, classificados) e o novo painel de
resultados-chave (qualidade dos leads e volume/conversao), reproduzindo o
print enviado.

Cabecalho, menu de frentes e a pastilha de status reaproveitam os tokens de
`deckstyle.py`, na mesma identidade visual do restante do deck. Tudo em
tabelas nativas do PowerPoint (o mini-painel de resultados e a grade de
status), editaveis celula a celula; nenhum dado tabular fica em caixa de
texto solta. Os selos de letra (A, B, C, D) ao lado do numero do entregavel
sao ovais desenhados por cima da tabela, como o selo numerado da frente em
build.py.

Conteudo desta pagina (snapshot de 25/set/26), abaixo. Para atualizar em um
proximo ciclo, edite so os dados no topo do arquivo.

    python3 build_status_frente2.py
"""
import json

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from deckstyle import (
    GREEN, GREEN_DEEP, BLACK, TEXT_DARK, TEXT_REG, TEXT_HELPER, NAV_IDLE,
    WHITE, GRAY_X_LIGHT, GRAY_LIGHT, GRAY_MEDIUM, STROKE, PLACEHOLDER,
    STATUS, FONT_NAME, CELL_MX,
    text_w, nlines, lh, split_bold,
    no_shadow, style_run, cell_box, cell_text, set_table_plain, set_row_heights,
)
from icons_bain import icon


def bain_icon(kind):
    """adapta icon(slide, kind, ...) de icons_bain.py para o icon_fn(slide,x,y,s,color)"""
    return lambda slide, x, y, s, color: icon(slide, kind, x, y, s, color)

OUT = '../output/260925__Status_frente2_resultados.pptx'

# ------------------------------------------------------------------ conteudo
FRENTE_ID = 2
TITULO = 'Tratamento, priorização e alocação de leads (1/2)'
OBJETIVO = ('Melhorar a experiência dos clientes em prospecção, o entendimento '
            'de suas necessidades, a priorização dos leads com maior potencial '
            'de conversão e a alocação às estruturas de vendas com maior '
            'chance de conversão.')
OBJ_BOLD = ['a experiência dos clientes em prospecção',
            'a priorização dos leads com maior potencial de conversão',
            'a alocação às estruturas de vendas com maior chance de conversão']
LIDERES = 'Priscila Martins'
TIME = 'Priscila Martins, Yure Cangussu, Thais Pimenta, Luma Gomes'

MESES = ['Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set*']

# icone (chave de icons_bain.py), cor, rotulo, legenda, valores mensais
AZUL_ICO = RGBColor(0x2F, 0x5B, 0xA8)
LARANJA_ICO = RGBColor(0xE0, 0x6C, 0x3C)
VERDE_ICO = RGBColor(0x2F, 0x8F, 0x5B)
QUALIDADE = {
    'titulo': 'Qualidade dos leads',
    'sub': '% de leads sem CEP (menor é melhor)',
    'linhas': [
        ('monitor',    BLACK,     'Site',    '% leads sem CEP', ['-', '-', '24%', '24%', '23%', '22%', '13%', '10%']),
        ('engrenagem', AZUL_ICO,  'Liza',    '% leads sem CEP', ['-', '-', '54%', '85%', '80%', '69%', '49%', '52%']),
        ('rede',       AZUL_ICO,  'Chatbot', '% leads sem CEP', ['-', '-', '57%', '58%', '56%', '54%', '41%', '24%']),
    ],
}
VOLUME = {
    'titulo': 'Volume e conversão',
    'sub': 'Leads classificados e taxa de conversão (Classif. → vendas)',
    'linhas': [
        ('pessoas', LARANJA_ICO, 'Leads Classif.', '', ['80', '79', '77', '93', '64', '68', '69', '14']),
        ('alvo',    VERDE_ICO,   'Conversão',      'Classif. → vendas',
         ['0,39%', '0,41%', '0,35%', '0,31%', '0,40%', '0,38%', '0,33%', '0,18%']),
    ],
}
NOTA = ('Nota: considera vendas que aconteceram em até 7 dias pós geração do lead; '
        '*Inclui somente leads gerados até dia 7 de Set; Fonte: base de vendas '
        'faturadas e base de funil dos leads')

# entregavel: id, selos (letras), descricao, trechos em negrito, responsavel,
# prazo, status (chave do semaforo), rotulo de status (None = padrao),
# progresso recente, proximos passos, pontos a escalar (listas = bullets)
STATUS_ROWS = [
    dict(id='2.1.1', selos=['A'],
         desc='Informações mínimas do cliente coletadas no site',
         bold=['Informações mínimas do cliente coletadas no site'],
         resp='Thais Pimenta', prazo='11/set', status='amarelo', rotulo=None,
         prog=['Telefone válido nas principais LPs',
               'Pesquisa da nova tela de loja de preferência lançada'],
         prox=['Criar formulário padrão para todas as LPs',
               'Publicar a 2ª página com a loja de preferência'],
         esc=['N/A']),
    dict(id='2.1.2', selos=['B', 'C'],
         desc='Informações mínimas do cliente coletadas nos canais Meta (Facebook e WhatsApp)',
         bold=['Informações mínimas do cliente coletadas nos canais Meta'],
         resp='João Azevedo + Isadora Barreto', prazo='11/set', status='amarelo', rotulo=None,
         prog=['Desenho do fluxo e início do desenvolv. de CEP obrigatório para '
               'geração de lead na Liza'],
         prox=['Finalizar desenvolv. de CEP obrigatório para geração de lead na Liza'],
         esc=['N/A']),
    dict(id='2.1.3', selos=['D'],
         desc='Informações mínimas do cliente coletadas via canais Classificados: '
              'CEP capturado para todos os leads',
         bold=['Classificados'],
         resp='Luma Gomes', prazo='25/set', status='amarelo', rotulo=None,
         prog=['Conversas para definição de prazos para obtenção de dados obrigatórios'],
         prox=['TBD'],
         esc=['Avaliar necessidade de acelerar a obtenção dos dados de CEP com os parceiros']),
    dict(id='2.1.4', selos=['D'],
         desc='Revisar modelo de clusters de anúncios',
         bold=[],
         resp='João Azevedo', prazo='25/set', status='amarelo', rotulo=None,
         prog=['Piloto de remoção de clusters de anúncios lançado na DVSUC'],
         prox=['Avaliar resultados dos leads classificados na DVSUC'],
         esc=['N/A']),
    dict(id='2.2.1', selos=[],
         desc='Outras informações já disponíveis do cliente (ex.: localização por '
              'IP, cidade, etc) integradas e disponíveis para todos os canais',
         bold=['integradas e disponíveis'],
         resp='TBD', prazo='25/set', status='vermelho', rotulo='Não iniciado',
         prog=['N/A'],
         prox=['Mapear dados disponíveis para uso',
               'Integrar novos dados às bases de clientes'],
         esc=['Alinhar prioridade e definir responsável e prazo']),
    dict(id='2.3.1', selos=[],
         desc='Lead score evoluído com novas variáveis (score B3 de propensão a '
              'financiamento, renda presumida, etc.)',
         bold=[],
         resp='João Azevedo', prazo='23/out', status='vermelho',
         rotulo='Não priorizado no curto prazo',
         prog=['N/A'], prox=['N/A'], esc=['N/A']),
]

# ------------------------------------------------------------------ heatmap
# escala de 3 cores classica (Excel), interpolada: verde (minimo da linha) ->
# amarelo (meio) -> vermelho (maximo da linha), igual a formatacao condicional
# de "escala de cores" do Excel/Sheets aplicada linha a linha
HEAT_GREEN = (RGBColor(0xC6, 0xEF, 0xCE), RGBColor(0x00, 0x61, 0x00))
HEAT_YELLOW = (RGBColor(0xFF, 0xEB, 0x9C), RGBColor(0x9C, 0x65, 0x00))
HEAT_RED = (RGBColor(0xFF, 0xC7, 0xCE), RGBColor(0x9C, 0x00, 0x06))


def _lerp_color(a, b, t):
    return RGBColor(*(int(a[i] + (b[i] - a[i]) * t) for i in range(3)))


def _heat_at(t):
    """t=0 -> verde, t=0.5 -> amarelo, t=1 -> vermelho"""
    if t <= 0.5:
        (bg0, fg0), (bg1, fg1), u = HEAT_GREEN, HEAT_YELLOW, t / 0.5
    else:
        (bg0, fg0), (bg1, fg1), u = HEAT_YELLOW, HEAT_RED, (t - 0.5) / 0.5
    return _lerp_color(bg0, bg1, u), _lerp_color(fg0, fg1, u)


HEAT_SPAN = 45.0   # pontos percentuais do minimo da linha at vermelho pleno


def heat_pct(vals):
    """verde no minimo da propria linha, vermelho a HEAT_SPAN p.p. dali.

    Ancorado no minimo de cada linha (nao no maximo) e com um passo fixo em
    pontos percentuais, para linhas com pouca variacao (Site, 10-24%) ficarem
    em tons proximos de verde/amarelo e linhas mais criticas (Liza, 49-85%)
    alcancarem o vermelho — em vez de esticar sempre verde->vermelho entre o
    minimo e o maximo de cada linha, o que faria o pior mes de qualquer linha
    parecer igualmente grave.
    """
    nums = [(i, float(v.replace('%', '').replace(',', '.')))
            for i, v in enumerate(vals) if v != '-']
    out = [None] * len(vals)
    if not nums:
        return out
    mn = min(v for _, v in nums)
    for i, v in nums:
        t = max(0.0, min(1.0, (v - mn) / HEAT_SPAN))
        out[i] = _heat_at(t)
    return out


def heat_last_only(vals):
    """sem heatmap mes a mes; so marca a coluna Set* (mes parcial) em vermelho"""
    return [HEAT_RED if i == len(vals) - 1 else None for i in range(len(vals))]


# ------------------------------------------------------------------ geometria
SW, SH = 13.3333, 7.5
ML = 0.42
CW = SW - 2 * ML
TOP = 0.26
BOTTOM = SH - 0.10

PT_TITLE, PT_OBJ = 14.0, 9.5
PT_LABEL, PT_VALUE = 7.5, 9.5
PT_NAV = 9.0
PT_PANEL, PT_SECT, PT_SUBSECT = 10.5, 9.5, 7.5
PT_HEAD, PT_BODY, PT_FOOT = 7.5, 8.0, 7.0
BODY_STEPS_STATUS = (8.0, 7.5, 7.0, 6.5, 6.0)

PANEL_GREEN = RGBColor(0x1F, 0x4E, 0x3D)
PANEL_GREEN_LT = RGBColor(0x9F, 0xD4, 0xB8)
CHIP_BG = RGBColor(0xE9, 0xEC, 0xF4)
CHIP_SET_BG = RGBColor(0xDD, 0xE1, 0xEC)
BAR = 0.075


# ------------------------------------------------------------------ helpers
def textbox(slide, x, y, w, h, text, size=9, color=TEXT_DARK, bold=False,
            italic=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, wrap=True):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    style_run(p.add_run(), text, size, color, bold, italic)
    return tb


def rect(slide, x, y, w, h, fill=None, line=None, lw=0.75,
         shape=MSO_SHAPE.RECTANGLE, adj=None):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line; s.line.width = Pt(lw)
    no_shadow(s)
    if adj is not None:
        try: s.adjustments[0] = adj
        except (IndexError, ValueError): pass
    tf = s.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return s


def funnel_icon(slide, x, y, s, color):
    """funil (filtro): trapezio + haste, para o titulo de "Qualidade dos leads" """
    top_w, bot_w, body_h = s, s * 0.30, s * 0.55
    pts = [(x, y), (x + top_w, y),
           (x + top_w / 2 + bot_w / 2, y + body_h),
           (x + top_w / 2 - bot_w / 2, y + body_h)]
    fb = slide.shapes.build_freeform(Inches(pts[0][0]), Inches(pts[0][1]))
    fb.add_line_segments([(Inches(px), Inches(py)) for px, py in pts[1:]], close=True)
    sh = fb.convert_to_shape()
    sh.fill.solid(); sh.fill.fore_color.rgb = color
    sh.line.fill.background()
    no_shadow(sh)
    rect(slide, x + top_w / 2 - bot_w * 0.22, y + body_h, bot_w * 0.44, s * 0.28, fill=color)


def list_icon(slide, x, y, s, color):
    """3 marcadores + 3 linhas, para o titulo de "Status das frentes em andamento" """
    dr = s * 0.09
    for frac in (0.12, 0.44, 0.76):
        cy = y + s * frac
        rect(slide, x, cy, dr * 2, dr * 2, fill=color, shape=MSO_SHAPE.OVAL)
        rect(slide, x + dr * 3, cy + dr * 0.3, s * 0.62, dr * 1.4, fill=color)


def new_table(slide, x, y, w, nrows, widths, heights):
    gf = slide.shapes.add_table(nrows, len(widths), Inches(x), Inches(y),
                                Inches(w), Inches(sum(heights)))
    tbl = gf.table
    set_table_plain(tbl)
    set_row_heights(tbl, heights)
    for i, cw in enumerate(widths):
        tbl.columns[i].width = Inches(cw)
    return tbl


def cell_paras(cell, paras, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
               mx=CELL_MX, mt=0.04, mb=0.04, space=1.5):
    """celula com varios paragrafos; paras = [(size, [(texto, negrito, cor)])]"""
    cell.margin_left = Inches(mx); cell.margin_right = Inches(mx)
    cell.margin_top = Inches(mt); cell.margin_bottom = Inches(mb)
    cell.vertical_anchor = anchor
    tf = cell.text_frame
    tf.word_wrap = True
    tf.clear()
    for i, (size, parts) in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if i:
            p.space_before = Pt(space)
        for seg, b, color in parts:
            style_run(p.add_run(), seg, size, color, b)
        p.font.size = Pt(size)
        p.font.name = FONT_NAME
    return tf


def bullet_list(cell, items, size, color, bold_map=None):
    vazio = len(items) == 1 and items[0] in ('N/A', 'TBD')
    if vazio:
        cell_paras(cell, [(size, [(items[0], False, PLACEHOLDER)])],
                   anchor=MSO_ANCHOR.MIDDLE)
        return
    multi = len(items) > 1
    paras = []
    for it in items:
        parts = split_bold(it, (bold_map or {}).get(it, []))
        parts = [(t, b, color) for t, b in parts]
        if multi:
            parts = [(chr(0x2022) + ' ', False, color)] + parts
        paras.append((size, parts))
    cell_paras(cell, paras, anchor=MSO_ANCHOR.MIDDLE, space=2.0)


def panel_header(slide, x, y, w, title, right_text=None, icon_fn=None):
    """faixa verde-escura do painel, como tabela 1x2 (editavel, nao textbox)"""
    h = 0.30
    tbl = new_table(slide, x, y, w, 1, [w * 0.55, w * 0.45] if right_text else [w], [h])
    c = tbl.cell(0, 0)
    cell_box(c, PANEL_GREEN)
    mx = 0.40 if icon_fn else 0.12
    cell_text(c, [(title, True)], PT_PANEL, WHITE, anchor=MSO_ANCHOR.MIDDLE, mx=mx)
    if icon_fn:
        icon_s = 0.20
        icon_fn(slide, x + 0.12, y + (h - icon_s) / 2.0, icon_s, PANEL_GREEN_LT)
    if right_text:
        c2 = tbl.cell(0, 1)
        cell_box(c2, PANEL_GREEN)
        cell_text(c2, [(right_text, False)], PT_SUBSECT, PANEL_GREEN_LT,
                  align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE, mx=0.12)
    return y + h


def section_header(slide, x, y, w, icon_fn, icon_color, titulo, sub):
    """chip claro com icone, titulo e subtitulo, como no print (tabela 1x1
    para o fundo arredondado, com o icone e os dois textos por cima)"""
    tit_n = nlines(titulo, PT_SECT, w - 0.55, True)
    sub_n = nlines(sub, PT_SUBSECT, w - 0.55)
    h = tit_n * lh(PT_SECT) + sub_n * lh(PT_SUBSECT) + 0.16
    rect(slide, x, y, w, h, fill=CHIP_BG, shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.12)
    icon_s = 0.24
    icon_fn(slide, x + 0.16, y + (h - icon_s) / 2.0, icon_s, icon_color)
    tx = x + 0.16 + icon_s + 0.14
    textbox(slide, tx, y + 0.08, w - (tx - x) - 0.12, tit_n * lh(PT_SECT),
            titulo, size=PT_SECT, color=BLACK, bold=True)
    textbox(slide, tx, y + 0.08 + tit_n * lh(PT_SECT), w - (tx - x) - 0.12,
            sub_n * lh(PT_SUBSECT), sub, size=PT_SUBSECT, color=TEXT_HELPER)
    return y + h + 0.10


def status_pill(cell, key, label, size):
    lbl_default, bg, fg = STATUS[key]
    cell_box(cell, bg, (WHITE, WHITE, WHITE, WHITE), lw=2.5)
    cell_text(cell, [(label or lbl_default, True)], size, fg,
              align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, mx=0.04)


def badge(slide, x, y, d, letter):
    s = rect(slide, x, y, d, d, fill=GREEN, shape=MSO_SHAPE.OVAL)
    tf = s.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    style_run(p.add_run(), letter, 6.5, BLACK, True)
    return s


# ------------------------------------------------------------------ cabecalho
def header(slide, frentes):
    y = TOP
    w_lid, w_time = 1.55, 2.55
    w_main = CW - BAR - w_lid - w_time
    BADGE, GAP = 0.28, 0.14
    MAIN_ML = BADGE + GAP + 0.18

    txt_w = w_main - MAIN_ML - CELL_MX
    tit_n = nlines(TITULO, PT_TITLE, txt_w, True)
    obj_n = nlines(OBJETIVO, PT_OBJ, txt_w)
    h_main = tit_n * lh(PT_TITLE) + 0.05 + obj_n * lh(PT_OBJ)
    side_n = max(nlines(LIDERES, PT_VALUE, w_lid - 2 * CELL_MX),
                 nlines(TIME, PT_VALUE, w_time - 2 * CELL_MX))
    h_side = lh(PT_LABEL) + 0.05 + side_n * lh(PT_VALUE)
    hdr_h = max(h_main, h_side) + 0.20

    hdr = new_table(slide, ML, y, CW, 1, [BAR, w_main, w_lid, w_time], [hdr_h])
    cell_box(hdr.cell(0, 0), GREEN)
    for c in (1, 2, 3):
        cell_box(hdr.cell(0, c), GRAY_X_LIGHT)

    c = hdr.cell(0, 1)
    tf = cell_text(c, [(TITULO, True)], PT_TITLE, BLACK, mt=0.10, mb=0.10)
    c.margin_left = Inches(MAIN_ML)
    p2 = tf.add_paragraph()
    p2.space_before = Pt(3)
    for seg, b in split_bold(OBJETIVO, OBJ_BOLD):
        style_run(p2.add_run(), seg, PT_OBJ, TEXT_REG, b)

    for col, lab, val in ((2, 'Líder(es)', LIDERES), (3, 'Time de trabalho', TIME)):
        c = hdr.cell(0, col)
        tf = cell_text(c, [(lab.upper(), True)], PT_LABEL, TEXT_HELPER, mt=0.10, mb=0.10)
        p2 = tf.add_paragraph()
        p2.space_before = Pt(2)
        style_run(p2.add_run(), val, PT_VALUE, TEXT_DARK)

    bx = ML + BAR + 0.16
    by = TOP + 0.10 + (lh(PT_TITLE) - BADGE) / 2.0
    rect(slide, bx, by, BADGE, BADGE, fill=GREEN, shape=MSO_SHAPE.OVAL)
    textbox(slide, bx, by + 0.045, BADGE, 0.20, str(FRENTE_ID), size=11,
            color=BLACK, bold=True, align=PP_ALIGN.CENTER)

    y += hdr_h + 0.11

    # ---- menu de frentes (tabs) ----
    labels = ['%d. %s' % (g['id'], g['nome']) for g in frentes]
    nav_pt = PT_NAV
    while nav_pt > 6.5:
        ws = [text_w(t, nav_pt, True) for t in labels]
        gap = (CW - sum(ws)) / (len(labels) - 1)
        if gap >= 0.14:
            break
        nav_pt -= 0.5
    nav_h = lh(nav_pt) + 0.06
    xx = ML
    for g, wgt in zip(frentes, ws):
        active = g['id'] == FRENTE_ID
        tb = textbox(slide, xx, y, wgt + 0.04, nav_h, '', size=nav_pt)
        style_run(tb.text_frame.paragraphs[0].add_run(), '%d. %s' % (g['id'], g['nome']),
                   nav_pt, GREEN_DEEP if active else NAV_IDLE, bold=active)
        if active:
            rect(slide, xx, y + nav_h - 0.01, wgt, 0.026, fill=GREEN)
        xx += wgt + gap
    y += nav_h + 0.06
    ln = slide.shapes.add_connector(1, Inches(ML), Inches(y), Inches(ML + CW), Inches(y))
    ln.line.color.rgb = STROKE
    ln.line.width = Pt(0.75)
    no_shadow(ln)
    return y + 0.14


# ------------------------------------------------------------------ painel esquerdo
def mini_table(slide, x, y, w, bloco, heat_fn):
    def fh(pt):
        return pt * 1.30 / 72.0

    n_meses = len(MESES)
    lab_w = w * 0.30
    col_w = (w - lab_w) / n_meses
    widths = [lab_w] + [col_w] * n_meses
    head_h = fh(PT_HEAD) + 0.14
    tem_legenda = any(l[2] for l in bloco['linhas'])
    row_h = fh(PT_BODY) + (fh(PT_HEAD - 0.5) + 0.04 if tem_legenda else 0) + 0.22
    heights = [head_h] + [row_h] * len(bloco['linhas'])

    tbl = new_table(slide, x, y, w, len(heights), widths, heights)
    cell_box(tbl.cell(0, 0), GRAY_X_LIGHT, (None, None, None, GRAY_MEDIUM), lw=0.5)
    for ci, mes in enumerate(MESES, start=1):
        c = tbl.cell(0, ci)
        cell_box(c, GRAY_X_LIGHT if ci < n_meses else CHIP_SET_BG,
                  (None, None, None, GRAY_MEDIUM), lw=0.5)
        cell_text(c, [(mes, True)], PT_HEAD, TEXT_HELPER, align=PP_ALIGN.CENTER,
                  anchor=MSO_ANCHOR.MIDDLE)

    icon_s = 0.13
    row_y = y + head_h
    for ri, (kind, cor, nome, legenda, vals) in enumerate(bloco['linhas'], start=1):
        sep = GRAY_MEDIUM if ri == len(bloco['linhas']) else GRAY_LIGHT
        c = tbl.cell(ri, 0)
        cell_box(c, WHITE, (None, None, None, sep), lw=0.5)
        paras = [(PT_BODY, [(nome, True, TEXT_DARK)])]
        if legenda:
            paras.append((PT_HEAD - 0.5, [(legenda, False, TEXT_HELPER)]))
        cell_paras(c, paras, anchor=MSO_ANCHOR.MIDDLE, mx=0.24)
        icon_y = row_y + (row_h - icon_s) / 2.0
        icon(slide, kind, x + 0.06, icon_y, icon_s, cor, bg=WHITE)

        cores = heat_fn(vals)
        for ci, (val, heat) in enumerate(zip(vals, cores), start=1):
            cc = tbl.cell(ri, ci)
            if heat:
                bg, fg = heat
                cell_box(cc, bg, (WHITE, WHITE, None, sep), lw=1.0)
            else:
                cell_box(cc, WHITE, (None, None, None, sep), lw=0.5)
                fg = PLACEHOLDER if val == '-' else TEXT_DARK
            cell_text(cc, [(val, heat is not None)], PT_BODY, fg,
                      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        row_y += row_h
    return y + sum(heights)


def left_panel(slide, x, y, w, bottom):
    y = panel_header(slide, x, y, w, 'Resultados-chave', 'Histórico mensal | 2026',
                      icon_fn=bain_icon('barras'))
    y += 0.10

    y = section_header(slide, x, y, w, funnel_icon, PANEL_GREEN,
                        QUALIDADE['titulo'], QUALIDADE['sub'])
    y = mini_table(slide, x, y, w, QUALIDADE, heat_pct)
    y += 0.16

    y = section_header(slide, x, y, w, bain_icon('barras'), PANEL_GREEN,
                        VOLUME['titulo'], VOLUME['sub'])
    y = mini_table(slide, x, y, w, VOLUME, heat_last_only)
    y += 0.12

    nota_n = nlines(NOTA, PT_FOOT, w)
    textbox(slide, x, bottom - nota_n * lh(PT_FOOT), w, nota_n * lh(PT_FOOT),
            NOTA, size=PT_FOOT, color=TEXT_HELPER)


# ------------------------------------------------------------------ painel direito
R_COLS = [
    ('#', 0.55, PP_ALIGN.LEFT),
    ('Entregável', 1.70, PP_ALIGN.LEFT),
    ('Responsável', 1.30, PP_ALIGN.LEFT),
    ('Prazo', 0.70, PP_ALIGN.LEFT),
    ('Status', 1.00, PP_ALIGN.CENTER),
    ('Progresso recente', 1.60, PP_ALIGN.LEFT),
    ('Próximos passos', 1.60, PP_ALIGN.LEFT),
    ('Pontos a escalar', 1.60, PP_ALIGN.LEFT),
]
PT_RHEAD = 6.5


def right_panel(slide, x, y, w, bottom):
    y = panel_header(slide, x, y, w, 'Status das frentes em andamento', icon_fn=list_icon)
    y += 0.10

    scale = w / sum(c[1] for c in R_COLS)
    widths = [c[1] * scale for c in R_COLS]

    head_h = 0.28
    avail = bottom - y
    BULLET_W = 0.13   # largura do marcador "- " descontada da 1a linha de cada item

    def fh(pt):
        return pt * 1.28 / 72.0

    def bullet_n(items, pt, width):
        w = width - (BULLET_W if len(items) > 1 else 0)
        return sum(nlines(it, pt, w) for it in items)

    for body_pt in BODY_STEPS_STATUS:
        heights = [head_h]
        for r in STATUS_ROWS:
            n_desc = nlines(r['desc'], body_pt, widths[1] - 2 * CELL_MX, True)
            n_prog = bullet_n(r['prog'], body_pt, widths[5] - 2 * CELL_MX)
            n_prox = bullet_n(r['prox'], body_pt, widths[6] - 2 * CELL_MX)
            n_esc = bullet_n(r['esc'], body_pt, widths[7] - 2 * CELL_MX)
            n = max(n_desc, n_prog, n_prox, n_esc, 2)
            heights.append(max(0.40, n * fh(body_pt) + 0.16))
        if sum(heights) <= avail:
            break

    if sum(heights) > avail:
        k = avail / sum(heights)
        heights = [max(h * k, 0.28) for h in heights]
    else:
        # distribui a folga entre as linhas de corpo (nao no cabecalho)
        slack = avail - sum(heights)
        if slack > 0.01:
            add = slack / (len(heights) - 1)
            heights = [heights[0]] + [h + add for h in heights[1:]]

    tbl = new_table(slide, x, y, w, len(heights), widths, heights)
    for ci, (name, _w, align) in enumerate(R_COLS):
        c = tbl.cell(0, ci)
        cell_box(c, GRAY_X_LIGHT, (None, None, None, BLACK), lw=1.0)
        cell_text(c, [(name.upper(), True)], PT_RHEAD, TEXT_HELPER, align=align,
                  anchor=MSO_ANCHOR.MIDDLE)

    row_y = [y + heights[0]]
    for h in heights[1:]:
        row_y.append(row_y[-1] + h)

    for ri, r in enumerate(STATUS_ROWS, start=1):
        sep = GRAY_MEDIUM if ri == len(STATUS_ROWS) else GRAY_LIGHT
        base = (None, None, None, sep)

        c = tbl.cell(ri, 0)
        cell_box(c, None, base, lw=0.5)
        cell_text(c, [(r['id'], True)], body_pt - 0.5, TEXT_HELPER,
                  anchor=MSO_ANCHOR.TOP, mt=0.05)
        c.text_frame.word_wrap = False

        c = tbl.cell(ri, 1)
        cell_box(c, None, base, lw=0.5)
        parts = [(t, b, BLACK) for t, b in split_bold(r['desc'], r['bold'])]
        cell_paras(c, [(body_pt, parts)], anchor=MSO_ANCHOR.MIDDLE)

        c = tbl.cell(ri, 2)
        cell_box(c, None, base, lw=0.5)
        cell_text(c, [(r['resp'], False)], body_pt, TEXT_DARK, anchor=MSO_ANCHOR.MIDDLE)

        c = tbl.cell(ri, 3)
        cell_box(c, None, base, lw=0.5)
        cell_text(c, [(r['prazo'], False)], body_pt, TEXT_DARK, anchor=MSO_ANCHOR.MIDDLE)

        status_pill(tbl.cell(ri, 4), r['status'], r['rotulo'], body_pt - 0.5)

        for ci, key in ((5, 'prog'), (6, 'prox'), (7, 'esc')):
            c = tbl.cell(ri, ci)
            cell_box(c, None, base, lw=0.5)
            bullet_list(c, r[key], body_pt, TEXT_DARK)

        # selos de letra, centrados sob o numero do entregavel
        selos = r['selos']
        if selos:
            d, gap = 0.16, 0.04
            total = len(selos) * d + (len(selos) - 1) * gap
            bx0 = x + (widths[0] - total) / 2.0
            by = row_y[ri - 1] + heights[ri] - d - 0.06
            for si, letra in enumerate(selos):
                badge(slide, bx0 + si * (d + gap), by, d, letra)

    return y + sum(heights)


# ------------------------------------------------------------------ main
def build(prs, frentes):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SW, SH, fill=WHITE)

    y = header(slide, frentes)

    gap = 0.18
    left_w = CW * 0.365
    right_w = CW - left_w - gap
    right_x = ML + left_w + gap

    left_panel(slide, ML, y, left_w, BOTTOM)
    right_panel(slide, right_x, y, right_w, BOTTOM)
    return slide


def main():
    frentes = json.load(open('frentes.json'))
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(SW), Inches(SH)
    build(prs, frentes)
    prs.core_properties.title = 'Tratamento, priorização e alocação de leads (1/2)'
    prs.save(OUT)
    print('ok:', OUT)


if __name__ == '__main__':
    main()
