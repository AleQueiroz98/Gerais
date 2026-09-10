# -*- coding: utf-8 -*-
"""Recap da agenda de valor Lead to Sales: conversao por informacao coletada.

Quatro paginas 16:9 no padrao Bain Core:

  1. Recap completo   - marimekko de 10 grupos (largura ~ volume, altura ~
                        conversao) com uma faixa nova de 'nivel de informacao
                        do lead' que agrega os grupos em tres degraus e mostra
                        a conversao media ponderada de cada um.
  2. Build alavanca 1 - pagina esmaecida, com a seta de migracao (apontando
                        para os grupos com mais informacao, nao ao contrario)
                        e o callout dos degraus.
  3. Build alavanca 2 - pagina esmaecida, com as linhas de alocacao vivas e o
                        gap intra-grupo anotado no grafico.
  4. Alavanca 1       - pagina dedicada: escada de tres degraus, volume a
                        migrar, cenarios de impacto em p.p. e os tres pontos
                        de captura da informacao (milestones 2.1.1 a 2.1.3).

Dados lidos do slide de origem; medias por grupo ponderadas pelo volume das
barras exibidas.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

# ------------------------------------------------------------------ tokens
RED    = RGBColor(0xCC, 0x00, 0x00)   # Bain red
RUBY   = RGBColor(0x99, 0x00, 0x00)   # negativo / gap
RED_BG = RGBColor(0xFB, 0xEA, 0xEA)
FOREST = RGBColor(0x10, 0x4C, 0x3E)   # Forest 1  - alocacao ideal
SAGE   = RGBColor(0x9C, 0xBF, 0xB3)   # Forest 4  - ideal com dados disponiveis
SAGE_D = RGBColor(0x3F, 0x74, 0x64)
GOLD   = RGBColor(0xC6, 0xAA, 0x3D)   # Sunset 2  - grande melhoria potencial
GOLD_D = RGBColor(0xAB, 0x89, 0x33)
BLACK  = RGBColor(0x00, 0x00, 0x00)
GREY1  = RGBColor(0x33, 0x33, 0x33)
GREY2  = RGBColor(0x5C, 0x5C, 0x5C)
GREY3  = RGBColor(0x85, 0x85, 0x85)
GREY4  = RGBColor(0xB4, 0xB4, 0xB4)
GREY5  = RGBColor(0xD9, 0xD9, 0xD9)
GREY6  = RGBColor(0xF2, 0xF2, 0xF2)
DIMTXT = RGBColor(0xC4, 0xC4, 0xC4)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
FONT   = 'Arial'
LANG   = 'pt-BR'

SW, SH = 13.333, 7.5


def mix(c, f):
    """clareia a cor em direcao ao branco (f=0 mantem, f=1 vira branco)"""
    return RGBColor(*(int(v + (255 - v) * f) for v in (c[0], c[1], c[2])))


# ------------------------------------------------------------------ helpers
def _lang(font):
    rPr = font._rPr
    rPr.set('lang', LANG)
    rPr.set('dirty', '0')


def _spc(font, hundredths):
    font._rPr.set('spc', str(int(hundredths)))


def txt(sl, x, y, w, h, lines, size=9, color=GREY1, bold=False,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spc=None, lsp=1.16):
    """lines: str | [(texto, opts)] | [[(texto, opts)], ...] (paragrafos)"""
    if isinstance(lines, str):
        lines = [[(lines, {})]]
    elif lines and isinstance(lines[0], tuple):
        lines = [lines]
    box = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    for i, runs in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = lsp
        for t, o in runs:
            r = p.add_run()
            r.text = t
            f = r.font
            f.name = FONT
            f.size = Pt(o.get('size', size))
            f.bold = o.get('bold', bold)
            f.color.rgb = o.get('color', color)
            if o.get('spc', spc):
                _spc(f, o.get('spc', spc))
            _lang(f)
    return box


def rect(sl, x, y, w, h, fill=None, line=None, lw=0.75,
         shape=MSO_SHAPE.RECTANGLE, adj=None, dash=None):
    sh = sl.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.shadow.inherit = False
    if adj is not None:
        sh.adjustments[0] = adj
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(lw)
        if dash:
            sh.line.dash_style = dash
    tf = sh.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = 0
    return sh


def label(sh, lines, size=9, color=WHITE, bold=False,
          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, spc=None, lsp=1.14):
    if isinstance(lines, str):
        lines = [[(lines, {})]]
    elif lines and isinstance(lines[0], tuple):
        lines = [lines]
    tf = sh.text_frame
    tf.vertical_anchor = anchor
    for i, runs in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = lsp
        for t, o in runs:
            r = p.add_run()
            r.text = t
            f = r.font
            f.name = FONT
            f.size = Pt(o.get('size', size))
            f.bold = o.get('bold', bold)
            f.color.rgb = o.get('color', color)
            if o.get('spc', spc):
                _spc(f, o.get('spc', spc))
            _lang(f)
    return sh


def line(sl, x1, y1, x2, y2, color=GREY4, lw=0.75, dash=None,
         head=False, tail=False):
    cn = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1),
                                 Inches(y1), Inches(x2), Inches(y2))
    cn.line.color.rgb = color
    cn.line.width = Pt(lw)
    if dash:
        cn.line.dash_style = dash
    ln = cn.line._get_or_add_ln()
    for flag, tag in ((head, 'a:headEnd'), (tail, 'a:tailEnd')):
        if flag:
            ln.append(ln.makeelement(
                qn(tag), {'type': 'triangle', 'w': 'sm', 'len': 'sm'}))
    return cn


def check(sl, cx, cy, r, color):
    """circulo + tique desenhado (independente de fonte)"""
    rect(sl, cx - r, cy - r, 2 * r, 2 * r, None, color, 0.75, MSO_SHAPE.OVAL)
    b = sl.shapes.build_freeform(Inches(cx - 0.52 * r), Inches(cy + 0.02 * r))
    b.add_line_segments([(Inches(cx - 0.14 * r), Inches(cy + 0.42 * r)),
                         (Inches(cx + 0.55 * r), Inches(cy - 0.44 * r))],
                        close=False)
    sh = b.convert_to_shape()
    sh.shadow.inherit = False
    sh.fill.background()
    sh.line.color.rgb = color
    sh.line.width = Pt(1.0)


def cross(sl, cx, cy, r, color):
    rect(sl, cx - r, cy - r, 2 * r, 2 * r, None, color, 0.75, MSO_SHAPE.OVAL)
    d = 0.48 * r
    line(sl, cx - d, cy - d, cx + d, cy + d, color, 1.0)
    line(sl, cx - d, cy + d, cx + d, cy - d, color, 1.0)


def bubble(sl, x, y, d, n, color):
    sh = rect(sl, x, y, d, d, WHITE, color, 1.0, MSO_SHAPE.OVAL)
    label(sh, n, size=9.5, color=color, bold=True)


def pct(v, nd=2):
    return ('%.*f' % (nd, v)).replace('.', ',') + '%'


def pp(v, nd=2, sign=True):
    s = ('%+.*f' if sign else '%.*f') % (nd, v)
    return s.replace('.', ',') + ' p.p.'


# ------------------------------------------------------------------ dados
# (conversao %, k leads/mes, categoria de alocacao, loja mais proxima, tinha carro)
BARS = [
    (0.86,  38, 'ideal', 'y', 'y'),
    (0.66,  46, 'gap',   'y', 'n'),
    (0.67,  30, 'gap',   'n', 'y'),
    (0.36,   7, 'gap',   'n', 'n'),
    (0.73,  62, 'avail', '-', 'y'),
    (0.44,  60, 'gap',   '-', 'n'),
    (0.43, 147, 'avail', 'y', '-'),
    (0.42,  39, 'gap',   'n', '-'),
    (0.30,  40, 'gap',   'n', '-'),
    (0.42, 134, 'gap',   '-', '-'),
]
FILL = {'ideal': FOREST, 'avail': SAGE, 'gap': GOLD}

GROUPS = [
    ('Carro escolhido + CEP', '', 0, 3, FOREST),
    ('Carro escolhido, sem CEP', '', 4, 5, SAGE_D),
    ('Carro não escolhido', 'com ou sem CEP', 6, 9, GREY2),
]

TOTAL = sum(b[1] for b in BARS)                      # 603k leads/mes


def gstats(a, b):
    vol = sum(x[1] for x in BARS[a:b + 1])
    conv = sum(x[0] * x[1] for x in BARS[a:b + 1]) / vol
    return vol, conv


GS = [gstats(a, b) for _, _, a, b, _ in GROUPS]
BASE = GS[2][1]                                      # 0,41% - grupo sem carro
MIGRAR = GS[0][0] + GS[1][0] - GS[0][0]              # placeholder, ver abaixo
MIGRAR = TOTAL - GS[0][0]                            # 482k sem informacao completa

# ------------------------------------------------------------------ grid
TITLE_X, TITLE_Y, TITLE_W, TITLE_H = 0.30, 0.14, 12.10, 0.78
RULE_Y = 0.96
BAND_Y, BAND_H = 1.02, 0.30
SUB_Y = 1.36
GB_Y, GB_H = 1.70, 0.64                              # faixa de nivel de informacao
ARR_Y, ARR_H = 2.36, 0.26                            # seta de migracao
CH_X, CH_W = 0.90, 9.50                              # area do marimekko
CH_TOP, CH_H, AXIS = 2.72, 2.50, 1.00                # topo, altura, eixo (%)
BASE_Y = CH_TOP + CH_H                               # 5.22
ROW_H = 0.30
PAN_X, PAN_W = 10.72, 2.31                           # painel das alavancas
GAP = 0.04
NOTE_Y, SRC_Y = 6.74, 7.06

SCALE = (CH_W - GAP * (len(BARS) - 1)) / TOTAL

_x, XS = CH_X, []
for _c, _v, _k, _a, _b in BARS:
    w = _v * SCALE
    XS.append((_x, w))
    _x += w + GAP

SUBTITLE = ('Conversão média dos leads por informação coletada e alocação '
            'final (% lead em venda, 2026)')
SRC = 'Fonte: bases internas (vendas e leads). Análises Bain'


# ------------------------------------------------------------------ cabecalho
def header(sl, title_runs, legend=True, dim=False):
    tc = DIMTXT if dim else GREY1
    txt(sl, TITLE_X, TITLE_Y, TITLE_W, TITLE_H, title_runs, size=19.5,
        color=tc, lsp=1.10)
    line(sl, TITLE_X, RULE_Y, SW - 0.30, RULE_Y,
         mix(RED, 0.72) if dim else RED, 1.5)

    band = rect(sl, TITLE_X, BAND_Y, 4.42, BAND_H,
                mix(BLACK, 0.90) if dim else RGBColor(0x1A, 0x1A, 0x1A),
                None, shape=MSO_SHAPE.PENTAGON, adj=0.10)
    label(band, 'A G E N D A   D E   V A L O R   L E A D   T O   S A L E S',
          size=8.5, color=WHITE, bold=True, align=PP_ALIGN.LEFT)
    band.text_frame.margin_left = Inches(0.12)

    if legend:
        ents = [(FOREST, 'Alocação ideal', 4.90, 1.06),
                (SAGE, 'Ideal com os dados disponíveis', 6.22, 1.84),
                (GOLD, 'Grande melhoria potencial', 8.28, 1.70)]
        for col, lab, x, w in ents:
            rect(sl, x, BAND_Y + 0.09, 0.14, 0.14,
                 mix(col, 0.78) if dim else col)
            txt(sl, x + 0.20, BAND_Y + 0.055, w, 0.22, lab, size=8,
                color=DIMTXT if dim else GREY1, anchor=MSO_ANCHOR.MIDDLE)

    txt(sl, PAN_X, BAND_Y + 0.03, PAN_W, 0.26,
        [('/', {'color': DIMTXT if dim else RED, 'bold': True}),
         ('P R E L I M I N A R', {'color': DIMTXT if dim else GREY1,
                                 'bold': True})],
        size=9, spc=60, anchor=MSO_ANCHOR.MIDDLE)

    txt(sl, TITLE_X, SUB_Y, 9.90, 0.24, SUBTITLE, size=9,
        color=DIMTXT if dim else GREY2)


def footer(sl, note, dim=False):
    txt(sl, TITLE_X, NOTE_Y, 12.10, 0.28, note, size=6.5,
        color=DIMTXT if dim else GREY3, lsp=1.20)
    txt(sl, TITLE_X, SRC_Y, 8.0, 0.18, SRC, size=6.5,
        color=DIMTXT if dim else GREY3)


# ------------------------------------------------------------------ paginas 1-3
TITLE1 = [
    [('Recap: ', {'bold': True}),
     ('80% dos leads chegam sem informação completa e convertem 0,41% a '
      '0,59%,', {})],
    [('contra 0,71% de quem informa carro e CEP; fechar o gap vale 300 a 950 '
      'carros por mês', {})],
]


def page_recap(prs, mode):
    """mode: 'full' | 'a1' | 'a2'"""
    d = mode != 'full'
    a1 = mode == 'a1'
    a2 = mode == 'a2'
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    header(sl, TITLE1, dim=d)

    # --- faixa de nivel de informacao -------------------------------------
    txt(sl, 0.14, GB_Y + 0.05, 0.72, 0.54,
        [[('NÍVEL DE', {})], [('INFORMAÇÃO', {})], [('DO LEAD', {})]],
        size=6.5, color=DIMTXT if (d and not a1) else GREY2, bold=True,
        spc=40, lsp=1.12)

    for i, (l1, l2, a, b, col) in enumerate(GROUPS):
        vol, conv = GS[i]
        x0 = XS[a][0]
        x1 = XS[b][0] + XS[b][1]
        vivid = a1 or not d
        c = col if vivid else mix(col, 0.80)
        bg = mix(col, 0.93 if vivid else 0.975)
        brd = RED if (a1 and i > 0) else (mix(col, 0.55) if vivid
                                         else mix(col, 0.85))
        box = rect(sl, x0, GB_Y, x1 - x0, GB_H, bg, brd,
                   1.5 if (a1 and i > 0) else 0.75)
        label(box, [
            [(l1, {'size': 9.5, 'bold': True, 'color': c})] +
            ([('  ' + l2, {'size': 8, 'color': c})] if l2 else []),
            [('%dk leads/mês · %d%%' % (vol, round(100.0 * vol / TOTAL)),
              {'size': 7.5, 'color': mix(c, 0.30) if vivid else c})],
            [('conversão média ', {'size': 8, 'color': c}),
             (pct(conv), {'size': 11, 'bold': True, 'color': c})],
        ], lsp=1.04)

    # --- seta de migracao --------------------------------------------------
    av = a1 or not d
    arr = rect(sl, CH_X, ARR_Y, CH_W, ARR_H,
               RED if av else mix(RED, 0.86), None,
               shape=MSO_SHAPE.LEFT_ARROW, adj=0.42)
    label(arr, [[('Alavanca 1: ', {'bold': True}),
                 ('migrar leads para a esquerda; hoje %dk leads/mês (%d%%) '
                  'chegam sem carro escolhido e/ou sem CEP'
                  % (MIGRAR, round(100.0 * MIGRAR / TOTAL)), {})]],
          size=8.5, color=WHITE)

    # --- barras ------------------------------------------------------------
    for i, (conv, vol, kind, _p, _c) in enumerate(BARS):
        x, w = XS[i]
        col = FILL[kind]
        rect(sl, x, BASE_Y - conv / AXIS * CH_H, w, conv / AXIS * CH_H,
             col if not d else mix(col, 0.86))

    line(sl, CH_X, BASE_Y, CH_X + CH_W, BASE_Y,
         GREY4 if not d else GREY5, 1.0)

    # nivel medio de cada grupo, tracejado por cima das barras
    # (omitido no build da alavanca 2, que usa o proprio tracejado do gap)
    if not a2:
        for i, (_l1, _l2, a, b, col) in enumerate(GROUPS):
            vivid = a1 or not d
            y = BASE_Y - GS[i][1] / AXIS * CH_H
            line(sl, XS[a][0], y, XS[b][0] + XS[b][1], y,
                 col if vivid else mix(col, 0.70), 1.5 if vivid else 0.75,
                 MSO_LINE_DASH_STYLE.DASH)

    # rotulos das barras por ultimo, em cartao branco, para nao serem
    # cobertos pelas barras vizinhas nem cortados pelos tracejados
    gmy = [BASE_Y - GS[i][1] / AXIS * CH_H for i in range(len(GROUPS))]
    for i, (conv, _v, _k, _p, _c) in enumerate(BARS):
        x, w = XS[i]
        tw = max(w, 0.46)
        ly = BASE_Y - conv / AXIS * CH_H - 0.23
        # se o tracejado da media do grupo cortaria o rotulo, sobe o rotulo
        # para logo acima da linha em vez de tapa-la
        for y in gmy:
            if ly - 0.02 <= y <= ly + 0.22:
                ly = y - 0.23
        # cartao branco so na barra estreita, cujo rotulo invade a vizinha
        if w < 0.30:
            rect(sl, x + w / 2 - tw / 2, ly, tw, 0.20, WHITE,
                 GREY5 if not d else None, 0.5)
        txt(sl, x + w / 2 - tw / 2, ly + 0.01, tw, 0.19, pct(conv),
            size=8, color=DIMTXT if d else GREY1, bold=True,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # --- tabela loja mais proxima / tinha o carro / volume -----------------
    sv = a2 or not d
    rows = [('Loja mais\npróxima', 3), ('Loja tinha\no carro', 4)]
    for ri, (lab, idx) in enumerate(rows):
        y = BASE_Y + ri * ROW_H
        txt(sl, 0.16, y + 0.03, 0.70, ROW_H - 0.04,
            [[(lab.split('\n')[0], {})], [(lab.split('\n')[1], {})]],
            size=7, color=GREY2 if sv else DIMTXT, lsp=1.06)
        for i, bar in enumerate(BARS):
            x, w = XS[i]
            cx, cy = x + w / 2, y + ROW_H / 2
            v = bar[idx]
            if v == 'y':
                check(sl, cx, cy, 0.075, SAGE_D if sv else GREY5)
            elif v == 'n':
                cross(sl, cx, cy, 0.075, RED if sv else mix(RED, 0.86))
            else:
                tw = max(w, 0.36)
                txt(sl, cx - tw / 2, cy - 0.09, tw, 0.18, 'N/A', size=7,
                    color=GREY3 if sv else GREY5, align=PP_ALIGN.CENTER,
                    anchor=MSO_ANCHOR.MIDDLE)
        line(sl, CH_X, y + ROW_H, CH_X + CH_W, y + ROW_H,
             GREY5 if sv else GREY6, 0.75)

    yv = BASE_Y + 2 * ROW_H
    txt(sl, 0.16, yv + 0.05, 0.70, 0.20, 'k leads/mês', size=7,
        color=GREY2 if (a1 or not d) else DIMTXT)
    for i, bar in enumerate(BARS):
        x, w = XS[i]
        tw = max(w, 0.34)
        txt(sl, x + w / 2 - tw / 2, yv + 0.05, tw, 0.20, '%d' % bar[1],
            size=7.5, color=GREY1 if (a1 or not d) else DIMTXT, bold=True,
            align=PP_ALIGN.CENTER)

    # --- callouts dos builds ----------------------------------------------
    if a1:
        cx0, cy0, cw = 5.34, 2.80, 5.00
        rect(sl, cx0, cy0, cw, 1.02, WHITE, RED, 1.5)
        txt(sl, cx0 + 0.14, cy0 + 0.09, cw - 0.28, 0.18,
            'Cada degrau de informação vale conversão', size=8.5,
            color=RED, bold=True)
        rw = [('Carro não escolhido', GS[2][1], None),
              ('Carro escolhido, sem CEP', GS[1][1], GS[1][1] - BASE),
              ('Carro escolhido + CEP', GS[0][1], GS[0][1] - BASE)]
        for j, (lab, cv, dl) in enumerate(rw):
            y = cy0 + 0.32 + j * 0.22
            txt(sl, cx0 + 0.14, y, 2.30, 0.20, lab, size=8, color=GREY1)
            txt(sl, cx0 + 2.50, y, 0.62, 0.20, pct(cv), size=8.5,
                color=GREY1, bold=True, align=PP_ALIGN.RIGHT)
            txt(sl, cx0 + 3.22, y, 1.60, 0.20,
                'base de comparação' if dl is None else pp(dl) + ' vs. base',
                size=8, color=GREY3 if dl is None else RED,
                bold=dl is not None)

    if a2:
        # gap intra-grupo: nivel do melhor grupo + seta ate a pior barra
        for i, (_l1, _l2, a, b, _c) in enumerate(GROUPS):
            best = max(range(a, b + 1), key=lambda k: BARS[k][0])
            worst = min(range(a, b + 1), key=lambda k: BARS[k][0])
            ytop = BASE_Y - BARS[best][0] / AXIS * CH_H
            ybot = BASE_Y - BARS[worst][0] / AXIS * CH_H
            line(sl, XS[a][0], ytop, XS[b][0] + XS[b][1], ytop, RUBY, 1.0,
                 MSO_LINE_DASH_STYLE.SQUARE_DOT)
            wx = XS[worst][0] + XS[worst][1] / 2
            # em barra estreita o rotulo fica em cartao branco acima da
            # barra: a seta para na borda do cartao, nao por cima dele
            if XS[worst][1] < 0.30:
                ybot -= 0.24
            line(sl, wx, ytop, wx, ybot, RUBY, 1.25, head=True, tail=True)

        cx0, cy0, cw = 5.34, 2.80, 5.00
        rect(sl, cx0, cy0, cw, 1.02, WHITE, RED, 1.5)
        txt(sl, cx0 + 0.14, cy0 + 0.09, cw - 0.28, 0.18,
            'Com a mesma informação, a alocação muda a conversão', size=8.5,
            color=RED, bold=True)
        for j, (_l1, _l2, a, b, _c) in enumerate(GROUPS):
            best = max(BARS[a:b + 1], key=lambda k: k[0])[0]
            worst = min(BARS[a:b + 1], key=lambda k: k[0])[0]
            y = cy0 + 0.32 + j * 0.22
            txt(sl, cx0 + 0.14, y, 2.30, 0.20, GROUPS[j][0] + ' ' +
                GROUPS[j][1], size=8, color=GREY1)
            txt(sl, cx0 + 2.50, y, 1.30, 0.20,
                '%s → %s' % (pct(best), pct(worst)), size=8.5, color=GREY1,
                bold=True, align=PP_ALIGN.RIGHT)
            txt(sl, cx0 + 3.92, y, 0.92, 0.20, pp(worst - best), size=8,
                color=RUBY, bold=True, align=PP_ALIGN.RIGHT)

    # --- painel das alavancas ---------------------------------------------
    txt(sl, PAN_X, GB_Y - 0.24, PAN_W, 0.42,
        [[('Alavancas de ganho', {})], [('de conversão', {})]], size=10.5,
        color=DIMTXT if d else GREY1, bold=True, lsp=1.08)
    line(sl, PAN_X, GB_Y + 0.22, SW - 0.30, GB_Y + 0.22,
         GREY5 if d else GREY4, 0.75)

    lev = [
        (1, [[('Aumento da ', {}),
              ('proporção de leads com informações mínimas necessárias',
               {'bold': True})],
             [('quais são: carro desejado e CEP',
               {'size': 8, 'color': GREY3})]]),
        (2, [[('Melhoria da ', {}),
              ('qualidade da alocação dos leads', {'bold': True}),
              (' com base nas preferências do cliente', {})],
             [('loja mais próxima e com o carro em estoque',
               {'size': 8, 'color': GREY3})]]),
    ]
    for j, (n, body) in enumerate(lev):
        y = GB_Y + 0.44 + j * 1.34
        on = (not d) or (a1 and n == 1) or (a2 and n == 2)
        acc = RED if (d and on) else (FOREST if on else DIMTXT)
        if d and on:
            rect(sl, PAN_X - 0.08, y - 0.12, PAN_W + 0.14, 1.14, None, RED,
                 1.75)
        bubble(sl, PAN_X, y, 0.26, str(n), acc)
        for para in body:
            for run in para:
                if not on:
                    run[1]['color'] = DIMTXT
                elif 'color' not in run[1]:
                    run[1]['color'] = GREY1
        txt(sl, PAN_X + 0.34, y - 0.03, PAN_W - 0.32, 1.02, body, size=9.5,
            lsp=1.16)

    # --- conclusao ---------------------------------------------------------
    box = rect(sl, CH_X, 6.22, CH_W + PAN_W + (PAN_X - CH_X - CH_W), 0.50,
               RED_BG if not d else mix(RED, 0.94), None)
    label(box, [[('Fechar 25% a 75% do gap até a maior conversão '
                  '(+0,1 a 0,3 p.p. sobre uma conversão média de 0,46%) '
                  'pode gerar', {})],
                [('R$ 20 a 60M por ano, ou 300 a 950 carros adicionais por '
                  'mês', {})]],
          size=10.5, color=RED if not d else mix(RED, 0.72), bold=True)

    footer(sl, [[('¹ Médias por grupo ponderadas pelo volume de leads das '
                  'barras exibidas; largura das barras proporcional ao volume '
                  'de leads/mês. ² Os grupos também diferem em intenção de '
                  'compra: capturar o ganho exige coleta ativa da informação, '
                  'não apenas segmentar os leads.', {})]], dim=d)
    return sl


# ------------------------------------------------------------------ pagina 4
STEPS = [
    ('Carro não escolhido', 'com ou sem CEP', GS[2][0], GS[2][1], GREY2,
     RGBColor(0xC9, 0xC9, 0xC9)),
    ('Carro escolhido', 'sem CEP', GS[1][0], GS[1][1], SAGE_D, SAGE),
    ('Carro escolhido', 'e CEP informado', GS[0][0], GS[0][1], FOREST,
     FOREST),
]
CAPTURA = [
    ('2.1.1', 'Site', 'validação dos dados coletados e 2ª página com a loja '
     'de preferência', 'S2 set/26'),
    ('2.1.2', 'Facebook e WhatsApp', 'CEP capturado para todos os leads e '
     '2ª página com a loja de preferência', 'S4 set/26'),
    ('2.1.3', 'Canais classificados', 'CEP capturado para todos os leads',
     'S3 set/26'),
]


def page_alavanca1(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    title = [
        [('Alavanca 1: ', {'bold': True}),
         ('60% dos leads não informam qual carro querem; levá-los a '
          'informar', {})],
        [('carro e CEP eleva a conversão de 0,41% para 0,71%', {})],
    ]
    global SUBTITLE
    keep, SUBTITLE = SUBTITLE, ('Conversão média por nível de informação do '
                                'lead (% lead em venda, 2026)')
    header(sl, title, legend=False)
    SUBTITLE = keep

    x0, w_tot = 0.90, 7.24
    gp = 0.10
    sc = (w_tot - gp * 2) / TOTAL
    top, ch, ax = 2.06, 2.34, 0.80
    base = top + ch                                   # 4.40

    x = x0
    geo = []
    for _l1, _l2, vol, conv, cd, cf in STEPS:
        w = vol * sc
        geo.append((x, w, conv, cd, cf))
        x += w + gp

    for i, (bx, bw, conv, cd, cf) in enumerate(geo):
        h = conv / ax * ch
        rect(sl, bx, base - h, bw, h, cf)
        txt(sl, bx, base - h - 0.42, bw, 0.38, pct(conv), size=21,
            color=cd, bold=True, align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.BOTTOM)
        if i:
            chip = rect(sl, bx + bw / 2 - 0.48, base - h + 0.10, 0.96, 0.24,
                        WHITE, cd, 1.0)
            label(chip, pp(conv - BASE, 2), size=8.5, color=cd, bold=True)
        l1, l2, vol = STEPS[i][0], STEPS[i][1], STEPS[i][2]
        txt(sl, bx, base + 0.07, bw, 0.52, [
            [(l1, {'size': 10.5, 'bold': True, 'color': cd})],
            [(l2, {'size': 8.5, 'color': GREY3})],
            [('%dk/mês · %d%% dos leads' % (vol, round(100.0 * vol / TOTAL))
              if i == 0 else '%dk/mês · %d%%'
              % (vol, round(100.0 * vol / TOTAL)),
              {'size': 8.5, 'color': GREY2, 'bold': True})],
        ], align=PP_ALIGN.CENTER, lsp=1.08)
    line(sl, x0, base, x0 + w_tot, base, GREY4, 1.0)

    # seta de migracao
    arr = rect(sl, x0, 5.08, w_tot, 0.34, RED, None,
               shape=MSO_SHAPE.RIGHT_ARROW, adj=0.40)
    label(arr, [[('%dk leads/mês (%d%%) ' % (MIGRAR,
                                             round(100.0 * MIGRAR / TOTAL)),
                  {'bold': True}),
                 ('a migrar para os níveis com mais informação', {})]],
          size=9.5, color=WHITE)

    # cenarios de impacto
    txt(sl, x0, 5.62, w_tot, 0.20,
        'Impacto na conversão média, por parcela dos leads migrada para '
        '"carro escolhido + CEP"¹', size=9, color=GREY1, bold=True)
    cw0, cwn = 2.44, 1.60
    hdr = ['Parcela dos leads migrada', '10 p.p.', '20 p.p.', '30 p.p.']
    for ci, h in enumerate(hdr):
        cx = x0 + (0 if ci == 0 else cw0 + (ci - 1) * cwn)
        cwd = cw0 if ci == 0 else cwn
        cell = rect(sl, cx, 5.88, cwd, 0.28,
                    GREY6 if ci == 0 else mix(FOREST, 0.90), None)
        label(cell, h, size=9, color=GREY1 if ci == 0 else FOREST, bold=True,
              align=PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER)
        cell.text_frame.margin_left = Inches(0.08)
    gain = [0.10, 0.20, 0.30]
    cell = rect(sl, x0, 6.16, cw0, 0.30, WHITE, GREY5)
    label(cell, 'Ganho na conversão média', size=9, color=GREY1,
          align=PP_ALIGN.LEFT)
    cell.text_frame.margin_left = Inches(0.08)
    for ci, g in enumerate(gain):
        cx = x0 + cw0 + ci * cwn
        cell = rect(sl, cx, 6.16, cwn, 0.30, WHITE, GREY5)
        label(cell, pp(g * (GS[0][1] - BASE), 2), size=11, color=RED,
              bold=True)

    # painel de captura
    px, pw = 8.42, 4.61
    rect(sl, px, GB_Y - 0.18, pw, 4.94, RGBColor(0xF7, 0xF9, 0xF8), GREY5)
    txt(sl, px + 0.16, GB_Y - 0.02, pw - 0.32, 0.24,
        'Onde a informação passa a ser capturada', size=10.5, color=FOREST,
        bold=True)
    line(sl, px + 0.16, GB_Y + 0.28, px + pw - 0.16, GB_Y + 0.28, GREY4, 0.75)
    for j, (mid, canal, det, prazo) in enumerate(CAPTURA):
        y = GB_Y + 0.46 + j * 1.24
        bubble(sl, px + 0.16, y, 0.26, str(j + 1), FOREST)
        txt(sl, px + 0.52, y - 0.02, pw - 0.90, 0.86, [
            [(canal, {'size': 10.5, 'bold': True, 'color': GREY1})],
            [(det, {'size': 9, 'color': GREY2})],
            [('milestone %s · %s · Thais Pimenta' % (mid, prazo),
              {'size': 8, 'color': GREY3})],
        ], lsp=1.14)
        if j < 2:
            line(sl, px + 0.16, y + 1.04, px + pw - 0.16, y + 1.04, GREY5,
                 0.75)

    box = rect(sl, px + 0.16, GB_Y + 4.22, pw - 0.32, 0.46, RED_BG, None)
    label(box, [[('Sem esses três pontos de captura, a alavanca 2 age sobre '
                  'apenas 20% dos leads', {})]], size=9, color=RED, bold=True)

    footer(sl, [[('¹ Ganho incremental de %s por lead migrado do grupo "carro '
                  'não escolhido" (%s) para "carro escolhido + CEP" (%s), '
                  'mantida a alocação atual; é parte do +0,1 a 0,3 p.p. '
                  'total, o restante vem da alavanca 2. ² Os grupos também '
                  'diferem em intenção de compra: capturar o ganho exige '
                  'coleta ativa da informação, não apenas segmentar os leads.'
                  % (pp(GS[0][1] - BASE), pct(BASE), pct(GS[0][1])), {})]])
    return sl


# ------------------------------------------------------------------ main
def main(dst):
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    for mode in ('full', 'a1', 'a2'):
        page_recap(prs, mode)
    page_alavanca1(prs)
    cp = prs.core_properties
    cp.title = 'Recap Lead to Sales: conversão por informação coletada'
    cp.author = 'ConsultingInnovation@bain.com'
    cp.last_modified_by = 'ConsultingInnovation@bain.com'
    prs.save(dst)
    print('ok %s (%d paginas)' % (dst, len(prs.slides.__iter__.__self__._sldIdLst)))
    print('total %dk leads/mes | grupos: %s' % (
        TOTAL, ' | '.join('%s %s (%dk, %d%%)' % (g[0], pct(GS[i][1]), GS[i][0],
                                                 round(100.0 * GS[i][0] / TOTAL))
                          for i, g in enumerate(GROUPS))))


if __name__ == '__main__':
    import sys
    main(sys.argv[1] if len(sys.argv) > 1
         else '../output/recap_alavancas_conversao.pptx')
