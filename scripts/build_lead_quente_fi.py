# -*- coding: utf-8 -*-
"""Cinco paginas 16:9 no padrao Bain Core, reproduzindo os prints enviados.

  1. LEAD QUENTE | NOSSA PROPOSTA       - quatro passos em cards
  2. F&I MULTIBANCO | VISAO GERAL       - ambicao, valor e cinco frentes
  3. F&I MULTIBANCO | PLANO DE EXECUCAO - tabela das cinco frentes
  4. L2S | LEAD QUENTE                  - variante com icones e logo Bain
  5. F&I MULTIBANCO                     - variante com icones e logo Bain

Tudo e montado com shapes, caixas de texto e tabela nativas do PowerPoint:
nada de imagem colada, todo elemento continua editavel no ppt. Os icones vem
de `icons_bain.py`, desenhados com freeforms e formas basicas.

    python3 build_lead_quente_fi.py ../output/lead_quente_fi_multibanco.pptx
"""
import sys

from PIL import ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from icons_bain import icon

# ------------------------------------------------------------------ tokens
RED    = RGBColor(0xCC, 0x00, 0x00)   # Bain red
RED_BG = RGBColor(0xFB, 0xEF, 0xEF)   # faixa de merito da pagina 4
RED_LN = RGBColor(0xE6, 0xB3, 0xB3)   # regua do kicker da pagina 5
BLACK  = RGBColor(0x00, 0x00, 0x00)
BODY   = RGBColor(0x26, 0x26, 0x26)   # texto corrente
SUB    = RGBColor(0x55, 0x55, 0x55)   # subtitulo
EYE    = RGBColor(0x8C, 0x8C, 0x8C)   # kicker e rodape
GREY4  = RGBColor(0xB4, 0xB4, 0xB4)
RULE   = RGBColor(0xD9, 0xD9, 0xD9)
CARD   = RGBColor(0xF2, 0xF2, 0xF2)
BAND   = RGBColor(0xEA, 0xEA, 0xEA)
BAND2  = RGBColor(0xF7, 0xF7, 0xF7)
THEAD  = RGBColor(0xE6, 0xE6, 0xE6)
TZEBRA = RGBColor(0xFA, 0xFA, 0xFA)
TCOL   = RGBColor(0xF4, 0xF4, 0xF4)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)

FONT   = 'Arial'
LANG   = 'pt-BR'
SW, SH = 13.333, 7.5
M      = 0.55                 # margem esquerda/direita
CW     = SW - 2 * M           # largura util

T_EYE  = 9.0                  # kicker
T_TIT  = 23.0                 # titulo de acao
T_SUB  = 11.5                 # subtitulo / pre-requisito
T_CARD = 14.5                 # titulo de card
T_BODY = 10.5                 # texto corrente
T_FOOT = 8.5                  # fonte e numero da pagina

_TTF = {False: '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
        True:  '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf'}
_fc = {}
_WARN = []


def text_w(s, pt, bold=False):
    """largura do texto em polegadas"""
    if bold not in _fc:
        _fc[bold] = ImageFont.truetype(_TTF[bold], 200)
    return _fc[bold].getlength(s) / 200.0 * pt / 72.0


def nlines(s, pt, width, bold=False):
    """quantas linhas o texto ocupa numa caixa de `width` polegadas"""
    n, cur = 1, ''
    for w in s.split():
        cand = (cur + ' ' + w).strip()
        if cur and text_w(cand, pt, bold) > width:
            n += 1
            cur = w
        else:
            cur = cand
    return n


def fits(tag, need, have):
    """registra transbordo em vez de deixar passar silenciosamente"""
    if need > have + 0.01:
        _WARN.append('%s: precisa de %.2f" e tem %.2f"' % (tag, need, have))


# ------------------------------------------------------------------ helpers
def no_shadow(shape):
    el = shape._element
    st = el.find(qn('p:style'))
    if st is not None:
        el.remove(st)
    spPr = el.spPr
    for e in spPr.findall(qn('a:effectLst')):
        spPr.remove(e)
    spPr.append(spPr.makeelement(qn('a:effectLst'), {}))


def shp(sl, kind, x, y, w, h, fill=None, line=None, lw=1.0, rot=0.0, adj=None):
    s = sl.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    no_shadow(s)
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(lw)
    if rot:
        s.rotation = rot
    if adj is not None:
        for i, v in enumerate(adj):
            try:
                s.adjustments[i] = v
            except (IndexError, ValueError):
                pass
    s.text_frame.word_wrap = True
    return s


def rect(sl, x, y, w, h, fill=None, line=None, lw=1.0, rot=0.0):
    return shp(sl, MSO_SHAPE.RECTANGLE, x, y, w, h, fill, line, lw, rot)


def ell(sl, x, y, w, h, fill=None, line=None, lw=1.0):
    return shp(sl, MSO_SHAPE.OVAL, x, y, w, h, fill, line, lw)


def hline(sl, x, y, w, color=RULE, thick=0.012):
    return rect(sl, x, y, w, thick, color, None)


def vline(sl, x, y, h, color=RULE, thick=0.010):
    return rect(sl, x, y, thick, h, color, None)


def _run(p, text, size, color, bold=False, spc=None):
    r = p.add_run()
    r.text = text
    f = r.font
    f.name = FONT
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = color
    rPr = r._r.get_or_add_rPr()
    rPr.set('lang', LANG)
    rPr.set('dirty', '0')
    if spc:
        rPr.set('spc', str(int(spc)))
    return r


def _norm(lines):
    """str | [(txt, bold)] | [[(txt, bold)], ...]  ->  lista de paragrafos"""
    if isinstance(lines, str):
        return [[(lines, False)]]
    if lines and isinstance(lines[0], tuple):
        return [list(lines)]
    return [list(p) if not isinstance(p, str) else [(p, False)] for p in lines]


def txt(host, x=None, y=None, w=None, h=None, lines='', size=T_BODY,
        color=BODY, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
        lsp=1.18, spc=None, space_before=0.0, wrap=True):
    """host: slide (cria textbox) ou shape (usa o text_frame existente)"""
    if x is None:
        tf = host.text_frame
        tf.word_wrap = wrap
    else:
        box = host.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.word_wrap = wrap
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    for i, parts in enumerate(_norm(lines)):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = lsp
        if i and space_before:
            p.space_before = Pt(space_before * 72)
        for seg, b in parts:
            _run(p, seg, size, color, bold or b, spc)
    return tf


def _bullet(p, indent):
    pPr = p._p.get_or_add_pPr()
    pPr.set('marL', str(int(indent * 914400)))
    pPr.set('indent', str(int(-indent * 914400)))
    pPr.append(pPr.makeelement(qn('a:buFont'), {'typeface': 'Arial'}))
    pPr.append(pPr.makeelement(qn('a:buChar'), {'char': '•'}))


def bullets(sl, x, y, w, items, size=T_BODY, gap=0.15, lsp=1.18, indent=0.15):
    """lista com marcador; devolve a altura ocupada, em polegadas"""
    tf = txt(sl, x, y, w, 0.4, items, size=size, color=BODY, lsp=lsp,
             space_before=gap)
    for p in tf.paragraphs:
        _bullet(p, indent)
    n = sum(nlines(''.join(s for s, _ in _norm(items)[i]), size, w - indent)
            for i in range(len(_norm(items))))
    return n * size * lsp / 72.0 + gap * (len(_norm(items)) - 1)


# ------------------------------------------------------------------ chrome
def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def eyebrow(sl, text):
    txt(sl, M, 0.30, CW, 0.22, [[(text.upper(), False)]], size=T_EYE,
        color=EYE, spc=180)


def title(sl, l1, l2, y=0.62):
    txt(sl, M, y, CW, 0.44, [[(l1, True)]], size=T_TIT, color=BLACK,
        bold=True, lsp=1.0)
    txt(sl, M, y + 0.395, CW, 0.44, [[(l2, True)]], size=T_TIT, color=RED,
        bold=True, lsp=1.0)


def footer(sl, source, page, logo=False):
    y = 7.08
    hline(sl, M, y - 0.10, CW, RULE, 0.008)
    txt(sl, M, y, 10.2, 0.3, [[('Fonte: ', True), (source, False)]],
        size=T_FOOT, color=EYE)
    r = SW - M
    if logo:
        txt(sl, r - 0.34, y, 0.34, 0.3, [[('%d' % page, False)]],
            size=T_FOOT, color=EYE, align=PP_ALIGN.RIGHT, wrap=False)
        txt(sl, r - 0.60, y, 0.16, 0.3, [[('|', False)]], size=T_FOOT,
            color=GREY4, align=PP_ALIGN.CENTER, wrap=False)
        ell(sl, r - 0.90, y + 0.015, 0.155, 0.155, RED, None)
        lw = text_w('BAIN & COMPANY', 8.5, True) + 0.24
        txt(sl, r - 1.00 - lw, y, lw, 0.3, [[('BAIN & COMPANY', True)]],
            size=8.5, color=BLACK, bold=True, spc=60, align=PP_ALIGN.RIGHT,
            wrap=False)
    else:
        txt(sl, r - 0.6, y, 0.6, 0.3, [[('%d' % page, False)]], size=T_FOOT,
            color=EYE, align=PP_ALIGN.RIGHT, wrap=False)


def numbubble(sl, cx, cy, d, n, size):
    c = ell(sl, cx, cy, d, d, RED, None)
    txt(c, lines=[[(str(n), True)]], size=size, color=WHITE, bold=True,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, lsp=1.0)
    return c


# ------------------------------------------------------------------ conteudo
PASSOS = [
    dict(n=1, tit='Comunicar\ndivisionais', ic='barras', dono='Glauco',
         prazo='Semana 1', dono_ic='pessoa', bullets=[
             'Glauco leva o número aos divisionais: leads quentes recebidos '
             'e conversão por regional',
             'Ranking nominal explícito: quem está bem e quem está mal',
             'Mensagem única: o acompanhamento passa a ser semanal']),
    dict(n=2, tit='Priorizar\nna ponta', ic='pessoas_cheio',
         dono='Regionais e GVs', prazo='Semanas 1 e 2', dono_ic='pessoas',
         bullets=[
             'Cascata completa: divisional para regional, regional para GV, '
             'GV para CV',
             'Cada CV recebe a lista nominal dos leads quentes do dia',
             'Instrução única: ligar em 100% da lista no mesmo dia']),
    dict(n=3, tit='Avisar da\nconsequência', ic='alerta',
         dono='Glauco e GVs', prazo='Semana 2', dono_ic='pessoa', bullets=[
             'Regra comunicada antes de valer: não converter lead quente '
             'reduz o volume recebido',
             'Conversão por CV acompanhada semana a semana',
             'Sem surpresa: aviso formal antes do primeiro corte']),
    dict(n=4, tit='Realocar\no lead quente', ic='ciclo',
         dono='L2S e comercial', prazo='Semanas 3 e 4, em piloto',
         dono_ic='pessoas', bullets=[
             'Alocação à loja segue o critério do cliente, como hoje',
             'Dentro da loja, o mix de quente, morno e frio segue a '
             'performance do CV',
             'Lead quente não tratado no prazo é realocado para quem atende']),
]

FRENTES = [
    dict(n=1, tit='Bancos', ic='banco',
         txt='Integrar em ondas, dos bancos menores aos maiores, à medida '
             'que a proposta de valor se confirma'),
    dict(n=2, tit='Tecnologia', ic='engrenagem',
         txt='APIs para simulação via Corban do lojista e envio da ficha a '
             'múltiplos bancos em um único disparo'),
    dict(n=3, tit='Interface\nda plataforma', ic='monitor',
         txt='Melhoria contínua guiada pelo uso do vendedor, com ciclos '
             'curtos de release'),
    dict(n=4, tit='Lojistas', ic='pessoas_cheio',
         txt='Piloto em pequena escala para provar a proposta de valor e, '
             'confirmada, escalar por onda'),
    dict(n=5, tit='Ecossistema', ic='rede',
         txt='Novos serviços e integração com Valoriza depois que o núcleo '
             'estiver estável'),
]

VALOR = [
    dict(tit='Lojista e seu cliente', ic='pessoas_cheio',
         txt='Mais aprovação, menos retrabalho e resposta em minutos'),
    dict(tit='Localiza', ic='carro',
         txt='Mais penetração de F&I e relação mais forte com a rede de '
             'lojistas'),
    dict(tit='Banco', ic='banco',
         txt='Originação qualificada com custo de aquisição menor'),
]

AMBICAO = ('Plataforma de F&I multibanco e ponta a ponta que faz o lojista '
           'vender mais, com simulação e envio de ficha em um só fluxo')

PLANO = [
    ('1', 'Bancos', 'Bancos integrados\nem cascata',
     '1 banco conectado,\nenvio manual', '4 a 5 bancos em\ncascata automática',
     '2T26'),
    ('2', 'Tecnologia', 'APIs de simulação\ne envio de ficha',
     'Simulação via Corban,\nficha manual',
     'Disparo único para\nmúltiplos bancos', '2T26'),
    ('3', 'Interface da\nplataforma', 'Jornada do vendedor\nponta a ponta',
     'Versão inicial em uso', 'Release contínuo\nguiado pelo uso',
     'Contínuo'),
    ('4', 'Lojistas', 'Piloto e plano\nde escala',
     'Amostra do piloto\nem definição', 'Piloto provado e\nescala por onda',
     '1T26'),
    ('5', 'Ecossistema', 'Integração com\nValoriza e novos serviços',
     'Escopo em discussão', 'Oferta ampliada na\nmesma jornada', '2027'),
]

SRC_LQ = 'análise Bain; discussão com liderança comercial Localiza Seminovos'
SRC_F1 = 'análise Bain. Preliminar, para discussão'
SRC_F2 = ('análise Bain. Conteúdo preliminar, para preenchimento com os '
          'donos de cada frente')

TIT_LQ = ('Converter o lead quente exige rodar quatro passos:',
          'comunicar, priorizar, avisar da consequência e realocar')
TIT_FI = ('Entregar a plataforma de F&I multibanco depende de cinco frentes,',
          'iniciadas por um piloto com lojistas')

MERITO = ('Lead quente passa a ser mérito, não direito:',
          ' quem converte recebe mais, quem não converte recebe menos')


# ------------------------------------------------------------------ pagina 1
def page_lead_quente(prs):
    sl = blank(prs)
    eyebrow(sl, 'Lead quente  |  Nossa proposta')
    title(sl, *TIT_LQ)
    txt(sl, M, 1.46, 11.6, 0.5,
        'Pré-requisito: o número existe e é nominal – leads quentes '
        'recebidos, conversão e tempo até o primeiro contato por divisional, '
        'loja e CV', size=T_SUB, color=SUB, lsp=1.20)

    cy, ch, gap = 1.98, 4.32, 0.34
    cwid = (CW - 3 * gap) / 4.0
    for i, p in enumerate(PASSOS):
        x = M + i * (cwid + gap)
        rect(sl, x, cy, cwid, ch, CARD, None)
        numbubble(sl, x + 0.20, cy + 0.24, 0.46, p['n'], 16)
        txt(sl, x + 0.82, cy + 0.26, cwid - 1.02, 0.76,
            [[(l, True)] for l in p['tit'].split('\n')],
            size=T_CARD, color=BLACK, bold=True, lsp=1.10)
        by = cy + ch - 1.08
        need = bullets(sl, x + 0.24, cy + 1.22, cwid - 0.48, p['bullets'],
                       size=10)
        fits('p1 card %d' % p['n'], need, by - 0.18 - (cy + 1.22))
        icon(sl, p['dono_ic'], x + 0.24, by + 0.03, 0.38, RED, CARD)
        txt(sl, x + 0.76, by, cwid - 0.96, 0.7,
            [[('Dono: ', True), (p['dono'], False)],
             [('Prazo: ', True), (p['prazo'], False)]],
            size=10, lsp=1.28, space_before=0.06)

    fy = 6.44
    rect(sl, M, fy, CW, 0.52, CARD, None)
    c = ell(sl, M + 0.12, fy + 0.09, 0.34, 0.34, RED, None)
    txt(c, lines=[[('›', True)]], size=17, color=WHITE, bold=True,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, lsp=1.0)
    txt(sl, M + 0.62, fy, CW - 0.76, 0.52,
        [(MERITO[0], True), (MERITO[1], False)], size=12, color=BLACK,
        anchor=MSO_ANCHOR.MIDDLE)

    footer(sl, SRC_LQ, 1)
    return sl


# ------------------------------------------------------------------ pagina 2
def page_fi_visao(prs):
    sl = blank(prs)
    eyebrow(sl, 'F&I multibanco  |  Visão geral')
    title(sl, *TIT_FI)

    ay, ah, lw = 1.48, 0.86, 1.95
    lab = rect(sl, M, ay, lw, ah, RED, None)
    txt(lab, lines=[[('Ambição', True)]], size=13, color=WHITE, bold=True,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    rect(sl, M + lw + 0.10, ay, CW - lw - 0.10, ah, BAND, None)
    txt(sl, M + lw + 0.40, ay, CW - lw - 0.70, ah, [(AMBICAO, True)],
        size=12.5, color=BODY, bold=True, anchor=MSO_ANCHOR.MIDDLE, lsp=1.22)

    vy, vh = 2.50, 1.42
    lab = rect(sl, M, vy, lw, vh, BAND, None)
    txt(lab, lines=[[('Valor para...', True)]], size=13, color=BODY,
        bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    vx0, vgap = M + lw + 0.10, 0.18
    vw = (CW - lw - 0.10 - 2 * vgap) / 3.0
    for i, v in enumerate(VALOR):
        x = vx0 + i * (vw + vgap)
        rect(sl, x, vy, vw, vh, CARD, None)
        icon(sl, v['ic'], x + 0.24, vy + 0.22, 0.36, RED, CARD)
        txt(sl, x + 0.74, vy + 0.20, vw - 0.94, 0.34, [(v['tit'], True)],
            size=12.5, color=BLACK, bold=True)
        txt(sl, x + 0.74, vy + 0.68, vw - 0.94, 0.70, v['txt'], size=T_BODY,
            lsp=1.20)

    dy = 4.16
    hline(sl, M, dy, CW, RED, 0.022)
    shp(sl, MSO_SHAPE.LEFT_ARROW, M - 0.06, dy - 0.085, 0.56, 0.19, RED, None,
        adj=(0.50, 0.42))
    lbl = 'cinco frentes, escaladas em ondas'
    lwd = text_w(lbl, 12.5, True) + 0.26
    rect(sl, M + 0.66, dy - 0.15, lwd, 0.30, WHITE, None)
    txt(sl, M + 0.79, dy - 0.135, lwd, 0.28, [(lbl, True)], size=12.5,
        color=RED, bold=True)

    fy, fw = 4.46, CW / 5.0
    for i, f in enumerate(FRENTES):
        x = M + i * fw
        if i:
            vline(sl, x - 0.005, fy + 0.06, 2.30, RULE)
        numbubble(sl, x + 0.14, fy, 0.40, f['n'], 14)
        icon(sl, f['ic'], x + 0.16, fy + 0.62, 0.38, RED, WHITE)
        txt(sl, x + 0.14, fy + 1.14, fw - 0.36, 0.66,
            [[(l, True)] for l in f['tit'].split('\n')],
            size=12.5, color=BLACK, bold=True, lsp=1.12)
        ty = fy + 1.62
        txt(sl, x + 0.14, ty, fw - 0.36, 1.1, f['txt'], size=T_BODY, lsp=1.20)
        fits('p2 frente %d' % f['n'],
             nlines(f['txt'], T_BODY, fw - 0.36) * T_BODY * 1.20 / 72.0,
             6.86 - ty)

    footer(sl, SRC_F1, 2)
    return sl


# ------------------------------------------------------------------ pagina 3
def page_fi_plano(prs):
    sl = blank(prs)
    eyebrow(sl, 'F&I multibanco  |  Plano de execução')
    title(sl, 'As cinco frentes têm entregável, ponto de partida, destino',
          'e prazo explícitos')

    cols = [0.66, 1.86, 2.60, 2.60, 3.05, 1.56]
    heads = ['#', 'Frente', 'Entregável', 'Onde estamos',
             'Onde queremos chegar', 'Quando']
    ty, hh, rh = 1.70, 0.56, 0.86
    gf = sl.shapes.add_table(len(PLANO) + 1, len(cols), Inches(M), Inches(ty),
                             Inches(sum(cols)), Inches(hh + rh * len(PLANO)))
    tbl = gf.table
    tbl.first_row = False
    tbl.horz_banding = False
    tblPr = tbl._tbl.tblPr
    for e in tblPr.findall(qn('a:tableStyleId')):
        tblPr.remove(e)
    sid = tblPr.makeelement(qn('a:tableStyleId'), {})
    sid.text = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'   # No Style, No Grid
    tblPr.append(sid)
    for i, w in enumerate(cols):
        tbl.columns[i].width = Emu(int(round(w * 914400)))
    tbl.rows[0].height = Emu(int(round(hh * 914400)))
    for r in range(1, len(PLANO) + 1):
        tbl.rows[r].height = Emu(int(round(rh * 914400)))

    def cell(r, c, parts, size, color, bold=False, align=PP_ALIGN.LEFT,
             fill=None):
        cl = tbl.cell(r, c)
        cl.margin_left = cl.margin_right = Inches(0.16)
        cl.margin_top = cl.margin_bottom = Inches(0.06)
        cl.vertical_anchor = MSO_ANCHOR.MIDDLE
        if fill is None:
            cl.fill.background()
        else:
            cl.fill.solid()
            cl.fill.fore_color.rgb = fill
        tf = cl.text_frame
        tf.word_wrap = True
        tf.clear()
        for i, line in enumerate(parts.split('\n')):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.line_spacing = 1.16
            _run(p, line, size, color, bold)
        return cl

    for c, h in enumerate(heads):
        cell(0, c, h, 11.5, BODY, True,
             PP_ALIGN.CENTER if c in (0, 5) else PP_ALIGN.LEFT, THEAD)
    for r, row in enumerate(PLANO, start=1):
        zebra = WHITE if r % 2 else TZEBRA
        for c, v in enumerate(row):
            if c == 0:
                cell(r, c, '', 11, BODY, False, PP_ALIGN.CENTER, zebra)
            elif c == 5:
                cell(r, c, v, 11, BODY, False, PP_ALIGN.CENTER, TCOL)
            else:
                cell(r, c, v, 11, BLACK if c == 1 else BODY, c == 1,
                     PP_ALIGN.LEFT, zebra)
        hline(sl, M, ty + hh + rh * (r - 1), sum(cols), RULE, 0.008)
        numbubble(sl, M + cols[0] / 2 - 0.19,
                  ty + hh + rh * (r - 1) + rh / 2 - 0.19, 0.38, row[0], 13.5)
    hline(sl, M, ty + hh + rh * len(PLANO), sum(cols), RULE, 0.008)

    footer(sl, SRC_F2, 3)
    return sl


# ------------------------------------------------------------------ pagina 4
def page_lead_quente_v2(prs):
    sl = blank(prs)
    eyebrow(sl, 'L2S  |  Lead quente')
    title(sl, *TIT_LQ)
    txt(sl, M, 1.46, 11.6, 0.5,
        'Pré-requisito: o número existe e é nominal: leads quentes '
        'recebidos, conversão e tempo até o primeiro contato por divisional, '
        'loja e CV', size=T_SUB, color=SUB, lsp=1.20)

    cy, ch, gap = 1.90, 4.44, 0.40
    cwid = (CW - 3 * gap) / 4.0
    hb = 0.88                        # faixa de cabecalho do card
    for i, p in enumerate(PASSOS):
        x = M + i * (cwid + gap)
        rect(sl, x, cy + 0.18, cwid, hb, BAND, None)
        icon(sl, p['ic'], x + cwid / 2 - 0.25, cy + 0.37, 0.50, RED, BAND)
        numbubble(sl, x - 0.11, cy, 0.46, p['n'], 16)
        txt(sl, x + 0.10, cy + 1.20, cwid - 0.20, 0.70,
            [[(l, True)] for l in p['tit'].split('\n')],
            size=13.5, color=BLACK, bold=True, lsp=1.10)
        by = cy + ch - 0.74
        need = bullets(sl, x + 0.10, cy + 1.90, cwid - 0.20, p['bullets'],
                       size=9.5, gap=0.14)
        fits('p4 card %d' % p['n'], need, by - 0.26 - (cy + 1.90))
        txt(sl, x + 0.10, by, cwid - 0.20, 0.7,
            [[('Dono: ', True), (p['dono'], False)],
             [('Prazo: ', True), (p['prazo'], False)]],
            size=9.5, lsp=1.28, space_before=0.06)
        if i < 3:
            shp(sl, MSO_SHAPE.CHEVRON, x + cwid + 0.09, cy + 0.48, 0.24, 0.28,
                GREY4, None, adj=(0.50,))

    fy = 6.46
    rect(sl, M, fy, CW, 0.50, RED_BG, None)
    icon(sl, 'trofeu', M + 0.18, fy + 0.09, 0.32, RED, RED_BG)
    txt(sl, M + 0.66, fy, CW - 0.80, 0.50, [(MERITO[0] + MERITO[1], True)],
        size=12, color=RED, bold=True, anchor=MSO_ANCHOR.MIDDLE)

    footer(sl, SRC_LQ, 1, logo=True)
    return sl


# ------------------------------------------------------------------ pagina 5
def page_fi_visao_v2(prs):
    sl = blank(prs)
    txt(sl, M, 0.28, CW, 0.22, [('F&I MULTIBANCO', False)], size=T_EYE,
        color=RED, spc=180)
    hline(sl, M, 0.54, CW, RED_LN, 0.012)
    title(sl, *TIT_FI, y=0.80)

    txt(sl, M, 1.62, 4.0, 0.28, [('Ambição', True)], size=12.5, color=RED,
        bold=True)
    ay, ah = 1.92, 0.78
    rect(sl, M, ay, CW, ah, BAND, None)
    icon(sl, 'alvo', M + 0.26, ay + 0.17, 0.44, RED, BAND)
    txt(sl, M + 0.94, ay, CW - 1.16, ah, [(AMBICAO, True)], size=12.5,
        color=BODY, bold=True, anchor=MSO_ANCHOR.MIDDLE, lsp=1.22)

    vy, vh, vgap = 2.88, 1.20, 0.20
    vw = (CW - 2 * vgap) / 3.0
    for i, v in enumerate(VALOR):
        x = M + i * (vw + vgap)
        rect(sl, x, vy, vw, vh, CARD, None)
        icon(sl, v['ic'], x + 0.24, vy + 0.18, 0.34, RED, CARD)
        icon_w = 0.24 + 0.34 + 0.16
        txt(sl, x + icon_w, vy + 0.18, vw - icon_w - 0.20, 0.32,
            [(v['tit'], True)], size=12.5, color=BLACK, bold=True)
        txt(sl, x + 0.24, vy + 0.66, vw - 0.48, 0.50, v['txt'], size=T_BODY,
            lsp=1.20)

    txt(sl, M, 4.24, 6.0, 0.28,
        [('Como: ', True), ('cinco frentes, escaladas em ondas', True)],
        size=12.5, color=RED, bold=True)

    fy, fgap = 4.58, 0.18
    fw = (CW - 4 * fgap) / 5.0
    for i, f in enumerate(FRENTES):
        x = M + i * (fw + fgap)
        rect(sl, x, fy, fw, 0.50, BAND, None)
        numbubble(sl, x + 0.10, fy + 0.07, 0.36, f['n'], 12.5)
        txt(sl, x + 0.54, fy, fw - 0.64, 0.50,
            [(f['tit'].replace('\n', ' '), True)], size=11.5, color=BLACK,
            bold=True, anchor=MSO_ANCHOR.MIDDLE, lsp=1.05)
        rect(sl, x, fy + 0.50, fw, 0.76, BAND2, None)
        icon(sl, f['ic'], x + fw / 2 - 0.21, fy + 0.67, 0.42, RED, BAND2)
        ty = fy + 1.40
        txt(sl, x + 0.02, ty, fw - 0.06, 1.1, f['txt'], size=10, lsp=1.20)
        fits('p5 frente %d' % f['n'],
             nlines(f['txt'], 10, fw - 0.06) * 10 * 1.20 / 72.0, 6.86 - ty)

    footer(sl, SRC_F1, 2, logo=True)
    return sl


# ------------------------------------------------------------------ main
def main(dst):
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    page_lead_quente(prs)
    page_fi_visao(prs)
    page_fi_plano(prs)
    page_lead_quente_v2(prs)
    page_fi_visao_v2(prs)
    cp = prs.core_properties
    cp.title = 'Lead quente e F&I multibanco'
    cp.author = 'ConsultingInnovation@bain.com'
    cp.last_modified_by = 'ConsultingInnovation@bain.com'
    prs.save(dst)
    print('ok %s (%d paginas)' % (dst, len(prs.slides._sldIdLst)))
    for w in _WARN:
        print('  ATENCAO transbordo -> %s' % w)


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1
         else '../output/lead_quente_fi_multibanco.pptx')
