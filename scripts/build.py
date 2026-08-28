# -*- coding: utf-8 -*-
"""Gera o PPT (6 paginas, uma por frente) replicando o layout do painel HTML."""
import json, copy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from PIL import ImageFont

# ---------- tokens Localiza (do :root do HTML) ----------
GREEN        = RGBColor(0x78, 0xDE, 0x1F)
GREEN_DEEP   = RGBColor(0x3F, 0x7C, 0x0E)
GREEN_TINT   = RGBColor(0xEF, 0xFA, 0xE3)
BLACK        = RGBColor(0x1A, 0x1A, 0x1A)
TEXT_DARK    = RGBColor(0x33, 0x33, 0x33)
TEXT_REG     = RGBColor(0x6B, 0x6B, 0x6B)
TEXT_HELPER  = RGBColor(0x75, 0x75, 0x75)
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
GRAY_X_LIGHT = RGBColor(0xF7, 0xF7, 0xF7)
GRAY_LIGHT   = RGBColor(0xEC, 0xEC, 0xEC)
GRAY_MEDIUM  = RGBColor(0xD5, 0xD5, 0xD5)
STROKE       = RGBColor(0xE2, 0xE2, 0xE2)
PLACEHOLDER  = RGBColor(0x82, 0x82, 0x82)

STATUS = {
    'verde':    ('Concluido',    RGBColor(0xE6,0xF7,0xEE), RGBColor(0x14,0x6A,0x3D)),
    'amarelo':  ('Em progresso', RGBColor(0xFF,0xF8,0xE1), RGBColor(0x6B,0x57,0x00)),
    'vermelho': ('Nao iniciado', RGBColor(0xFC,0xEA,0xEA), RGBColor(0xA3,0x00,0x00)),
}
STATUS_LABEL = {'verde': 'Concluído', 'amarelo': 'Em progresso', 'vermelho': 'Não iniciado'}

FONT_NAME = 'Arial'
_TTF = {False: '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
        True:  '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf'}
_fc = {}
def _f(bold):
    if bold not in _fc:
        _fc[bold] = ImageFont.truetype(_TTF[bold], 200)
    return _fc[bold]

def text_w(s, pt, bold=False):
    """largura do texto em polegadas"""
    return _f(bold).getlength(s) / 200.0 * pt / 72.0

def wrap(s, pt, width_in, bold=False):
    """quebra em linhas que cabem em width_in"""
    words, lines, cur = s.split(), [], ''
    for w in words:
        cand = (cur + ' ' + w).strip()
        if cur and text_w(cand, pt, bold) > width_in:
            lines.append(cur); cur = w
        else:
            cur = cand
    if cur:
        lines.append(cur)
    return lines or ['']

def nlines(s, pt, width_in, bold=False):
    return len(wrap(s, pt, width_in, bold))

def lh(pt):
    return pt * 1.22 / 72.0   # altura de linha em polegadas

# ---------- helpers de desenho ----------
def textbox(slide, x, y, w, h, runs, size=8, color=TEXT_DARK, bold=False,
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.22, caps=False):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    if isinstance(runs, str):
        runs = [(runs, {})]
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = spacing
    for txt, ov in runs:
        r = p.add_run()
        r.text = txt.upper() if caps else txt
        fnt = r.font
        fnt.name = FONT_NAME
        fnt.size = Pt(ov.get('size', size))
        fnt.bold = ov.get('bold', bold)
        fnt.italic = ov.get('italic', False)
        fnt.color.rgb = ov.get('color', color)
        if caps or ov.get('caps'):
            r.text = r.text.upper()
    return tb

def no_shadow(shape):
    """remove a sombra herdada do tema (p:style -> effectRef)"""
    el = shape._element
    st = el.find(qn('p:style'))
    if st is not None:
        el.remove(st)
    spPr = el.spPr
    for e in spPr.findall(qn('a:effectLst')):
        spPr.remove(e)
    spPr.append(spPr.makeelement(qn('a:effectLst'), {}))


def rect(slide, x, y, w, h, fill=None, line=None, lw=0.75, shape=MSO_SHAPE.RECTANGLE, adj=None):
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
        except Exception: pass
    tf = s.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return s

def hline(slide, x, y, w, color, lw=0.75):
    ln = slide.shapes.add_connector(1, Inches(x), Inches(y), Inches(x + w), Inches(y))
    ln.line.color.rgb = color
    ln.line.width = Pt(lw)
    no_shadow(ln)
    return ln

# ---------- geometria da pagina ----------
SW, SH = 13.3333, 7.5
ML = 0.42
CW = SW - 2 * ML          # 12.4933
TOP = 0.30
BOTTOM = 7.5 - 0.34

COLS = [
    ('#',                  0.70, PP_ALIGN.LEFT),
    ('Milestone',          3.10, PP_ALIGN.LEFT),
    ('Responsável',        1.36, PP_ALIGN.LEFT),
    ('Prazo',              0.86, PP_ALIGN.LEFT),
    ('Status',             0.95, PP_ALIGN.CENTER),
    ('Progresso recente',  1.84, PP_ALIGN.LEFT),
    ('Próximos passos',    1.84, PP_ALIGN.LEFT),
    ('Pontos a escalar',   1.84, PP_ALIGN.LEFT),
]
_scale = CW / sum(c[1] for c in COLS)
COLS = [(n, w * _scale, a) for n, w, a in COLS]
CX = []
_x = ML
for _n, _w, _a in COLS:
    CX.append(_x); _x += _w

PADX = 0.075     # padding horizontal da celula
PADY = 0.055     # padding vertical da celula
LEAD = 8.0       # pt do cabecalho da tabela

# ---------- construcao ----------
frentes = json.load(open('frentes.json'))

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(SW), Inches(SH)
blank = prs.slide_layouts[6]

for f in frentes:
    slide = prs.slides.add_slide(blank)

    # fundo
    rect(slide, 0, 0, SW, SH, fill=WHITE)

    # ---- kicker ----
    textbox(slide, ML, TOP, CW * 0.7, 0.16,
            [('Lead-to-Sales  ·  Painel de Controle', {'bold': True, 'color': BLACK}),
             ('   |   Aceleração Seminovos', {'color': TEXT_HELPER})], size=7.5)
    textbox(slide, ML + CW - 2.2, TOP, 2.2, 0.16,
            'Frente %d de 6' % f['id'], size=7.5, color=TEXT_HELPER, align=PP_ALIGN.RIGHT)

    y = TOP + 0.26

    # ================= HEADER DA FRENTE =================
    HPAD_L, HPAD_R, HPAD_T, HPAD_B = 0.26, 0.20, 0.16, 0.16
    BAR = 0.075
    BADGE = 0.28

    right_w_lid, right_w_time, right_gap = 1.72, 2.42, 0.22
    right_total = right_w_lid + right_w_time + right_gap
    x_right = ML + CW - HPAD_R - right_total
    x_txt = ML + BAR + HPAD_L
    left_w = x_right - x_txt - 0.30

    tit_pt, obj_pt = 12.5, 8.5
    text_w_col = left_w - BADGE - 0.14
    tit_lines = wrap(f['nome'], tit_pt, text_w_col, True)
    obj_lines = wrap(f['objetivo'], obj_pt, text_w_col)
    h_left = max(BADGE, len(tit_lines) * lh(tit_pt)) + 0.07 + len(obj_lines) * lh(obj_pt)

    lab_pt, val_pt = 6.5, 8.0
    n_lid = nlines(f['lideres'], val_pt, right_w_lid)
    n_tim = nlines(f['time'], val_pt, right_w_time)
    h_right = lh(lab_pt) + 0.05 + max(n_lid, n_tim) * lh(val_pt)

    hdr_h = max(h_left, h_right) + HPAD_T + HPAD_B

    rect(slide, ML, y, CW, hdr_h, fill=GRAY_X_LIGHT, line=STROKE)
    rect(slide, ML, y, BAR, hdr_h, fill=GREEN)

    # badge numerada
    by = y + HPAD_T + (max(BADGE, len(tit_lines) * lh(tit_pt)) - BADGE) / 2.0
    rect(slide, x_txt, by, BADGE, BADGE, fill=GREEN, shape=MSO_SHAPE.OVAL)
    textbox(slide, x_txt, by + 0.045, BADGE, 0.20, str(f['id']),
            size=11, color=BLACK, bold=True, align=PP_ALIGN.CENTER)

    tx = x_txt + BADGE + 0.14
    textbox(slide, tx, y + HPAD_T, text_w_col,
            len(tit_lines) * lh(tit_pt), f['nome'],
            size=tit_pt, color=BLACK, bold=True)
    y_obj = y + HPAD_T + max(BADGE, len(tit_lines) * lh(tit_pt)) + 0.07
    textbox(slide, tx, y_obj, text_w_col,
            len(obj_lines) * lh(obj_pt), f['objetivo'], size=obj_pt, color=TEXT_REG)

    for xoff, wcol, lab, val in (
            (0, right_w_lid, 'Líder(es)', f['lideres']),
            (right_w_lid + right_gap, right_w_time, 'Time de trabalho', f['time'])):
        xx = x_right + xoff
        textbox(slide, xx, y + HPAD_T, wcol, lh(lab_pt), lab,
                size=lab_pt, color=TEXT_HELPER, bold=True, caps=True)
        textbox(slide, xx, y + HPAD_T + lh(lab_pt) + 0.05, wcol,
                max(n_lid, n_tim) * lh(val_pt), val, size=val_pt, color=TEXT_DARK)

    y += hdr_h + 0.20

    # ================= TABELA DE MILESTONES =================
    textbox(slide, ML, y, 3.0, 0.17, 'Milestones', size=9.5, color=BLACK, bold=True)
    y += 0.24

    rows = f['entregaveis']
    avail = BOTTOM - y

    # escolhe o corpo de fonte que cabe na pagina
    for body_pt in (8.0, 7.5, 7.0, 6.5, 6.0, 5.5):
        id_pt   = body_pt - 0.5
        head_h  = lh(LEAD) + 0.13
        heights = []
        for r in rows:
            n = max(nlines(r['id'],    id_pt,   COLS[0][1] - 2 * PADX, True),
                    nlines(r['desc'],  body_pt, COLS[1][1] - 2 * PADX),
                    nlines(r['owner'], body_pt, COLS[2][1] - 2 * PADX),
                    nlines(r['prazo'], body_pt, COLS[3][1] - 2 * PADX))
            heights.append(max(0.355, n * lh(body_pt) + 2 * PADY + 0.055))
        total = head_h + sum(heights)
        if total <= avail:
            break

    # distribui o espaco que sobra entre as linhas (limite: 2x a altura natural)
    caps = [max(h, 1.50) for h in heights]
    for _ in range(40):
        slack = avail - (head_h + sum(heights))
        movable = [i for i, h in enumerate(heights) if h < caps[i] - 1e-4]
        if slack <= 0.01 or not movable:
            break
        step = slack / len(movable)
        for i in movable:
            heights[i] = min(caps[i], heights[i] + step)

    # cabecalho
    for i, (name, wcol, align) in enumerate(COLS):
        textbox(slide, CX[i] + PADX, y + 0.015, wcol - 2 * PADX, lh(LEAD),
                name, size=LEAD, color=TEXT_HELPER, bold=True, align=align)
    yy = y + head_h
    hline(slide, ML, yy, CW, BLACK, 1.0)

    for r, rh in zip(rows, heights):
        # id
        textbox(slide, CX[0] + PADX, yy + PADY, COLS[0][1] - 2 * PADX, rh - 2 * PADY,
                r['id'], size=id_pt, color=TEXT_HELPER, bold=True)
        # milestone / responsavel / prazo
        for i, key, col in ((1, 'desc', TEXT_DARK), (2, 'owner', TEXT_DARK), (3, 'prazo', TEXT_REG)):
            textbox(slide, CX[i] + PADX, yy + PADY, COLS[i][1] - 2 * PADX, rh - 2 * PADY,
                    r[key], size=body_pt, color=col)
        # pill de status
        lbl = STATUS_LABEL[r['status']]
        _, bg, fg = STATUS[r['status']]
        pw = min(COLS[4][1] - 2 * PADX, text_w(lbl, body_pt - 0.5, True) + 0.22)
        ph = 0.185
        px = CX[4] + (COLS[4][1] - pw) / 2.0
        rect(slide, px, yy + PADY + 0.012, pw, ph, fill=bg,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.5)
        textbox(slide, px, yy + PADY + 0.012, pw, ph, lbl, size=body_pt - 0.5,
                color=fg, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # caixas para preencher (TBD)
        for i in (5, 6, 7):
            bx, bw = CX[i] + PADX * 0.6, COLS[i][1] - 1.2 * PADX
            bh = rh - 2 * PADY + 0.02
            rect(slide, bx, yy + PADY - 0.01, bw, bh, fill=WHITE, line=GRAY_MEDIUM, lw=0.5)
            textbox(slide, bx + 0.06, yy + PADY + 0.02, bw - 0.12, bh - 0.06, 'TBD',
                    size=body_pt, color=PLACEHOLDER, bold=False)
        yy += rh
        last = r is rows[-1]
        hline(slide, ML, yy, CW, GRAY_MEDIUM if last else GRAY_LIGHT, 0.5)

    # ---- rodape ----
    hline(slide, ML, SH - 0.30, CW, STROKE, 0.5)
    textbox(slide, ML, SH - 0.255, CW * 0.8,
            0.14, 'Fonte: Painel de Controle Lead-to-Sales — Aceleração Seminovos',
            size=6.5, color=TEXT_HELPER)
    textbox(slide, ML + CW - 1.0, SH - 0.255, 1.0, 0.14, str(f['id']),
            size=6.5, color=TEXT_HELPER, align=PP_ALIGN.RIGHT)

prs.core_properties.title = 'Lead-to-Sales — Frentes e Milestones'
prs.save('frentes_milestones.pptx')
print('ok', len(prs.slides.__iter__.__self__._sldIdLst))
