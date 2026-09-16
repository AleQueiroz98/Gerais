# -*- coding: utf-8 -*-
"""Pagina 'Jornada do lead: qualificacao, alocacao e priorizacao' (redesign).

Reformula, em um layout mais bonito e com icones vetoriais, o fluxograma de
3 etapas (qualificar -> alocar -> priorizar/gerir) do tratamento de leads,
mantendo os 4 callouts numerados originais. Segue as convencoes visuais ja
usadas nos outros builds do repo (Arial, pt-BR, fundo branco, vermelho Bain
como cor de destaque, sem sombra padrao do PowerPoint).
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

# ---------------------------------------------------------------- tokens
RED    = RGBColor(0xCC, 0x00, 0x00)   # Bain red - destaque / callouts
RED_BG = RGBColor(0xFC, 0xEA, 0xEA)
GREEN  = RGBColor(0x10, 0x4C, 0x3E)   # Forest 1 - etapa 3 / positivo
GREEN_BG = RGBColor(0xE9, 0xF3, 0xEE)
INK1   = RGBColor(0x1A, 0x1A, 0x1A)   # etapa 1 (header escuro)
INK2   = RGBColor(0x3B, 0x3B, 0x3B)   # etapa 2 (header medio)
GREY1  = RGBColor(0x33, 0x33, 0x33)
GREY2  = RGBColor(0x5C, 0x5C, 0x5C)
GREY3  = RGBColor(0x85, 0x85, 0x85)
GREY4  = RGBColor(0xB4, 0xB4, 0xB4)
LINE   = RGBColor(0xE0, 0xE0, 0xE0)
CARDBG = RGBColor(0xF7, 0xF8, 0xF8)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "Arial"
LANG = "pt-BR"

SW, SH = 13.333, 7.5

TITLE = "Tratamos, alocamos e priorizamos melhor os leads ao longo de toda a jornada"
SUB = ("Do primeiro contato ao atendimento, cada etapa aumenta a chance de o lead certo "
       "chegar ao vendedor certo, no momento certo")

COL_X = [0.34, 4.68, 9.02]
COL_W = 3.98
HEAD_Y, HEAD_H = 1.62, 0.56
CARD_Y, CARD_H = 2.30, 4.66
GAP_Y = 6.86

# ---------------------------------------------------------------- helpers
def _lang(run):
    run.font._rPr.set('lang', LANG)


def txt(slide, x, y, w, h, parts, size=10, color=GREY1, bold=False,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line=None, wrap=True):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
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
        _lang(r)
    return tb


def shape(slide, sh, x, y, w, h, fill=None, line_color=None, line_w=0.75, rot=0):
    s = slide.shapes.add_shape(sh, Inches(x), Inches(y), Inches(w), Inches(h))
    s.shadow.inherit = False
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line_color is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line_color
        s.line.width = Pt(line_w)
    if rot:
        s.rotation = rot
    s.text_frame.word_wrap = False
    return s


def soft_shadow(sh, blur=0.06, dist=0.03, alpha=78):
    """sombra suave (substitui a sombra padrao, mais discreta)."""
    sp = sh._element.spPr
    for e in sp.findall(qn('a:effectLst')):
        sp.remove(e)
    eff = sp.makeelement(qn('a:effectLst'), {})
    sh_el = eff.makeelement(qn('a:outerShdw'), {
        'blurRad': str(int(blur * 914400)), 'dist': str(int(dist * 914400)),
        'dir': '5400000', 'rotWithShape': '0'})
    clr = sh_el.makeelement(qn('a:srgbClr'), {'val': '1A1A1A'})
    alpha_el = clr.makeelement(qn('a:alpha'), {'val': str(alpha * 1000)})
    clr.append(alpha_el)
    sh_el.append(clr)
    eff.append(sh_el)
    sp.append(eff)


def hline(slide, x1, y, x2, color=GREY4, w=0.75, dash=None):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y),
                                    Inches(x2), Inches(y))
    c.line.color.rgb = color
    c.line.width = Pt(w)
    if dash:
        ln = c.line._get_or_add_ln()
        d = ln.makeelement(qn('a:prstDash'), {'val': dash})
        ln.append(d)
    return c


def vline(slide, x, y1, y2, color=GREY4, w=0.75, dash=None):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y1),
                                    Inches(x), Inches(y2))
    c.line.color.rgb = color
    c.line.width = Pt(w)
    if dash:
        ln = c.line._get_or_add_ln()
        d = ln.makeelement(qn('a:prstDash'), {'val': dash})
        ln.append(d)
    return c


def arrow(slide, x1, y1, x2, y2, color=GREY3, w=1.25, dash=None):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1),
                                    Inches(x2), Inches(y2))
    c.line.color.rgb = color
    c.line.width = Pt(w)
    ln = c.line._get_or_add_ln()
    ln.append(ln.makeelement(qn('a:tailEnd'), {'type': 'triangle', 'w': 'sm', 'len': 'sm'}))
    if dash:
        d = ln.makeelement(qn('a:prstDash'), {'val': dash})
        ln.append(d)
    return c


def badge(slide, cx, cy, d, number, fill=RED, text_color=WHITE):
    s = shape(slide, MSO_SHAPE.OVAL, cx - d / 2, cy - d / 2, d, d, fill=fill)
    soft_shadow(s, blur=0.05, dist=0.02, alpha=55)
    txt(slide, cx - d / 2, cy - d / 2, d, d, str(number), size=12, color=text_color,
        bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return s


def callout(slide, x, y, w, h, number, text, accent=RED, accent_bg=RED_BG):
    box = shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=accent_bg)
    box.adjustments[0] = 0.10
    bar = shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, 0.07, h, fill=accent)
    bar.adjustments[0] = 0.5
    badge(slide, x + 0.30, y + h / 2, 0.30, number, fill=accent)
    txt(slide, x + 0.56, y + 0.04, w - 0.70, h - 0.08, text, size=10.5, color=GREY1,
        bold=True, anchor=MSO_ANCHOR.MIDDLE, line=1.05)


def header(slide, x, y, w, h, label, fill, last=False):
    sh_type = MSO_SHAPE.ROUNDED_RECTANGLE if last else MSO_SHAPE.PENTAGON
    tip = 0.0 if last else 0.30
    s = shape(slide, sh_type, x, y, w + (0 if last else tip), h, fill=fill)
    if last:
        s.adjustments[0] = 0.10
    soft_shadow(s, blur=0.05, dist=0.025, alpha=45)
    txt(slide, x + 0.22, y, w - 0.40, h, label, size=13, color=WHITE, bold=True,
        anchor=MSO_ANCHOR.MIDDLE, line=0.98)
    return s


def card(slide, x, y, w, h):
    s = shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=WHITE,
              line_color=LINE, line_w=1.0)
    s.adjustments[0] = 0.035
    soft_shadow(s, blur=0.09, dist=0.035, alpha=32)
    return s


def stage_num(slide, x, y, n):
    s = shape(slide, MSO_SHAPE.OVAL, x, y, 0.30, 0.30, fill=WHITE)
    txt(slide, x, y, 0.30, 0.30, str(n), size=11, color=INK1, bold=True,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ---------------------------------------------------------------- icones
def icon_person(slide, cx, top, size, color=GREY1):
    """circulo (cabeca) + trapezio (corpo), pictograma generico de pessoa."""
    hd = size * 0.36
    shape(slide, MSO_SHAPE.OVAL, cx - hd / 2, top, hd, hd, fill=color)
    bw_top, bw_bot, bh = size * 0.34, size * 0.62, size * 0.50
    body = shape(slide, MSO_SHAPE.TRAPEZOID, cx - bw_bot / 2, top + hd + size * 0.06,
                 bw_bot, bh, fill=color)
    body.rotation = 180
    body.adjustments[0] = 1 - (bw_top / bw_bot)


def icon_store(slide, cx, top, size, color=GREY1, highlight=False):
    """telhado triangular + fachada retangular + porta, pictograma de loja."""
    fill = RED if highlight else color
    roof_w, roof_h = size * 0.92, size * 0.34
    shape(slide, MSO_SHAPE.ISOSCELES_TRIANGLE, cx - roof_w / 2, top, roof_w, roof_h, fill=fill)
    body_w, body_h = size * 0.74, size * 0.50
    body_y = top + roof_h - 0.02
    shape(slide, MSO_SHAPE.RECTANGLE, cx - body_w / 2, body_y, body_w, body_h, fill=fill)
    door_w, door_h = size * 0.20, size * 0.30
    shape(slide, MSO_SHAPE.RECTANGLE, cx - door_w / 2, body_y + body_h - door_h,
          door_w, door_h, fill=WHITE)


def icon_gear(slide, cx, top, size, color=GREY1):
    shape(slide, MSO_SHAPE.GEAR_9, cx - size / 2, top, size, size, fill=color)


def icon_robot(slide, cx, top, size, color=GREEN):
    """cabeca arredondada + antena + olhos, pictograma da Liza (IA)."""
    ant_h = size * 0.14
    vline(slide, cx, top - ant_h, top, color=color, w=1.5)
    shape(slide, MSO_SHAPE.OVAL, cx - size * 0.06, top - ant_h - size * 0.12,
          size * 0.12, size * 0.12, fill=color)
    head = shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, cx - size / 2, top, size, size * 0.78,
                 fill=color)
    head.adjustments[0] = 0.32
    eye_w, eye_h = size * 0.14, size * 0.20
    eye_y = top + size * 0.28
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, cx - size * 0.24, eye_y, eye_w, eye_h, fill=WHITE)
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, cx + size * 0.10, eye_y, eye_w, eye_h, fill=WHITE)


def info_tag(slide, x, y, w, h, label, color=GREEN):
    tag = shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=CARDBG)
    tag.adjustments[0] = 0.5
    dot = shape(slide, MSO_SHAPE.OVAL, x + 0.10, y + h / 2 - 0.045, 0.09, 0.09, fill=color)
    txt(slide, x + 0.26, y, w - 0.32, h, label, size=8.5, color=GREY1,
        anchor=MSO_ANCHOR.MIDDLE)


def mini_list(slide, x, y, w, h, title, title_color, rows):
    """rows: list of (texto, destacado?)"""
    s = shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=WHITE, line_color=LINE, line_w=0.75)
    s.adjustments[0] = 0.06
    hd = shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, 0.30, fill=title_color)
    hd.adjustments[0] = 0.30
    shape(slide, MSO_SHAPE.RECTANGLE, x, y + 0.14, w, 0.16, fill=title_color)
    txt(slide, x, y, w, 0.30, title, size=8.5, color=WHITE, bold=True,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    ry = y + 0.40
    rh = (h - 0.46) / max(len(rows), 1)
    for text, hot in rows:
        col = RED if hot else GREY2
        dot = shape(slide, MSO_SHAPE.OVAL, x + 0.12, ry + rh / 2 - 0.035, 0.07, 0.07,
                    fill=col if hot else GREY4)
        txt(slide, x + 0.26, ry, w - 0.32, rh, text, size=8, color=col,
            bold=hot, anchor=MSO_ANCHOR.MIDDLE)
        ry += rh
    return s


# ---------------------------------------------------------------- build
prs = Presentation()
prs.slide_width, prs.slide_height = Inches(SW), Inches(SH)
slide = prs.slides.add_slide(prs.slide_layouts[6])

# fundo
bg = shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SW, SH, fill=WHITE)

# titulo + subtitulo + regua
txt(slide, 0.34, 0.30, 12.66, 0.44, TITLE, size=19, color=GREY1, bold=False, line=1.05)
txt(slide, 0.34, 0.78, 12.66, 0.34, SUB, size=11, color=GREY3, line=1.05)
hline(slide, 0.34, 1.30, SW - 0.34, RED, 2.25)

# ---- setas de fluxo entre as 3 etapas (acima dos headers) --------------
for i in range(2):
    xa = COL_X[i] + COL_W + 0.02
    xb = COL_X[i + 1] - 0.06
    arrow(slide, xa, HEAD_Y + HEAD_H / 2, xb, HEAD_Y + HEAD_H / 2, color=GREY4, w=1.5)

# ---- headers -------------------------------------------------------------
header(slide, COL_X[0], HEAD_Y, COL_W, HEAD_H, "1  Qualificar melhor o lead", INK1)
header(slide, COL_X[1], HEAD_Y, COL_W, HEAD_H, "2  Alocar melhor o lead", INK2)
header(slide, COL_X[2], HEAD_Y, COL_W, HEAD_H,
       "3  Melhorar gestão e priorização dos leads", GREEN, last=True)

# ---- cards de conteudo ----------------------------------------------------
for x in COL_X:
    card(slide, x, CARD_Y, COL_W, CARD_H)

PAD = 0.28

# === Etapa 1: qualificar ===================================================
x0 = COL_X[0]
icon_person(slide, x0 + 0.95, CARD_Y + 0.34, 0.95, color=INK1)
arrow(slide, x0 + 1.55, CARD_Y + 0.78, x0 + 1.95, CARD_Y + 0.78, color=GREY4, w=1.25)

labels = ["Info A", "Info B", "Info C", "Info D", "CPF", "CEP"]
tag_w, tag_h = 0.98, 0.30
tx0, ty0 = x0 + 2.00, CARD_Y + 0.30
for i, lb in enumerate(labels):
    r, c = divmod(i, 2)
    info_tag(slide, tx0 + c * (tag_w + 0.10), ty0 + r * (tag_h + 0.09), tag_w, tag_h, lb)

txt(slide, x0 + PAD, CARD_Y + 1.55, COL_W - 2 * PAD, 0.5,
    "Capturar o máximo de informações do cliente logo no primeiro contato",
    size=9.5, color=GREY2, align=PP_ALIGN.LEFT, line=1.15)

callout(slide, x0 + 0.18, CARD_Y + CARD_H - 1.02, COL_W - 0.36, 0.86, 1,
        "Obter informações completas dos clientes, para qualificar e priorizar com mais precisão")

# === Etapa 2: alocar ========================================================
x1 = COL_X[1]
gx, gy = x1 + 0.95, CARD_Y + 0.30
icon_gear(slide, gx, gy, 0.62, color=INK2)
icon_gear(slide, gx + 0.36, gy + 0.30, 0.38, color=GREY3)

stores_y = [CARD_Y + 0.30, CARD_Y + 0.95, CARD_Y + 1.60]
store_x = x1 + 2.55
for i, sy in enumerate(stores_y):
    hot = (i == 1)
    arrow(slide, gx + 0.30, gy + 0.55, store_x - 0.18, sy + 0.30,
          color=RED if hot else GREY4, w=1.6 if hot else 1.0)
    icon_store(slide, store_x, sy, 0.52, color=GREY2, highlight=hot)

txt(slide, x1 + PAD, CARD_Y + 2.05, COL_W - 2 * PAD, 0.5,
    "O motor de alocação direciona cada lead à loja com maior chance de conversão",
    size=9.5, color=GREY2, line=1.15)

callout(slide, x1 + 0.18, CARD_Y + CARD_H - 1.02, COL_W - 0.36, 0.86, 2,
        "Alocar o lead na loja mais adequada, considerando estoque, distância e histórico")

# === Etapa 3: priorizar/gerir ===============================================
x2 = COL_X[2]
persons = [("Vendedor 1", INK1), ("Vendedor 2", INK2), ("Liza", GREEN)]
pw = (COL_W - 2 * PAD - 2 * 0.14) / 3
py = CARD_Y + 0.26
for i, (label, col) in enumerate(persons):
    px_ = x2 + PAD + i * (pw + 0.14)
    cx = px_ + pw / 2
    if label == "Liza":
        icon_robot(slide, cx, py + 0.10, 0.40, color=col)
    else:
        icon_person(slide, cx, py, 0.40, color=col)
    txt(slide, px_, py + 0.50, pw, 0.20, label, size=8.5, color=col, bold=True,
        align=PP_ALIGN.CENTER)

rows1 = [("Lead 1", True), ("Lead 2", False), ("Lead 3", True)]
rows2 = [("Lead 1", True), ("Lead 2", False), ("Lead 3", False)]
rows3 = [("Lead 1", False), ("Lead 2", False), ("Lead 3", False)]
list_y = py + 0.78
list_h = 0.95
list_xs = [x2 + PAD + i * (pw + 0.14) for i in range(3)]
for i, rows in enumerate((rows1, rows2, rows3)):
    title_col = persons[i][1]
    mini_list(slide, list_xs[i], list_y, pw, list_h, "Leads", title_col, rows)

row_h = (list_h - 0.46) / 3
row1_cy = list_y + 0.40 + row_h * 0.5
row3_cy = list_y + 0.40 + row_h * 2.5

# lead quente redistribuido entre vendedores (linha "Lead 1", destacada nas 2 listas)
arrow(slide, list_xs[0] + pw, row1_cy, list_xs[1] + 0.02, row1_cy, color=RED, w=1.5)
# lead frio direcionado para a Liza (linha "Lead 3", base das listas 2 e 3)
arrow(slide, list_xs[1] + pw, row3_cy, list_xs[2] + 0.02, row3_cy, color=GREEN, w=1.5)

txt(slide, x2 + PAD, list_y + list_h + 0.10, COL_W - 2 * PAD, 0.32,
    "Leads quentes são redistribuídos entre vendedores; leads frios seguem para a Liza",
    size=8.5, color=GREY2, align=PP_ALIGN.LEFT, line=1.1)

CO_Y2 = CARD_Y + CARD_H - 0.72
CO_Y1 = CO_Y2 - 0.62
callout(slide, x2 + 0.18, CO_Y1, COL_W - 0.36, 0.56, 3,
        "Garantir atendimento dos leads quentes", accent=RED, accent_bg=RED_BG)
callout(slide, x2 + 0.18, CO_Y2, COL_W - 0.36, 0.56, 4,
        "Otimizar tempo do CV, alocando leads frios na Liza", accent=GREEN, accent_bg=GREEN_BG)

OUT = "/home/user/Gerais/output/jornada_lead_qualificacao_alocacao_priorizacao.pptx"
prs.save(OUT)
print("saved", OUT)
