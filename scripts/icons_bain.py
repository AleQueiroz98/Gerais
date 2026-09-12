# -*- coding: utf-8 -*-
"""Biblioteca de icones vetoriais em formas nativas do PowerPoint.

Cada icone e desenhado dentro de uma caixa quadrada de lado `s` (polegadas),
so com freeforms, elipses, retangulos e arcos: nada de imagem. Continua
editavel, recolorivel e redimensionavel no ppt.

    icon(slide, 'banco', x, y, 0.42, RED, bg=CARD)

`bg` e a cor do fundo em que o icone esta apoiado; alguns icones usam formas
nessa cor para abrir "furos" (o miolo da engrenagem, o vao das rodas).
"""
import math

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

WHITE = RGBColor(0xFF, 0xFF, 0xFF)
EMU = 914400


# ------------------------------------------------------------------ base
def _clean(shape, fill, line, lw):
    el = shape._element
    st = el.find(qn('p:style'))
    if st is not None:
        el.remove(st)
    spPr = el.spPr
    for e in spPr.findall(qn('a:effectLst')):
        spPr.remove(e)
    spPr.append(spPr.makeelement(qn('a:effectLst'), {}))
    if fill is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(lw)
        ln = spPr.find(qn('a:ln'))
        ln.set('cap', 'rnd')
        for t in ('a:round', 'a:bevel', 'a:miter'):
            for e in ln.findall(qn(t)):
                ln.remove(e)
        ln.append(ln.makeelement(qn('a:round'), {}))
    shape.text_frame.word_wrap = True
    return shape


def _shp(sl, kind, x, y, w, h, fill=None, line=None, lw=1.0, rot=0.0,
         adj=None):
    s = sl.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if adj:
        for i, v in enumerate(adj):
            try:
                s.adjustments[i] = v
            except (IndexError, ValueError):
                pass
    if rot:
        s.rotation = rot
    return _clean(s, fill, line, lw)


def _rect(sl, x, y, w, h, fill=None, line=None, lw=1.0, rot=0.0):
    return _shp(sl, MSO_SHAPE.RECTANGLE, x, y, w, h, fill, line, lw, rot)


def _ell(sl, x, y, w, h, fill=None, line=None, lw=1.0):
    return _shp(sl, MSO_SHAPE.OVAL, x, y, w, h, fill, line, lw)


def _ff(sl, pts, fill=None, line=None, lw=1.0):
    """freeform a partir de pontos em polegadas"""
    e = [(Emu(int(round(a * EMU))), Emu(int(round(b * EMU)))) for a, b in pts]
    bld = sl.shapes.build_freeform(e[0][0], e[0][1])
    bld.add_line_segments(e[1:], close=True)
    return _clean(bld.convert_to_shape(), fill, line, lw)


def _arc(sl, x, y, w, h, a0, a1, thick, fill):
    """arco em bloco: angulos em graus, 0 = leste, sentido horario"""
    s = _shp(sl, MSO_SHAPE.BLOCK_ARC, x, y, w, h, fill, None)
    av = s._element.spPr.find(qn('a:prstGeom')).find(qn('a:avLst'))
    for e in list(av):
        av.remove(e)
    for name, val in (('adj1', int(a0 * 60000)), ('adj2', int(a1 * 60000)),
                      ('adj3', int(thick * 100000))):
        g = av.makeelement(qn('a:gd'), {'name': name, 'fmla': 'val %d' % val})
        av.append(g)
    return s


def _seg(sl, x0, y0, x1, y1, thick, color):
    """segmento de reta como retangulo rotacionado"""
    dx, dy = x1 - x0, y1 - y0
    ln = math.hypot(dx, dy)
    ang = math.degrees(math.atan2(dy, dx))
    return _rect(sl, (x0 + x1) / 2 - ln / 2, (y0 + y1) / 2 - thick / 2,
                 ln, thick, color, None, rot=ang)


def _tri(sl, cx, cy, size, ang, color):
    """triangulo equilatero apontando para `ang` graus"""
    pts = []
    for a in (0, 140, 220):
        r = size if a == 0 else size * 0.62
        t = math.radians(ang + a)
        pts.append((cx + r * math.cos(t), cy + r * math.sin(t)))
    return _ff(sl, pts, color, None)


# ------------------------------------------------------------------ icones
def barras(sl, x, y, s, c, bg=WHITE):
    """grafico de barras cheio, ascendente"""
    for i, hf in enumerate((0.40, 0.68, 1.00)):
        _rect(sl, x + i * 0.37 * s, y + s * (1 - hf), s * 0.26, s * hf, c)


def pessoa(sl, x, y, s, c, bg=WHITE):
    """uma pessoa, contorno"""
    lw = s * 5.2
    _ell(sl, x + s * 0.29, y + s * 0.03, s * 0.42, s * 0.42, None, c, lw)
    _shp(sl, MSO_SHAPE.ROUND_2_SAME_RECTANGLE, x + s * 0.08, y + s * 0.56,
         s * 0.84, s * 0.44, None, c, lw, adj=(0.50,))


def pessoas(sl, x, y, s, c, bg=WHITE):
    """tres pessoas, contorno, a do meio a frente"""
    lw = s * 4.6
    for dx, dy, sc in ((0.00, 0.07, 0.84), (0.66, 0.07, 0.84), (0.33, 0.0, 1.0)):
        bw, bh = s * 0.34 * sc, s * 0.40 * sc
        bx, by = x + dx * s, y + (0.52 + dy) * s
        d = s * 0.28 * sc
        if sc == 1.0:            # halo do da frente, para separar
            _shp(sl, MSO_SHAPE.ROUND_2_SAME_RECTANGLE, bx - s * 0.035,
                 by - s * 0.035, bw + s * 0.07, bh + s * 0.07, bg, None,
                 adj=(0.50,))
            _ell(sl, bx + (bw - d) / 2 - s * 0.035, by - d - s * 0.02,
                 d + s * 0.07, d + s * 0.07, bg, None)
        _ell(sl, bx + (bw - d) / 2, by - d + s * 0.015, d, d, None, c, lw)
        _shp(sl, MSO_SHAPE.ROUND_2_SAME_RECTANGLE, bx, by, bw, bh, None, c,
             lw, adj=(0.50,))


def pessoas_cheio(sl, x, y, s, c, bg=WHITE):
    """tres pessoas, preenchidas"""
    for dx, dy, sc in ((0.00, 0.07, 0.84), (0.66, 0.07, 0.84), (0.33, 0.0, 1.0)):
        bw, bh = s * 0.34 * sc, s * 0.40 * sc
        bx, by = x + dx * s, y + (0.52 + dy) * s
        d = s * 0.28 * sc
        if sc == 1.0:
            g = s * 0.045
            _shp(sl, MSO_SHAPE.ROUND_2_SAME_RECTANGLE, bx - g, by - g,
                 bw + 2 * g, bh + 2 * g, bg, None, adj=(0.50,))
            _ell(sl, bx + (bw - d) / 2 - g, by - d + s * 0.015 - g,
                 d + 2 * g, d + 2 * g, bg, None)
        _ell(sl, bx + (bw - d) / 2, by - d + s * 0.015, d, d, c)
        _shp(sl, MSO_SHAPE.ROUND_2_SAME_RECTANGLE, bx, by, bw, bh, c, None,
             adj=(0.50,))


def alerta(sl, x, y, s, c, bg=WHITE):
    """triangulo de atencao, contorno, com exclamacao"""
    lw = s * 5.2
    _shp(sl, MSO_SHAPE.ISOSCELES_TRIANGLE, x + s * 0.02, y + s * 0.06,
         s * 0.96, s * 0.86, None, c, lw, adj=(0.5,))
    _rect(sl, x + s * 0.465, y + s * 0.40, s * 0.07, s * 0.24, c)
    _ell(sl, x + s * 0.462, y + s * 0.69, s * 0.076, s * 0.076, c)


def ciclo(sl, x, y, s, c, bg=WHITE):
    """duas setas circulares (realocar)"""
    th = 0.17
    m = s * 0.05                  # folga para as pontas das setas
    d = s - 2 * m
    r = d * 0.5
    rm = r * (1 - th / 2)
    for a0, a1 in ((25, 155), (205, 335)):
        _arc(sl, x + m, y + m, d, d, a0, a1, th, c)
        t = math.radians(a1)
        cx, cy = x + m + r + rm * math.cos(t), y + m + r + rm * math.sin(t)
        _tri(sl, cx, cy, s * 0.155, a1 + 90, c)


def banco(sl, x, y, s, c, bg=WHITE):
    """predio classico: frontao, colunas e base"""
    _ff(sl, [(x + s * 0.5, y), (x + s, y + s * 0.26), (x, y + s * 0.26)], c)
    _rect(sl, x + s * 0.02, y + s * 0.28, s * 0.96, s * 0.07, c)
    for i in range(4):
        _rect(sl, x + s * (0.10 + i * 0.245), y + s * 0.39, s * 0.11,
              s * 0.42, c)
    _rect(sl, x + s * 0.04, y + s * 0.83, s * 0.92, s * 0.07, c)
    _rect(sl, x, y + s * 0.92, s, s * 0.08, c)


def engrenagem(sl, x, y, s, c, bg=WHITE):
    _shp(sl, MSO_SHAPE.GEAR_9, x, y, s, s, c, None)
    _ell(sl, x + s * 0.325, y + s * 0.325, s * 0.35, s * 0.35, bg)


def monitor(sl, x, y, s, c, bg=WHITE):
    """tela com linhas e pe"""
    lw = s * 5.0
    _rect(sl, x, y + s * 0.04, s, s * 0.62, None, c, lw)
    _rect(sl, x + s * 0.13, y + s * 0.18, s * 0.44, s * 0.065, c)
    _rect(sl, x + s * 0.13, y + s * 0.31, s * 0.74, s * 0.065, c)
    _rect(sl, x + s * 0.13, y + s * 0.44, s * 0.33, s * 0.065, c)
    _rect(sl, x + s * 0.43, y + s * 0.66, s * 0.14, s * 0.18, c)
    _rect(sl, x + s * 0.24, y + s * 0.84, s * 0.52, s * 0.09, c)


def rede(sl, x, y, s, c, bg=WHITE):
    """tres nos ligados (ecossistema)"""
    d = s * 0.30
    cl = (x + d / 2, y + s * 0.50)
    ct = (x + s - d / 2, y + d / 2)
    cb = (x + s - d / 2, y + s - d / 2)
    _seg(sl, cl[0], cl[1], ct[0], ct[1], s * 0.065, c)
    _seg(sl, cl[0], cl[1], cb[0], cb[1], s * 0.065, c)
    for cx, cy in (cl, ct, cb):
        _ell(sl, cx - d / 2, cy - d / 2, d, d, c)


def alvo(sl, x, y, s, c, bg=WHITE):
    _shp(sl, MSO_SHAPE.DONUT, x, y, s, s, c, None, adj=(0.13,))
    _shp(sl, MSO_SHAPE.DONUT, x + s * 0.25, y + s * 0.25, s * 0.50, s * 0.50,
         c, None, adj=(0.22,))
    _ell(sl, x + s * 0.40, y + s * 0.40, s * 0.20, s * 0.20, c)


def trofeu(sl, x, y, s, c, bg=WHITE):
    _shp(sl, MSO_SHAPE.DONUT, x + s * 0.04, y + s * 0.08, s * 0.26, s * 0.28,
         c, None, adj=(0.24,))
    _shp(sl, MSO_SHAPE.DONUT, x + s * 0.70, y + s * 0.08, s * 0.26, s * 0.28,
         c, None, adj=(0.24,))
    _ff(sl, [(x + s * 0.22, y), (x + s * 0.78, y), (x + s * 0.70, y + s * 0.38),
             (x + s * 0.58, y + s * 0.50), (x + s * 0.42, y + s * 0.50),
             (x + s * 0.30, y + s * 0.38)], c)
    _rect(sl, x + s * 0.44, y + s * 0.50, s * 0.12, s * 0.20, c)
    _rect(sl, x + s * 0.28, y + s * 0.70, s * 0.44, s * 0.09, c)
    _rect(sl, x + s * 0.18, y + s * 0.82, s * 0.64, s * 0.11, c)


def carro(sl, x, y, s, c, bg=WHITE):
    """carro de perfil (Localiza)"""
    _ff(sl, [(x + s * 0.00, y + s * 0.62), (x + s * 0.03, y + s * 0.46),
             (x + s * 0.20, y + s * 0.43), (x + s * 0.33, y + s * 0.20),
             (x + s * 0.67, y + s * 0.20), (x + s * 0.81, y + s * 0.43),
             (x + s * 0.97, y + s * 0.48), (x + s * 1.00, y + s * 0.62),
             (x + s * 1.00, y + s * 0.74), (x + s * 0.00, y + s * 0.74)], c)
    for cx in (0.25, 0.75):
        _ell(sl, x + s * (cx - 0.13), y + s * 0.61, s * 0.26, s * 0.26, c)
        _ell(sl, x + s * (cx - 0.052), y + s * 0.688, s * 0.104, s * 0.104, bg)


# ------------------------------- icones de linha (line art, viewBox 32x32)
# Mesma familia de traco dos SVGs da pagina F&I multibanco: contorno fino,
# sem preenchimento. Coexistem com os icones cheios acima --- escolha um
# estilo por pagina.
def _p32(x, y, s, a, b):
    """ponto do viewBox 32x32 -> polegadas"""
    return (x + a / 32.0 * s, y + b / 32.0 * s)


def _lw32(s, w=2.0):
    """stroke-width `w` do viewBox, em pt"""
    return max(0.75, w / 32.0 * s * 72.0)


def _cx32(sl, x, y, s, p0, p1, c, w=2.0):
    """traco reto: conector (freeform de bounding box nula nao renderiza)"""
    (x0, y0), (x1, y1) = _p32(x, y, s, *p0), _p32(x, y, s, *p1)
    cn = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x0), Inches(y0),
                                 Inches(x1), Inches(y1))
    cn.line.color.rgb = c
    cn.line.width = Pt(_lw32(s, w))
    cn.line._get_or_add_ln().set('cap', 'rnd')
    return cn


def _ln32(sl, x, y, s, pts, c, close=False, w=2.0):
    """polilinha no viewBox 32x32"""
    if len(pts) == 2 and not close:
        return _cx32(sl, x, y, s, pts[0], pts[1], c, w)
    e = [(Emu(int(round(a * EMU))), Emu(int(round(b * EMU))))
         for a, b in (_p32(x, y, s, a, b) for a, b in pts)]
    bld = sl.shapes.build_freeform(e[0][0], e[0][1])
    bld.add_line_segments(e[1:], close=close)
    return _clean(bld.convert_to_shape(), None, c, _lw32(s, w))


def _el32(sl, x, y, s, cx, cy, r, c, w=2.0):
    px, py = _p32(x, y, s, cx - r, cy - r)
    d = 2 * r / 32.0 * s
    return _ell(sl, px, py, d, d, None, c, _lw32(s, w))


def _bez32(p0, p1, p2, p3, n=14):
    """amostra a curva de Bezier cubica (os `c`/`s` dos paths SVG)"""
    out = []
    for i in range(1, n + 1):
        t = i / float(n)
        u = 1 - t
        out.append((u ** 3 * p0[0] + 3 * u * u * t * p1[0]
                    + 3 * u * t * t * p2[0] + t ** 3 * p3[0],
                    u ** 3 * p0[1] + 3 * u * u * t * p1[1]
                    + 3 * u * t * t * p2[1] + t ** 3 * p3[1]))
    return out


def _svgarc(p0, p1, r, large, sweep, n=16):
    """arco SVG `a r r 0 large sweep x y` -> polilinha (sem repetir p0)"""
    (x1, y1), (x2, y2) = p0, p1
    dx, dy = (x1 - x2) / 2.0, (y1 - y2) / 2.0
    lam = (dx * dx + dy * dy) / (r * r)
    if lam > 1:
        r *= math.sqrt(lam)
    q = r * r - (dx * dx + dy * dy)
    co = math.sqrt(max(0.0, q / (dx * dx + dy * dy))) * (1 if large != sweep
                                                        else -1)
    cx = co * dy + (x1 + x2) / 2.0
    cy = -co * dx + (y1 + y2) / 2.0
    t1 = math.atan2(y1 - cy, x1 - cx)
    dt = math.atan2(y2 - cy, x2 - cx) - t1
    if sweep == 0 and dt > 0:
        dt -= 2 * math.pi
    if sweep == 1 and dt < 0:
        dt += 2 * math.pi
    return [(cx + r * math.cos(t1 + dt * i / float(n)),
             cy + r * math.sin(t1 + dt * i / float(n))) for i in range(1, n + 1)]


def _v(pts, view=64):
    """pontos de um viewBox `view` -> o 32x32 da biblioteca"""
    k = 32.0 / view
    return [(a * k, b * k) for a, b in pts]


def _semi32(cx, cy, r, n=12):
    """semicircunferencia de barriga para cima, da direita para a esquerda"""
    return [(cx + r * math.cos(math.pi * i / n),
             cy - r * math.sin(math.pi * i / n)) for i in range(n + 1)]


def pessoas_linha(sl, x, y, s, c, bg=WHITE):
    """duas pessoas, so contorno"""
    _el32(sl, x, y, s, 11, 10, 4, c)
    _el32(sl, x, y, s, 21, 10, 4, c)
    _ln32(sl, x, y, s, [(4, 25)] + _bez32((4, 25), (4.5, 19), (8, 16), (12, 16))
          + _bez32((12, 16), (16, 16), (19.5, 19), (20, 25)), c)
    _ln32(sl, x, y, s, [(13, 25)]
          + _bez32((13, 25), (13.5, 20), (16.5, 17.5), (20, 17.5))
          + _bez32((20, 17.5), (23.5, 17.5), (26.5, 20), (27, 25)), c)


def banco_linha(sl, x, y, s, c, bg=WHITE):
    """predio classico, so contorno"""
    _ln32(sl, x, y, s, [(5, 12), (16, 5), (27, 12)], c)
    _ln32(sl, x, y, s, [(7, 13), (25, 13)], c)
    for cx in (9, 15, 21):
        _ln32(sl, x, y, s, [(cx, 13), (cx, 25)], c)
    _ln32(sl, x, y, s, [(5, 26), (27, 26)], c)


def carro_linha(sl, x, y, s, c, bg=WHITE):
    """carro de frente, so contorno, com os vaos das rodas"""
    body = [(4, 19), (7, 12), (25, 12), (28, 19), (28, 25), (25, 25)] \
        + _semi32(21, 25, 4)[1:] + [(15, 25)] + _semi32(11, 25, 4)[1:] + [(4, 25)]
    _ln32(sl, x, y, s, body, c, close=True)
    _ln32(sl, x, y, s, [(8, 12), (11, 8), (21, 8), (24, 12)], c)
    _el32(sl, x, y, s, 11, 24, 2.5, c)
    _el32(sl, x, y, s, 21, 24, 2.5, c)


def engrenagem_linha(sl, x, y, s, c, bg=WHITE):
    """circulos concentricos com raios (tecnologia)"""
    _el32(sl, x, y, s, 16, 16, 5, c)
    _el32(sl, x, y, s, 16, 16, 10, c)
    for a, b, d, e in ((16, 3, 16, 7), (16, 25, 16, 29), (3, 16, 7, 16),
                       (25, 16, 29, 16), (7, 7, 10, 10), (22, 22, 25, 25),
                       (25, 7, 22, 10), (10, 22, 7, 25)):
        _ln32(sl, x, y, s, [(a, b), (d, e)], c)


def monitor_linha(sl, x, y, s, c, bg=WHITE):
    """tela com pe, so contorno"""
    _ln32(sl, x, y, s, [(5, 5), (27, 5), (27, 21), (5, 21)], c, close=True)
    _ln32(sl, x, y, s, [(12, 27), (20, 27)], c)
    _ln32(sl, x, y, s, [(16, 21), (16, 27)], c)


def rede_linha(sl, x, y, s, c, bg=WHITE):
    """tres nos ligados, so contorno (ecossistema)"""
    _el32(sl, x, y, s, 8, 16, 3, c)
    _el32(sl, x, y, s, 23, 8, 3, c)
    _el32(sl, x, y, s, 23, 24, 3, c)
    _ln32(sl, x, y, s, [(11, 15), (20, 9)], c)
    _ln32(sl, x, y, s, [(11, 17), (20, 23)], c)


def barras_linha(sl, x, y, s, c, bg=WHITE):
    """grafico de barras ascendente, so contorno"""
    for bx, by, bw, bh in ((10, 35, 8, 15), (27, 22, 8, 28), (44, 12, 8, 38)):
        _ln32(sl, x, y, s, _v([(bx, by), (bx + bw, by), (bx + bw, by + bh),
                               (bx, by + bh)]), c, close=True, w=1.25)


def pessoas_grupo_linha(sl, x, y, s, c, bg=WHITE):
    """tres pessoas lado a lado, so contorno (a cascata ate a ponta)"""
    for cx, cy, r in ((32, 19, 7), (18, 24, 6), (46, 24, 6)):
        _el32(sl, x, y, s, cx / 2.0, cy / 2.0, r / 2.0, c, w=1.25)
    corpo = [(20, 47), (20, 40)] + _svgarc((20, 40), (44, 40), 12, 0, 1) \
        + [(44, 47)]
    esq = [(7, 47), (7, 42)] + _svgarc((7, 42), (17, 32), 10, 0, 1)
    dir_ = [(57, 47), (57, 42)] + _svgarc((57, 42), (47, 32), 10, 0, 0)
    for pts in (corpo, esq, dir_):
        _ln32(sl, x, y, s, _v(pts), c, w=1.25)


def alerta_linha(sl, x, y, s, c, bg=WHITE):
    """triangulo de atencao com exclamacao, so contorno"""
    _ln32(sl, x, y, s, _v([(32, 10), (54, 50), (10, 50)]), c, close=True,
          w=1.25)
    _ln32(sl, x, y, s, _v([(32, 24), (32, 37)]), c, w=1.25)
    _ln32(sl, x, y, s, _v([(32, 43), (32, 45)]), c, w=1.25)


def ciclo_linha(sl, x, y, s, c, bg=WHITE):
    """duas setas de ciclo, so contorno (realocar)"""
    cima = [(47, 18)] + _svgarc((47, 18), (19, 20), 18, 0, 0) + [(14, 21),
                                                                 (15, 15)]
    baixo = [(17, 46)] + _svgarc((17, 46), (45, 44), 18, 0, 0) + [(50, 43),
                                                                  (49, 49)]
    for pts in (cima, baixo):
        _ln32(sl, x, y, s, _v(pts), c, w=1.25)
    _ln32(sl, x, y, s, _v([(45, 18), (39, 18)]), c, w=1.25)
    _ln32(sl, x, y, s, _v([(19, 46), (25, 46)]), c, w=1.25)


def trofeu_linha(sl, x, y, s, c, bg=WHITE):
    """trofeu, so contorno (faixa de merito)"""
    taca = ([(10, 5), (22, 5), (22, 11)] + _svgarc((22, 11), (16, 19), 8, 0, 1)
            + _svgarc((16, 19), (10, 11), 8, 0, 1))
    _ln32(sl, x, y, s, taca, c, close=True)
    _ln32(sl, x, y, s, [(10, 8), (5, 8)] + _svgarc((5, 8), (11, 16), 8, 0, 0), c)
    _ln32(sl, x, y, s, [(22, 8), (27, 8)] + _svgarc((27, 8), (21, 16), 8, 0, 1), c)
    _ln32(sl, x, y, s, [(16, 19), (16, 24)], c)
    _ln32(sl, x, y, s, [(13, 24), (19, 24)], c)
    _ln32(sl, x, y, s, [(11, 28), (21, 28)], c)


ICONS = {'barras': barras, 'pessoa': pessoa, 'pessoas': pessoas,
         'pessoas_cheio': pessoas_cheio, 'alerta': alerta, 'ciclo': ciclo,
         'banco': banco, 'engrenagem': engrenagem, 'monitor': monitor,
         'rede': rede, 'alvo': alvo, 'trofeu': trofeu, 'carro': carro,
         'pessoas_linha': pessoas_linha, 'banco_linha': banco_linha,
         'carro_linha': carro_linha, 'engrenagem_linha': engrenagem_linha,
         'monitor_linha': monitor_linha, 'rede_linha': rede_linha,
         'barras_linha': barras_linha,
         'pessoas_grupo_linha': pessoas_grupo_linha,
         'alerta_linha': alerta_linha, 'ciclo_linha': ciclo_linha,
         'trofeu_linha': trofeu_linha}


def icon(sl, kind, x, y, s, c, bg=WHITE):
    ICONS[kind](sl, x, y, s, c, bg)
