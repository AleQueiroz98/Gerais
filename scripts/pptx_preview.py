# -*- coding: utf-8 -*-
"""Preview HTML de um .pptx, lido da geometria real do arquivo.

Serve para conferir o layout sem depender do PowerPoint ou do LibreOffice:
percorre shapes e tabelas, le posicao, tamanho, preenchimento e runs de texto
e desenha tudo em divs posicionados (1" = 100px).

    python3 pptx_preview.py ../output/deck.pptx ../output/deck_preview.html

E uma aproximacao: fontes e quebra de linha ficam a cargo do navegador, entao
serve para checar estrutura, alturas e transbordo, nao para prova de cor.
"""
import html
import sys

from pptx import Presentation
from pptx.enum.dml import MSO_FILL
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Emu

PX = 100.0   # pixels por polegada


def inches(v):
    return Emu(v).inches if v is not None else 0.0


def px(v):
    return inches(v) * PX


def color_of(fmt):
    try:
        if fmt.type is None:
            return None
        return '#' + str(fmt.fore_color.rgb)
    except Exception:
        return None


def gradient_of(shape):
    """fundo CSS de um shape com preenchimento em gradiente"""
    try:
        if shape.fill.type != MSO_FILL.GRADIENT:
            return None
        stops = ', '.join('#%s %.0f%%' % (st.color.rgb, st.position * 100)
                          for st in shape.fill.gradient_stops)
    except Exception:
        return None
    try:
        ang = shape.fill.gradient_angle
    except Exception:
        ang = 90.0
    return 'linear-gradient(%.0fdeg, %s)' % (ang + 90.0, stops)


def prst_of(shape):
    """nome da geometria pronta (ellipse, roundRect, ...) ou None"""
    try:
        prst = shape._element.spPr.find(qn('a:prstGeom'))
        return prst.get('prst') if prst is not None else None
    except Exception:
        return None


def freeform_svg(shape, w, h):
    """desenha o custGeom (freeform) como SVG, preservando o traco"""
    path = shape._element.spPr.find(qn('a:custGeom'))
    if path is None:
        return None
    path = path.find(qn('a:pathLst')).find(qn('a:path'))
    pw = float(path.get('w') or 0) or 1.0   # linhas retas tem w ou h = 0
    ph = float(path.get('h') or 0) or 1.0
    pts, closed = [], False
    for el in path:
        tag = el.tag.split('}')[1]
        if tag == 'close':
            closed = True
            continue
        pt = el.find(qn('a:pt'))
        if pt is None:
            continue
        pts.append((float(pt.get('x')) / pw * w, float(pt.get('y')) / ph * h))
    if not pts:
        return None
    try:
        stroke = '#' + str(shape.line.color.rgb)
        lw = max((shape.line.width.pt if shape.line.width else 0.75) * PX / 72.0, 1.0)
    except Exception:
        stroke, lw = '#333333', 1.0
    return ('<svg style="position:absolute;left:0;top:0;overflow:visible" '
            'width="%.1f" height="%.1f"><%s points="%s" fill="none" '
            'stroke="%s" stroke-width="%.1f" stroke-linejoin="round" '
            'stroke-linecap="round"/></svg>'
            % (max(w, 1.0), max(h, 1.0), 'polygon' if closed else 'polyline',
               ' '.join('%.1f,%.1f' % p for p in pts), stroke, lw))


def connector_svg(shape, w, h):
    """conector (linha) desenhado como SVG, respeitando flipH/flipV"""
    el = shape._element
    if not el.tag.endswith('}cxnSp'):
        return None
    xfrm = el.spPr.find(qn('a:xfrm'))
    fh = xfrm is not None and xfrm.get('flipH') == '1'
    fv = xfrm is not None and xfrm.get('flipV') == '1'
    x1, x2 = (w, 0) if fh else (0, w)
    y1, y2 = (h, 0) if fv else (0, h)
    try:
        stroke = '#' + str(shape.line.color.rgb)
        lw = max((shape.line.width.pt if shape.line.width else 0.75) * PX / 72.0, 1.0)
    except Exception:
        stroke, lw = '#333333', 1.0
    return ('<svg style="position:absolute;left:0;top:0;overflow:visible" '
            'width="%.1f" height="%.1f"><line x1="%.1f" y1="%.1f" x2="%.1f" '
            'y2="%.1f" stroke="%s" stroke-width="%.1f" stroke-linecap="round"/>'
            '</svg>' % (max(w, 1.0), max(h, 1.0), x1, y1, x2, y2, stroke, lw))


def run_color(run):
    try:
        return '#' + str(run.font.color.rgb)
    except Exception:
        return '#333333'


ALIGN = {PP_ALIGN.CENTER: 'center', PP_ALIGN.RIGHT: 'right',
         PP_ALIGN.LEFT: 'left', None: 'left'}
ANCHOR = {MSO_ANCHOR.MIDDLE: 'center', MSO_ANCHOR.BOTTOM: 'flex-end',
          MSO_ANCHOR.TOP: 'flex-start', None: 'flex-start'}


def text_html(tf, default_pt=9.0):
    out = []
    for p in tf.paragraphs:
        runs = []
        for r in p.runs:
            size = r.font.size.pt if r.font.size else default_pt
            style = 'font-size:%.1fpx;line-height:%.1fpx;color:%s;%s' % (
                size * PX / 72.0, size * 1.24 * PX / 72.0, run_color(r),
                'font-weight:700;' if r.font.bold else '')
            runs.append('<span style="%s">%s</span>'
                        % (style, html.escape(r.text)))
        out.append('<div style="text-align:%s">%s</div>'
                   % (ALIGN.get(p.alignment, 'left'), ''.join(runs) or '&nbsp;'))
    return ''.join(out)


def frame_style(tf):
    return 'justify-content:%s' % ANCHOR.get(tf.vertical_anchor, 'flex-start')


def render_table(shape, parts):
    tbl = shape.table
    xs, ys = [px(shape.left)], [px(shape.top)]
    for c in tbl.columns:
        xs.append(xs[-1] + px(c.width))
    for r in tbl.rows:
        ys.append(ys[-1] + px(r.height))

    for ri in range(len(tbl.rows)):
        for ci in range(len(tbl.columns)):
            cell = tbl.cell(ri, ci)
            if cell.is_spanned:
                continue
            w = xs[ci + cell.span_width] - xs[ci]
            h = ys[ri + cell.span_height] - ys[ri]
            fill = color_of(cell.fill)
            style = ('position:absolute;left:%.1fpx;top:%.1fpx;width:%.1fpx;'
                     'height:%.1fpx;box-sizing:border-box;overflow:hidden;'
                     'display:flex;flex-direction:column;%s;'
                     'padding:%.1fpx %.1fpx;%s'
                     % (xs[ci], ys[ri], w, h, frame_style(cell.text_frame),
                        px(cell.margin_top), px(cell.margin_left),
                        'background:%s;' % fill if fill else ''))
            parts.append('<div style="%s">%s</div>'
                         % (style, text_html(cell.text_frame)))
    # grade de referencia, para ver a altura real de cada linha
    for y in ys[1:-1]:
        parts.append('<div style="position:absolute;left:%.1fpx;top:%.1fpx;'
                     'width:%.1fpx;height:1px;background:#E6E6E6"></div>'
                     % (xs[0], y, xs[-1] - xs[0]))


def border_of(shape):
    try:
        line = shape.line
        if line.fill.type is None:
            return ''
        w = line.width.pt if line.width else 0.75
        return 'border:%.1fpx solid %s;' % (max(w * PX / 72.0, 1.0),
                                            '#' + str(line.color.rgb))
    except Exception:
        return ''


def render_shape(shape, parts):
    if shape.has_table:
        render_table(shape, parts)
        return
    fill = None
    try:
        fill = color_of(shape.fill)
    except Exception:
        pass
    fill = fill or gradient_of(shape)
    w, h = px(shape.width), px(shape.height)
    body = freeform_svg(shape, w, h) or connector_svg(shape, w, h) or ''
    free = bool(body)          # o traco do freeform ja vai dentro do SVG
    if not body and shape.has_text_frame:
        body = text_html(shape.text_frame)
    radius = 'border-radius:50%;' if prst_of(shape) == 'ellipse' else ''
    rot = shape.rotation or 0
    style = ('position:absolute;left:%.1fpx;top:%.1fpx;width:%.1fpx;'
             'height:%.1fpx;box-sizing:border-box;display:flex;'
             'flex-direction:column;%s;%s%s%s%s'
             % (px(shape.left), px(shape.top), w, h,
                frame_style(shape.text_frame)
                if shape.has_text_frame else 'justify-content:flex-start',
                'background:%s;' % fill if fill else '',
                '' if free else border_of(shape), radius, 'transform:rotate(%.1fdeg);' % rot if rot else ''))
    parts.append('<div style="%s">%s</div>' % (style, body))


def main(src, dst):
    prs = Presentation(src)
    w, h = px(prs.slide_width), px(prs.slide_height)
    pages = []
    for i, slide in enumerate(prs.slides, start=1):
        parts = []
        for shape in slide.shapes:
            render_shape(shape, parts)
        pages.append('<section style="position:relative;width:%.0fpx;'
                     'height:%.0fpx;background:#fff;margin:0 auto 24px;'
                     'box-shadow:0 2px 12px rgba(0,0,0,.18)">%s'
                     '<div style="position:absolute;right:6px;bottom:2px;'
                     'font:10px Arial;color:#BBB">preview p%d</div>'
                     '</section>' % (w, h, ''.join(parts), i))
    open(dst, 'w', encoding='utf-8').write(
        '<!DOCTYPE html><meta charset="utf-8">'
        '<body style="margin:0;padding:24px;background:#EEE;'
        'font-family:Arial,Helvetica,sans-serif">%s</body>' % ''.join(pages))
    print('preview: %s (%d paginas, %.0fx%.0f px)' % (dst, len(pages), w, h))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
