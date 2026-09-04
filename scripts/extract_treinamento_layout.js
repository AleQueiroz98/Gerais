/* Mede no Chromium o layout real de cada pagina do material de treinamento e
 * grava um "display list" (posicao, estilo e trechos de texto de cada
 * elemento) que o build_treinamento_ppt.py usa para montar o PPT.
 *
 *     node extract_treinamento_layout.js <html> <saida.json> <pasta-telas>
 *
 * A medicao usa Liberation Sans (metrica proxima da Segoe UI usada no HTML),
 * porque a Segoe UI nao existe no container.
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { chromium } = require('playwright');

const [HTML, OUT, TELAS] = process.argv.slice(2);

const PROP = [
  'display', 'position', 'color', 'fontSize', 'fontWeight', 'fontStyle',
  'lineHeight', 'letterSpacing', 'textTransform', 'textAlign', 'fontFamily',
  'backgroundColor', 'backgroundImage', 'boxShadow', 'borderRadius',
  'borderTopWidth', 'borderRightWidth', 'borderBottomWidth', 'borderLeftWidth',
  'borderTopColor', 'borderRightColor', 'borderBottomColor', 'borderLeftColor',
  'borderTopStyle', 'borderRightStyle', 'borderBottomStyle', 'borderLeftStyle',
  'paddingTop', 'paddingRight', 'paddingBottom', 'paddingLeft',
  'marginTop', 'left', 'top', 'right', 'bottom', 'width', 'height', 'content',
  'fontVariant', 'whiteSpace', 'verticalAlign', 'alignItems', 'justifyContent',
];

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewportSize: { width: 1400, height: 900 } });
  await page.goto('file://' + path.resolve(HTML));
  // todas as paginas visiveis e sem o scale do palco, para medir em 1:1
  await page.addStyleTag({ content: `
    #palco{transform:none!important}
    .slide{display:flex!important}
    body{font-family:'Liberation Sans',sans-serif!important}
    #nav,#barra{display:none!important}
  `});
  await page.waitForTimeout(400);

  const data = await page.evaluate((PROP) => {
    const px = (v) => { const n = parseFloat(v); return Number.isFinite(n) ? n : null; };
    const vis = (c) => c && c !== 'transparent' && !/rgba\(\s*0,\s*0,\s*0,\s*0\)/.test(c);
    const origin = document.getElementById('palco').getBoundingClientRect();
    const rel = (r) => ({ x: r.left - origin.left, y: r.top - origin.top, w: r.width, h: r.height });

    const styleOf = (el, pseudo) => {
      const cs = getComputedStyle(el, pseudo || null);
      const o = {};
      for (const p of PROP) o[p] = cs[p];
      return o;
    };

    // trechos de texto de um elemento com filhos apenas inline
    const inlineOnly = (el) => [...el.children].every(
      (c) => getComputedStyle(c).display === 'inline');

    const runsOf = (el) => {
      const out = [];
      const walk = (node, st) => {
        for (const ch of node.childNodes) {
          if (ch.nodeType === 3) {
            if (ch.nodeValue) out.push({ t: ch.nodeValue, ...st });
          } else if (ch.nodeType === 1) {
            if (ch.tagName === 'BR') { out.push({ br: true }); continue; }
            const cs = getComputedStyle(ch);
            walk(ch, {
              color: cs.color, size: parseFloat(cs.fontSize),
              bold: parseInt(cs.fontWeight, 10) >= 600,
              italic: cs.fontStyle === 'italic',
              spacing: px(cs.letterSpacing) || 0,
              caps: cs.textTransform === 'uppercase',
            });
          }
        }
      };
      const cs = getComputedStyle(el);
      walk(el, {
        color: cs.color, size: parseFloat(cs.fontSize),
        bold: parseInt(cs.fontWeight, 10) >= 600,
        italic: cs.fontStyle === 'italic',
        spacing: px(cs.letterSpacing) || 0,
        caps: cs.textTransform === 'uppercase',
      });
      // colapso de espacos como no CSS (sem tocar no nbsp)
      for (const r of out) if (r.t) r.t = r.t.replace(/[ \t\n\r]+/g, ' ');
      while (out.length && out[0].t !== undefined && !out[0].t.trim()) out.shift();
      while (out.length && out[out.length - 1].t !== undefined
             && !out[out.length - 1].t.trim()) out.pop();
      if (out.length && out[0].t) out[0].t = out[0].t.replace(/^ /, '');
      const last = out[out.length - 1];
      if (last && last.t) last.t = last.t.replace(/ $/, '');
      return out.filter((r) => r.br || r.t !== '');
    };

    // caixa justa do texto (uniao das linhas), util onde o flex centraliza
    const textRect = (el) => {
      const rg = document.createRange();
      rg.selectNodeContents(el);
      const rs = [...rg.getClientRects()].filter((r) => r.width > 0 && r.height > 0);
      if (!rs.length) return null;
      const l = Math.min(...rs.map((r) => r.left)), t = Math.min(...rs.map((r) => r.top));
      const rr = Math.max(...rs.map((r) => r.right)), b = Math.max(...rs.map((r) => r.bottom));
      return rel({ left: l, top: t, width: rr - l, height: b - t });
    };

    const slides = [];
    for (const sl of document.querySelectorAll('#palco > .slide')) {
      const nodes = [];
      const visit = (el, depth) => {
        const st = styleOf(el);
        const r = el.getBoundingClientRect();
        const box = rel(r);
        const cbox = {
          x: box.x + px(st.borderLeftWidth) + px(st.paddingLeft),
          y: box.y + px(st.borderTopWidth) + px(st.paddingTop),
          w: box.w - px(st.borderLeftWidth) - px(st.paddingLeft)
                   - px(st.borderRightWidth) - px(st.paddingRight),
          h: box.h - px(st.borderTopWidth) - px(st.paddingTop)
                   - px(st.borderBottomWidth) - px(st.paddingBottom),
        };
        const node = {
          tag: el.tagName.toLowerCase(), cls: el.className || '', depth,
          box, cbox, style: st,
        };
        if (el.tagName === 'IMG') node.img = el.getAttribute('src');
        for (const ps of ['::before', '::after']) {
          const p = styleOf(el, ps);
          const hasContent = p.content && p.content !== 'none' && p.content !== 'normal';
          const hasBox = vis(p.backgroundColor) || px(p.borderLeftWidth) > 0;
          if (hasContent || hasBox) node[ps === '::before' ? 'before' : 'after'] = p;
        }
        const text = [...el.childNodes].some(
          (c) => c.nodeType === 3 && c.nodeValue.trim());
        if ((text || el.children.length) && inlineOnly(el)
            && el.tagName !== 'IMG' && el.textContent.trim()) {
          node.runs = runsOf(el);
          node.tbox = textRect(el);
          nodes.push(node);
          return;                     // inline ja consumido nos runs
        }
        nodes.push(node);
        for (const ch of el.children) visit(ch, depth + 1);
      };
      visit(sl, 0);
      slides.push({ n: sl.getAttribute('data-n'), cls: sl.className, nodes });
    }
    return { slides };
  }, PROP);

  // telas do produto: base64 -> arquivo em disco
  fs.mkdirSync(TELAS, { recursive: true });
  let i = 0;
  for (const sl of data.slides) {
    for (const nd of sl.nodes) {
      if (!nd.img) continue;
      const b64 = nd.img.split(',')[1];
      const buf = Buffer.from(b64, 'base64');
      const h = crypto.createHash('md5').update(buf).digest('hex').slice(0, 8);
      const name = `tela_${String(++i).padStart(2, '0')}_${h}.png`;
      fs.writeFileSync(path.join(TELAS, name), buf);
      nd.img = name;
    }
  }

  fs.writeFileSync(OUT, JSON.stringify(data, null, 1));
  console.log('paginas:', data.slides.length,
              'nos:', data.slides.reduce((a, s) => a + s.nodes.length, 0),
              'telas:', i);
  await browser.close();
})();
