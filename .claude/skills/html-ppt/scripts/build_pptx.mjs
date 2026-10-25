#!/usr/bin/env node
// HTML deck -> native, editable PPTX.
// usage: node build_pptx.mjs <deck.html> <out.pptx> [--hard-wrap] [--font-map "Pretendard=Malgun Gothic,Inter=Arial"]
// Every <section class="slide"> (1280x720 CSS px) becomes one slide. The DOM is walked in
// document order; boxes become shapes, text becomes text boxes, and img/svg/canvas (or
// anything marked data-ppt="raster") is rasterized as a picture. See SKILL.md for the contract.
import { chromium } from 'playwright';
import PptxGenJS from 'pptxgenjs';
import JSZip from 'jszip';
import { embedFonts, restrictedFamilies } from './embed_fonts.mjs';
import { makeFontLoader, outlineGroups, shapeOf } from './outline.mjs';
import { writeFileSync, readFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';
import { resolve } from 'node:path';

const argv = process.argv.slice(2);
const flag = (n) => argv.includes(n);
const opt = (n) => { const i = argv.indexOf(n); return i >= 0 ? argv[i + 1] : undefined; };
const pos = argv.filter((a, i) => !a.startsWith('--') && !['--font-map', '--settle', '--template'].includes(argv[i - 1]));
const [input, output] = pos;
if (!input || !output) {
  console.error('usage: node build_pptx.mjs <deck.html> <out.pptx> [--hard-wrap] [--font-map "A=B,C=D"]');
  process.exit(1);
}
const HARD = flag('--hard-wrap');
const NO_SVG = flag('--no-svg'); // force PNG pictures instead of vector SVG
const SETTLE_MS = Number(opt('--settle') ?? 2000); // wait after load so runtime-animated content (fps readouts, canvases) settles
const NO_EMBED = flag('--no-embed-fonts'); // by default @font-face fonts (TTF/WOFF) are embedded in the pptx
const FONT_MAP = Object.fromEntries((opt('--font-map') || '').split(',').filter(Boolean).map((p) => p.split('=').map((s) => s.trim())));

const W = 1280, H = 720;
// --template: a .pptx whose slide master, layouts and theme the output adopts (slides pick a layout by data-layout)
const TEMPLATE = opt('--template');
let TPL_W = 13.333, TPL_H = 7.5, TPL_ZIP = null;
const TPL_LAYOUTS = [];
if (TEMPLATE) {
  TPL_ZIP = await JSZip.loadAsync(readFileSync(resolve(TEMPLATE)));
  const pres = await TPL_ZIP.file('ppt/presentation.xml').async('string');
  const sz = pres.match(/<p:sldSz cx="(\d+)" cy="(\d+)"/);
  TPL_W = +sz[1] / 914400; TPL_H = +sz[2] / 914400;
  const files = Object.keys(TPL_ZIP.files).filter((f) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(f))
    .sort((a, b) => +a.match(/(\d+)\.xml$/)[1] - +b.match(/(\d+)\.xml$/)[1]);
  for (const f of files) {
    const x = await TPL_ZIP.file(f).async('string');
    // placeholders in layout order: type (absent = "obj") and idx
    const phs = [...x.matchAll(/<p:ph\b([^>]*?)\/?>/g)].map((m) => ({ type: (m[1].match(/type="(\w+)"/) || [, 'obj'])[1], idx: (m[1].match(/idx="(\d+)"/) || [])[1] }));
    TPL_LAYOUTS.push({ file: f.split('/').pop(), name: (x.match(/<p:cSld name="([^"]*)"/) || [, ''])[1], phs });
  }
}
// each slide uses the layout its section asked for (default: "Title Only", then "Blank", then the first layout)
const pick = (want) => {
  if (want) { const hit = TPL_LAYOUTS.find((l) => l.name.toLowerCase() === want.toLowerCase()) || TPL_LAYOUTS.find((l) => l.name.toLowerCase().includes(want.toLowerCase())); if (hit) return hit; }
  return TPL_LAYOUTS.find((l) => /title[ _-]?only/i.test(l.name)) || TPL_LAYOUTS.find((l) => /blank/i.test(l.name)) || TPL_LAYOUTS[0];
};
// first layout placeholder that takes data-ph="body" / data-ph="subtitle"
const layoutPh = (layout, kind) => layout.phs.find((p) => (kind === 'body' ? p.type === 'body' || p.type === 'obj' : p.type === 'subTitle'));
const SCALE = TEMPLATE ? Math.min(TPL_W / 13.333, TPL_H / 7.5) : 1;
const PX = 1 / 96; // css px -> inch

// Runs inside the page: returns, per slide, a flat z-ordered list of draw items.
function extract(opts) {
  const slides = [...document.querySelectorAll('section.slide')];
  let rid = 0;

  const rgba = (c) => {
    // SVG attributes such as flood-color are written as #rgb / #rrggbb; computed styles are rgb()
    const hx = c && c.trim().match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/i);
    if (hx) { const h = hx[1].length === 3 ? hx[1].replace(/./g, (x) => x + x) : hx[1]; return { hex: h.toUpperCase(), a: 1 }; }
    const m = c.match(/rgba?\(([^)]+)\)/);
    if (!m) return null;
    const p = m[1].split(/[ ,/]+/).filter(Boolean).map(Number);
    const a = p.length > 3 ? p[3] : 1;
    if (a === 0) return null;
    const hex = p.slice(0, 3).map((v) => Math.round(v).toString(16).padStart(2, '0')).join('').toUpperCase();
    return { hex, a };
  };
  // The font the browser actually drew with: first family of the stack that is installed or loaded as a web font.
  // (An uninstalled "Pretendard" falls through to the next family, so PowerPoint matches the preview.)
  const fontProbe = document.createElement('canvas').getContext('2d');
  const probeCache = {};
  window.__webFonts = window.__webFonts || [];
  const isInstalled = (f) => {
    if (f in probeCache) return probeCache[f];
    const sample = 'mmmmmmmmmmlliWM한글あ漢';
    const w = (font) => { fontProbe.font = `72px ${font}`; return fontProbe.measureText(sample).width; };
    const hit = ['monospace', 'serif', 'sans-serif'].some((b) => w(`"${f}", ${b}`) !== w(b));
    return (probeCache[f] = hit);
  };
  const family = (ff) => {
    const generic = /^(sans-serif|serif|monospace|system-ui|ui-.*|-apple-system|blinkmacsystemfont|cursive|fantasy)$/i;
    const fams = ff.split(',').map((s) => s.trim().replace(/^["']|["']$/g, '')).filter((f) => !generic.test(f));
    const found = fams.find(isInstalled);
    if (found) {
      if ([...document.fonts].some((d) => d.family.replace(/["']/g, '') === found && d.status === 'loaded') && !window.__webFonts.includes(found)) window.__webFonts.push(found);
      return found;
    }
    return fams.length ? fams[fams.length - 1] : 'Arial';
  };
  const parseShadow = (str, inset = false) => {
    if (!str || str === 'none') return null;
    const parts = str.split(/,(?![^(]*\))/);
    for (const p of parts) {
      if (/inset/.test(p) !== inset) continue;
      const m = p.match(/(rgba?\([^)]+\))\s+(-?[\d.]+)px\s+(-?[\d.]+)px(?:\s+([\d.]+)px)?/);
      const c = m && rgba(m[1]);
      if (c) return { color: c.hex, a: c.a, ox: +m[2], oy: +m[3], blur: +(m[4] || 0) };
    }
    return null;
  };
  // linear/radial-gradient(...) -> {kind, ang(OOXML deg), stops:[{pos,color,alpha}]}; null when unsupported
  const parseGradient = (str) => {
    const m = str.match(/^(linear|radial)-gradient\((.*)\)$/s);
    if (!m) return null;
    const parts = m[2].split(/,(?![^(]*\))/).map((s) => s.trim());
    let ang = 90; // CSS default: to bottom (180deg) -> OOXML 90
    if (m[1] === 'linear') {
      const head = parts[0];
      if (/deg$/.test(head)) { ang = parseFloat(head) - 90; parts.shift(); }
      else if (/^to /.test(head)) {
        const map = { 'to top': 270, 'to right': 0, 'to bottom': 90, 'to left': 180, 'to top right': 315, 'to right top': 315, 'to bottom right': 45, 'to right bottom': 45, 'to bottom left': 135, 'to left bottom': 135, 'to top left': 225, 'to left top': 225 };
        if (!(head in map)) return null;
        ang = map[head]; parts.shift();
      } else ang = 90;
    } else if (/^(circle|ellipse|closest|farthest|at )/.test(parts[0])) parts.shift();
    const stops = [];
    for (const p of parts) {
      const cm = p.match(/^(rgba?\([^)]+\))\s*(-?[\d.]+%)?$/);
      if (!cm) return null;
      const c = rgba(cm[1]);
      stops.push({ pos: cm[2] ? parseFloat(cm[2]) : null, color: c ? c.hex : '000000', alpha: c ? c.a : 0 });
    }
    if (stops.length < 2) return null;
    stops[0].pos ??= 0; stops[stops.length - 1].pos ??= 100;
    for (let i = 1; i < stops.length - 1; i++) if (stops[i].pos === null) {
      let j = i; while (stops[j].pos === null) j++;
      const a = stops[i - 1].pos, b = stops[j].pos;
      for (let k = i; k < j; k++) stops[k].pos = a + (b - a) * (k - i + 1) / (j - i + 1);
    }
    return { kind: m[1], ang: ((ang % 360) + 360) % 360, stops };
  };
  const RASTER_TAGS = new Set(['IMG', 'SVG', 'CANVAS', 'VIDEO']);
  const isRasterEl = (el) => RASTER_TAGS.has(el.tagName.toUpperCase()) || el.getAttribute('data-ppt') === 'raster';
  const hasTransform = (cs) => cs.transform && cs.transform !== 'none';
  const hasBoxStyle = (cs) => !!rgba(cs.backgroundColor) || cs.backgroundImage !== 'none' ||
    ['Top', 'Right', 'Bottom', 'Left'].some((s) => parseFloat(cs['border' + s + 'Width']) > 0 && cs['border' + s + 'Style'] !== 'none');
  // inline content that can be folded into its parent's text runs
  const isInlineMergeable = (el) => {
    if (el.tagName === 'BR') return true;
    if (isRasterEl(el)) return false;
    const cs = getComputedStyle(el);
    if (cs.display === 'inline') return [...el.children].every(isInlineMergeable);
    if (cs.display === 'inline-block') return !hasBoxStyle(cs) && !hasTransform(cs) && [...el.children].every(isInlineMergeable);
    return false;
  };
  const opacityOf = (el, root) => {
    let o = 1;
    for (let n = el; n && n !== root.parentElement; n = n.parentElement) o *= parseFloat(getComputedStyle(n).opacity);
    return o;
  };
  const snap = (cs, hl) => {
    const col = rgba(cs.color);
    const sw = parseFloat(cs.webkitTextStrokeWidth) || 0, sc = sw > 0 && rgba(cs.webkitTextStrokeColor);
    const fnt = family(cs.fontFamily), bold = parseInt(cs.fontWeight) >= 600 || cs.fontWeight === 'bold', italic = cs.fontStyle === 'italic';
    window.__fontUse = window.__fontUse || [];
    if (!window.__fontUse.some((u) => u.family === fnt && u.bold === bold && u.italic === italic)) window.__fontUse.push({ family: fnt, bold, italic });
    return {
      highlight: hl,
      stroke: sc ? { size: sw, color: sc.hex } : undefined,
      font: fnt,
      size: parseFloat(cs.fontSize),
      bold,
      italic,
      underline: cs.textDecorationLine.includes('underline'),
      strike: cs.textDecorationLine.includes('line-through'),
      color: col ? col.hex : '000000',
      alpha: col ? col.a : 1,
      spacing: cs.letterSpacing === 'normal' ? 0 : parseFloat(cs.letterSpacing) || 0,
    };
  };
  // a <filter> with exactly one primitive that has a CSS/native equivalent -> {blur} | {shadow} | {colors}; else null
  // colour pipeline pieces: {mat:[20 numbers], lin} on RGBA (0..1), or {ch:0..3, f, lin} per-channel transfer (3 = alpha)
  const clamp01 = (x) => Math.max(0, Math.min(1, x));
  const toLin = (c) => (c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
  const toSrgb = (c) => (c <= 0.0031308 ? 12.92 * c : 1.055 * c ** (1 / 2.4) - 0.055);
  const m3 = (M) => [M[0], M[1], M[2], 0, 0, M[3], M[4], M[5], 0, 0, M[6], M[7], M[8], 0, 0, 0, 0, 0, 1, 0];
  // CSS filter functions (always sRGB)
  const cssOps = (fn, v) => {
    if (fn === 'grayscale') { const k = 1 - Math.min(1, v); return [{ mat: m3([0.2126 + 0.7874 * k, 0.7152 - 0.7152 * k, 0.0722 - 0.0722 * k, 0.2126 - 0.2126 * k, 0.7152 + 0.2848 * k, 0.0722 - 0.0722 * k, 0.2126 - 0.2126 * k, 0.7152 - 0.7152 * k, 0.0722 + 0.9278 * k]), lin: false }]; }
    if (fn === 'sepia') { const k = 1 - Math.min(1, v); return [{ mat: m3([0.393 + 0.607 * k, 0.769 - 0.769 * k, 0.189 - 0.189 * k, 0.349 - 0.349 * k, 0.686 + 0.314 * k, 0.168 - 0.168 * k, 0.272 - 0.272 * k, 0.534 - 0.534 * k, 0.131 + 0.869 * k]), lin: false }]; }
    if (fn === 'saturate') return [{ mat: m3([0.213 + 0.787 * v, 0.715 - 0.715 * v, 0.072 - 0.072 * v, 0.213 - 0.213 * v, 0.715 + 0.285 * v, 0.072 - 0.072 * v, 0.213 - 0.213 * v, 0.715 - 0.715 * v, 0.072 + 0.928 * v]), lin: false }];
    if (fn === 'hue-rotate') { const c = Math.cos(v * Math.PI / 180), n = Math.sin(v * Math.PI / 180); return [{ mat: m3([0.213 + c * 0.787 - n * 0.213, 0.715 - c * 0.715 - n * 0.715, 0.072 - c * 0.072 + n * 0.928, 0.213 - c * 0.213 + n * 0.143, 0.715 + c * 0.285 + n * 0.14, 0.072 - c * 0.072 - n * 0.283, 0.213 - c * 0.213 - n * 0.787, 0.715 - c * 0.715 + n * 0.715, 0.072 + c * 0.928 + n * 0.072]), lin: false }]; }
    if (fn === 'brightness') return [0, 1, 2].map((ch) => ({ ch, f: (x) => x * v, lin: false }));
    if (fn === 'contrast') return [0, 1, 2].map((ch) => ({ ch, f: (x) => (x - 0.5) * v + 0.5, lin: false }));
    if (fn === 'invert') { const k = Math.min(1, v); return [0, 1, 2].map((ch) => ({ ch, f: (x) => x * (1 - 2 * k) + k, lin: false })); }
    if (fn === 'opacity') return [{ ch: 3, f: (x) => x * v, lin: false }];
    return [];
  };
  // one feFuncR/G/B/A transfer function -> f(x), or null when unsupported
  const transferFn = (node) => {
    const type = node.getAttribute('type') || 'identity';
    const n = (a, d) => (node.getAttribute(a) === null ? d : parseFloat(node.getAttribute(a)));
    if (type === 'identity') return (x) => x;
    if (type === 'linear') { const sl = n('slope', 1), ic = n('intercept', 0); return (x) => sl * x + ic; }
    if (type === 'gamma') { const A = n('amplitude', 1), E = n('exponent', 1), O = n('offset', 0); return (x) => A * Math.pow(x, E) + O; }
    if (type === 'table' || type === 'discrete') {
      const tv = (node.getAttribute('tableValues') || '').split(/[\s,]+/).filter((q) => q !== '').map(Number);
      if (!tv.length) return (x) => x;
      const m = tv.length;
      if (type === 'discrete') return (x) => tv[Math.min(m - 1, Math.floor(x * m))];
      return (x) => { if (m === 1) return tv[0]; const k = Math.min(m - 2, Math.floor(x * (m - 1))); return tv[k] + (x * (m - 1) - k) * (tv[k + 1] - tv[k]); };
    }
    return null;
  };
  // a <filter> element -> {blur, shadow, ops} that PowerPoint draws natively, or null (picture)
  const svgFilterAsCss = (f) => {
    if (!f || f.tagName.toLowerCase() !== 'filter') return null;
    const prims = [...f.children];
    const colourOf = (node) => rgba(node.getAttribute('flood-color') || 'rgb(0,0,0)') || { hex: '000000', a: 1 };
    const numOf = (node, a, d = 0) => (node.getAttribute(a) === null ? d : parseFloat(node.getAttribute(a).split(/[\s,]+/)[0]));
    // SVG filters work in linearRGB unless color-interpolation-filters says sRGB (CSS shorthands always use sRGB)
    const ciOf = (node) => { const e = node.closest('[color-interpolation-filters]'); return e ? e.getAttribute('color-interpolation-filters') : null; };
    const lin = ciOf(f) !== 'sRGB';
    if (prims.length === 1 && prims[0].tagName === 'feDropShadow') {
      const k = prims[0], col = colourOf(k);
      return { blur: 0, ops: [], shadow: { color: col.hex, a: numOf(k, 'flood-opacity', 1) * col.a, ox: numOf(k, 'dx', 2), oy: numOf(k, 'dy', 2), blur: numOf(k, 'stdDeviation', 0) * 2 } };
    }
    // the same drop shadow spelled out: blur(SourceAlpha) -> offset -> flood -> composite -> merge(SourceGraphic)
    const blurP = prims.find((p) => p.tagName === 'feGaussianBlur' && p.getAttribute('in') === 'SourceAlpha');
    const offP = prims.find((p) => p.tagName === 'feOffset');
    const floodP = prims.find((p) => p.tagName === 'feFlood');
    const mergeP = prims.find((p) => p.tagName === 'feMerge');
    if (blurP && offP && floodP && mergeP && prims.every((p) => ['feGaussianBlur', 'feOffset', 'feFlood', 'feComposite', 'feMerge'].includes(p.tagName))) {
      const merged = [...mergeP.children].map((c) => c.getAttribute('in'));
      if (merged.includes('SourceGraphic') && offP.getAttribute('in') !== 'SourceGraphic') {
        const col = colourOf(floodP);
        return { blur: 0, ops: [], shadow: { color: col.hex, a: numOf(floodP, 'flood-opacity', 1) * col.a, ox: numOf(offP, 'dx', 0), oy: numOf(offP, 'dy', 0), blur: numOf(blurP, 'stdDeviation', 0) * 2 } };
      }
    }
    // general chain of colour primitives plus a blur of the source graphic
    const out = { blur: 0, shadow: null, ops: [] };
    let prev = null;
    for (const p of prims) {
      const inA = p.getAttribute('in');
      if (inA !== null && inA !== 'SourceGraphic' && inA !== prev) return null;
      const t = p.tagName;
      if (t === 'feGaussianBlur') out.blur += numOf(p, 'stdDeviation');
      else if (t === 'feColorMatrix') {
        const type = p.getAttribute('type') || 'matrix';
        const vals = (p.getAttribute('values') || '').split(/[\s,]+/).filter((q) => q !== '').map(Number);
        if (type === 'matrix') { if (vals.length === 20) out.ops.push({ mat: vals, lin }); else if (vals.length) return null; }
        else if (type === 'saturate') out.ops.push({ mat: cssOps('saturate', isNaN(vals[0]) ? 1 : vals[0])[0].mat, lin });
        else if (type === 'hueRotate') out.ops.push({ mat: cssOps('hue-rotate', isNaN(vals[0]) ? 0 : vals[0])[0].mat, lin });
        else if (type === 'luminanceToAlpha') out.ops.push({ mat: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0.2125, 0.7154, 0.0721, 0, 0], lin });
        else return null;
      } else if (t === 'feComponentTransfer') {
        for (const c of p.children) {
          const ch = { feFuncR: 0, feFuncG: 1, feFuncB: 2, feFuncA: 3 }[c.tagName];
          if (ch === undefined) return null;
          const fn = transferFn(c);
          if (!fn) return null;
          out.ops.push({ ch, f: fn, lin });
        }
      } else return null; // feTurbulence, feFlood, feImage, ... are not colour pipelines
      prev = p.getAttribute('result') || '\u0000';
    }
    if (!out.blur && !out.ops.length) return null;
    return out;
  };
  // applies colour operations to one colour: hex + alpha (0..1) -> {hex, a}
  const applyOps = (hex, alpha, ops) => {
    if (!ops.length) return { hex, a: alpha };
    let c = [0, 2, 4].map((i) => parseInt(hex.substr(i, 2), 16) / 255), a = alpha, lin = false;
    for (const op of ops) {
      if (op.lin !== lin) { c = c.map(op.lin ? toLin : toSrgb); lin = op.lin; }
      if (op.mat) {
        const v = [c[0], c[1], c[2], a], M = op.mat, o = [];
        for (let r = 0; r < 4; r++) o.push(M[r * 5] * v[0] + M[r * 5 + 1] * v[1] + M[r * 5 + 2] * v[2] + M[r * 5 + 3] * v[3] + M[r * 5 + 4]);
        c = o.slice(0, 3).map(clamp01); a = clamp01(o[3]);
      } else if (op.ch === 3) a = clamp01(op.f(a));
      else c[op.ch] = clamp01(op.f(clamp01(c[op.ch])));
    }
    if (lin) c = c.map(toSrgb);
    return { hex: c.map((x) => Math.round(clamp01(x) * 255).toString(16).padStart(2, '0')).join('').toUpperCase(), a: clamp01(a) };
  };
  // separable blend functions (values 0..1), per channel
  const blendChannel = (mode, b, s) => {
    const screen = (x, y) => x + y - x * y;
    const hard = (x, y) => (y <= 0.5 ? x * 2 * y : screen(x, 2 * y - 1));
    switch (mode) {
      case 'multiply': return b * s;
      case 'screen': return screen(b, s);
      case 'overlay': return hard(s, b);
      case 'darken': return Math.min(b, s);
      case 'lighten': return Math.max(b, s);
      case 'color-dodge': return b === 0 ? 0 : s >= 1 ? 1 : Math.min(1, b / (1 - s));
      case 'color-burn': return b >= 1 ? 1 : s <= 0 ? 0 : 1 - Math.min(1, (1 - b) / s);
      case 'hard-light': return hard(b, s);
      case 'soft-light': {
        const D = b <= 0.25 ? ((16 * b - 12) * b + 4) * b : Math.sqrt(b);
        return s <= 0.5 ? b - (1 - 2 * s) * b * (1 - b) : b + (2 * s - 1) * (D - b);
      }
      case 'difference': return Math.abs(b - s);
      case 'exclusion': return b + s - 2 * b * s;
      default: return s;
    }
  };
  // non-separable blend modes (W3C compositing): operate on whole RGB triples (0..1)
  const lumOf = (c) => 0.3 * c[0] + 0.59 * c[1] + 0.11 * c[2];
  const clipColor = (c) => {
    const l = lumOf(c), n = Math.min(...c), x = Math.max(...c);
    let r = c.slice();
    if (n < 0) r = r.map((v) => l + ((v - l) * l) / (l - n));
    if (x > 1) r = r.map((v) => l + ((v - l) * (1 - l)) / (x - l));
    return r;
  };
  const setLum = (c, l) => { const d = l - lumOf(c); return clipColor(c.map((v) => v + d)); };
  const satOf = (c) => Math.max(...c) - Math.min(...c);
  const setSat = (c, sv) => {
    const mx = Math.max(...c), mn = Math.min(...c);
    return mx > mn ? c.map((v) => ((v - mn) * sv) / (mx - mn)) : [0, 0, 0];
  };
  const nonSeparable = {
    hue: (B, S) => setLum(setSat(S, satOf(B)), lumOf(B)),
    saturation: (B, S) => setLum(setSat(B, satOf(S)), lumOf(B)),
    color: (B, S) => setLum(S, lumOf(B)),
    luminosity: (B, S) => setLum(B, lumOf(S)),
  };

  // A blended solid box is drawn in two layers: (1) its own colour blended against the slide background, (2) for each
  // solid box beneath it, the overlap rectangle coloured with the blend against that box (lowest first, so the higher
  // box wins where they overlap). PowerPoint has no blend modes; this is exact for opaque rectangular boxes.
  const applyBlends = (items, bgHex) => {
    const extra = [];
    const hex = (c) => [0, 2, 4].map((i) => parseInt(c.substr(i, 2), 16) / 255);
    const toHex = (v) => v.map((x) => Math.round(Math.max(0, Math.min(1, x)) * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
    const compose = (mode, B, S, a) => {
      if (nonSeparable[mode]) { const F = nonSeparable[mode](B, S); return B.map((bv, k) => (1 - a) * bv + a * F[k]); }
      return B.map((bv, k) => (1 - a) * bv + a * blendChannel(mode, bv, S[k]));
    };
    const bgRGB = hex(bgHex || 'FFFFFF');
    items.forEach((it, idx) => {
      if (!it.blend || it.rot || it.type !== 'shape') return;
      const S = hex(it.blend.color), a = it.blend.a, mode = it.blend.mode;
      it.fill = { color: toHex(compose(mode, bgRGB, S, a)), transparency: 0 };
      const unders = [];
      for (let j = 0; j < idx; j++) {
        const b = items[j];
        if (b.type !== 'shape' || !b.fill || b.fill.grad || b.fill.transparency >= 100 || b.rot || b.poly || b.blend) continue;
        if (b.z > it.z) continue;
        const x0 = Math.max(it.x, b.x), y0 = Math.max(it.y, b.y), x1 = Math.min(it.x + it.w, b.x + b.w), y1 = Math.min(it.y + it.h, b.y + b.h);
        if (x1 - x0 > 0.5 && y1 - y0 > 0.5) unders.push({ b, rect: { x: x0, y: y0, w: x1 - x0, h: y1 - y0 } });
      }
      unders.sort((p, q) => p.b.z - q.b.z || p.b.order - q.b.order);
      unders.forEach(({ b, rect }, k) => {
        const ab = 1 - b.fill.transparency / 100;
        const Bc = hex(b.fill.color).map((c, i) => ab * c + (1 - ab) * bgRGB[i]);
        extra.push({ type: 'shape', ...rect, radius: 0, radii: [0, 0, 0, 0], rot: 0, shadow: null, line: null, fill: { color: toHex(compose(mode, Bc, S, a)), transparency: 0 }, z: it.z, order: it.order + 0.5 + k * 0.01 });
      });
    });
    items.push(...extra);
  };
  const xform = (t, cs) => {
    if (cs.textTransform === 'uppercase') return t.toUpperCase();
    if (cs.textTransform === 'lowercase') return t.toLowerCase();
    return t;
  };
  // text of one node with whitespace collapsed; with opts.hard, "\n" is inserted where the browser wrapped
  const textOf = (n, cs, st) => {
    const raw = n.textContent;
    if (/^(pre|pre-wrap)/.test(cs.whiteSpace)) return xform(raw, cs);
    if (!opts.hard) return xform(raw.replace(/\s+/g, ' '), cs);
    const rg = document.createRange();
    const fs = parseFloat(cs.fontSize);
    let out = '';
    for (let i = 0; i < raw.length; i++) {
      const c = raw[i];
      if (/\s/.test(c)) { if (!st.ws) out += ' '; st.ws = true; continue; }
      rg.setStart(n, i); rg.setEnd(n, i + 1);
      const top = rg.getBoundingClientRect().top;
      if (st.top !== null && top - st.top > fs * 0.6) { out = out.replace(/ +$/, '') + '\n'; st.top = top; }
      else if (st.top === null) st.top = top;
      out += c; st.ws = false;
    }
    return xform(out, cs);
  };
  const runsOf = (nodes, base) => {
    const runs = [];
    const st = { top: null, ws: true };
    const walk = (list, cs, hl, tg) => {
      for (const n of list) {
        if (n.nodeType === 3) {
          const t = textOf(n, cs, st);
          if (t) runs.push({ text: t, cs: tg ? { ...snap(cs, hl), color: tg.stops[0].color, alpha: 1 } : snap(cs, hl), tg });
        } else if (n.nodeType === 1) {
          if (n.tagName === 'BR') { runs.push({ text: '\n', cs: snap(cs, hl) }); st.top = null; st.ws = true; }
          else {
            const ncs = getComputedStyle(n);
            const bgc = rgba(ncs.backgroundColor);
            const ntg = ncs.backgroundImage !== 'none' && /text/.test(ncs.webkitBackgroundClip || ncs.backgroundClip) ? parseGradient(ncs.backgroundImage) : tg;
            const nr = n.getBoundingClientRect();
            // runs have no margins: an inline element's horizontal margin becomes spaces of about that width
            const gap = (px) => ' '.repeat(Math.max(1, Math.round(px / (parseFloat(ncs.fontSize) * 0.28))));
            const ml = parseFloat(ncs.marginLeft), mr = parseFloat(ncs.marginRight);
            if (ml >= 1) runs.push({ text: gap(ml), cs: snap(ncs, hl) });
            walk(n.childNodes, ncs, bgc ? bgc.hex : hl, ntg ? { ...ntg, span: { x0: nr.left, x1: nr.right } } : undefined);
            if (mr >= 1) runs.push({ text: gap(mr), cs: snap(ncs, hl) });
          }
        }
      }
    };
    walk(nodes, base);
    runs.forEach((r, i) => { if (r.text.startsWith('\n') && i && runs[i - 1].text.endsWith(' ')) runs[i - 1].text = runs[i - 1].text.replace(/ +$/, ''); });
    if (runs.length) {
      runs[0].text = runs[0].text.replace(/^[ \t]+/, '');
      runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/\s+$/, '');
    }
    // explicit lines: pptxgenjs splits "\n" inside a run unreliably (leading indentation became blank paragraphs),
    // so each line is emitted as its own runs and the last run of a line carries breakLine
    const lines = [[]];
    let styleRun = null;
    for (const r of runs) {
      r.text.split('\n').forEach((p, i) => {
        if (i > 0) lines.push([]);
        if (p) lines[lines.length - 1].push({ ...r, text: p });
        styleRun = r;
      });
    }
    if (!styleRun) return [];
    if (lines.length > 1 && !lines[lines.length - 1].length) lines.pop();
    const out = [];
    lines.forEach((ln, li) => {
      if (!ln.length) ln.push({ ...styleRun, text: ' ' });
      ln[ln.length - 1].brk = li < lines.length - 1;
      out.push(...ln);
    });
    return out.filter((r) => r.text);
  };

  return slides.map((slide, si) => {
    const sr = slide.getBoundingClientRect();
    const items = [];
    const phOk = (opts.ph && opts.ph[si]) || {}; // body / subtitle placeholders the slide's template layout has
    let phBody = null; // data-ph="body" element being visited: its list items get the placeholder's bullets, not "•"
    const notes = [...slide.querySelectorAll('.notes')].map((n) => n.textContent.trim()).join('\n');
    const scs = getComputedStyle(slide);
    const bg = rgba(scs.backgroundColor);
    let curZ = 0;
    const push = (it) => { it.z = curZ; it.order = items.length; items.push(it); return it; };
    const slideGrad = scs.backgroundImage !== 'none' && parseGradient(scs.backgroundImage);
    if (slideGrad) {
      push({ type: 'shape', bgLayer: true, x: 0, y: 0, w: W(), h: H(), radius: 0, rot: 0, shadow: null, line: null, name: 'hp' + ++rid,
        fill: { color: slideGrad.stops[0].color, transparency: 0, grad: slideGrad } });
    } else if (scs.backgroundImage !== 'none') {
      slide.dataset.pptRid = 'r' + ++rid;
      push({ type: 'raster', bgLayer: true, id: slide.dataset.pptRid, slide: true, x: 0, y: 0, w: W(), h: H(), hideChildren: true, op: 1 });
    }
    function W() { return 1280; } function H() { return 720; }
    const rel = (r) => ({ x: r.left - sr.left, y: r.top - sr.top, w: r.width, h: r.height });
    const SVG_PROPS = ['fill', 'fill-opacity', 'stroke', 'stroke-width', 'stroke-opacity', 'stroke-dasharray', 'stroke-linecap', 'stroke-linejoin', 'opacity', 'font-family', 'font-size', 'font-weight', 'text-anchor', 'stop-color', 'stop-opacity'];
    // standalone SVG markup with computed paint inlined, so it renders as a vector picture without the page's CSS
    const svgMarkup = (svg) => {
      if ([...svg.querySelectorAll('image')].some((i) => !/^data:/.test(i.getAttribute('href') || i.getAttribute('xlink:href') || ''))) return null;
      const clone = svg.cloneNode(true);
      const a = [svg, ...svg.querySelectorAll('*')], b = [clone, ...clone.querySelectorAll('*')];
      a.forEach((n, i) => { const c = getComputedStyle(n); b[i].setAttribute('style', SVG_PROPS.map((p) => `${p}:${c.getPropertyValue(p)}`).join(';')); });
      clone.querySelectorAll('style').forEach((n) => n.remove()); // computed paint is inlined, class rules no longer needed
      clone.querySelectorAll('foreignObject').forEach((n) => n.remove()); // HTML inside is emitted as native text/shapes
      // PowerPoint's SVG renderer draws these filter primitives correctly; feDropShadow it drops entirely, so that one is
      // rewritten as blur + offset + colour matrix + merge. Any other primitive means the SVG goes in as a picture.
      const OK = new Set(['feGaussianBlur', 'feOffset', 'feMerge', 'feMergeNode', 'feColorMatrix', 'feFlood', 'feComposite']);
      const NS = 'http://www.w3.org/2000/svg';
      for (const f of clone.querySelectorAll('filter')) {
        for (const c of [...f.children]) {
          if (c.tagName === 'feDropShadow') {
            const num = (a, d) => (c.getAttribute(a) === null ? d : parseFloat(c.getAttribute(a)));
            const col = rgba(c.style.getPropertyValue('flood-color') || c.getAttribute('flood-color') || 'rgb(0,0,0)') || { hex: '000000', a: 1 };
            const op = num('flood-opacity', 1) * col.a;
            const mk = (tag, attrs) => { const e = document.createElementNS(NS, tag); Object.entries(attrs).forEach(([k, v]) => e.setAttribute(k, v)); return e; };
            const rgbN = [0, 2, 4].map((i) => (parseInt(col.hex.substr(i, 2), 16) / 255).toFixed(4));
            const chain = [
              mk('feGaussianBlur', { in: 'SourceAlpha', stdDeviation: num('stdDeviation', 2), result: 'ds_b' }),
              mk('feOffset', { in: 'ds_b', dx: num('dx', 2), dy: num('dy', 2), result: 'ds_o' }),
              mk('feColorMatrix', { in: 'ds_o', type: 'matrix', values: `0 0 0 0 ${rgbN[0]} 0 0 0 0 ${rgbN[1]} 0 0 0 0 ${rgbN[2]} 0 0 0 ${op} 0`, result: 'ds_c' }),
              (() => { const m = mk('feMerge', {}); m.appendChild(mk('feMergeNode', { in: 'ds_c' })); m.appendChild(mk('feMergeNode', { in: 'SourceGraphic' })); return m; })(),
            ];
            chain.forEach((n) => f.insertBefore(n, c)); c.remove();
          } else if (c.tagName === 'feFlood') {
            // PowerPoint applies flood-opacity differently: flood opaque, then scale alpha with a colour matrix
            const o = parseFloat(c.getAttribute('flood-opacity') ?? c.style.getPropertyValue('flood-opacity') ?? '1');
            if (!isNaN(o) && o < 1) {
              const res = c.getAttribute('result');
              c.setAttribute('flood-opacity', '1'); c.style.setProperty('flood-opacity', '1');
              if (res) c.setAttribute('result', res + '_raw');
              const m = document.createElementNS(NS, 'feColorMatrix');
              m.setAttribute('type', 'matrix'); m.setAttribute('values', `1 0 0 0 0 0 1 0 0 0 0 0 1 0 0 0 0 0 ${o} 0`);
              if (res) { m.setAttribute('in', res + '_raw'); m.setAttribute('result', res); }
              c.after(m);
            }
          } else if (!OK.has(c.tagName)) return null;
        }
      }
      const r = svg.getBoundingClientRect();
      clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
      clone.setAttribute('width', r.width); clone.setAttribute('height', r.height);
      if (!clone.getAttribute('viewBox') && svg.viewBox && svg.viewBox.baseVal && svg.viewBox.baseVal.width) { const v = svg.viewBox.baseVal; clone.setAttribute('viewBox', `${v.x} ${v.y} ${v.width} ${v.height}`); }
      return new XMLSerializer().serializeToString(clone);
    };
    const raster = (el, box, extra = {}) => {
      el.dataset.pptRid = 'r' + ++rid;
      const svg = el.tagName.toLowerCase() === 'svg' && !extra.iso ? svgMarkup(el) : null;
      push({ type: 'raster', id: el.dataset.pptRid, ...box, op: opacityOf(el, slide), svg, ...extra });
    };
    const inside = (b, c) => b.x >= c.x - 1 && b.y >= c.y - 1 && b.x + b.w <= c.x + c.w + 1 && b.y + b.h <= c.y + c.h + 1;
    const intersect = (a, b) => {
      if (!b) return a;
      const x = Math.max(a.x, b.x), y = Math.max(a.y, b.y);
      return { x, y, w: Math.max(0, Math.min(a.x + a.w, b.x + b.w) - x), h: Math.max(0, Math.min(a.y + a.h, b.y + b.h) - y) };
    };
    const ownTextOutside = (el, clip) => {
      const rg = document.createRange();
      for (const n of el.childNodes) {
        if (n.nodeType !== 3 || !n.textContent.trim()) continue;
        rg.selectNodeContents(n);
        if (!inside(rel(rg.getBoundingClientRect()), clip)) return true;
      }
      return false;
    };
    const padOf = (cs) => ({
      l: parseFloat(cs.paddingLeft) + parseFloat(cs.borderLeftWidth), r: parseFloat(cs.paddingRight) + parseFloat(cs.borderRightWidth),
    });
    let curTextGrad = null, curFsh = null; // gradient-filled text / drop-shadow filter of the element being visited
    const textShadowOf = (cs, op) => {
      const s = parseShadow(cs.textShadow) || curFsh;
      return s && { color: s.color, opacity: Math.min(1, s.a * op), blur: s.blur, offset: Math.hypot(s.ox, s.oy), angle: (Math.atan2(s.oy, s.ox) * 180 / Math.PI + 360) % 360 };
    };

    // per-character glyph origins (slide px, baseline) straight from the browser layout, for vector-outline text
    const glyphsOf = (nodes) => {
      const out = [], rg = document.createRange();
      const walk = (list) => {
        for (const n of list) {
          if (n.nodeType === 3) {
            const ecs = getComputedStyle(n.parentElement), st = snap(ecs), raw = n.textContent;
            fontProbe.font = `${st.italic ? 'italic ' : ''}${st.bold ? 'bold ' : ''}${st.size}px "${st.font}"`;
            const asc = fontProbe.measureText('Hg').fontBoundingBoxAscent;
            for (let i = 0; i < raw.length; i++) {
              if (/\s/.test(raw[i])) continue;
              rg.setStart(n, i); rg.setEnd(n, i + 1);
              const r = rg.getBoundingClientRect();
              if (!r.width && !r.height) continue;
              out.push({ c: xform(raw[i], ecs), x: r.left - sr.left, y: r.top - sr.top + asc, size: st.size, bold: st.bold, italic: st.italic, font: st.font, color: st.color, alpha: st.alpha });
            }
          } else if (n.nodeType === 1 && n.tagName !== 'BR') walk(n.childNodes);
        }
      };
      walk(nodes);
      return out;
    };

    // one text item for a run of sibling nodes inside an element
    const emitText = (cs, nodes, box, op, markerEl) => {
      const runs = runsOf(nodes, cs);
      let markerW = 0; // the browser draws the marker outside the content box; PowerPoint text starts at the marker
      if (markerEl && runs.length && getComputedStyle(markerEl).listStyleType !== 'none' && !(phBody && phBody.contains(markerEl))) {
        const lst = markerEl.parentElement;
        const idx = [...lst.children].indexOf(markerEl) + 1;
        const mtext = lst.tagName === 'OL' ? idx + '. ' : '•  ';
        runs.unshift({ text: mtext, cs: runs[0].cs });
        if (getComputedStyle(markerEl).listStylePosition !== 'inside') {
          fontProbe.font = `${runs[0].cs.bold ? 'bold' : 'normal'} ${runs[0].cs.size}px "${runs[0].cs.font}"`;
          markerW = fontProbe.measureText(mtext).width;
        }
      }
      if (!runs.length) return;
      const range = document.createRange();
      range.setStartBefore(nodes[0]); range.setEndAfter(nodes[nodes.length - 1]);
      const rects = [...range.getClientRects()].filter((r) => r.width > 0 && r.height > 0);
      const tr = rel(range.getBoundingClientRect());
      const lineTops = new Set(rects.map((r) => Math.round(r.top / 4)));
      const fs = parseFloat(cs.fontSize);
      const lh = cs.lineHeight === 'normal' ? fs * 1.2 : parseFloat(cs.lineHeight);
      const hard = opts.hard && runs.some((r) => r.text.includes('\n'));
      const single = (lineTops.size <= 1 && !runs.some((r) => r.text.includes('\n')));
      const align = cs.textAlign === 'start' ? 'left' : cs.textAlign === 'end' ? 'right' : cs.textAlign;
      const p = padOf(cs);
      // real text rect keeps flex/grid centering; block width for wrapped/aligned text
      const t = single && align === 'left' ? { x: tr.x, y: tr.y, w: tr.w, h: tr.h }
        : { x: box.x + p.l, y: tr.y, w: box.w - p.l - p.r, h: tr.h };
      // group opacity: text on a filled box that fades with it is "text over fill" first, then the group blends with the
      // backdrop, so the text's final colour is mix(text, backdrop) and must stay opaque; loose text just fades.
      let boxed = false;
      if (op < 1) {
        for (let n = nodes[0].parentElement; n && n !== slide.parentElement; n = n.parentElement) {
          const ncs = getComputedStyle(n);
          if (rgba(ncs.backgroundColor) || ncs.backgroundImage !== 'none') { boxed = true; break; }
          if (parseFloat(ncs.opacity) < 1) break;
        }
      }
      // backdrop colour behind the text: slide colour, or the slide gradient evaluated at the text centre
      let back = (bg && bg.hex) || 'FFFFFF';
      if (slideGrad) {
        const g = slideGrad, cx = t.x + t.w / 2, cy = t.y + t.h / 2;
        let u;
        if (g.kind === 'linear') {
          const a = g.ang * Math.PI / 180, dx = Math.cos(a), dy = Math.sin(a);
          u = ((cx - 640) * dx + (cy - 360) * dy) / (Math.abs(1280 * dx) + Math.abs(720 * dy)) + 0.5;
        } else u = Math.hypot(cx - 640, cy - 360) / Math.hypot(640, 360);
        u = Math.max(0, Math.min(1, u)) * 100;
        let k = 1; while (k < g.stops.length - 1 && g.stops[k].pos < u) k++;
        const s0 = g.stops[k - 1], s1 = g.stops[k], f = s1.pos === s0.pos ? 0 : Math.max(0, Math.min(1, (u - s0.pos) / (s1.pos - s0.pos)));
        back = [0, 2, 4].map((i) => Math.round(parseInt(s0.color.substr(i, 2), 16) * (1 - f) + parseInt(s1.color.substr(i, 2), 16) * f).toString(16).padStart(2, '0')).join('').toUpperCase();
      }
      const mix = (c) => [0, 2, 4].map((i) => Math.round(parseInt(c.substr(i, 2), 16) * op + parseInt(back.substr(i, 2), 16) * (1 - op)).toString(16).padStart(2, '0')).join('').toUpperCase();
      if (markerW) { t.x -= markerW; t.w += markerW; }
      // ink box of a single line (PowerPoint's text warp fits the glyph ink, not the line box)
      let ink = null;
      if (single) {
        const r0 = runs[0].cs;
        fontProbe.font = `${r0.italic ? 'italic ' : ''}${r0.bold ? 'bold ' : ''}${r0.size}px "${r0.font}"`;
        const mm = fontProbe.measureText(runs.map((r) => r.text).join(''));
        ink = { x: tr.x - mm.actualBoundingBoxLeft, y: tr.y + mm.fontBoundingBoxAscent - mm.actualBoundingBoxAscent, w: mm.actualBoundingBoxRight + mm.actualBoundingBoxLeft, h: mm.actualBoundingBoxAscent + mm.actualBoundingBoxDescent };
      }
      const tg = curTextGrad;
      const topRun = runs.reduce((a, r) => (r.cs.size > a.cs.size ? r : a), runs[0]);
      fontProbe.font = `${topRun.cs.italic ? 'italic ' : ''}${topRun.cs.bold ? 'bold ' : ''}${topRun.cs.size}px "${topRun.cs.font}"`;
      const asc = fontProbe.measureText('Hg').fontBoundingBoxAscent;
      push({ type: 'text', ...t, asc, tr, ink, gl: markerW ? null : glyphsOf(nodes), tm: [1, 0, 0, 1, 0, 0], align, lineHeight: lh, single, hard, shadow: textShadowOf(cs, boxed ? 1 : op), rot: 0,
        name: (tg || runs.some((r) => r.tg)) ? 'hp' + ++rid : undefined, tgrad: tg || undefined,
        tgRuns: runs.filter((r) => r.tg).map((r) => ({ text: r.text, grad: { ...r.tg, span: r.tg.span ? { x0: r.tg.span.x0 - sr.left, x1: r.tg.span.x1 - sr.left } : null } })),
        tbox: { x: t.x, w: t.w },
        runs: runs.map((r) => (tg ? { text: r.text, brk: r.brk, ...r.cs, color: tg.stops[0].color, alpha: 1 }
          : boxed ? { text: r.text, brk: r.brk, ...r.cs, color: mix(r.cs.color) } : { text: r.text, brk: r.brk, ...r.cs, alpha: r.cs.alpha * op })) });
    };

    // CSS multi-column text -> one native text box per column
    const emitColumns = (el, cs, box, op) => {
      const n = parseInt(cs.columnCount);
      const gap = cs.columnGap === 'normal' ? parseFloat(cs.fontSize) : parseFloat(cs.columnGap);
      const p = padOf(cs);
      const left = box.x + p.l, colW = (box.w - p.l - p.r - gap * (n - 1)) / n;
      const cols = Array.from({ length: n }, () => ({ runs: [], top: Infinity, bottom: -Infinity, block: null, ws: true }));
      const rg = document.createRange();
      const tw = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
      const blockOf = (t) => { let e = t.parentElement; while (e !== el && getComputedStyle(e).display === 'inline') e = e.parentElement; return e; };
      for (let t; (t = tw.nextNode());) {
        const ecs = getComputedStyle(t.parentElement);
        const style = snap(ecs);
        const raw = t.textContent;
        for (let i = 0; i < raw.length; i++) {
          rg.setStart(t, i); rg.setEnd(t, i + 1);
          const r = rg.getBoundingClientRect();
          if (!r.width && !r.height) continue;
          const rr = rel(r);
          const ci = Math.max(0, Math.min(n - 1, Math.floor((rr.x - left + 2) / (colW + gap))));
          const c = cols[ci];
          const blk = blockOf(t);
          const ws = /\s/.test(raw[i]);
          if (c.block && c.block !== blk) { c.runs.push({ text: '\n', ...style }); c.ws = true; }
          c.block = blk;
          if (ws && c.ws) continue;
          c.ws = ws;
          const ch = xform(ws ? ' ' : raw[i], ecs);
          const last = c.runs[c.runs.length - 1];
          if (last && last.node === t) last.text += ch; else c.runs.push({ node: t, text: ch, ...style });
          if (!ws) { c.top = Math.min(c.top, rr.y); c.bottom = Math.max(c.bottom, rr.y + rr.h); }
        }
      }
      const fs = parseFloat(cs.fontSize);
      const lh = cs.lineHeight === 'normal' ? fs * 1.2 : parseFloat(cs.lineHeight);
      cols.forEach((c, i) => {
        c.runs.forEach((r) => delete r.node);
        if (c.runs.length) { c.runs[0].text = c.runs[0].text.replace(/^\s+/, ''); c.runs[c.runs.length - 1].text = c.runs[c.runs.length - 1].text.replace(/\s+$/, ''); }
        const runs = c.runs.filter((r) => r.text);
        if (!runs.length) return;
        push({ type: 'text', x: left + i * (colW + gap), y: c.top, w: colW, h: c.bottom - c.top, align: cs.textAlign === 'start' ? 'left' : cs.textAlign,
          lineHeight: lh, single: false, hard: false, shadow: null, rot: 0, runs: runs.map((r) => ({ ...r, alpha: r.alpha * op })) });
      });
    };

    // apply a CSS transform matrix to the items produced for an element subtree
    const applyMatrix = (from, m, O) => {
      const sx = Math.hypot(m.a, m.b), sy = Math.hypot(m.c, m.d), s = Math.min(sx, sy), ang = Math.atan2(m.b, m.a) * 180 / Math.PI;
      const skew = Math.abs(m.a * m.c + m.b * m.d) > 0.01 * sx * sy;
      const tp = (x, y) => { const dx = x - O.x, dy = y - O.y; return [O.x + m.a * dx + m.c * dy + m.e, O.y + m.b * dx + m.d * dy + m.f]; };
      const cornersOf = (it) => {
        if (it.poly) return it.poly;
        const cx = it.x + it.w / 2, cy = it.y + it.h / 2, r = (it.rot || 0) * Math.PI / 180, co = Math.cos(r), si = Math.sin(r);
        return [[-1, -1], [1, -1], [1, 1], [-1, 1]].map(([a, b]) => { const dx = a * it.w / 2, dy = b * it.h / 2; return [cx + dx * co - dy * si, cy + dx * si + dy * co]; });
      };
      for (let i = from; i < items.length; i++) {
        const it = items[i];
        if (it.type === 'path') { it.cmds = it.cmds.map((c) => ({ op: c.op, pts: c.pts.map(([x, y]) => tp(x, y)) })); if (it.line) it.line.width *= s; continue; }
        if (skew && it.type === 'shape') { // skewed boxes become polygons (corner radii are dropped)
          const pts = cornersOf(it).map(([x, y]) => tp(x, y));
          const xs = pts.map((p) => p[0]), ys = pts.map((p) => p[1]);
          it.x = Math.min(...xs); it.y = Math.min(...ys); it.w = Math.max(...xs) - it.x; it.h = Math.max(...ys) - it.y;
          it.poly = pts; it.rot = 0; it.radius = 0; it.radii = [0, 0, 0, 0];
          if (it.line) it.line.width *= s;
          if (it.shadow) { it.shadow.blur *= s; it.shadow.offset *= s; }
          continue;
        }
        if (it.type === 'text' && it.gl) { // keep the glyph outlines' own affine in step with the box
          const A = [m.a, m.b, m.c, m.d, O.x - m.a * O.x - m.c * O.y + m.e, O.y - m.b * O.x - m.d * O.y + m.f], T = it.tm;
          it.tm = [A[0] * T[0] + A[2] * T[1], A[1] * T[0] + A[3] * T[1], A[0] * T[2] + A[2] * T[3], A[1] * T[2] + A[3] * T[3], A[0] * T[4] + A[2] * T[5] + A[4], A[1] * T[4] + A[3] * T[5] + A[5]];
          if (skew) it.outline = true; // slanted glyphs only exist as outlines
        }
        const cx = it.x + it.w / 2, cy = it.y + it.h / 2;
        const [nx, ny] = tp(cx, cy);
        it.w *= sx; it.h *= sy; it.x = nx - it.w / 2; it.y = ny - it.h / 2;
        it.rot = (it.rot || 0) + ang;
        if (it.type === 'text') {
          it.lineHeight *= s;
          if (skew && m.c < -0.1) it.runs.forEach((r) => { r.italic = true; }); // glyphs cannot be slanted: lean-right skew -> italic
          if (it.single && it.ink && it.ink.w > 0 && !it.warp && sx > s * 1.001) {
            // PowerPoint cannot stretch glyphs on its own: a single line is boxed to its stretched ink size and warped to fill it
            const k = it.ink, [wx, wy] = tp(k.x + k.w / 2, k.y + k.h / 2);
            it.w = k.w * sx; it.h = k.h * sy; it.x = wx - it.w / 2; it.y = wy - it.h / 2; it.noShift = true;
            it.warp = true; it.align = 'center'; it.name = it.name || 'hp' + ++rid;
            it.runs.forEach((r) => { r.size *= s; r.spacing *= s; });
          } else {
            const widen = sx > s ? (sx / s - 1) * 0.55 : 0; // wrapped text: widen glyph spacing instead
            it.runs.forEach((r) => { r.size *= s; r.spacing = r.spacing * s + widen * r.size; });
          }
          if (it.shadow) { it.shadow.blur *= s; it.shadow.offset *= s; }
        } else if (it.type === 'shape') {
          it.radius *= s;
          if (it.radii) it.radii = it.radii.map((v) => v * s);
          if (it.poly) it.poly = it.poly.map(([x, y]) => tp(x, y));
          if (it.line) it.line.width *= s;
          if (it.shadow) { it.shadow.blur *= s; it.shadow.offset *= s; }
        }
      }
    };

    // CSS colour filter functions applied to the colours of produced items
    const recolor = (it, ops) => {
      if (!ops.length) return;
      const hexOnly = (h) => applyOps(h, 1, ops).hex;
      if (it.type === 'shape') {
        if (it.fill && it.fill.grad) it.fill.grad.stops.forEach((st) => { const r = applyOps(st.color, st.alpha, ops); st.color = r.hex; st.alpha = r.a; });
        if (it.fill && !it.fill.grad) { const r = applyOps(it.fill.color, 1 - (it.fill.transparency || 0) / 100, ops); it.fill.color = r.hex; it.fill.transparency = Math.round((1 - r.a) * 100); }
        if (it.line) { const r = applyOps(it.line.color, 1 - (it.line.transparency || 0) / 100, ops); it.line.color = r.hex; it.line.transparency = Math.round((1 - r.a) * 100); }
        if (it.shadow) { const r = applyOps(it.shadow.color, it.shadow.opacity, ops); it.shadow.color = r.hex; it.shadow.opacity = r.a; }
        if (it.inner) { const r = applyOps(it.inner.color, it.inner.opacity, ops); it.inner.color = r.hex; it.inner.opacity = r.a; }
      } else if (it.type === 'path') {
        if (it.fill && it.fill.grad) it.fill.grad.stops.forEach((st) => { const r = applyOps(st.color, st.alpha, ops); st.color = r.hex; st.alpha = r.a; });
        if (it.fill && !it.fill.grad) { const r = applyOps(it.fill.color, 1 - (it.fill.transparency || 0) / 100, ops); it.fill.color = r.hex; it.fill.transparency = Math.round((1 - r.a) * 100); }
        if (it.line) { const r = applyOps(it.line.color, 1 - (it.line.transparency || 0) / 100, ops); it.line.color = r.hex; it.line.transparency = Math.round((1 - r.a) * 100); }
      } else if (it.type === 'text') {
        it.runs.forEach((r) => {
          const c = applyOps(r.color, r.alpha, ops); r.color = c.hex; r.alpha = c.a;
          if (r.highlight) r.highlight = hexOnly(r.highlight);
          if (r.stroke) r.stroke.color = hexOnly(r.stroke.color);
        });
        if (it.tgrad) it.tgrad.stops.forEach((st) => { const r = applyOps(st.color, st.alpha ?? 1, ops); st.color = r.hex; st.alpha = r.a; });
        if (it.shadow) { const r = applyOps(it.shadow.color, it.shadow.opacity, ops); it.shadow.color = r.hex; it.shadow.opacity = r.a; }
      }
    };

    const clips = [];
    // per-corner radii [tl, tr, br, bl] in px, clamped the way CSS does
    const radiiOf = (cs, b) => {
      const f = (v) => (/%/.test(v) ? (parseFloat(v) / 100) * Math.min(b.w, b.h) : parseFloat(v) || 0);
      const r = [cs.borderTopLeftRadius, cs.borderTopRightRadius, cs.borderBottomRightRadius, cs.borderBottomLeftRadius].map(f);
      const k = Math.min(1, b.w / ((r[0] + r[1]) || 1e9), b.w / ((r[3] + r[2]) || 1e9), b.h / ((r[0] + r[3]) || 1e9), b.h / ((r[1] + r[2]) || 1e9));
      return r.map((v) => v * k);
    };
    // does `b` stick out of a rounded clipping container's corner arc?
    const roundedCut = (b, cl) => {
      if (!cl || !cl.radii) return false;
      const R = cl.rect, r = cl.radii;
      const ctr = [[R.x + r[0], R.y + r[0]], [R.x + R.w - r[1], R.y + r[1]], [R.x + R.w - r[2], R.y + R.h - r[2]], [R.x + r[3], R.y + R.h - r[3]]];
      for (let k = 0; k < 4; k++) {
        if (!(r[k] > 0)) continue;
        const px = k === 0 || k === 3 ? b.x : b.x + b.w, py = k < 2 ? b.y : b.y + b.h;
        const inSquare = (k === 0 || k === 3 ? px < ctr[k][0] : px > ctr[k][0]) && (k < 2 ? py < ctr[k][1] : py > ctr[k][1]);
        if (inSquare && Math.hypot(px - ctr[k][0], py - ctr[k][1]) > r[k] + 0.5) return true;
      }
      return false;
    };
    // ---- SVG decomposition: every drawable element of an inline SVG becomes its own native shape, freeform or text box ----
    const SVG_BLOCKERS = new Set(['clippath', 'mask', 'pattern', 'filter', 'image', 'foreignobject', 'use', 'symbol', 'marker', 'switch', 'textpath']);
    const svgDecomposable = (svg) => ![...svg.querySelectorAll('*')].some((e) => SVG_BLOCKERS.has(e.localName.toLowerCase()));
    // elliptical arc (SVG endpoint form) -> cubic segments, F.6 of the SVG spec
    const arcToCubics = (x1, y1, rx, ry, phiDeg, fa, fs, x2, y2) => {
      if (x1 === x2 && y1 === y2) return [];
      rx = Math.abs(rx); ry = Math.abs(ry);
      if (!rx || !ry) return [['L', x2, y2]];
      const phi = (phiDeg * Math.PI) / 180, cp = Math.cos(phi), sp = Math.sin(phi);
      const dx = (x1 - x2) / 2, dy = (y1 - y2) / 2;
      const x1p = cp * dx + sp * dy, y1p = -sp * dx + cp * dy;
      const lam = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
      if (lam > 1) { rx *= Math.sqrt(lam); ry *= Math.sqrt(lam); }
      const num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
      const den = rx * rx * y1p * y1p + ry * ry * x1p * x1p;
      let co = Math.sqrt(Math.max(0, num / den));
      if (fa === fs) co = -co;
      const cxp = (co * rx * y1p) / ry, cyp = (-co * ry * x1p) / rx;
      const cx = cp * cxp - sp * cyp + (x1 + x2) / 2, cy = sp * cxp + cp * cyp + (y1 + y2) / 2;
      const angle = (ux, uy, vx, vy) => {
        const dot = ux * vx + uy * vy, len = Math.hypot(ux, uy) * Math.hypot(vx, vy);
        let a = Math.acos(Math.max(-1, Math.min(1, dot / len)));
        if (ux * vy - uy * vx < 0) a = -a;
        return a;
      };
      const th1 = angle(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
      let dth = angle((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);
      if (!fs && dth > 0) dth -= 2 * Math.PI;
      else if (fs && dth < 0) dth += 2 * Math.PI;
      const n = Math.ceil(Math.abs(dth) / (Math.PI / 2)), d = dth / n, k = (4 / 3) * Math.tan(d / 4);
      const map = (u, v) => [cp * rx * u - sp * ry * v + cx, sp * rx * u + cp * ry * v + cy];
      const segs = [];
      let t = th1;
      for (let i = 0; i < n; i++) {
        const c1 = Math.cos(t), s1 = Math.sin(t), c2 = Math.cos(t + d), s2 = Math.sin(t + d);
        const p1 = map(c1 - k * s1, s1 + k * c1), p2 = map(c2 + k * s2, s2 - k * c2), p3 = map(c2, s2);
        segs.push(['C', ...p1, ...p2, ...p3]);
        t += d;
      }
      return segs;
    };
    // SVG path data -> absolute M/L/C/Q/Z commands in user units
    const parsePath = (d) => {
      const out = [];
      let p = 0, cx = 0, cy = 0, sx = 0, sy = 0, cmd = '', lastC = null, lastQ = null;
      const skip = () => { while (p < d.length && /[\s,]/.test(d[p])) p++; };
      const num = () => {
        skip();
        const re = /[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/y;
        re.lastIndex = p;
        const r = re.exec(d);
        if (!r) return null;
        p = re.lastIndex;
        return parseFloat(r[0]);
      };
      const flag = () => { skip(); const c = d[p]; if (c === '0' || c === '1') { p++; return c === '1'; } return null; };
      const more = () => { skip(); return p < d.length && /[-+.\d]/.test(d[p]); };
      while (p < d.length) {
        skip();
        if (p >= d.length) break;
        const before = p;
        if (/[A-Za-z]/.test(d[p])) cmd = d[p++];
        else if (!cmd) break;
        const U = cmd.toUpperCase(), rel = cmd !== U;
        const X = (v) => (rel ? cx + v : v), Y = (v) => (rel ? cy + v : v);
        if (U === 'Z') { out.push({ op: 'Z', pts: [] }); cx = sx; cy = sy; lastC = lastQ = null; cmd = ''; continue; }
        if (U === 'M') {
          const x = num(), y = num();
          if (x === null || y === null) break;
          cx = X(x); cy = Y(y); sx = cx; sy = cy;
          out.push({ op: 'M', pts: [[cx, cy]] });
          cmd = rel ? 'l' : 'L';
          lastC = lastQ = null;
          continue;
        }
        let guard = 0;
        do {
          if (guard++ > 1e6) break;
          if (U === 'L' || U === 'T' || U === 'H' || U === 'V') {
            if (U === 'H') { const x = num(); if (x === null) break; cx = rel ? cx + x : x; out.push({ op: 'L', pts: [[cx, cy]] }); }
            else if (U === 'V') { const y = num(); if (y === null) break; cy = rel ? cy + y : y; out.push({ op: 'L', pts: [[cx, cy]] }); }
            else if (U === 'L') { const x = num(), y = num(); if (x === null || y === null) break; cx = X(x); cy = Y(y); out.push({ op: 'L', pts: [[cx, cy]] }); }
            else {
              const x = num(), y = num();
              if (x === null || y === null) break;
              const ex = X(x), ey = Y(y);
              const c = lastQ ? [2 * cx - lastQ[0], 2 * cy - lastQ[1]] : [cx, cy];
              out.push({ op: 'Q', pts: [c, [ex, ey]] });
              lastQ = c; lastC = null; cx = ex; cy = ey;
              continue;
            }
            lastC = lastQ = null;
          } else if (U === 'C' || U === 'S') {
            const x1 = U === 'C' ? num() : null, y1 = U === 'C' ? num() : null;
            const x2 = num(), y2 = num(), x = num(), y = num();
            if (x2 === null || y2 === null || x === null || y === null || (U === 'C' && (x1 === null || y1 === null))) break;
            const c1 = U === 'C' ? [X(x1), Y(y1)] : (lastC ? [2 * cx - lastC[0], 2 * cy - lastC[1]] : [cx, cy]);
            const c2 = [X(x2), Y(y2)], e = [X(x), Y(y)];
            out.push({ op: 'C', pts: [c1, c2, e] });
            lastC = c2; lastQ = null; cx = e[0]; cy = e[1];
          } else if (U === 'Q') {
            const x1 = num(), y1 = num(), x = num(), y = num();
            if (x1 === null || y1 === null || x === null || y === null) break;
            const c = [X(x1), Y(y1)], e = [X(x), Y(y)];
            out.push({ op: 'Q', pts: [c, e] });
            lastQ = c; lastC = null; cx = e[0]; cy = e[1];
          } else if (U === 'A') {
            const rx = num(), ry = num(), rot = num(), fa = flag(), fs = flag(), x = num(), y = num();
            if ([rx, ry, rot, fa, fs, x, y].some((v) => v === null)) break;
            const ex = X(x), ey = Y(y);
            for (const seg of arcToCubics(cx, cy, rx, ry, rot, fa, fs, ex, ey)) {
              if (seg[0] === 'L') out.push({ op: 'L', pts: [[seg[1], seg[2]]] });
              else out.push({ op: 'C', pts: [[seg[1], seg[2]], [seg[3], seg[4]], [seg[5], seg[6]]] });
            }
            cx = ex; cy = ey; lastC = lastQ = null;
          } else break;
        } while (more());
        if (p === before) break; // malformed data: stop instead of spinning
      }
      return out;
    };
    const ellipseCmds = (cx, cy, rx, ry) => {
      const k = 0.5523;
      return [
        { op: 'M', pts: [[cx + rx, cy]] },
        { op: 'C', pts: [[cx + rx, cy + k * ry], [cx + k * rx, cy + ry], [cx, cy + ry]] },
        { op: 'C', pts: [[cx - k * rx, cy + ry], [cx - rx, cy + k * ry], [cx - rx, cy]] },
        { op: 'C', pts: [[cx - rx, cy - k * ry], [cx - k * rx, cy - ry], [cx, cy - ry]] },
        { op: 'C', pts: [[cx + k * rx, cy - ry], [cx + rx, cy - k * ry], [cx + rx, cy]] },
        { op: 'Z', pts: [] },
      ];
    };
    // fill / stroke paint of an SVG element -> {color, transparency, grad?} or null
    const svgPaintOf = (svg, val, op) => {
      if (!val || val === 'none') return null;
      const u = /url\(["']?#([^"')]+)/.exec(val);
      if (u) {
        const g = svg.ownerDocument.getElementById(u[1]);
        if (!g) return null;
        const stops = [...g.querySelectorAll('stop')].map((st) => {
          const sc2 = getComputedStyle(st), c = rgba(sc2.stopColor) || { hex: '000000', a: 1 };
          const off = st.getAttribute('offset') || '0';
          return { pos: /%/.test(off) ? parseFloat(off) : parseFloat(off) * 100, color: c.hex, alpha: c.a * parseFloat(sc2.stopOpacity) * op };
        });
        if (!stops.length) return null;
        const kind = g.localName.toLowerCase() === 'radialgradient' ? 'radial' : 'linear';
        const f = (a, d) => { const v = g.getAttribute(a); if (v === null) return d; return /%/.test(v) ? parseFloat(v) / 100 : parseFloat(v); };
        const ang = kind === 'linear' ? (Math.atan2(f('y2', 0) - f('y1', 0), f('x2', 1) - f('x1', 0)) * 180 / Math.PI + 360) % 360 : 0;
        return { color: stops[0].color, transparency: 0, grad: { kind, ang, stops } };
      }
      const c = rgba(val);
      if (!c) return null;
      return { color: c.hex, transparency: Math.round((1 - c.a * op) * 100) };
    };
    // walks an SVG and pushes one native item per drawable element, in document order
    const decomposeSvg = (svg) => {
      const emitShape = (it) => { if (it.fill && it.fill.grad) it.name = 'hp' + ++rid; push(it); };
      const walk = (node, gop) => {
        for (const el of [...node.children]) {
          const tag = el.localName.toLowerCase();
          if (['defs', 'lineargradient', 'radialgradient', 'style', 'title', 'desc', 'metadata', 'stop', 'symbol'].includes(tag)) continue;
          const cs = getComputedStyle(el);
          if (cs.display === 'none' || cs.visibility === 'hidden') continue;
          const op = gop * parseFloat(cs.opacity);
          if (tag === 'g' || tag === 'svg') { walk(el, op); continue; }
          const m = el.getScreenCTM();
          if (!m) continue;
          const sc = Math.sqrt(Math.abs(m.a * m.d - m.b * m.c));
          const toS = (u, v) => [m.a * u + m.c * v + m.e - sr.left, m.b * u + m.d * v + m.f - sr.top];
          const A = (k, dflt = 0) => { const v = parseFloat(el.getAttribute(k)); return Number.isNaN(v) ? dflt : v; };
          const fillP = (tag === 'line' || tag === 'polyline') ? null : svgPaintOf(svg, cs.fill, op * parseFloat(cs.fillOpacity));
          const sp = svgPaintOf(svg, cs.stroke, op * parseFloat(cs.strokeOpacity));
          const sw = parseFloat(cs.strokeWidth) * sc;
          const lineP = sp && sw > 0 ? { color: sp.color, width: sw, dash: cs.strokeDasharray && cs.strokeDasharray !== 'none' ? 'dash' : 'solid', transparency: sp.transparency } : null;
          const axis = Math.abs(m.b) < 1e-6 && Math.abs(m.c) < 1e-6;
          if (tag === 'text') {
            const txt = el.textContent.replace(/\s+/g, ' ').trim();
            if (!txt || !fillP) continue;
            const [bx, by] = toS(A('x'), A('y'));
            const fsP = parseFloat(cs.fontSize) * sc, fam = family(cs.fontFamily);
            const bold = parseInt(cs.fontWeight) >= 600 || cs.fontWeight === 'bold', italic = cs.fontStyle === 'italic';
            fontProbe.font = `${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}${fsP}px "${fam}"`;
            const tw = fontProbe.measureText(txt).width, asc = fontProbe.measureText('Hg').fontBoundingBoxAscent;
            const align = cs.textAnchor === 'middle' ? 'center' : cs.textAnchor === 'end' ? 'right' : 'left';
            const w = tw + 8;
            const x0 = align === 'center' ? bx - w / 2 : align === 'right' ? bx - w : bx;
            const col = { color: fillP.color, alpha: 1 - fillP.transparency / 100 };
            push({ type: 'text', x: x0, y: by - asc, w, h: fsP * 1.2, tr: { x: x0, y: by - asc, w, h: fsP * 1.2 }, asc, align, lineHeight: fsP * 1.2, single: true, hard: false, shadow: null, rot: 0, gl: null, tm: [1, 0, 0, 1, 0, 0], tgRuns: [], tbox: { x: x0, w }, valign: 'top',
              runs: [{ text: txt, font: fam, size: fsP, bold, italic, underline: false, strike: false, color: col.color, alpha: col.alpha, spacing: 0, highlight: undefined }] });
            continue;
          }
          if (tag === 'rect') {
            const w = A('width'), h = A('height');
            if (!(w > 0 && h > 0)) continue;
            const x = A('x'), y = A('y'), rx = Math.min(A('rx', A('ry')), w / 2, h / 2);
            if (axis) {
              const [x0, y0] = toS(x, y);
              const r = rx * Math.abs(m.a);
              emitShape({ type: 'shape', x: x0, y: y0, w: w * Math.abs(m.a), h: h * Math.abs(m.d), radius: r, radii: [r, r, r, r], rot: 0, shadow: null, fill: fillP, line: lineP });
              continue;
            }
            const corners = [[x, y], [x + w, y], [x + w, y + h], [x, y + h]];
            emitPath(corners.map((p, i) => ({ op: i ? 'L' : 'M', pts: [p] })).concat([{ op: 'Z', pts: [] }]), fillP, lineP, m, toS);
            continue;
          }
          let local = null;
          if (tag === 'circle') { const r = A('r'); if (r > 0) local = ellipseCmds(A('cx'), A('cy'), r, r); }
          else if (tag === 'ellipse') { const rx = A('rx'), ry = A('ry'); if (rx > 0 && ry > 0) local = ellipseCmds(A('cx'), A('cy'), rx, ry); }
          else if (tag === 'line') local = [{ op: 'M', pts: [[A('x1'), A('y1')]] }, { op: 'L', pts: [[A('x2'), A('y2')]] }];
          else if (tag === 'polyline' || tag === 'polygon') {
            const nums = (el.getAttribute('points') || '').trim().split(/[\s,]+/).map(Number).filter((v) => !Number.isNaN(v));
            const pts = []; for (let i = 0; i + 1 < nums.length; i += 2) pts.push([nums[i], nums[i + 1]]);
            if (pts.length) { local = pts.map((p, i) => ({ op: i ? 'L' : 'M', pts: [p] })); if (tag === 'polygon') local.push({ op: 'Z', pts: [] }); }
          } else if (tag === 'path') local = parsePath(el.getAttribute('d') || '');
          if (local && local.length) {
            const pat = cs.strokeDasharray && cs.strokeDasharray !== 'none' ? cs.strokeDasharray.split(/[\s,]+/).map(parseFloat).filter((v) => !Number.isNaN(v)) : [];
            if (lineP && pat.length) {
              if (fillP) emitPath(local, fillP, null, m, toS);
              const dashes = dashPolys(flattenCmds(local), pat.length % 2 ? [...pat, ...pat] : pat);
              for (const poly of dashes) emitPath([{ op: 'M', pts: [poly[0]] }, ...poly.slice(1).map((q) => ({ op: 'L', pts: [q] }))], null, { ...lineP, dash: 'solid' }, m, toS);
            } else emitPath(local, fillP, lineP, m, toS);
          }
        }
      };
      // stroke-dasharray: the outline is flattened and cut into real dashes, each dash one native freeform
      const flattenCmds = (cmds) => {
        const subs = [];
        let cur = null, cx = 0, cy = 0, sx = 0, sy = 0;
        for (const c of cmds) {
          if (c.op === 'M') { cur = [[c.pts[0][0], c.pts[0][1]]]; subs.push(cur); cx = sx = c.pts[0][0]; cy = sy = c.pts[0][1]; }
          else if (!cur) continue;
          else if (c.op === 'L') { cur.push([c.pts[0][0], c.pts[0][1]]); cx = c.pts[0][0]; cy = c.pts[0][1]; }
          else if (c.op === 'C' || c.op === 'Q') {
            const p1 = c.pts[0], p2 = c.op === 'C' ? c.pts[1] : c.pts[0], e = c.op === 'C' ? c.pts[2] : c.pts[1];
            for (let i = 1; i <= 24; i++) {
              const t = i / 24, mt = 1 - t;
              const x = c.op === 'C' ? mt * mt * mt * cx + 3 * mt * mt * t * p1[0] + 3 * mt * t * t * p2[0] + t * t * t * e[0] : mt * mt * cx + 2 * mt * t * p1[0] + t * t * e[0];
              const y = c.op === 'C' ? mt * mt * mt * cy + 3 * mt * mt * t * p1[1] + 3 * mt * t * t * p2[1] + t * t * t * e[1] : mt * mt * cy + 2 * mt * t * p1[1] + t * t * e[1];
              cur.push([x, y]);
            }
            cx = e[0]; cy = e[1];
          } else if (c.op === 'Z') { cur.push([sx, sy]); cx = sx; cy = sy; }
        }
        return subs.filter((q) => q.length > 1);
      };
      const dashPolys = (subs, pattern) => {
        const out = [];
        if (!pattern.length || pattern.every((v) => v <= 0)) return out;
        for (const poly of subs) {
          let idx = 0, remain = pattern[0], on = true, cur = [poly[0]];
          for (let i = 1; i < poly.length; i++) {
            const [x0, y0] = poly[i - 1], [x1, y1] = poly[i];
            const seg = Math.hypot(x1 - x0, y1 - y0);
            let pos = 0;
            while (seg - pos > remain) {
              pos += remain;
              const t = pos / seg, px = x0 + (x1 - x0) * t, py = y0 + (y1 - y0) * t;
              if (on) { cur.push([px, py]); if (cur.length > 1) out.push(cur); cur = []; }
              else cur = [[px, py]];
              on = !on; idx = (idx + 1) % pattern.length; remain = pattern[idx];
            }
            remain -= seg - pos;
            if (on) cur.push([x1, y1]);
          }
          if (on && cur.length > 1) out.push(cur);
        }
        return out;
      };
      const emitPath = (local, fillP, lineP, m, toS) => {
        const cmds = local.map((c) => ({ op: c.op, pts: c.pts.map(([u, v]) => toS(u, v)) }));
        push({ type: 'path', cmds, fill: fillP, line: lineP, name: fillP && fillP.grad ? 'hp' + ++rid : undefined });
      };
      walk(svg, 1);
    };

    const visit = (el) => {
      if (el.classList.contains('notes') || el.hasAttribute('data-ppt-skip')) return;
      const cs = getComputedStyle(el);
      if (cs.display === 'none' || cs.visibility === 'hidden') return;
      const op = opacityOf(el, slide);

      // CSS filter list: drop-shadow -> native shadow, blur -> native blur effect, colour functions -> recolour the
      // produced items; anything else (url(), backdrop-filter, mix-blend-mode: PowerPoint has no blend modes) -> picture
      let fsh = null, blurPx = 0, bad = false;
      const colorOps = [];
      const re = /([\w-]+)\(((?:[^()]|\([^()]*\))*)\)/g;
      for (let fm; cs.filter && cs.filter !== 'none' && (fm = re.exec(cs.filter));) {
        const [, fn, v] = fm;
        if (fn === 'drop-shadow') {
          const mm = v.match(/(rgba?\([^)]+\))\s+(-?[\d.]+)px\s+(-?[\d.]+)px(?:\s+([\d.]+)px)?/), c = mm && rgba(mm[1]);
          if (c) fsh = { color: c.hex, a: c.a, ox: +mm[2], oy: +mm[3], blur: +(mm[4] || 0) }; else bad = true;
        } else if (fn === 'url') {
          // SVG filter referenced from CSS: single-primitive filters have exact native equivalents
                    // computed style returns an absolute URL (file:///...#id), so the id is what follows the last '#'
          const sp = svgFilterAsCss(document.getElementById(v.split('#').pop().replace(/["'\s)]/g, '')));
          if (!sp) bad = true;
          else {
            blurPx += sp.blur || 0;
            if (sp.shadow) fsh = sp.shadow;
            colorOps.push(...sp.ops);
          }
        } else if (fn === 'blur') blurPx += parseFloat(v) || 0;
        else if (['grayscale', 'sepia', 'brightness', 'contrast', 'saturate', 'hue-rotate', 'invert', 'opacity'].includes(fn)) colorOps.push(...cssOps(fn, fn === 'hue-rotate' ? parseFloat(v) : /%/.test(v) ? parseFloat(v) / 100 : parseFloat(v)));
        else bad = true;
      }
      // blend modes PowerPoint cannot express: solid plain boxes are computed against the solid box beneath them
      // (see applyBlends); anything else is a picture rendered with its backdrop
      const BLEND = ['multiply', 'screen', 'overlay', 'darken', 'lighten', 'color-dodge', 'color-burn', 'hard-light', 'soft-light', 'difference', 'exclusion', 'hue', 'saturation', 'color', 'luminosity'];
      const blendable = BLEND.includes(cs.mixBlendMode) && !el.children.length && !el.textContent.trim() && cs.backgroundImage === 'none' && !hasTransform(cs) && cs.filter === 'none';
      if (bad || (cs.mixBlendMode !== 'normal' && !blendable)) {
        const rr = el.getBoundingClientRect();
        // blend modes need what is behind them, so this picture is rendered with the backdrop included
        if (rr.width > 0.5 && rr.height > 0.5) raster(el, rel(rr), cs.mixBlendMode !== 'normal' ? { noIso: true } : {});
        return;
      }
      // backdrop-filter (frosted glass): the blurred backdrop becomes a picture under the element, which itself stays native
      if (cs.backdropFilter && cs.backdropFilter !== 'none') {
        const rr = el.getBoundingClientRect();
        if (rr.width > 0.5 && rr.height > 0.5 && !hasTransform(cs)) raster(el, rel(rr), { noIso: true, backdrop: true });
      }

      // transforms: measure in the untransformed layout, then apply the matrix to the produced items
      let matrix = null, saved = null;
      if (hasTransform(cs)) {
        const m = new DOMMatrix(cs.transform);
        const sx = Math.hypot(m.a, m.b), sy = Math.hypot(m.c, m.d);
        // rotate / scale / skew all map: shapes become polygons under skew, text keeps the rotation component
        const ok = Math.abs(m.a * m.d - m.b * m.c) > 1e-6 && sx > 0 && sy > 0 && !el.closest('svg');
        if (!ok) { raster(el, rel(el.getBoundingClientRect())); return; }
        matrix = m; saved = el.style.getPropertyValue('transform');
        el.style.setProperty('transform', 'none', 'important');
      }
      const restore = () => { if (matrix) { saved ? el.style.setProperty('transform', saved) : el.style.removeProperty('transform'); } };
      const r = el.getBoundingClientRect();
      if (r.width < 0.5 || r.height < 0.5) { restore(); return; }
      const box = rel(r);
      const start = items.length;
      const prevZ = curZ, prevTG = curTextGrad, prevFsh = curFsh;
      if (cs.position !== 'static' && cs.zIndex !== 'auto' && !isNaN(parseInt(cs.zIndex))) curZ = parseInt(cs.zIndex);
      const phA = el.getAttribute('data-ph'), prevPhBody = phBody;
      const phKind = phA === 'body' && phOk.body ? 'body' : phA === 'subtitle' && phOk.subTitle ? 'subTitle' : phA === 'title' ? 'title' : null;
      if (phKind === 'body') phBody = el;

      const finish = () => {
        curZ = prevZ; curTextGrad = prevTG; curFsh = prevFsh; phBody = prevPhBody;
        // data-ph body / subtitle: the subtree's text becomes one placeholder (body: one paragraph per text item;
        // subtitle / title: one paragraph; title items on the same line join without a break); line breaks inside an item stay line breaks
        const phTexts = phKind ? items.slice(start).filter((i) => i.type === 'text') : [];
        if (phKind === 'title' && phTexts.length === 1) { phTexts[0].ph = 'title'; phTexts[0].name = phTexts[0].name || 'hp' + ++rid; }
        else if (phTexts.length) {
          const m = phTexts[0], runs = [];
          phTexts.forEach((t, ti) => t.runs.forEach((r, ri) => {
            const nr = { ...r, brk: false };
            const pv = phTexts[ti - 1];
            if (ri === 0 && ti > 0) { if (phKind === 'body') runs[runs.length - 1].brk = true; else if (phKind !== 'title' || t.y >= pv.y + pv.h / 2) nr.soft = true; }
            else if (ri > 0 && t.runs[ri - 1].brk) nr.soft = true;
            runs.push(nr);
          }));
          const x0 = Math.min(...phTexts.map((t) => t.x)), y0 = Math.min(...phTexts.map((t) => t.y));
          const x1 = Math.max(...phTexts.map((t) => t.x + t.w)), y1 = Math.max(...phTexts.map((t) => t.y + t.h));
          Object.assign(m, { x: x0, y: y0, w: x1 - x0, h: y1 - y0, runs, tgRuns: phTexts.flatMap((t) => t.tgRuns || []), single: false, hard: phTexts.some((t) => t.hard),
            gl: null, ink: null, warp: undefined, tbox: { x: x0, w: x1 - x0 }, ph: phKind, name: m.name || 'hp' + ++rid });
          for (const t of phTexts.slice(1)) items.splice(items.indexOf(t), 1);
        }
        let flat = false; // true when the subtree has to be one picture
        if (colorOps.length || blurPx) {
          const sub = items.slice(start);
          if (sub.some((i) => i.type === 'raster')) flat = true;
          else {
            sub.forEach((i) => recolor(i, colorOps));
            if (blurPx) {
              sub.forEach((i) => { if (i.type === 'shape' || i.type === 'path') { i.blur = blurPx; i.name = i.name || 'hp' + ++rid; } });
              // PowerPoint cannot blur live text: the text becomes vector outlines that carry the native blur
              // (a blurred picture of the text is the fallback when a font file cannot be read)
              const texts = sub.filter((i) => i.type === 'text');
              if (texts.length) {
                el.dataset.pptRid = el.dataset.pptRid || 'r' + ++rid;
                const rb = rel(el.getBoundingClientRect());
                texts.forEach((i) => { i.outline = true; i.blur = blurPx; i.fb = { id: el.dataset.pptRid, ...rb }; });
              }
            }
          }
        }
        // pictures inside a transformed group (or a filtered group that cannot be recoloured): rasterize the whole group as it appears
        const asPicture = flat || (matrix && items.slice(start).some((i) => i.type === 'raster'));
        if (asPicture) {
          items.length = start;
          restore();
          raster(el, rel(el.getBoundingClientRect()));
        } else if (matrix) {
          const ox = cs.transformOrigin.split(' ').map(parseFloat);
          applyMatrix(start, matrix, { x: box.x + ox[0], y: box.y + ox[1] });
          restore();
        }
        clips.length && clips[clips.length - 1].popFor === el && clips.pop();
      };

      // SVG with foreignObject: the vector part becomes a picture, the HTML inside becomes native shapes and text
      if (el.tagName.toLowerCase() === 'svg' && el.querySelector('foreignObject') && svgMarkup(el)) {
        raster(el, rel(el.getBoundingClientRect()));
        for (const fo of el.querySelectorAll('foreignObject')) for (const c of [...fo.children]) visit(c);
        return finish();
      }

      const clip = clips.length ? clips[clips.length - 1].rect : null;
      const clips_ = /hidden|clip/.test(cs.overflowX + cs.overflowY);
      const eff = clips_ ? intersect(box, clip) : clip;
      const own = clips.length ? clips[clips.length - 1] : null; // nearest clipping ancestor (before this element's own clip)
      const radii = radiiOf(cs, box);
      if (clips_) clips.push({ rect: eff, popFor: el, radii: radii.some((v) => v > 0) && (!clip || inside(box, clip)) ? radii : null });

      // content that would be cut off by an ancestor's overflow:hidden is rasterized with the clip applied
      if (clip && !inside(box, clip) && !clips_) {
        const b = intersect(box, clip);
        if (b.w > 0 && b.h > 0) raster(el, b, { iso: true });
        return finish();
      }
      if (el.localName.toLowerCase() === 'svg' && svgDecomposable(el)) { decomposeSvg(el); return finish(); }
      if (isRasterEl(el)) {
        // pictures poking out of a rounded clipping ancestor keep that rounded cut (isolated render honours the clip)
        if (own && !clips_ && roundedCut(box, own)) raster(el, intersect(box, own.rect), { iso: true });
        else raster(el, box);
        return finish();
      }
      if (eff && ownTextOutside(el, eff)) {
        const b = intersect(box, eff);
        if (b.w > 0 && b.h > 0) raster(el, b, { iso: true });
        return finish();
      }

      // background / border / shadow
      const fill = rgba(cs.backgroundColor);
      const inset = parseShadow(cs.boxShadow, true);
      const grad = cs.backgroundImage !== 'none' ? parseGradient(cs.backgroundImage) : null;
      const clipText = /text/.test(cs.webkitBackgroundClip || cs.backgroundClip);
      const textGrad = clipText && grad ? grad : null; // gradient-filled live text
      if (clipText && !grad) { raster(el, box); return finish(); }
      curTextGrad = textGrad;
      if (cs.backgroundImage !== 'none' && !grad) {
        raster(el, box, { hideChildren: true });
      } else if (!textGrad) {
        const bw = ['Top', 'Right', 'Bottom', 'Left'].map((s) => ({
          w: parseFloat(cs['border' + s + 'Width']) || 0,
          st: cs['border' + s + 'Style'],
          c: rgba(cs['border' + s + 'Color']),
        })).map((b) => (b.st === 'none' || !b.c ? { ...b, w: 0 } : b));
        const uniform = bw.every((b) => b.w > 0 && b.w === bw[0].w && b.c.hex === bw[0].c.hex);
        // flush corners of a rounded clipping parent inherit its radius
        if (own && own.radii) {
          const R = own.rect, tol = 1.5;
          const cc = [[R.x, R.y], [R.x + R.w, R.y], [R.x + R.w, R.y + R.h], [R.x, R.y + R.h]];
          const mc = [[box.x, box.y], [box.x + box.w, box.y], [box.x + box.w, box.y + box.h], [box.x, box.y + box.h]];
          for (let k = 0; k < 4; k++) if (own.radii[k] > 0 && Math.abs(mc[k][0] - cc[k][0]) < tol && Math.abs(mc[k][1] - cc[k][1]) < tol) radii[k] = Math.max(radii[k], own.radii[k]);
        }
        const radius = radii[0];
        if (!(fill || uniform || grad || inset)) curFsh = fsh; else curFsh = null;
        const sh = parseShadow(cs.boxShadow) || fsh;
        const shadow = sh && { color: sh.color, opacity: Math.min(1, sh.a * op), blur: sh.blur, offset: Math.hypot(sh.ox, sh.oy), angle: (Math.atan2(sh.oy, sh.ox) * 180 / Math.PI + 360) % 360 };
        if (fill || uniform || grad || inset) {
          const gfill = grad && { color: grad.stops[0].color, transparency: 0, grad: { ...grad, stops: grad.stops.map((s) => ({ ...s, alpha: s.alpha * op })) } };
          push({
            type: 'shape', ...box, radius, radii, rot: 0, shadow,
            blend: blendable && fill ? { mode: cs.mixBlendMode, color: fill.hex, a: fill.a * op } : undefined,
            name: grad || inset ? 'hp' + ++rid : undefined,
            inner: inset && { color: inset.color, opacity: Math.min(1, inset.a * op), blur: inset.blur, offset: Math.hypot(inset.ox, inset.oy), angle: (Math.atan2(inset.oy, inset.ox) * 180 / Math.PI + 360) % 360 },
            fill: gfill || (fill && { color: fill.hex, transparency: Math.round((1 - fill.a * op) * 100) }),
            line: uniform ? { color: bw[0].c.hex, width: bw[0].w, dash: bw[0].st === 'dashed' ? 'dash' : bw[0].st === 'dotted' ? 'sysDot' : 'solid', transparency: Math.round((1 - bw[0].c.a * op) * 100) } : null,
          });
        }
        if (!uniform) {
          const [t, rt, b, l] = bw;
          const side = (bd, x, y, w, h) => bd.w > 0 && push({ type: 'shape', x, y, w, h, radius: 0, rot: 0, fill: { color: bd.c.hex, transparency: Math.round((1 - bd.c.a * op) * 100) }, line: null });
          side(t, box.x, box.y, box.w, t.w);
          side(b, box.x, box.y + box.h - b.w, box.w, b.w);
          side(l, box.x, box.y, l.w, box.h);
          side(rt, box.x + box.w - rt.w, box.y, rt.w, box.h);
        }
      }

      if (cs.columnCount !== 'auto' && parseInt(cs.columnCount) > 1) { emitColumns(el, cs, box, op); return finish(); }

      // children: fold text + inline content into text boxes, recurse into block children
      let buf = [];
      let first = true;
      const isLi = el.tagName === 'LI' && el.parentElement && /^(UL|OL)$/.test(el.parentElement.tagName);
      const flush = () => {
        while (buf.length && buf[0].nodeType === 1 && buf[0].tagName === 'BR') buf.shift();
        if (buf.some((n) => n.textContent.trim())) {
          emitText(cs, buf, box, op, first && isLi ? el : null);
          first = false;
        }
        buf = [];
      };
      for (const n of el.childNodes) {
        if (n.nodeType === 3) { if (n.textContent.trim() || buf.length) buf.push(n); }
        else if (n.nodeType === 1) {
          if (isInlineMergeable(n)) buf.push(n);
          else { flush(); visit(n); }
        }
      }
      flush();
      finish();
    };
    for (const k of slide.children) visit(k);
    applyBlends(items, bg && bg.hex);
    items.sort((a, b) => a.z - b.z || a.order - b.order);
    return { bg: bg && bg.hex, notes, items, layout: slide.getAttribute('data-layout') };
  });
}

const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: W, height: H }, deviceScaleFactor: 2 });
await page.goto(pathToFileURL(resolve(input)).href, { waitUntil: 'networkidle' });
await page.evaluate(() => document.fonts.ready);
// let scripts that animate or measure at runtime (fps readouts, canvases, charts) settle before the DOM is read
await page.waitForTimeout(SETTLE_MS);
// neutralize on-screen deck chrome so slides sit at their natural 1280x720 size; settle CSS animations at their end state
await page.addStyleTag({ content: 'section.slide{margin:0 !important;transform:none !important;box-shadow:none !important}' });
await page.evaluate(() => document.getAnimations().forEach((a) => { try { a.finish(); } catch {} }));

// which of data-ph="body" / "subtitle" each slide's template layout can take (otherwise they stay text boxes)
const phAvail = TEMPLATE ? (await page.evaluate(() => [...document.querySelectorAll('section.slide')].map((s) => s.getAttribute('data-layout'))))
  .map((l) => { const lay = pick(l); return { body: !!layoutPh(lay, 'body'), subTitle: !!layoutPh(lay, 'subTitle') }; }) : [];
const slidesData = await page.evaluate(extract, { hard: HARD, ph: phAvail });
const webFonts = await page.evaluate(() => window.__webFonts || []);
const fontUse = await page.evaluate(() => window.__fontUse || []);
const faces = await page.evaluate(() => {
  const out = [], rank = (u) => (/\.ttf(\?|#|$)/i.test(u) ? 0 : /\.woff(\?|#|$)/i.test(u) ? 1 : /^data:/.test(u) ? 2 : 3);
  for (const sh of document.styleSheets) {
    let rules; try { rules = sh.cssRules; } catch { continue; }
    for (const r of rules) {
      if (r.type !== CSSRule.FONT_FACE_RULE) continue;
      const urls = [...r.style.getPropertyValue('src').matchAll(/url\(["']?([^"')]+)["']?\)/g)].map((m) => new URL(m[1], sh.href || document.baseURI).href).sort((a, b) => rank(a) - rank(b));
      if (urls.length) out.push({ family: r.style.getPropertyValue('font-family').replace(/["']/g, '').trim(), weight: r.style.getPropertyValue('font-weight') || '400', style: r.style.getPropertyValue('font-style') || 'normal', url: urls[0] });
    }
  }
  return out;
});
// replaces the generated master, layouts and theme with the template's; slides point at the chosen layout
async function mergeTemplate(zip, slides) {
  for (const f of Object.keys(zip.files)) if (/^ppt\/(slideLayouts|slideMasters)\//.test(f)) zip.remove(f);
  const mediaNames = new Map();
  for (const f of Object.keys(TPL_ZIP.files)) {
    if (!/^ppt\/(slideLayouts|slideMasters)\//.test(f) && f !== 'ppt/theme/theme1.xml') continue;
    if (TPL_ZIP.files[f].dir) continue;
    let body = await TPL_ZIP.file(f).async('string');
    if (f.endsWith('.rels')) {
      body = body.replace(/Target="\.\.\/media\/([^"]+)"/g, (m, name) => { mediaNames.set(name, 'tpl_' + name); return `Target="../media/tpl_${name}"`; });
    }
    zip.file(f, body);
  }
  for (const [name, dst] of mediaNames) zip.file('ppt/media/' + dst, await TPL_ZIP.file('ppt/media/' + name).async('nodebuffer'));
  // content types for the copied parts
  let ct = await zip.file('[Content_Types].xml').async('string');
  ct = ct.replace(/<Override PartName="\/ppt\/(slideLayouts|slideMasters)\/[^>]*\/>/g, '');
  const add = Object.keys(TPL_ZIP.files).filter((f) => /^ppt\/(slideLayouts\/slideLayout\d+|slideMasters\/slideMaster\d+)\.xml$/.test(f));
  for (const f of add) {
    const type = f.includes('slideLayouts') ? 'slideLayout' : 'slideMaster';
    ct = ct.replace('</Types>', `<Override PartName="/${f}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.${type}+xml"/></Types>`);
  }
  zip.file('[Content_Types].xml', ct);
  // slide size of the template
  let pres = await zip.file('ppt/presentation.xml').async('string');
  const tpres = await TPL_ZIP.file('ppt/presentation.xml').async('string');
  pres = pres.replace(/<p:sldSz[^>]*\/>/, tpres.match(/<p:sldSz[^>]*\/>/)[0]);
  zip.file('ppt/presentation.xml', pres);
  for (let i = 0; i < slides.length; i++) {
    const rels = `ppt/slides/_rels/slide${i + 1}.xml.rels`;
    let body = await zip.file(rels).async('string');
    const layout = pick(slides[i].layout);
    body = body.replace(/\.\.\/slideLayouts\/slideLayout\d+\.xml/, '../slideLayouts/' + layout.file);
    zip.file(rels, body);
    const sx = `ppt/slides/slide${i + 1}.xml`;
    let sxml = await zip.file(sx).async('string');
    const lx = await TPL_ZIP.file('ppt/slideLayouts/' + layout.file).async('string');
    if (sxml.includes('<p:ph type="title"/>') && /<p:ph type="(title|ctrTitle)"/.test(lx)) {
      const k = sxml.indexOf('<p:ph type="title"/>');
      const a = sxml.lastIndexOf('<p:sp>', k), b = sxml.indexOf('</p:sp>', k);
      sxml = sxml.slice(0, a) + sxml.slice(a, b).replace(/<a:xfrm[^>]*>.*?<\/a:xfrm>/s, '') + sxml.slice(b);
      zip.file(sx, sxml);
    }
    // body / subtitle placeholders: point at the layout's placeholder (type + idx) and take its position
    for (const kind of ['body', 'subTitle']) {
      const tag = `<p:ph type="${kind}"/>`, lp = layoutPh(layout, kind);
      if (!lp) continue;
      const real = `<p:ph${lp.type === 'obj' ? '' : ` type="${lp.type}"`}${lp.idx ? ` idx="${lp.idx}"` : ''}/>`;
      for (let k; (k = sxml.indexOf(tag)) >= 0;) {
        const a = sxml.lastIndexOf('<p:sp>', k), b = sxml.indexOf('</p:sp>', k);
        sxml = sxml.slice(0, a) + sxml.slice(a, b).replace(tag, real).replace(/<a:xfrm[^>]*>.*?<\/a:xfrm>/s, '') + sxml.slice(b);
      }
      zip.file(sx, sxml);
    }
  }
}
const slideEls = await page.$$('section.slide');
if (!slideEls.length) throw new Error('No <section class="slide"> found in ' + input);

const fontOf = makeFontLoader(faces);
const restricted = NO_EMBED ? new Set() : await restrictedFamilies(fontUse.filter((u) => webFonts.includes(u.family)), faces);
const fbDone = new Set();
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'HTMLPPT', width: TEMPLATE ? TPL_W : W * PX, height: TEMPLATE ? TPL_H : H * PX });
pptx.layout = 'HTMLPPT';
const inch = (v) => +(v * PX * SCALE).toFixed(4);
// content is centred on the template's slide (zero offset for 16:9 templates)
const OFF_X = TEMPLATE ? (TPL_W - W * PX * SCALE) / 2 : 0, OFF_Y = TEMPLATE ? (TPL_H - H * PX * SCALE) / 2 : 0;
const ix = (v) => +(inch(v) + OFF_X).toFixed(4), iy = (v) => +(inch(v) + OFF_Y).toFixed(4);
const pt = (px) => +(px * SCALE * 0.75).toFixed(2);
const fonts = new Set();

// rasterize one item in isolation: everything else hidden/transparent, clipped to the item's rect
async function rasterize(origin, it) {
  const t = `section.slide [data-ppt-rid="${it.id}"]`;
  let css = 'html,body{background:transparent !important}';
  if (it.backdrop) css = `${t}{background:transparent !important;border-color:transparent !important;box-shadow:none !important;color:transparent !important} ${t} *{visibility:hidden !important}`;
  else if (it.noIso) css = '';
  else css += it.slide ? `section.slide:not([data-ppt-rid="${it.id}"]){background:transparent !important} section.slide *{visibility:hidden !important}`
    : `section.slide{background:transparent !important} section.slide *{visibility:hidden !important} ${t}, ${t} *{visibility:visible !important}`;
  if (it.textOnly) css += ` ${t}, ${t} *{background:transparent !important;background-image:none !important;border-color:transparent !important;box-shadow:none !important}`;
  if (it.hideChildren) css += it.slide ? '' : ` ${t} *{visibility:hidden !important} ${t}{color:transparent !important}`;
  const style = css ? await page.addStyleTag({ content: css }) : null;
  const absY = origin.y + it.y;
  const sy = await page.evaluate((y) => { scrollTo(0, y); return scrollY; }, absY);
  const data = await page.screenshot({ clip: { x: origin.x + it.x, y: absY - sy, width: it.w, height: it.h }, omitBackground: true });
  if (style) await style.evaluate((n) => n.remove());
  return data;
}

const stats = { shapes: 0, texts: 0, images: 0, svg: 0, outlined: 0 };
let outlineN = 0;
const patches = {}; // slide index -> shapes needing XML-only features (gradient fill, inner shadow)
for (let i = 0; i < slidesData.length; i++) {
  const sd = slidesData[i];
  const slide = pptx.addSlide();
  if (sd.bg && !TEMPLATE) slide.background = { color: sd.bg };
  if (sd.notes) slide.addNotes(sd.notes);
  const origin = await slideEls[i].evaluate((e) => { const r = e.getBoundingClientRect(); return { x: r.left + scrollX, y: r.top + scrollY }; });
  for (const it of sd.items) {
    if (TEMPLATE && it.bgLayer) continue;
    if (it.type === 'raster') {
      const vector = it.svg && !NO_SVG;
      const data = vector ? Buffer.from(it.svg) : await rasterize(origin, it);
      slide.addImage({ data: (vector ? 'image/svg+xml;base64,' : 'image/png;base64,') + data.toString('base64'), x: ix(it.x), y: iy(it.y), w: inch(it.w), h: inch(it.h), transparency: it.op < 1 ? Math.round((1 - it.op) * 100) : undefined });
      stats.images++; if (vector) stats.svg++;
    } else if (it.type === 'shape') {
      const opts = { x: ix(it.x), y: iy(it.y), w: inch(it.w), h: inch(it.h) };
      opts.fill = it.fill ? { color: it.fill.color, transparency: it.fill.transparency } : { type: 'none' };
      opts.line = it.line ? { color: it.line.color, width: pt(it.line.width), dashType: it.line.dash, transparency: it.line.transparency } : { type: 'none' };
      if (it.rot) opts.rotate = ((it.rot % 360) + 360) % 360;
      if (it.shadow) opts.shadow = { type: 'outer', color: it.shadow.color, opacity: it.shadow.opacity, blur: pt(it.shadow.blur), offset: pt(it.shadow.offset), angle: it.shadow.angle };
      if (it.name) { opts.objectName = it.name; (patches[i] ||= []).push({ name: it.name, grad: it.fill && it.fill.grad, inner: it.inner, blur: it.blur }); }
      const full = it.radius >= Math.min(it.w, it.h) / 2 - 0.5 && it.radius > 0;
      const rd = it.radii || [it.radius, it.radius, it.radius, it.radius];
      const mixed = rd.some((v) => Math.abs(v - rd[0]) > 0.5);
      if (it.poly) {
        // skewed box: polygon in inches relative to the shape's bounding box
        const pts = it.poly.map(([x, y], k) => ({ x: inch(x - it.x), y: inch(y - it.y), moveTo: k === 0 }));
        pts.push({ close: true });
        slide.addShape(pptx.ShapeType.custGeom, { ...opts, points: pts });
      } else if (mixed) {
        // different corner radii: custom geometry (lines + quarter arcs), in inches relative to the shape
        const w = inch(it.w), h = inch(it.h), [tl, tr, br, bl] = rd.map(inch);
        const pts = [{ x: tl, y: 0, moveTo: true }, { x: w - tr, y: 0 }];
        if (tr) pts.push({ x: w, y: tr, curve: { type: 'arc', wR: tr, hR: tr, stAng: 270, swAng: 90 } });
        pts.push({ x: w, y: h - br });
        if (br) pts.push({ x: w - br, y: h, curve: { type: 'arc', wR: br, hR: br, stAng: 0, swAng: 90 } });
        pts.push({ x: bl, y: h });
        if (bl) pts.push({ x: 0, y: h - bl, curve: { type: 'arc', wR: bl, hR: bl, stAng: 90, swAng: 90 } });
        pts.push({ x: 0, y: tl });
        if (tl) pts.push({ x: tl, y: 0, curve: { type: 'arc', wR: tl, hR: tl, stAng: 180, swAng: 90 } });
        pts.push({ close: true });
        slide.addShape(pptx.ShapeType.custGeom, { ...opts, points: pts });
      } else if (full && Math.abs(it.w - it.h) < 1) slide.addShape(pptx.ShapeType.ellipse, opts);
      else if (it.radius > 0) {
        opts.rectRadius = inch(Math.min(it.radius, Math.min(it.w, it.h) / 2));
        slide.addShape(pptx.ShapeType.roundRect, opts);
      } else slide.addShape(pptx.ShapeType.rect, opts);
      stats.shapes++;
    } else if (it.type === 'path') {
      let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
      for (const c of it.cmds) for (const p of c.pts) { x0 = Math.min(x0, p[0]); y0 = Math.min(y0, p[1]); x1 = Math.max(x1, p[0]); y1 = Math.max(y1, p[1]); }
      if (!isFinite(x0)) continue;
      const R = (p) => ({ x: inch(p[0] - x0), y: inch(p[1] - y0) });
      const points = [];
      for (const c of it.cmds) {
        if (c.op === 'M') points.push({ ...R(c.pts[0]), moveTo: true });
        else if (c.op === 'L') points.push(R(c.pts[0]));
        else if (c.op === 'C') { const a = R(c.pts[0]), b = R(c.pts[1]), e = R(c.pts[2]); points.push({ x: e.x, y: e.y, curve: { type: 'cubic', x1: a.x, y1: a.y, x2: b.x, y2: b.y } }); }
        else if (c.op === 'Q') { const a = R(c.pts[0]), e = R(c.pts[1]); points.push({ x: e.x, y: e.y, curve: { type: 'quadratic', x1: a.x, y1: a.y } }); }
        else if (c.op === 'Z') points.push({ close: true });
      }
      const opts = { x: ix(x0), y: iy(y0), w: inch(Math.max(x1 - x0, 0.01)), h: inch(Math.max(y1 - y0, 0.01)), points };
      opts.fill = it.fill ? { color: it.fill.color, transparency: it.fill.transparency } : { type: 'none' };
      opts.line = it.line ? { color: it.line.color, width: pt(it.line.width), dashType: it.line.dash, transparency: it.line.transparency } : { type: 'none' };
      if (it.name) { opts.objectName = it.name; (patches[i] ||= []).push({ name: it.name, grad: it.fill && it.fill.grad, blur: it.blur }); }
      slide.addShape(pptx.ShapeType.custGeom, opts);
      stats.shapes++;
    } else if (it.type === 'text') {
      // vector outlines: blur / skew / fonts whose licence forbids embedding
      if ((it.outline || it.runs.some((r) => restricted.has(r.font))) && it.gl && it.gl.length) {
        const groups = await outlineGroups(it, fontOf);
        if (groups) {
          for (const g of groups) {
            const sh = shapeOf(g, PX * SCALE);
            if (!sh) continue;
            const name = 'hp' + (++outlineN);
            slide.addShape(pptx.ShapeType.custGeom, {
              x: ix(sh.x), y: iy(sh.y), w: inch(sh.w), h: inch(sh.h), points: sh.points, objectName: name,
              fill: { color: g.color, transparency: g.alpha < 1 ? Math.round((1 - g.alpha) * 100) : 0 }, line: { type: 'none' },
            });
            (patches[i] ||= []).push({ name, grad: it.tgrad ? { ...it.tgrad } : undefined, blur: it.blur });
            stats.shapes++;
          }
          stats.outlined++;
          continue;
        }
        if (it.fb) { // font file unreadable: blurred picture of the text instead
          if (!fbDone.has(it.fb.id)) {
            fbDone.add(it.fb.id);
            const data = await rasterize(origin, { id: it.fb.id, x: it.fb.x, y: it.fb.y, w: it.fb.w, h: it.fb.h, iso: true, textOnly: true });
            slide.addImage({ data: 'image/png;base64,' + data.toString('base64'), x: ix(it.fb.x), y: iy(it.fb.y), w: inch(it.fb.w), h: inch(it.fb.h) });
            stats.images++;
          }
          continue;
        }
      }
      const runs = it.runs.map((r) => {
        const face = FONT_MAP[r.font] || r.font;
        fonts.add(face);
        return {
          text: r.text,
          options: {
            breakLine: r.brk || undefined, softBreakBefore: r.soft || undefined,
            fontFace: face, fontSize: pt(r.size), bold: r.bold, italic: r.italic,
            underline: r.underline ? { style: 'sng' } : undefined, strike: r.strike ? 'sngStrike' : undefined,
            color: r.color, highlight: r.highlight, outline: r.stroke ? { size: pt(r.stroke.size), color: r.stroke.color } : undefined, transparency: r.alpha < 1 ? Math.round((1 - r.alpha) * 100) : undefined,
            charSpacing: r.spacing ? pt(r.spacing) : undefined,
          },
        };
      });
      const slack = it.single || it.hard ? 0 : 4; // soften wrap differences between browser and PowerPoint
      const o = {
        // PowerPoint seats exact-spaced lines ~0.13em higher than the browser's line box
        x: ix(it.x - (it.align === 'center' ? slack / 2 : it.align === 'right' ? slack : 0)), y: iy(it.y + (it.noShift ? 0 : it.asc !== undefined ? it.asc - (0.33 * it.lineHeight + 0.49 * it.asc + 2.4) : 0.13 * Math.max(...it.runs.map((r) => r.size)))), w: inch(it.w + slack), h: inch(it.h),
        margin: 0, valign: it.warp ? 'middle' : 'top', align: it.align === 'justify' ? 'justify' : it.align,
        wrap: !(it.single || it.hard || it.warp), fit: 'none', rotate: it.rot ? ((it.rot % 360) + 360) % 360 : undefined,
        lineSpacing: pt(it.lineHeight), paraSpaceBefore: 0, paraSpaceAfter: 0,
      };
      if (it.shadow) o.shadow = { type: 'outer', color: it.shadow.color, opacity: it.shadow.opacity, blur: pt(it.shadow.blur), offset: pt(it.shadow.offset), angle: it.shadow.angle };
      if (it.name) { o.objectName = it.name; (patches[i] ||= []).push({ name: it.name, tgrad: it.tgrad, warp: it.warp, tgRuns: it.tgRuns, tbox: it.tbox, ph: it.ph }); }
      slide.addText(runs, o);
      stats.texts++;
    }
  }
}
// pptxgenjs has no gradient fill / inner shadow: patch those into the saved XML by shape name
const emu = (ptv) => Math.round(ptv * 12700);
const gradXml = (g) => {
  const gs = g.stops.map((s) => `<a:gs pos="${Math.round(s.pos * 1000)}"><a:srgbClr val="${s.color}">${s.alpha < 1 ? `<a:alpha val="${Math.round(s.alpha * 100000)}"/>` : ''}</a:srgbClr></a:gs>`).join('');
  const shade = g.kind === 'linear' ? `<a:lin ang="${Math.round(g.ang * 60000)}" scaled="0"/>` : '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>';
  return `<a:gradFill rotWithShape="1"><a:gsLst>${gs}</a:gsLst>${shade}</a:gradFill>`;
};
const innerXml = (s) => `<a:innerShdw blurRad="${emu(pt(s.blur))}" dist="${emu(pt(s.offset))}" dir="${Math.round(s.angle * 60000)}"><a:srgbClr val="${s.color}"><a:alpha val="${Math.round(s.opacity * 100000)}"/></a:srgbClr></a:innerShdw>`;
const buf = await pptx.write({ outputType: 'nodebuffer' });
const zip = await JSZip.loadAsync(buf);
for (const [idx, list] of Object.entries(patches)) {
  const file = `ppt/slides/slide${+idx + 1}.xml`;
  let xml = await zip.file(file).async('string');
  for (const p of list) {
    const at = xml.indexOf(`name="${p.name}"`);
    if (at < 0) continue;
    if (p.ph) {
      // the placeholder keeps its text and style, and takes its position from the layout
      const end1 = xml.indexOf('</p:sp>', at);
      let seg1 = xml.slice(at, end1).replace('<p:nvPr></p:nvPr>', `<p:nvPr><p:ph type="${p.ph}"/></p:nvPr>`);
      // body paragraphs take the layout's bullets and indents (pptxgenjs writes buNone + zero indent on every paragraph)
      if (p.ph === 'body') seg1 = seg1.replace(/ indent="0" marL="0"/g, '').replace(/<a:buNone\/>/g, '');
      // mixed-weight title: the non-bold runs say so (pptxgenjs omits b="0", and layout titles are often bold)
      if (p.ph === 'title' && /<a:rPr\b[^>]*\bb="1"/.test(seg1)) seg1 = seg1.replace(/<a:rPr\b(?![^>]*\bb=")/g, '<a:rPr b="0"');
      xml = xml.slice(0, at) + seg1 + xml.slice(end1);
    }
    if (p.tgRuns && p.tgRuns.length) {
      const end0 = xml.indexOf('</p:sp>', at);
      let seg0 = xml.slice(at, end0);
      for (const { text, grad: g0 } of p.tgRuns) {
        let grad = g0;
        if (g0.span && p.tbox && g0.ang === 0) {
          const { x0, x1 } = g0.span;
          grad = { ...g0, stops: g0.stops.map((st) => ({ ...st, pos: Math.max(0, Math.min(100, ((x0 + (st.pos / 100) * (x1 - x0) - p.tbox.x) / p.tbox.w) * 100)) })) };
        }
        const esc = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        const lit = esc.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const re = new RegExp('<a:r>(<a:rPr[^>]*>)((?:(?!<\\/a:rPr>)[\\s\\S])*?)(<\\/a:rPr><a:t>' + lit + '<\\/a:t><\\/a:r>)');
        seg0 = seg0.replace(re, (m, open, inner, close) => `<a:r>${open}${inner.replace(/<a:solidFill>.*?<\/a:solidFill>/s, gradXml(grad))}${close}`);
      }
      xml = xml.slice(0, at) + seg0 + xml.slice(end0);
    }
    if (p.warp || p.tgrad) {
      const end = xml.indexOf('</p:sp>', at);
      let seg = xml.slice(at, end);
      // stretched single-line text: "plain" text warp makes PowerPoint scale the glyphs to the box
      if (p.warp) seg = seg.replace(/<a:bodyPr([^>]*?)(\/?)>/, (m, attrs, self) => self ? `<a:bodyPr${attrs}><a:prstTxWarp prst="textPlain"><a:avLst/></a:prstTxWarp></a:bodyPr>` : `<a:bodyPr${attrs}><a:prstTxWarp prst="textPlain"><a:avLst/></a:prstTxWarp>`);
      if (p.tgrad) seg = seg.replace(/<a:rPr([^>]*)>(.*?)<\/a:rPr>/gs, (m, attrs, inner) => `<a:rPr${attrs}>${inner.replace(/<a:solidFill>.*?<\/a:solidFill>/s, gradXml(p.tgrad))}</a:rPr>`);
      xml = xml.slice(0, at) + seg + xml.slice(end);
      continue;
    }
    const s = xml.indexOf('<p:spPr>', at), e = xml.indexOf('</p:spPr>', s);
    let sp = xml.slice(s, e);
    if (p.grad) sp = sp.replace(/<a:solidFill>.*?<\/a:solidFill>/s, gradXml(p.grad));
    if (p.inner) sp = /<a:effectLst>/.test(sp) ? sp.replace('<a:effectLst>', '<a:effectLst>' + innerXml(p.inner)) : sp + `<a:effectLst>${innerXml(p.inner)}</a:effectLst>`;
    if (p.blur) { const blur = `<a:blur rad="${Math.round(p.blur * 3 * 9525)}" grow="0"/>`; sp = /<a:effectLst>/.test(sp) ? sp.replace('<a:effectLst>', '<a:effectLst>' + blur) : sp + `<a:effectLst>${blur}</a:effectLst>`; }
    xml = xml.slice(0, s) + sp + xml.slice(e);
  }
  zip.file(file, xml);
}
let embedReport = [];
if (!NO_EMBED && webFonts.length) ({ report: embedReport } = await embedFonts(zip, fontUse.filter((u) => webFonts.includes(u.family)), faces));
if (TEMPLATE) await mergeTemplate(zip, slidesData);
writeFileSync(resolve(output), await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
await browser.close();
console.log(`wrote ${output}: ${slidesData.length} slides, ${stats.texts} text boxes, ${stats.shapes} shapes, ${stats.images} pictures (${stats.svg} vector SVG), ${stats.outlined} text items as vector outlines`);
console.log(`fonts used (must be installed where the pptx is opened): ${[...fonts].join(', ')}`);
embedReport.forEach((l) => console.log('font: ' + l));
const unembedded = webFonts.filter((f) => !restricted.has(f) && !embedReport.some((l) => l.startsWith(`embedded ${f} `)));
if (unembedded.length) console.log(`WARNING web fonts not embedded (PowerPoint needs them installed, or map with --font-map): ${unembedded.join(', ')}`);
