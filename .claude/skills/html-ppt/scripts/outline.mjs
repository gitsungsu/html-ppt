// Text as vector outlines: glyph paths -> native PowerPoint freeform shapes (cubic/quadratic curves kept exact).
// Used where live text cannot express the look (blur, skew, fonts whose licence forbids embedding).
import opentype from 'opentype.js';
import { existsSync, readFileSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { join } from 'node:path';
import { loadSfnt } from './embed_fonts.mjs';

let regCache = null;
// installed font files by registry name (Windows) -> path
function windowsFonts() {
  if (regCache) return regCache;
  regCache = new Map();
  for (const key of ['HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts', 'HKCU\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts']) {
    const r = spawnSync('reg', ['query', key], { encoding: 'utf8' });
    if (r.error || r.status !== 0) continue;
    for (const line of r.stdout.split(/\r?\n/)) {
      const m = line.match(/^\s+(.+?)\s+\(\w+\)\s+REG_SZ\s+(.+)$/);
      if (!m) continue;
      const file = /^[A-Za-z]:\\/.test(m[2]) ? m[2] : join(process.env.WINDIR || 'C:\\Windows', 'Fonts', m[2]);
      regCache.set(m[1].toLowerCase(), file.trim());
    }
  }
  return regCache;
}

function systemFontFile(family, bold, italic) {
  if (process.platform === 'win32') {
    const m = windowsFonts();
    const f = family.toLowerCase();
    const names = [];
    if (bold && italic) names.push(`${f} bold italic`);
    if (bold) names.push(`${f} bold`);
    if (italic) names.push(`${f} italic`);
    names.push(f);
    for (const n of names) if (m.has(n)) return m.get(n);
    return null;
  }
  const r = spawnSync('fc-match', ['-f', '%{file}', `${family}:${bold ? 'bold' : ''}${italic ? ':italic' : ''}`], { encoding: 'utf8' });
  return r.status === 0 && r.stdout && existsSync(r.stdout) ? r.stdout : null;
}

const toArrayBuffer = (b) => b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);

// (family, bold, italic) -> opentype.js Font, or null when the file cannot be read
export function makeFontLoader(faces) {
  const cache = new Map();
  return async function fontOf(family, bold, italic) {
    const key = `${family}|${bold ? 1 : 0}|${italic ? 1 : 0}`;
    if (cache.has(key)) return cache.get(key);
    let font = null;
    try {
      let buf = null;
      const cand = faces.filter((f) => f.family === family);
      if (cand.length) {
        const want = (f) => (parseInt(f.weight) >= 600 || f.weight === 'bold' ? 1 : 0) === (bold ? 1 : 0) && (f.style === 'italic' || f.style === 'oblique') === !!italic;
        const face = cand.find(want) || cand[0];
        buf = (await loadSfnt(face.url)).buf;
      } else {
        const file = systemFontFile(family, bold, italic);
        if (file && !/\.ttc$/i.test(file)) buf = readFileSync(file);
      }
      if (buf) font = opentype.parse(toArrayBuffer(buf));
    } catch { font = null; }
    cache.set(key, font);
    return font;
  };
}

// item.gl: [{c,x,y,size,bold,italic,font,color,alpha}] (x,y = pen origin, baseline, slide px)
// item.tm: affine [a,b,c,d,e,f] applied to slide px (rotation / scale / skew of the enclosing element)
// -> [{color, alpha, contours:[[{op:'M'|'L'|'Q'|'C'|'Z', pts:[{x,y}]}]]}] grouped by colour, or null when a font/glyph is missing
export async function outlineGroups(item, fontOf) {
  const tm = item.tm || [1, 0, 0, 1, 0, 0];
  const groups = new Map();
  for (const g of item.gl) {
    const font = await fontOf(g.font, g.bold, g.italic);
    if (!font) return null;
    const glyph = font.charToGlyph(g.c);
    if (!glyph || glyph.index === 0 && g.c !== '\u0000') return null;
    const path = glyph.getPath(g.x, g.y, g.size).commands;
    const key = g.color + '|' + g.alpha;
    if (!groups.has(key)) groups.set(key, { color: g.color, alpha: g.alpha, contours: [] });
    const map = (p) => ({ x: tm[0] * p.x + tm[2] * p.y + tm[4], y: tm[1] * p.x + tm[3] * p.y + tm[5] });
    let cur = null;
    for (const c of path) {
      if (c.type === 'M') { cur = []; groups.get(key).contours.push(cur); cur.push({ op: 'M', pts: [map(c)] }); }
      else if (c.type === 'L') cur.push({ op: 'L', pts: [map(c)] });
      else if (c.type === 'Q') cur.push({ op: 'Q', pts: [map({ x: c.x1, y: c.y1 }), map(c)] });
      else if (c.type === 'C') cur.push({ op: 'C', pts: [map({ x: c.x1, y: c.y1 }), map({ x: c.x2, y: c.y2 }), map(c)] });
      else if (c.type === 'Z') cur.push({ op: 'Z', pts: [] });
    }
  }
  return [...groups.values()];
}

// groups -> pptxgenjs custGeom: bounding box (slide px) + points relative to it, in inches
export function shapeOf(group, px) {
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  for (const c of group.contours) for (const o of c) for (const p of o.pts) { x0 = Math.min(x0, p.x); y0 = Math.min(y0, p.y); x1 = Math.max(x1, p.x); y1 = Math.max(y1, p.y); }
  if (!isFinite(x0)) return null;
  const r = (p) => ({ x: +((p.x - x0) * px).toFixed(4), y: +((p.y - y0) * px).toFixed(4) });
  const points = [];
  for (const c of group.contours) {
    for (const o of c) {
      if (o.op === 'M') points.push({ ...r(o.pts[0]), moveTo: true });
      else if (o.op === 'L') points.push(r(o.pts[0]));
      else if (o.op === 'Q') { const ctl = r(o.pts[0]), end = r(o.pts[1]); points.push({ x: end.x, y: end.y, curve: { type: 'quadratic', x1: ctl.x, y1: ctl.y } }); }
      else if (o.op === 'C') { const c1 = r(o.pts[0]), c2 = r(o.pts[1]), end = r(o.pts[2]); points.push({ x: end.x, y: end.y, curve: { type: 'cubic', x1: c1.x, y1: c1.y, x2: c2.x, y2: c2.y } }); }
      else if (o.op === 'Z') points.push({ close: true });
    }
  }
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0, points };
}
