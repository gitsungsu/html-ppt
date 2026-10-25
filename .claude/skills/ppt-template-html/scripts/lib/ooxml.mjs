// OOXML helpers: guarded zip access (AC7), a small XML parser, relationships, colour maths.
import { readFile } from 'node:fs/promises';
import { statSync } from 'node:fs';
import { posix as pp } from 'node:path';
import JSZip from 'jszip';

export class InputError extends Error {}

export const LIMITS = {
  fileBytes: 50 * 1024 * 1024,
  entries: 2000,
  uncompressedBytes: 200 * 1024 * 1024,
};

// ---------------------------------------------------------------- zip (AC7)

export async function openPackage(file) {
  let st;
  try { st = statSync(file); } catch { throw new InputError(`cannot read ${file}`); }
  if (!st.isFile()) throw new InputError(`not a file: ${file}`);
  if (st.size > LIMITS.fileBytes) throw new InputError(`${file}: file is larger than 50MB`);
  let buf;
  try { buf = await readFile(file); } catch { throw new InputError(`cannot read ${file}`); }

  // Entry count from the end-of-central-directory record, before JSZip builds its entry table.
  const tail = Math.max(0, buf.length - 65557);
  for (let i = buf.length - 22; i >= tail; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) {
      if (buf.readUInt16LE(i + 10) > LIMITS.entries) {
        throw new InputError(`${file}: more than ${LIMITS.entries} zip entries`);
      }
      break;
    }
  }

  let zip;
  try { zip = await JSZip.loadAsync(buf); }
  catch (e) { throw new InputError(`${file}: not a valid zip package (${e.message})`); }

  const names = Object.keys(zip.files);
  if (names.length > LIMITS.entries) throw new InputError(`${file}: more than ${LIMITS.entries} zip entries`);
  let declared = 0;
  for (const n of names) {
    const sz = zip.files[n]._data && zip.files[n]._data.uncompressedSize;
    if (typeof sz === 'number') declared += sz;
  }
  if (declared > LIMITS.uncompressedBytes) {
    throw new InputError(`${file}: uncompressed size exceeds 200MB`);
  }

  let readTotal = 0;
  return {
    names,
    has: (name) => !!zip.file(name.replace(/^\//, '')),
    async read(name) {
      const f = zip.file(name.replace(/^\//, ''));
      if (!f) return null;
      let data;
      try { data = await f.async('nodebuffer'); }
      catch (e) { throw new InputError(`${file}: cannot read ${name} (${e.message})`); }
      readTotal += data.length;
      if (readTotal > LIMITS.uncompressedBytes) throw new InputError(`${file}: uncompressed size exceeds 200MB`);
      return data;
    },
  };
}

// ---------------------------------------------------------------- XML

const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
const decode = (s) => (s.indexOf('&') < 0 ? s : s.replace(/&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z]+);/g, (m, e) => {
  if (e[0] === '#') {
    const cp = e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
    try { return String.fromCodePoint(cp); } catch { return m; }
  }
  return ENT[e] ?? m;
}));
const localName = (n) => n.slice(n.indexOf(':') + 1);

function parseAttrs(str) {
  const attrs = {};
  if (!str) return attrs;
  const re = /([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  let m;
  while ((m = re.exec(str))) attrs[m[1]] = decode(m[2] ?? m[3]);
  return attrs;
}

// Element names are stored without their namespace prefix; attribute names keep it.
export function parseXml(src) {
  if (src.charCodeAt(0) === 0xFEFF) src = src.slice(1);
  const root = { name: '#root', attrs: {}, children: [], text: '' };
  const stack = [root];
  const re = /<!--[\s\S]*?-->|<!\[CDATA\[([\s\S]*?)\]\]>|<\?[\s\S]*?\?>|<!DOCTYPE[^>]*>|<\/\s*([^\s>]+)\s*>|<([^\s\/>!?][^\s\/>]*)((?:\s+[^\s=\/>]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|([^<]+)/g;
  let m;
  while ((m = re.exec(src))) {
    const top = stack[stack.length - 1];
    if (m[1] !== undefined) top.text += m[1];
    else if (m[2] !== undefined) { if (stack.length > 1) stack.pop(); }
    else if (m[3] !== undefined) {
      const node = { name: localName(m[3]), attrs: parseAttrs(m[4]), children: [], text: '' };
      top.children.push(node);
      if (!m[5]) stack.push(node);
    } else if (m[6] !== undefined) top.text += decode(m[6]);
  }
  if (!root.children[0]) throw new InputError('empty or invalid XML part');
  return root.children[0];
}

export const kids = (n, name) => (n ? n.children.filter((c) => c.name === name) : []);
export const kid = (n, name) => (n ? n.children.find((c) => c.name === name) : undefined);
export const dig = (n, ...names) => names.reduce((x, nm) => kid(x, nm), n);
// Relationship-namespaced attribute (r:id, r:embed ...) by local name.
export function rattr(n, local) {
  if (!n) return undefined;
  for (const k of Object.keys(n.attrs)) {
    if (k.includes(':') && k.slice(k.indexOf(':') + 1) === local && !k.startsWith('xmlns')) return n.attrs[k];
  }
  return undefined;
}

export async function readXml(pkg, part) {
  const buf = await pkg.read(part);
  if (!buf) throw new InputError(`package part missing: ${part}`);
  try { return parseXml(buf.toString('utf8')); }
  catch (e) { throw new InputError(`cannot parse ${part}: ${e.message}`); }
}

// ---------------------------------------------------------------- relationships

export function resolveTarget(part, target) {
  let t = target;
  try { t = decodeURIComponent(target); } catch { /* keep raw */ }
  if (t.startsWith('/')) return pp.normalize(t.slice(1));
  return pp.normalize(pp.join(pp.dirname(part), t));
}

// part '' = package root. Returns Map id -> {type, target}. Missing rels file -> empty map.
export async function readRels(pkg, part) {
  const relsPath = part === '' ? '_rels/.rels' : pp.join(pp.dirname(part), '_rels', pp.basename(part) + '.rels');
  const map = new Map();
  const buf = await pkg.read(relsPath);
  if (!buf) return map;
  const root = parseXml(buf.toString('utf8'));
  for (const r of kids(root, 'Relationship')) {
    if (r.attrs.TargetMode === 'External') continue;
    map.set(r.attrs.Id, { type: r.attrs.Type || '', target: resolveTarget(part, r.attrs.Target || '') });
  }
  return map;
}

export function findRel(rels, typeSuffix) {
  for (const r of rels.values()) if (r.type.endsWith(typeSuffix)) return r.target;
  return undefined;
}

// ---------------------------------------------------------------- colour

export const COLOR_EL = new Set(['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'hslClr', 'scrgbClr']);
export const colorChild = (n) => (n ? n.children.find((c) => COLOR_EL.has(c.name)) : undefined);

const PRESET = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', lime: '00FF00', blue: '0000FF',
  yellow: 'FFFF00', cyan: '00FFFF', aqua: '00FFFF', magenta: 'FF00FF', fuchsia: 'FF00FF',
  gray: '808080', grey: '808080', silver: 'C0C0C0', orange: 'FFA500', purple: '800080',
  maroon: '800000', navy: '000080', teal: '008080', olive: '808000', brown: 'A52A2A',
  pink: 'FFC0CB', dkGray: 'A9A9A9', dkGrey: 'A9A9A9', ltGray: 'D3D3D3', ltGrey: 'D3D3D3',
};
const SYS = { windowText: '000000', window: 'FFFFFF' };

export function hexToRgb(h) {
  const s = String(h || '000000').replace('#', '').padEnd(6, '0');
  return [parseInt(s.slice(0, 2), 16) || 0, parseInt(s.slice(2, 4), 16) || 0, parseInt(s.slice(4, 6), 16) || 0];
}
export function rgbToHex(rgb) {
  return '#' + rgb.map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('').toUpperCase();
}

function rgbToHsl([r, g, b]) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, s, l];
}
function hslToRgb([h, s, l]) {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [f(h + 1 / 3) * 255, f(h) * 255, f(h - 1 / 3) * 255];
}
const toLin = (c) => { c /= 255; return c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4; };
const fromLin = (c) => 255 * (c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055);
const clamp01 = (x) => Math.max(0, Math.min(1, x));

// ctx: { theme: {dk1:'RRGGBB',...}, clrMap: {bg1:'lt1',...}, phClr: [r,g,b]|null }
// Returns [r,g,b] floats (unrounded) so chained modifiers do not drift.
export function colorRgb(node, ctx) {
  let rgb;
  const a = node.attrs;
  switch (node.name) {
    case 'srgbClr': rgb = hexToRgb(a.val); break;
    case 'sysClr': rgb = hexToRgb(a.lastClr ?? SYS[a.val] ?? '000000'); break;
    case 'prstClr': rgb = hexToRgb(PRESET[a.val] ?? '000000'); break;
    case 'hslClr': rgb = hslToRgb([(+a.hue || 0) / 21600000, (+a.sat || 0) / 100000, (+a.lum || 0) / 100000]); break;
    case 'scrgbClr': rgb = [fromLin((+a.r || 0) / 100000), fromLin((+a.g || 0) / 100000), fromLin((+a.b || 0) / 100000)]; break;
    case 'schemeClr': {
      if (a.val === 'phClr') { rgb = ctx.phClr ? [...ctx.phClr] : [0, 0, 0]; break; }
      const slot = (ctx.clrMap && ctx.clrMap[a.val]) || a.val;
      rgb = hexToRgb(ctx.theme && ctx.theme[slot]);
      break;
    }
    default: rgb = [0, 0, 0];
  }
  for (const m of node.children) {
    const v = (+m.attrs.val || 0) / 100000;
    switch (m.name) {
      case 'lumMod': { const h = rgbToHsl(rgb); h[2] = clamp01(h[2] * v); rgb = hslToRgb(h); break; }
      case 'lumOff': { const h = rgbToHsl(rgb); h[2] = clamp01(h[2] + v); rgb = hslToRgb(h); break; }
      case 'satMod': { const h = rgbToHsl(rgb); h[1] = clamp01(h[1] * v); rgb = hslToRgb(h); break; }
      case 'satOff': { const h = rgbToHsl(rgb); h[1] = clamp01(h[1] + v); rgb = hslToRgb(h); break; }
      case 'hueMod': { const h = rgbToHsl(rgb); h[0] = (((h[0] * v) % 1) + 1) % 1; rgb = hslToRgb(h); break; }
      case 'hueOff': { const h = rgbToHsl(rgb); h[0] = ((((h[0] + (+m.attrs.val || 0) / 21600000) % 1) + 1) % 1); rgb = hslToRgb(h); break; }
      // ECMA-376 20.1.2.3.34 / .31: tint mixes with white, shade with black, in linear RGB.
      case 'tint': rgb = rgb.map((c) => fromLin(toLin(c) * v + (1 - v))); break;
      case 'shade': rgb = rgb.map((c) => fromLin(toLin(c) * v)); break;
      default: break; // alpha and others do not change the opaque colour
    }
  }
  return rgb;
}
export const colorHex = (node, ctx) => rgbToHex(colorRgb(node, ctx));

// ---------------------------------------------------------------- text style

// A source is { algn, b, sz, fill } (raw attribute strings, sz in hundredths of a point, and an a:solidFill node,
// each possibly undefined), ordered most specific first. Each field takes the first source that defines it.
// Returns { color: '#RRGGBB'|null, bold: true|false|null, align: 'l'|'ctr'|'r'|null, size: points|null }.
export function resolveTextStyle(sources, ctx) {
  const first = (f) => sources.find((s) => s && s[f] !== undefined)?.[f];
  const algn = first('algn');
  const b = first('b');
  const sz = +first('sz');
  const c = colorChild(first('fill'));
  return {
    color: c ? colorHex(c, ctx) : null,
    bold: b === '1' || b === 'true' ? true : b === '0' || b === 'false' ? false : null,
    align: algn === 'l' || algn === 'ctr' || algn === 'r' ? algn : null,
    size: sz > 0 ? sz / 100 : null,
  };
}
