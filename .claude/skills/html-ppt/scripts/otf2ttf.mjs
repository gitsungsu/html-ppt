// Pure-JS CFF (OpenType/.otf) -> TrueType converter, so PowerPoint can embed the font.
// Cubic CFF curves are approximated by quadratic TrueType curves within 0.5 font units.
// Reads with opentype.js; writes head/hhea/maxp/cmap/loca/glyf/hmtx/name/OS2/post directly.
import opentype from 'opentype.js';

const TOL = 0.5;

function cubicToQuads(p0, p1, p2, p3, out, depth = 0) {
  const q = { x: (3 * p1.x + 3 * p2.x - p0.x - p3.x) / 4, y: (3 * p1.y + 3 * p2.y - p0.y - p3.y) / 4 };
  const cm = { x: (p0.x + 3 * p1.x + 3 * p2.x + p3.x) / 8, y: (p0.y + 3 * p1.y + 3 * p2.y + p3.y) / 8 };
  const qm = { x: (p0.x + 2 * q.x + p3.x) / 4, y: (p0.y + 2 * q.y + p3.y) / 4 };
  if (Math.hypot(cm.x - qm.x, cm.y - qm.y) <= TOL || depth >= 8) { out.push({ c: q, e: p3 }); return out; }
  const mid = (a, b) => ({ x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 });
  const a = mid(p0, p1), b = mid(p1, p2), c = mid(p2, p3), d = mid(a, b), e = mid(b, c), f = mid(d, e);
  cubicToQuads(p0, a, d, f, out, depth + 1);
  cubicToQuads(f, e, c, p3, out, depth + 1);
  return out;
}

// opentype.js path commands -> TrueType contours [{x,y,on}]
function contoursOf(path) {
  const contours = [];
  let cur = null, start = null, last = null;
  for (const c of path.commands) {
    if (c.type === 'M') { cur = [{ x: c.x, y: c.y, on: true }]; contours.push(cur); start = last = { x: c.x, y: c.y }; }
    else if (c.type === 'L') { cur.push({ x: c.x, y: c.y, on: true }); last = { x: c.x, y: c.y }; }
    else if (c.type === 'Q') { cur.push({ x: c.x1, y: c.y1, on: false }, { x: c.x, y: c.y, on: true }); last = { x: c.x, y: c.y }; }
    else if (c.type === 'C') {
      for (const q of cubicToQuads(last, { x: c.x1, y: c.y1 }, { x: c.x2, y: c.y2 }, { x: c.x, y: c.y }, [])) {
        cur.push({ x: q.c.x, y: q.c.y, on: false }, { x: q.e.x, y: q.e.y, on: true });
      }
      last = { x: c.x, y: c.y };
    } else if (c.type === 'Z') last = start;
  }
  for (const k of contours) {
    if (k.length > 1) {
      const a = k[0], b = k[k.length - 1];
      if (a.on && b.on && Math.abs(a.x - b.x) < 0.01 && Math.abs(a.y - b.y) < 0.01) k.pop();
    }
  }
  return contours.filter((k) => k.length >= 2);
}

function simpleGlyph(contours) {
  const pts = contours.flat();
  if (!pts.length) return null;
  const R = pts.map((p) => ({ x: Math.round(p.x), y: Math.round(p.y), on: p.on }));
  const xMin = Math.min(...R.map((p) => p.x)), yMin = Math.min(...R.map((p) => p.y));
  const xMax = Math.max(...R.map((p) => p.x)), yMax = Math.max(...R.map((p) => p.y));
  const endPts = [];
  let n = 0;
  for (const k of contours) { n += k.length; endPts.push(n - 1); }
  const body = [];
  const head = Buffer.alloc(10);
  head.writeInt16BE(contours.length, 0);
  head.writeInt16BE(xMin, 2); head.writeInt16BE(yMin, 4); head.writeInt16BE(xMax, 6); head.writeInt16BE(yMax, 8);
  body.push(head);
  const ep = Buffer.alloc(endPts.length * 2 + 2);
  endPts.forEach((e, i) => ep.writeUInt16BE(e, i * 2));
  ep.writeUInt16BE(0, endPts.length * 2); // instructionLength = 0
  body.push(ep);
  body.push(Buffer.from(R.map((p) => (p.on ? 1 : 0))));
  const xs = Buffer.alloc(R.length * 2), ys = Buffer.alloc(R.length * 2);
  let px = 0, py = 0;
  R.forEach((p, i) => { xs.writeInt16BE(p.x - px, i * 2); ys.writeInt16BE(p.y - py, i * 2); px = p.x; py = p.y; });
  body.push(xs, ys);
  let buf = Buffer.concat(body);
  if (buf.length % 4) buf = Buffer.concat([buf, Buffer.alloc(4 - (buf.length % 4))]);
  return { buf, bbox: [xMin, yMin, xMax, yMax], npts: R.length, ncont: contours.length };
}

function cmapFormat4(map) { // Map<codepoint, gid>, BMP only
  const codes = [...map.keys()].filter((c) => c < 0xffff).sort((a, b) => a - b);
  const segs = [];
  for (const c of codes) {
    const g = map.get(c), d = (g - c + 65536) & 0xffff;
    const last = segs[segs.length - 1];
    if (last && last.end === c - 1 && last.delta === d) last.end = c;
    else segs.push({ start: c, end: c, delta: d });
  }
  segs.push({ start: 0xffff, end: 0xffff, delta: 1 });
  const segX2 = segs.length * 2;
  const pow2 = 2 ** Math.floor(Math.log2(segs.length));
  const len = 16 + segX2 * 4;
  const b = Buffer.alloc(len);
  b.writeUInt16BE(4, 0); b.writeUInt16BE(len, 2); b.writeUInt16BE(0, 4); b.writeUInt16BE(segX2, 6);
  b.writeUInt16BE(pow2 * 2, 8); b.writeUInt16BE(Math.log2(pow2), 10); b.writeUInt16BE(segX2 - pow2 * 2, 12);
  let o = 14;
  segs.forEach((s) => { b.writeUInt16BE(s.end, o); o += 2; });
  b.writeUInt16BE(0, o); o += 2;
  segs.forEach((s) => { b.writeUInt16BE(s.start, o); o += 2; });
  segs.forEach((s) => { b.writeUInt16BE(s.delta, o); o += 2; });
  segs.forEach(() => { b.writeUInt16BE(0, o); o += 2; });
  return b;
}

function cmapTable(map) {
  const sub = cmapFormat4(map);
  const b = Buffer.alloc(12);
  b.writeUInt16BE(0, 0); b.writeUInt16BE(1, 2); b.writeUInt16BE(3, 4); b.writeUInt16BE(1, 6); b.writeUInt32BE(12, 8);
  return Buffer.concat([b, sub]);
}

function nameTable(strings) {
  const recs = [];
  const data = [];
  let off = 0;
  for (const [id, s] of strings) {
    const d = Buffer.from(s, 'utf16le').swap16();
    recs.push({ id, len: d.length, off });
    data.push(d);
    off += d.length;
  }
  const head = Buffer.alloc(6 + recs.length * 12);
  head.writeUInt16BE(0, 0); head.writeUInt16BE(recs.length, 2); head.writeUInt16BE(head.length, 4);
  recs.forEach((r, i) => {
    const o = 6 + i * 12;
    head.writeUInt16BE(3, o); head.writeUInt16BE(1, o + 2); head.writeUInt16BE(0x409, o + 4);
    head.writeUInt16BE(r.id, o + 6); head.writeUInt16BE(r.len, o + 8); head.writeUInt16BE(r.off, o + 10);
  });
  return Buffer.concat([head, ...data]);
}

function os2Table(font) {
  const o = font.tables.os2 || {};
  const b = Buffer.alloc(86);
  b.writeUInt16BE(1, 0);
  b.writeInt16BE(o.xAvgCharWidth || 0, 2);
  b.writeUInt16BE(o.usWeightClass || 400, 4);
  b.writeUInt16BE(5, 6);
  b.writeUInt16BE(o.fsType || 0, 8);
  const sub = [o.ySubscriptXSize, o.ySubscriptYSize, o.ySubscriptXOffset, o.ySubscriptYOffset, o.ySuperscriptXSize, o.ySuperscriptYSize, o.ySuperscriptXOffset, o.ySuperscriptYOffset, o.yStrikeoutSize, o.yStrikeoutPosition];
  sub.forEach((v, i) => b.writeInt16BE(v || 0, 10 + i * 2));
  if (Array.isArray(o.panose)) o.panose.slice(0, 10).forEach((v, i) => b.writeUInt8(v, 32 + i));
  [o.ulUnicodeRange1, o.ulUnicodeRange2, o.ulUnicodeRange3, o.ulUnicodeRange4].forEach((v, i) => b.writeUInt32BE((v || 0) >>> 0, 42 + i * 4));
  b.write('NONE', 58, 'latin1');
  b.writeUInt16BE((o.fsSelection || 0x40) & 0xffff, 62);
  b.writeUInt16BE(o.usFirstCharIndex || 0, 64);
  b.writeUInt16BE(o.usLastCharIndex || 0xffff, 66);
  b.writeInt16BE(o.sTypoAscender || font.ascender, 68);
  b.writeInt16BE(o.sTypoDescender || font.descender, 70);
  b.writeInt16BE(o.sTypoLineGap || 0, 72);
  b.writeUInt16BE(o.usWinAscent || font.ascender, 74);
  b.writeUInt16BE(o.usWinDescent || -font.descender, 76);
  b.writeUInt32BE(o.ulCodePageRange1 || 1, 78);
  b.writeUInt32BE(o.ulCodePageRange2 || 0, 82);
  return b;
}

function checksum(buf) {
  let s = 0;
  const p = Buffer.concat([buf, Buffer.alloc((4 - (buf.length % 4)) % 4)]);
  for (let i = 0; i < p.length; i += 4) s = (s + p.readUInt32BE(i)) >>> 0;
  return s;
}

// name strings straight from the source file's name table (opentype.js does not expose every platform's names)
function readNames(buf) {
  const out = {};
  const n = buf.readUInt16BE(4);
  for (let i = 0; i < n; i++) {
    const o = 12 + i * 16;
    if (buf.slice(o, o + 4).toString('latin1') !== 'name') continue;
    const t = buf.readUInt32BE(o + 8);
    const count = buf.readUInt16BE(t + 2), strOff = t + buf.readUInt16BE(t + 4);
    for (let r = 0; r < count; r++) {
      const q = t + 6 + r * 12;
      const plat = buf.readUInt16BE(q), lang = buf.readUInt16BE(q + 4), id = buf.readUInt16BE(q + 6), len = buf.readUInt16BE(q + 8), off = strOff + buf.readUInt16BE(q + 10);
      if (plat !== 3 || lang !== 0x409) continue; // en-US Windows records, the set PowerPoint reads
      const raw = buf.slice(off, off + len);
      const s = plat === 3 || plat === 0 ? Buffer.from(raw).swap16().toString('utf16le') : raw.toString('latin1');
      out[id] = s;
    }
  }
  return out;
}

// OpenType (CFF) font bytes -> TrueType font bytes
export function otfToTtf(otfBuf) {
  const font = opentype.parse(otfBuf.buffer.slice(otfBuf.byteOffset, otfBuf.byteOffset + otfBuf.byteLength));
  const n = font.glyphs.length;
  const glyphs = [], offsets = [0], advances = [], lsbs = [];
  const cmap = new Map();
  let maxPts = 0, maxCont = 0;
  let bb = [Infinity, Infinity, -Infinity, -Infinity];
  for (let i = 0; i < n; i++) {
    const g = font.glyphs.get(i);
    // CFF outer contours run counter-clockwise, TrueType expects clockwise: reverse each contour
    const contours = contoursOf(g.path).map((k) => k.reverse());
    const sg = simpleGlyph(contours);
    if (sg) {
      glyphs.push(sg.buf);
      maxPts = Math.max(maxPts, sg.npts); maxCont = Math.max(maxCont, sg.ncont);
      bb = [Math.min(bb[0], sg.bbox[0]), Math.min(bb[1], sg.bbox[1]), Math.max(bb[2], sg.bbox[2]), Math.max(bb[3], sg.bbox[3])];
      lsbs.push(sg.bbox[0]);
    } else { glyphs.push(Buffer.alloc(0)); lsbs.push(0); }
    offsets.push(offsets[offsets.length - 1] + glyphs[glyphs.length - 1].length);
    advances.push(Math.round(g.advanceWidth || 0));
    for (const u of g.unicodes || (g.unicode !== undefined ? [g.unicode] : [])) if (!cmap.has(u)) cmap.set(u, i);
  }
  if (bb[0] === Infinity) bb = [0, 0, 0, 0];
  const glyf = Buffer.concat(glyphs);
  const loca = Buffer.alloc(offsets.length * 4);
  offsets.forEach((v, i) => loca.writeUInt32BE(v, i * 4));
  const upem = font.unitsPerEm;
  const headT = Buffer.alloc(54);
  headT.writeUInt32BE(0x00010000, 0); headT.writeUInt32BE(0x00010000, 4); headT.writeUInt32BE(0, 8); headT.writeUInt32BE(0x5f0f3cf5, 12);
  headT.writeUInt16BE(0x000b, 16); headT.writeUInt16BE(upem, 18);
  const now = Math.floor(Date.now() / 1000) + 2082844800; // seconds since 1904
  headT.writeBigUInt64BE(BigInt(now), 20); headT.writeBigUInt64BE(BigInt(now), 28);
  headT.writeInt16BE(bb[0], 36); headT.writeInt16BE(bb[1], 38); headT.writeInt16BE(bb[2], 40); headT.writeInt16BE(bb[3], 42);
  headT.writeUInt16BE(0, 44); headT.writeUInt16BE(8, 46); headT.writeInt16BE(2, 48); headT.writeInt16BE(1, 50); headT.writeInt16BE(0, 52);
  const hhea = Buffer.alloc(36);
  hhea.writeUInt32BE(0x00010000, 0); hhea.writeInt16BE(font.ascender, 4); hhea.writeInt16BE(font.descender, 6);
  hhea.writeInt16BE(0, 8); hhea.writeUInt16BE(Math.max(...advances), 10);
  hhea.writeInt16BE(Math.min(...lsbs), 12); hhea.writeInt16BE(0, 14); hhea.writeInt16BE(bb[2], 16);
  hhea.writeInt16BE(1, 18); hhea.writeInt16BE(0, 20); hhea.writeUInt16BE(n, 34); // numberOfHMetrics
  const hmtx = Buffer.alloc(n * 4);
  advances.forEach((a, i) => { hmtx.writeUInt16BE(a, i * 4); hmtx.writeInt16BE(lsbs[i], i * 4 + 2); });
  const maxp = Buffer.alloc(32);
  maxp.writeUInt32BE(0x00010000, 0); maxp.writeUInt16BE(n, 4); maxp.writeUInt16BE(maxPts, 6); maxp.writeUInt16BE(maxCont, 8);
  maxp.writeUInt16BE(1, 14); // maxZones
  const post = Buffer.alloc(32);
  post.writeUInt32BE(0x00030000, 0);
  const names = readNames(otfBuf);
  const fam = names[1] || 'Font', sub = names[2] || 'Regular';
  const ps = (names[6] || `${fam}-${sub}`).replace(/\s/g, '');
  const recs = Object.entries(names).map(([id, v]) => [Number(id), v]);
  if (!names[1]) recs.push([1, fam]);
  if (!names[2]) recs.push([2, sub]);
  if (!names[4]) recs.push([4, `${fam} ${sub}`]);
  if (!names[5]) recs.push([5, 'Version 1.000']);
  if (!names[6]) recs.push([6, ps]);
  const name = nameTable(recs.sort((a, b) => a[0] - b[0]));
  const tables = {
    cmap: cmapTable(cmap), glyf, head: headT, hhea, hmtx, loca, maxp, name, 'OS/2': os2Table(font), post,
  };
  const order = Object.keys(tables).sort((a, b) => (a.toLowerCase() < b.toLowerCase() ? -1 : a.toLowerCase() > b.toLowerCase() ? 1 : 0));
  const nt = order.length;
  const dirLen = 12 + nt * 16;
  let off = dirLen;
  const dir = Buffer.alloc(dirLen), datas = [];
  dir.writeUInt32BE(0x00010000, 0); dir.writeUInt16BE(nt, 4);
  const pow = 2 ** Math.floor(Math.log2(nt));
  dir.writeUInt16BE(pow * 16, 6); dir.writeUInt16BE(Math.log2(pow), 8); dir.writeUInt16BE(nt * 16 - pow * 16, 10);
  order.forEach((tag, i) => {
    const t = tables[tag];
    const padded = Buffer.concat([t, Buffer.alloc((4 - (t.length % 4)) % 4)]);
    if (tag === 'head') padded.writeUInt32BE(0, 8);
    const o = 12 + i * 16;
    dir.write(tag.padEnd(4, ' '), o, 'latin1');
    dir.writeUInt32BE(checksum(padded), o + 4);
    dir.writeUInt32BE(off, o + 8);
    dir.writeUInt32BE(t.length, o + 12);
    datas.push(padded);
    off += padded.length;
  });
  const out = Buffer.concat([dir, ...datas]);
  const headOff = 12 + order.indexOf('head') * 16 + 8;
  const adj = (0xb1b0afba - checksum(out)) >>> 0;
  out.writeUInt32BE(adj, out.readUInt32BE(headOff) + 8);
  return out;
}
