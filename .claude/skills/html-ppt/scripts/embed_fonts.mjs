// Embed web fonts (@font-face TTF/WOFF) into a PPTX so PowerPoint can render them without installing anything.
// PowerPoint stores embedded fonts as EOT-wrapped TrueType in ppt/fonts/*.fntdata.
import { inflateSync } from 'node:zlib';
import { readFileSync } from 'node:fs';
import { otfToTtf } from './otf2ttf.mjs';
import { fileURLToPath } from 'node:url';

const u16 = (b, o) => b.readUInt16BE(o), u32 = (b, o) => b.readUInt32BE(o);

// WOFF 1 -> plain sfnt (TrueType)
function woffToSfnt(b) {
  const n = u16(b, 12), flavor = u32(b, 4);
  const dir = [];
  let off = 12 + n * 16;
  for (let i = 0; i < n; i++) {
    const o = 44 + i * 20;
    dir.push({ tag: b.slice(o, o + 4), offset: u32(b, o + 4), comp: u32(b, o + 8), orig: u32(b, o + 12), sum: u32(b, o + 16) });
  }
  const parts = [];
  const head = Buffer.alloc(12 + n * 16);
  head.writeUInt32BE(flavor, 0); head.writeUInt16BE(n, 4);
  const pow = 2 ** Math.floor(Math.log2(n)); head.writeUInt16BE(pow * 16, 6); head.writeUInt16BE(Math.log2(pow), 8); head.writeUInt16BE(n * 16 - pow * 16, 10);
  dir.forEach((t, i) => {
    const raw = b.slice(t.offset, t.offset + t.comp);
    const data = t.comp < t.orig ? inflateSync(raw) : raw;
    const pad = (4 - (data.length % 4)) % 4;
    const e = 12 + i * 16;
    t.tag.copy(head, e); head.writeUInt32BE(t.sum, e + 4); head.writeUInt32BE(off, e + 8); head.writeUInt32BE(data.length, e + 12);
    parts.push(data, Buffer.alloc(pad)); off += data.length + pad;
  });
  return Buffer.concat([head, ...parts]);
}

function tables(b) {
  const n = u16(b, 4), t = {};
  for (let i = 0; i < n; i++) { const o = 12 + i * 16; t[b.slice(o, o + 4).toString('latin1')] = { offset: u32(b, o + 8), length: u32(b, o + 12) }; }
  return t;
}

function nameOf(b, nameTab, id) {
  if (!nameTab) return '';
  const base = nameTab.offset, count = u16(b, base + 2), strOff = base + u16(b, base + 4);
  let best = '';
  for (let i = 0; i < count; i++) {
    const r = base + 6 + i * 12;
    if (u16(b, r + 6) !== id) continue;
    const plat = u16(b, r), len = u16(b, r + 8), o = strOff + u16(b, r + 10);
    const raw = b.slice(o, o + len);
    const s = plat === 3 || plat === 0 ? Buffer.from(raw).swap16().toString('utf16le') : raw.toString('latin1');
    if (plat === 3) return s;
    best = best || s;
  }
  return best;
}

// uncompressed EOT v2.1 around the sfnt bytes
function toEot(sfnt) {
  const t = tables(sfnt), os2 = t['OS/2'], head = t.head;
  const fsType = u16(sfnt, os2.offset + 8);
  const w = u16(sfnt, os2.offset + 4), fsSel = u16(sfnt, os2.offset + 62);
  const version = u16(sfnt, os2.offset);
  const str = (s) => { const d = Buffer.from(s, 'utf16le'); const h = Buffer.alloc(2); h.writeUInt16LE(d.length); return Buffer.concat([h, d]); };
  const pad = Buffer.alloc(2);
  const names = [nameOf(sfnt, t.name, 1), nameOf(sfnt, t.name, 2), nameOf(sfnt, t.name, 5), nameOf(sfnt, t.name, 4)];
  const h = Buffer.alloc(82);
  let o = 4; // EOTSize patched below
  h.writeUInt32LE(sfnt.length, o); o += 4;
  h.writeUInt32LE(0x00020001, o); o += 4;
  h.writeUInt32LE(0, o); o += 4; // flags: uncompressed, not XOR-obfuscated
  sfnt.copy(h, o, os2.offset + 32, os2.offset + 42); o += 10; // PANOSE
  h.writeUInt8(1, o++); h.writeUInt8(fsSel & 1, o++);
  h.writeUInt32LE(w, o); o += 4;
  h.writeUInt16LE(fsType, o); o += 2;
  h.writeUInt16LE(0x504c, o); o += 2;
  for (let i = 0; i < 4; i++) { h.writeUInt32LE(u32(sfnt, os2.offset + 42 + i * 4), o); o += 4; }
  for (let i = 0; i < 2; i++) { h.writeUInt32LE(version >= 1 ? u32(sfnt, os2.offset + 78 + i * 4) : 0, o); o += 4; }
  h.writeUInt32LE(u32(sfnt, head.offset + 8), o); o += 4; // checkSumAdjustment
  o += 16; // reserved
  const body = Buffer.concat([pad, str(names[0]), pad, str(names[1]), pad, str(names[2]), pad, str(names[3]), pad, Buffer.alloc(2)]);
  const out = Buffer.concat([h.slice(0, o), body, sfnt]);
  out.writeUInt32LE(out.length, 0);
  return { eot: out, fsType, family: names[0] };
}

// font file (ttf/otf/woff/woff2, any scheme) -> raw sfnt bytes + 'ttf' | 'otf'
export async function loadSfnt(url) {
  let buf = await load(url);
  const magic = buf.slice(0, 4).toString('latin1');
  if (magic === 'wOFF') buf = woffToSfnt(buf);
  else if (magic === 'wOF2') buf = Buffer.from(await (await import('wawoff2')).default.decompress(buf));
  return { buf, type: buf.slice(0, 4).toString('latin1') === 'OTTO' ? 'otf' : 'ttf' };
}

async function load(url) {
  if (url.startsWith('data:')) return Buffer.from(decodeURIComponent(url.split(',')[1]), /;base64/.test(url.split(',')[0]) ? 'base64' : 'utf8');
  if (url.startsWith('file:')) return readFileSync(fileURLToPath(url));
  const r = await fetch(url); if (!r.ok) throw new Error(`${r.status} ${url}`);
  return Buffer.from(await r.arrayBuffer());
}

// which of the used web fonts have an embedding-restricted licence (OS/2 fsType bit 1)
export async function restrictedFamilies(used, faces) {
  const out = new Set();
  for (const u of used) {
    const face = faces.find((f) => f.family === u.family);
    if (!face || out.has(u.family)) continue;
    try { const { buf } = await loadSfnt(face.url); const t = tables(buf); if (t['OS/2'] && (u16(buf, t['OS/2'].offset + 8) & 0x0002)) out.add(u.family); } catch { /* unreadable: handled at embed time */ }
  }
  return out;
}

// used: [{family,bold,italic}], faces: [{family,weight,style,url}] -> adds parts to the zip, returns report lines
export async function embedFonts(zip, used, faces) {
  const report = [];
  const restricted = new Set();
  const slots = {}; // family -> {regular,bold,italic,boldItalic: partName}
  let n = 0;
  for (const u of used) {
    const cand = faces.filter((f) => f.family === u.family);
    if (!cand.length) continue;
    const want = (f) => (parseInt(f.weight) >= 600 || f.weight === 'bold' ? 1 : 0) === (u.bold ? 1 : 0) && (f.style === 'italic' || f.style === 'oblique') === !!u.italic;
    const face = cand.find(want) || cand[0];
    const slot = u.bold && u.italic ? 'boldItalic' : u.bold ? 'bold' : u.italic ? 'italic' : 'regular';
    if (slots[u.family] && slots[u.family][slot]) continue;
    try {
      let { buf } = await loadSfnt(face.url);
      if (buf.slice(0, 4).toString('latin1') === 'OTTO') buf = otfToTtf(buf); // PowerPoint only embeds TrueType outlines
      const { eot, fsType, family } = toEot(buf);
      if (fsType & 0x0002) { report.push(`skip ${u.family}: font license forbids embedding (text is converted to vector outlines instead)`); restricted.add(u.family); continue; }
      n++;
      zip.file(`ppt/fonts/font${n}.fntdata`, eot);
      (slots[u.family] ||= {})[slot] = { part: `font${n}.fntdata`, rid: `rIdEmbF${n}` };
      report.push(`embedded ${u.family} (${slot})`);
    } catch (e) { report.push(`skip ${u.family}: ${e.message}`); }
  }
  if (!n) return { report, restricted };
  // content type, relationships, presentation.xml
  let ct = await zip.file('[Content_Types].xml').async('string');
  if (!/Extension="fntdata"/.test(ct)) ct = ct.replace('<Default ', '<Default Extension="fntdata" ContentType="application/x-fontdata"/><Default ');
  zip.file('[Content_Types].xml', ct);
  let rels = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
  let lst = '<p:embeddedFontLst>';
  for (const [fam, s] of Object.entries(slots)) {
    lst += `<p:embeddedFont><p:font typeface="${fam.replace(/"/g, '&quot;')}"/>`;
    // the schema is a sequence: regular, bold, italic, boldItalic
    for (const [slot, { part, rid }] of ['regular', 'bold', 'italic', 'boldItalic'].filter((k) => s[k]).map((k) => [k, s[k]])) {
      rels = rels.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/${part}"/></Relationships>`);
      lst += `<p:${slot} r:id="${rid}"/>`;
    }
    lst += '</p:embeddedFont>';
  }
  lst += '</p:embeddedFontLst>';
  zip.file('ppt/_rels/presentation.xml.rels', rels);
  let pres = await zip.file('ppt/presentation.xml').async('string');
  pres = pres.replace(/(<p:notesSz[^>]*\/>)/, `$1${lst}`).replace('<p:presentation ', '<p:presentation embedTrueTypeFonts="1" ');
  zip.file('ppt/presentation.xml', pres);
  return { report, restricted };
}
