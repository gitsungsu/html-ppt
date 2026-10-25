#!/usr/bin/env node
// Learn a slide template from 1+ sample PPTX files and write template-spec.json + assets/.
// Usage: node extract-template.mjs <a.pptx> <b.pptx> [more...] --out <templateDir> [--primary <x.pptx>]
// Exit: 0 ok, 2 input error (incl. AC7 limits), 3 tie conflict recorded (output still written).
import { createHash } from 'node:crypto';
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import {
  InputError, openPackage, readXml, readRels, findRel, kid, kids, dig, rattr,
  colorChild, colorRgb, colorHex, resolveTextStyle,
} from './lib/ooxml.mjs';

const PALETTE_KEYS = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
const FILL_EL = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);
const UNSUPPORTED_IMG = new Set(['.emf', '.wmf', '.emz', '.wmz']);

// ---------------------------------------------------------------- layout kind

function layoutKind(name) {
  const n = name.replace(/([a-z])([A-Z])/g, '$1 $2').toLowerCase();
  const tokens = n.split(/[^a-z0-9]+/).filter(Boolean);
  const en = (...w) => w.some((x) => tokens.includes(x));
  const ko = (...w) => w.some((x) => n.includes(x));
  if (en('two', 'comparison', 'compare') || ko('두 개', '두개', '2단', '2열', '비교', '양쪽', '좌우')) return 'two';
  if (en('table', 'tbl') || ko('표', '테이블')) return 'table';
  if (en('picture', 'image', 'photo', 'img', 'pic', 'figure') || ko('그림', '이미지', '사진')) return 'image';
  if (en('blank', 'empty') || ko('빈 화면', '빈화면', '공백', '백지')) return 'other';
  if (en('content', 'body', 'text', 'list') || ko('본문', '내용', '목록', '텍스트')) return 'content';
  if (en('title', 'cover', 'section', 'header') || ko('제목', '표지', '타이틀', '커버', '구역', '섹션')) return 'title';
  return 'content';
}

// ---------------------------------------------------------------- per-sample extraction

async function extractSample(file) {
  const pkg = await openPackage(file);
  const base = path.basename(file);
  const warnings = new Set();
  const media = new Map();        // asset file -> { data, source }
  const mediaByPath = new Map();  // package path -> asset file | null

  async function registerMedia(target) {
    if (mediaByPath.has(target)) return mediaByPath.get(target);
    const ext = path.posix.extname(target).toLowerCase();
    let out = null;
    if (UNSUPPORTED_IMG.has(ext)) {
      warnings.add(`${target}: ${ext.slice(1).toUpperCase()} picture is not supported and was skipped`);
    } else {
      const data = await pkg.read(target);
      if (!data) warnings.add(`${target}: referenced media is missing from the package`);
      else {
        const hash = createHash('sha1').update(data).digest('hex').slice(0, 8);
        const stem = path.posix.basename(target, ext).replace(/[^A-Za-z0-9_-]/g, '_');
        out = `${stem}-${hash}${ext}`;
        if (!media.has(out)) media.set(out, { data, source: `${base}!${target}` });
      }
    }
    mediaByPath.set(target, out);
    return out;
  }

  // ---- package walk: _rels/.rels -> presentation -> master -> theme
  const rootRels = await readRels(pkg, '');
  const presPath = findRel(rootRels, '/officeDocument');
  if (!presPath) throw new InputError(`${file}: no officeDocument relationship (not a PPTX?)`);
  const pres = await readXml(pkg, presPath);
  const sz = kid(pres, 'sldSz');
  if (!sz) throw new InputError(`${file}: presentation.xml has no sldSz`);
  const slideSize = { cx: +sz.attrs.cx, cy: +sz.attrs.cy };
  const presRels = await readRels(pkg, presPath);
  const masterIds = kids(kid(pres, 'sldMasterIdLst'), 'sldMasterId');
  if (!masterIds.length) throw new InputError(`${file}: no slide master`);
  if (masterIds.length > 1) warnings.add('more than one slide master: only the first master and its layouts are used');
  const masterRel = presRels.get(rattr(masterIds[0], 'id'));
  if (!masterRel) throw new InputError(`${file}: slide master relationship not found`);
  const masterPath = masterRel.target;
  const master = await readXml(pkg, masterPath);
  const masterRels = await readRels(pkg, masterPath);
  const themePath = findRel(masterRels, '/theme');
  if (!themePath) throw new InputError(`${file}: slide master has no theme`);
  const theme = await readXml(pkg, themePath);
  const themeRels = await readRels(pkg, themePath);
  const themeEls = kid(theme, 'themeElements');

  // palette
  const clrScheme = kid(themeEls, 'clrScheme');
  const themeHex = {};
  const palette = {};
  for (const k of PALETTE_KEYS) {
    const c = colorChild(kid(clrScheme, k));
    if (!c) { warnings.add(`theme colour ${k} is missing, using #000000`); themeHex[k] = '000000'; }
    else themeHex[k] = colorHex(c, { theme: {} }).slice(1);
    palette[k] = '#' + themeHex[k];
  }

  // fonts
  const fontScheme = kid(themeEls, 'fontScheme');
  const readFont = (el) => {
    const latin = kid(el, 'latin')?.attrs.typeface || '';
    let ea = kid(el, 'ea')?.attrs.typeface || '';
    if (!ea) ea = kids(el, 'font').find((f) => f.attrs.script === 'Hang')?.attrs.typeface || '';
    return { latin, ea };
  };
  const fonts = { major: readFont(kid(fontScheme, 'majorFont')), minor: readFont(kid(fontScheme, 'minorFont')) };

  // theme fill styles (for bgRef / fillRef)
  const fmt = kid(themeEls, 'fmtScheme');
  const fillStyles = kid(fmt, 'fillStyleLst')?.children ?? [];
  const bgFillStyles = kid(fmt, 'bgFillStyleLst')?.children ?? [];

  // ---- fills
  async function fillFromNode(f, ctx, rels) {
    switch (f.name) {
      case 'noFill': return { kind: 'none' };
      case 'solidFill': {
        const c = colorChild(f);
        return c ? { kind: 'color', hex: colorHex(c, ctx) } : null;
      }
      case 'gradFill': {
        const c = colorChild(kids(kid(f, 'gsLst'), 'gs')[0]);
        if (!c) return null;
        warnings.add('gradient fill approximated by its first stop colour');
        return { kind: 'color', hex: colorHex(c, ctx) };
      }
      case 'pattFill': {
        const c = colorChild(kid(f, 'fgClr'));
        return c ? { kind: 'color', hex: colorHex(c, ctx) } : null;
      }
      case 'blipFill': {
        const rel = rels.get(rattr(kid(f, 'blip'), 'embed'));
        const file2 = rel ? await registerMedia(rel.target) : null;
        return file2 ? { kind: 'image', file: file2 } : { kind: 'none' };
      }
      default: return null;
    }
  }
  const fillFromParent = async (parent, ctx, rels) => {
    const f = parent && parent.children.find((c) => FILL_EL.has(c.name));
    return f ? fillFromNode(f, ctx, rels) : null;
  };

  async function resolveBg(bg, ctx, rels) {
    let f = null;
    const pr = kid(bg, 'bgPr');
    if (pr) f = await fillFromParent(pr, ctx, rels);
    else {
      const ref = kid(bg, 'bgRef');
      if (ref) {
        const idx = +ref.attrs.idx || 0;
        const c = colorChild(ref);
        const phClr = c ? colorRgb(c, ctx) : null;
        const style = idx >= 1001 ? bgFillStyles[idx - 1001] : idx >= 1 && idx <= 999 ? fillStyles[idx - 1] : null;
        f = style ? await fillFromNode(style, { ...ctx, phClr }, themeRels) : { kind: 'none' };
      }
    }
    if (!f || f.kind === 'none') return { type: 'none', value: null };
    if (f.kind === 'color') return { type: 'color', value: f.hex };
    return { type: 'image', value: f.file };
  }

  // ---- shapes
  const phOf = (el) => dig(el, 'nvSpPr', 'nvPr', 'ph') || dig(el, 'nvPicPr', 'nvPr', 'ph');
  const hiddenOf = (el) => (dig(el, 'nvSpPr', 'cNvPr') || dig(el, 'nvPicPr', 'cNvPr') || dig(el, 'nvCxnSpPr', 'cNvPr'))?.attrs.hidden === '1';
  const boxOf = (el) => {
    const x = dig(el, 'spPr', 'xfrm');
    const off = kid(x, 'off'), ext = kid(x, 'ext');
    if (!off || !ext) return null;
    return { x: +off.attrs.x || 0, y: +off.attrs.y || 0, w: +ext.attrs.cx || 0, h: +ext.attrs.cy || 0 };
  };
  const spTreeOf = (part) => dig(part, 'cSld', 'spTree');

  async function shapeFill(sp, ctx, rels) {
    let f = await fillFromParent(kid(sp, 'spPr'), ctx, rels);
    if (f === null) {
      const fr = dig(sp, 'style', 'fillRef');
      const idx = +fr?.attrs.idx || 0;
      if (idx >= 1 && idx <= 999 && fillStyles[idx - 1]) {
        const c = colorChild(fr);
        f = await fillFromNode(fillStyles[idx - 1], { ...ctx, phClr: c ? colorRgb(c, ctx) : null }, themeRels);
      }
    }
    if (f && f.kind === 'color') return f.hex;
    if (f && f.kind === 'image') warnings.add('picture-filled shape: the fill was skipped');
    return null;
  }

  const paraText = (p) => p.children.map((r) => (r.name === 'br' ? '\n' : (r.name === 'r' || r.name === 'fld') ? (kid(r, 't')?.text ?? '') : '')).join('');
  const textOf = (sp) => kids(kid(sp, 'txBody'), 'p').map(paraText).join('\n').trim();
  function textStyle(sp, ctx) {
    for (const p of kids(kid(sp, 'txBody'), 'p')) {
      const r = p.children.find((c) => (c.name === 'r' || c.name === 'fld') && kid(c, 't')?.text);
      if (!r) continue;
      const rPr = kid(r, 'rPr');
      const c = colorChild(kid(rPr, 'solidFill')) || colorChild(dig(sp, 'style', 'fontRef'));
      return { textColor: c ? colorHex(c, ctx) : null, fontSize: rPr?.attrs.sz ? +rPr.attrs.sz / 100 : null };
    }
    return { textColor: null, fontSize: null };
  }

  async function collectDecorations(tree, ctx, rels) {
    const out = [];
    for (const el of tree?.children ?? []) {
      if (hiddenOf(el)) continue;
      if (el.name === 'grpSp') { warnings.add('group shapes are not extracted'); continue; }
      if (el.name === 'graphicFrame') { warnings.add('tables/charts/SmartArt in the master or a layout are not extracted'); continue; }
      if (el.name === 'pic') {
        if (phOf(el)) continue;
        const box = boxOf(el);
        const rel = rels.get(rattr(kid(kid(el, 'blipFill'), 'blip'), 'embed'));
        if (!box || !rel) { warnings.add('a picture without position or image was skipped'); continue; }
        const asset = await registerMedia(rel.target);
        if (asset) out.push({ kind: 'pic', ...box, fill: null, text: null, asset });
      } else if (el.name === 'sp' || el.name === 'cxnSp') {
        if (phOf(el)) continue;
        const box = boxOf(el);
        if (!box) continue;
        const geom = kid(kid(el, 'spPr'), 'prstGeom')?.attrs.prst;
        let fill = await shapeFill(el, ctx, rels);
        if (el.name === 'cxnSp' && !fill) {
          const c = colorChild(dig(el, 'spPr', 'ln', 'solidFill'));
          fill = c ? colorHex(c, ctx) : null;
        }
        const text = el.name === 'sp' ? textOf(el) : '';
        const d = { kind: text ? 'text' : 'shape', ...box, fill, text: text || null, asset: null };
        if (geom && geom !== 'rect') d.geom = geom;
        if (text) Object.assign(d, textStyle(el, ctx));
        if (!text && !fill) continue; // invisible: no fill and no text
        out.push(d);
      }
    }
    return out;
  }

  function collectPlaceholders(tree) {
    const out = [];
    for (const el of tree?.children ?? []) {
      const ph = (el.name === 'sp' || el.name === 'pic') && phOf(el);
      if (!ph) continue;
      out.push({ type: ph.attrs.type || 'obj', idx: +ph.attrs.idx || 0, box: boxOf(el), el });
    }
    return out;
  }

  // text style sources, most specific first: layout lstStyle, the placeholder's own first paragraph/run, master txStyles
  const lvlSource = (lvl) => {
    const d = kid(lvl, 'defRPr');
    return { algn: lvl?.attrs.algn, b: d?.attrs.b, sz: d?.attrs.sz, fill: kid(d, 'solidFill') };
  };
  const ownSource = (el) => {
    const p = kids(kid(el, 'txBody'), 'p')[0];
    const pPr = kid(p, 'pPr');
    const rPr = kid(p?.children.find((c) => c.name === 'r' || c.name === 'fld'), 'rPr');
    return {
      algn: pPr?.attrs.algn,
      b: rPr?.attrs.b ?? kid(pPr, 'defRPr')?.attrs.b,
      sz: kid(pPr, 'defRPr')?.attrs.sz ?? rPr?.attrs.sz,
      fill: kid(rPr, 'solidFill') || kid(kid(pPr, 'defRPr'), 'solidFill'),
    };
  };
  const masterTx = kid(master, 'txStyles');
  const placeholderTextStyle = (ph, ctx) => {
    const tx = kid(masterTx, ph.type === 'title' || ph.type === 'ctrTitle' ? 'titleStyle' : 'bodyStyle');
    return resolveTextStyle([
      lvlSource(dig(ph.el, 'txBody', 'lstStyle', 'lvl1pPr')),
      ownSource(ph.el),
      lvlSource(kid(tx, 'lvl1pPr')),
    ], ctx);
  };

  const masterTree = spTreeOf(master);
  const masterPhs = collectPlaceholders(masterTree);
  const masterClrMap = kid(master, 'clrMap')?.attrs ?? {};
  const masterBg = dig(master, 'cSld', 'bg');

  function inheritBox(ph) {
    const t = ph.type === 'ctrTitle' ? 'title' : ph.type;
    const want = (t === 'title' || t === 'dt' || t === 'ftr' || t === 'sldNum') ? t : 'body';
    const cands = masterPhs.filter((m) => m.type === want && m.box);
    return (cands.find((m) => m.idx === ph.idx) || cands[0])?.box ?? null;
  }

  // ---- layouts
  const layouts = [];
  const layoutByPath = new Map(); // layout part path -> layout entry (for slide learning)
  const layoutRels = kids(kid(master, 'sldLayoutIdLst'), 'sldLayoutId');
  for (const lid of layoutRels) {
    const rel = masterRels.get(rattr(lid, 'id'));
    if (!rel) continue;
    const lx = await readXml(pkg, rel.target);
    const lrels = await readRels(pkg, rel.target);
    const name = dig(lx, 'cSld')?.attrs.name || path.posix.basename(rel.target, '.xml');
    const showMasterSp = !['0', 'false'].includes(lx.attrs.showMasterSp);
    const ovr = dig(lx, 'clrMapOvr', 'overrideClrMapping');
    const ctx = { theme: themeHex, clrMap: ovr ? ovr.attrs : masterClrMap, phClr: null };

    const layoutBg = dig(lx, 'cSld', 'bg');
    const background = layoutBg
      ? await resolveBg(layoutBg, ctx, lrels)
      : masterBg ? await resolveBg(masterBg, ctx, masterRels) : { type: 'none', value: null };

    const placeholders = [];
    for (const ph of collectPlaceholders(spTreeOf(lx))) {
      let box = ph.box;
      if (!box) {
        box = inheritBox(ph);
        if (!box) { warnings.add(`layout "${name}": placeholder ${ph.type}:${ph.idx} has no position in layout or master`); box = { x: 0, y: 0, w: 0, h: 0 }; }
      }
      placeholders.push({ type: ph.type, idx: ph.idx, ...box, textStyle: placeholderTextStyle(ph, ctx) });
    }

    const decorations = [
      ...(showMasterSp ? await collectDecorations(masterTree, ctx, masterRels) : []),
      ...await collectDecorations(spTreeOf(lx), ctx, lrels),
    ];
    const signature = placeholders.map((p) => `${p.type}:${p.idx}`).sort().join(',');
    layouts.push({ name, kind: layoutKind(name), showMasterSp, background, placeholders, decorations, signature });
    layoutByPath.set(rel.target, layouts[layouts.length - 1]);
  }

  // ---- slides (AC8): text runs and pictures placed directly on the slides
  const runs = [];           // { title, latin, ea, color, n }
  const slideCounts = {};    // layout name -> number of slides
  const notes = new Set();
  const picsByLayout = new Map(); // layout entry -> [Map(media path -> first box on that slide)]
  const sctx = { theme: themeHex, clrMap: masterClrMap, phClr: null };
  const face = (n) => { const t = n?.attrs.typeface; return t && !t.startsWith('+') ? t : null; };
  let groupPics = false;
  const styleRecs = [];      // AC9 a: { layout, role, n, x?, y?, w?, h?, align?, size?, bold?, colour?, latin? }
  let tableHeader = null;    // AC9 b: { fills: [...], texts: [...] } from the first table
  for (const sid of kids(kid(pres, 'sldIdLst'), 'sldId')) {
    const srel = presRels.get(rattr(sid, 'id'));
    if (!srel) continue;
    const sx = await readXml(pkg, srel.target);
    const srels = await readRels(pkg, srel.target);
    const layout = layoutByPath.get(findRel(srels, '/slideLayout'));
    if (!layout) { notes.add(`${path.posix.basename(srel.target)}: layout is not under the first slide master, slide skipped`); continue; }
    slideCounts[layout.name] = (slideCounts[layout.name] ?? 0) + 1;

    const walk = (n, isTitle) => {
      for (const c of n.children) {
        let t = isTitle;
        if (c.name === 'sp') { const ph = phOf(c); t = !!ph && (ph.attrs.type === 'title' || ph.attrs.type === 'ctrTitle'); }
        if (c.name === 'pic' && n.name === 'grpSp') groupPics = true;
        if (c.name === 'r') {
          const len = [...(kid(c, 't')?.text ?? '')].length;
          if (!len) continue;
          const rPr = kid(c, 'rPr');
          const cc = colorChild(kid(rPr, 'solidFill'));
          runs.push({ title: t, latin: face(kid(rPr, 'latin')), ea: face(kid(rPr, 'ea')), color: cc ? colorHex(cc, sctx) : null, n: len });
        } else walk(c, t);
      }
    };
    const tree = spTreeOf(sx);
    if (tree) walk(tree, false);

    // AC9 a: slide-level title/body placeholder style (top-level shapes only; group children are skipped)
    const styleOf = (el, role) => {
      const rec = { layout: layout.name, role, n: Math.max(1, [...textOf(el)].length) };
      const box = boxOf(el);
      if (box) Object.assign(rec, box);
      const ps = kids(kid(el, 'txBody'), 'p');
      const algn = kid(ps[0], 'pPr')?.attrs.algn;
      if (algn === 'l' || algn === 'ctr' || algn === 'r') rec.align = algn;
      const r = ps.flatMap((p) => p.children).find((c) => (c.name === 'r' || c.name === 'fld') && kid(c, 't')?.text);
      const rPr = kid(r, 'rPr');
      if (+rPr?.attrs.sz > 0) rec.size = +rPr.attrs.sz / 100;
      if (rPr?.attrs.b !== undefined) rec.bold = rPr.attrs.b === '1' || rPr.attrs.b === 'true';
      const cc = colorChild(kid(rPr, 'solidFill'));
      if (cc) rec.colour = colorHex(cc, sctx);
      const lf = face(kid(rPr, 'latin'));
      if (lf) rec.latin = lf;
      return rec;
    };
    let titleBox = null;
    for (const el of tree?.children ?? []) {
      const ph = el.name === 'sp' && phOf(el);
      if (!ph) continue;
      const t = ph.attrs.type || 'obj';
      const role = t === 'title' || t === 'ctrTitle' ? 'title' : ['body', 'obj', 'subTitle'].includes(t) ? 'body' : null;
      if (!role) continue;
      const rec = styleOf(el, role);
      if (role === 'title') titleBox ??= Number.isFinite(rec.x) ? rec : null;
      styleRecs.push(rec);
    }

    // AC10: top-level non-placeholder text boxes. The nearest one directly below the title whose horizontal
    // range overlaps the title's is a subtitle candidate; the largest other one below the title is the text box.
    titleBox ??= layout.placeholders.find((p) => (p.type === 'title' || p.type === 'ctrTitle') && p.w > 0) ?? null;
    const boxes = (tree?.children ?? [])
      .filter((el) => el.name === 'sp' && !phOf(el) && !hiddenOf(el) && textOf(el))
      .map((el) => styleOf(el, 'textbox'))
      .filter((r) => Number.isFinite(r.x));
    const below = titleBox ? boxes.filter((r) => r.y >= titleBox.y + titleBox.h) : boxes;
    const sub = titleBox
      ? below.filter((r) => r.x < titleBox.x + titleBox.w && r.x + r.w > titleBox.x).sort((a, b) => a.y - b.y)[0]
      : null;
    if (sub) styleRecs.push({ ...sub, role: 'subtitle' });
    const tb = below.filter((r) => r !== sub).sort((a, b) => b.n - a.n)[0];
    if (tb) styleRecs.push(tb);

    // AC9 b: first table's header row (cell fills and first-run text colours)
    if (!tableHeader && tree) {
      const findTbl = (n) => (n.name === 'tbl' ? n : n.children.reduce((a, c) => a ?? findTbl(c), null));
      const tbl = findTbl(tree);
      if (tbl) {
        tableHeader = { fills: [], texts: [] };
        for (const tc of kids(kids(tbl, 'tr')[0], 'tc')) {
          const fc = colorChild(dig(tc, 'tcPr', 'solidFill'));
          if (fc) tableHeader.fills.push(colorHex(fc, sctx));
          const run = kids(kid(tc, 'txBody'), 'p').flatMap((p) => p.children).find((c) => c.name === 'r' && kid(c, 't')?.text);
          const tcc = colorChild(dig(run, 'rPr', 'solidFill'));
          if (tcc) tableHeader.texts.push(colorHex(tcc, sctx));
        }
      }
    }

    const pics = new Map();
    for (const el of tree?.children ?? []) {
      if (el.name !== 'pic' || phOf(el) || hiddenOf(el)) continue;
      const box = boxOf(el);
      const rel = srels.get(rattr(dig(el, 'blipFill', 'blip'), 'embed'));
      if (box && rel && !pics.has(rel.target)) pics.set(rel.target, box);
    }
    if (!picsByLayout.has(layout)) picsByLayout.set(layout, []);
    picsByLayout.get(layout).push(pics);
  }
  if (groupPics) notes.add('pictures inside group shapes are not counted as repeated pictures');

  // A layout picture that the slides of that layout place at the same position and size is a copy of slide content,
  // not layout art (the original Blank layout carries the contents-slide picture, which would put "01 02 03" into the
  // master). The bytes can differ (the slide and the layout store separate media parts), so the box decides. The
  // repeated-picture rule below adds the picture back if it really repeats on the layout's slides.
  for (const [layout, list] of picsByLayout) {
    const boxes = list.flatMap((pics) => [...pics.values()]);
    layout.decorations = layout.decorations.filter((d) => {
      if (d.kind !== 'pic') return true;
      const copy = boxes.some((b) => b.x === d.x && b.y === d.y && b.w === d.w && b.h === d.h);
      if (copy) notes.add(`layout "${layout.name}": picture ${d.asset} is slide content at the same place, not layout art; dropped`);
      return !copy;
    });
  }

  // A picture that a later opaque (JPEG) full-slide picture of the same layout covers is never visible, so it is not kept.
  // The original Title Only layout hides the contents-slide picture under the cover picture this way. The tolerance
  // absorbs the small offsets the original file uses for full-slide pictures.
  const tol = 0.01 * slideSize.cx;
  for (const layout of layouts) {
    layout.decorations = layout.decorations.filter((d, i) => {
      if (d.kind !== 'pic') return true;
      const hidden = layout.decorations.slice(i + 1).some((o) => o.kind === 'pic' && /\.jpe?g$/i.test(o.asset)
        && o.x <= d.x + tol && o.y <= d.y + tol && o.x + o.w >= d.x + d.w - tol && o.y + o.h >= d.y + d.h - tol);
      if (hidden) notes.add(`layout "${layout.name}": picture ${d.asset} is covered by a later opaque picture; dropped`);
      return !hidden;
    });
  }

  // Repeated pictures (>= 60% of a layout's slides) become decorations; a picture covering >= 90% of the slide becomes the background.
  const slideArea = slideSize.cx * slideSize.cy;
  for (const [layout, list] of picsByLayout) {
    const seen = new Map(); // media path -> { n, box (first occurrence) }
    for (const pics of list) for (const [t, box] of pics) { const s = seen.get(t) ?? { n: 0, box }; s.n++; seen.set(t, s); }
    for (const [t, { n, box }] of seen) {
      if (n < 0.6 * list.length) continue;
      const asset = await registerMedia(t);
      if (!asset || layout.decorations.some((d) => d.asset === asset)) continue;
      if (box.w * box.h >= 0.9 * slideArea) {
        if (layout.background.type === 'image' && layout.background.value === asset) continue;
        layout.background = { type: 'image', value: asset };
        notes.add(`layout "${layout.name}": background picture ${asset} from ${n} of ${list.length} slide(s)`);
      } else {
        layout.decorations.push({ kind: 'pic', ...box, fill: null, text: null, asset });
        notes.add(`layout "${layout.name}": repeated picture ${asset} from ${n} of ${list.length} slide(s)`);
      }
    }
  }

  return { file, base, slideSize, palette, fonts, layouts, media, runs, slideCounts, styleRecs, tableHeader, notes: [...notes], warnings: [...warnings] };
}

// ---------------------------------------------------------------- merge (majority vote, D1 ties)

// entries: values in ordered-sample order (undefined = this sample has no value). Unique max wins;
// a tie goes to conflicts[] and the first ordered sample (primary, else filename order) among the tied values wins.
function vote(vpath, entries, conflicts) {
  const present = entries.filter((v) => v !== undefined);
  if (!present.length) return undefined;
  const groups = new Map();
  for (const v of present) {
    const k = JSON.stringify(v);
    if (!groups.has(k)) groups.set(k, { value: v, n: 0 });
    groups.get(k).n++;
  }
  const max = Math.max(...[...groups.values()].map((g) => g.n));
  const tops = [...groups.entries()].filter(([, g]) => g.n === max).map(([k]) => k);
  if (tops.length === 1) return groups.get(tops[0]).value;
  const chosen = present.find((v) => tops.includes(JSON.stringify(v)));
  conflicts.push({
    path: vpath,
    values: [...groups.values()].map((g) => g.value),
    chosen,
    reason: `tie: ${tops.length} values with ${max} vote(s) each out of ${present.length} sample(s); chose the primary/first sample's value`,
  });
  return chosen;
}

// lists[i] = items of member i. Items are matched by keyOf (+ ordinal); presence is a majority vote too.
function mergeList(label, lists, keyOf, build, conflicts) {
  const n = lists.length;
  const keyed = new Map();
  lists.forEach((list, mi) => {
    const seen = new Map();
    for (const item of list) {
      const b = keyOf(item);
      const c = (seen.get(b) ?? 0) + 1;
      seen.set(b, c);
      const key = c > 1 ? `${b}#${c}` : b;
      if (!keyed.has(key)) keyed.set(key, new Array(n).fill(undefined));
      keyed.get(key)[mi] = item;
    }
  });
  const out = [];
  for (const [key, items] of keyed) {
    const c = items.filter(Boolean).length;
    if (2 * c < n) continue;
    if (2 * c === n) {
      const chosenPresent = items[0] !== undefined;
      conflicts.push({
        path: `${label}[${key}]`,
        values: ['present', 'absent'],
        chosen: chosenPresent ? 'present' : 'absent',
        reason: `tie: present in ${c} of ${n} sample(s)`,
      });
      if (!chosenPresent) continue;
    }
    out.push(build(`${label}[${key}]`, items));
  }
  return out;
}

function mergeLayouts(samples, conflicts) {
  const N = samples.length;
  const groups = [];
  const addTo = (g, si, layout) => { g.sis.add(si); g.members.push({ si, layout }); };
  const newGroup = (si, layout) => { const g = { sis: new Set(), members: [] }; addTo(g, si, layout); groups.push(g); };

  samples.forEach((s, si) => {
    if (si === 0) { s.layouts.forEach((l) => newGroup(si, l)); return; }
    const rest = [];
    for (const l of s.layouts) {
      const g = groups.find((x) => !x.sis.has(si) && x.members.some((m) => m.layout.name === l.name));
      if (g) addTo(g, si, l); else rest.push(l);
    }
    const rest2 = [];
    for (const l of rest) {
      const g = groups.find((x) => !x.sis.has(si) && x.members.some((m) => m.layout.signature === l.signature));
      if (g) addTo(g, si, l); else rest2.push(l);
    }
    rest2.forEach((l) => newGroup(si, l));
  });

  const out = [];
  for (const g of groups) {
    const ls = g.members.map((m) => m.layout);
    const label = `layouts[${ls[0].name}]`;
    const cnt = g.members.length;
    if (2 * cnt < N) continue;
    if (2 * cnt === N) {
      const chosenPresent = g.sis.has(0);
      conflicts.push({
        path: label,
        values: ['present', 'absent'],
        chosen: chosenPresent ? 'present' : 'absent',
        reason: `tie: layout found in ${cnt} of ${N} sample(s)`,
      });
      if (!chosenPresent) continue;
    }
    const name = vote(`${label}.name`, ls.map((l) => l.name), conflicts);
    const placeholders = mergeList(
      `${label}.placeholders`,
      ls.map((l) => l.placeholders),
      (p) => `${p.type}:${p.idx}`,
      (p, items) => {
        const its = items.filter(Boolean);
        const box = vote(`${p}.box`, its.map(({ x, y, w, h }) => ({ x, y, w, h })), conflicts);
        const textStyle = vote(`${p}.textStyle`, its.map((i) => i.textStyle), conflicts);
        return { type: its[0].type, idx: its[0].idx, ...box, textStyle };
      },
      conflicts,
    );
    const decorations = mergeList(
      `${label}.decorations`,
      ls.map((l) => l.decorations),
      (d) => `${d.kind}|${d.asset ?? d.text ?? d.geom ?? ''}`,
      (p, items) => {
        const its = items.filter(Boolean);
        const box = vote(`${p}.box`, its.map(({ x, y, w, h }) => ({ x, y, w, h })), conflicts);
        const d = {
          kind: its[0].kind, ...box,
          fill: vote(`${p}.fill`, its.map((i) => i.fill), conflicts) ?? null,
          text: its[0].text ?? null,
          asset: its[0].asset ?? null,
        };
        if (its[0].geom) d.geom = its[0].geom;
        if (its[0].kind === 'text') {
          d.textColor = vote(`${p}.textColor`, its.map((i) => i.textColor), conflicts) ?? null;
          d.fontSize = vote(`${p}.fontSize`, its.map((i) => i.fontSize), conflicts) ?? null;
        }
        return d;
      },
      conflicts,
    );
    out.push({
      name,
      kind: vote(`${label}.kind`, ls.map((l) => l.kind), conflicts),
      showMasterSp: vote(`${label}.showMasterSp`, ls.map((l) => l.showMasterSp), conflicts),
      background: vote(`${label}.background`, ls.map((l) => l.background), conflicts),
      placeholders,
      decorations,
    });
  }
  return out;
}

// ---------------------------------------------------------------- observed slide style (AC8 a, b)

// Character-weighted mode; a tie keeps the value seen first (sample order, then slide order).
function weightedMode(runs, key) {
  const m = new Map();
  for (const r of runs) { const v = key(r); if (v) m.set(v, (m.get(v) ?? 0) + r.n); }
  let best = null, bn = 0;
  for (const [v, n] of m) if (n > bn) { best = v; bn = n; }
  return best;
}

// Fonts vote by family: weight words at the end of the name are dropped ("나눔고딕 ExtraBold" -> "나눔고딕").
// The winner is the family name if runs use it as is, else its most used member name.
const WEIGHT = /(\s+(thin|hairline|extralight|ultralight|light|regular|normal|book|medium|semibold|demibold|bold|extrabold|ultrabold|heavy|black))+$/i;
function fontMode(runs, key, label, notes) {
  const fam = weightedMode(runs, (r) => key(r) && key(r).replace(WEIGHT, ''));
  if (!fam) return null;
  const members = runs.filter((r) => key(r) && key(r).replace(WEIGHT, '') === fam);
  const total = members.reduce((a, r) => a + r.n, 0);
  const name = members.some((r) => key(r) === fam) ? fam : weightedMode(members, key);
  notes.push(`${label}: family "${fam}" (${total} characters) -> "${name}"`);
  return name;
}

function observeSlides(samples) {
  const runs = samples.flatMap((s) => s.runs);
  const title = runs.filter((r) => r.title);
  const body = runs.filter((r) => !r.title);
  const chroma = (h) => { const c = [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16)); return Math.max(...c) - Math.min(...c); };
  const slides = {};
  for (const s of samples) for (const [k, v] of Object.entries(s.slideCounts)) slides[k] = (slides[k] ?? 0) + v;
  const notes = [...new Set(samples.flatMap((s) => s.notes))];
  if (!runs.length) notes.unshift('no text runs found on the slides');
  const fonts = {
    title: { latin: fontMode(title, (r) => r.latin, 'title latin', notes), ea: fontMode(title, (r) => r.ea, 'title ea', notes) },
    body: { latin: fontMode(body, (r) => r.latin, 'body latin', notes), ea: fontMode(body, (r) => r.ea, 'body ea', notes) },
  };
  // AC9 a / AC10: per-layout title/body/subtitle/textbox style: median position and size, character-weighted mode of the rest.
  const median = (a) => { const s = [...a].sort((x, y) => x - y); const m = s.length >> 1; return s.length % 2 ? s[m] : Math.round((s[m - 1] + s[m]) / 2); };
  const layoutStyles = {};
  for (const r of samples.flatMap((s) => s.styleRecs)) {
    layoutStyles[r.layout] ??= {};
    (layoutStyles[r.layout][r.role] ??= []).push(r);
  }
  for (const roles of Object.values(layoutStyles)) {
    for (const [role, recs] of Object.entries(roles)) {
      const boxed = recs.filter((r) => Number.isFinite(r.x));
      const pos = (f) => (boxed.length ? median(boxed.map((r) => r[f])) : null);
      const bold = weightedMode(recs, (r) => (r.bold === undefined ? null : String(r.bold)));
      roles[role] = {
        x: pos('x'), y: pos('y'), w: pos('w'), h: pos('h'),
        align: weightedMode(recs, (r) => r.align),
        size: weightedMode(recs, (r) => r.size),
        bold: bold === null ? null : bold === 'true',
        colour: weightedMode(recs, (r) => r.colour),
        latin: weightedMode(recs, (r) => r.latin),
      };
    }
  }
  // AC9 b: table header colours from the first sample that has a table.
  const th = samples.find((s) => s.tableHeader)?.tableHeader;
  const mode = (list) => weightedMode((list ?? []).map((v) => ({ v, n: 1 })), (r) => r.v);
  return {
    fonts,
    colours: {
      title: weightedMode(title, (r) => r.color),
      body: weightedMode(body, (r) => r.color),
      accent: weightedMode(runs, (r) => (r.color && chroma(r.color) > 40 ? r.color : null)),
      tableHeader: mode(th?.fills),
      tableHeaderText: mode(th?.texts),
    },
    layoutStyles,
    slides,
    notes,
  };
}

// ---------------------------------------------------------------- CLI

function parseArgs(argv) {
  const samples = [];
  let out, primary;
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--out') out = argv[++i];
    else if (a === '--primary') primary = argv[++i];
    else if (a.startsWith('--')) throw new InputError(`unknown option ${a}`);
    else samples.push(a);
  }
  if (!out) throw new InputError('--out <templateDir> is required');
  if (samples.length < 1) throw new InputError('at least 1 sample PPTX file is required');
  if (primary !== undefined) {
    if (!samples.some((s) => path.resolve(s) === path.resolve(primary))) {
      throw new InputError('--primary must be one of the sample files');
    }
  }
  return { samples, out, primary };
}

async function main() {
  const { samples: paths, out, primary } = parseArgs(process.argv.slice(2));

  // Order: --primary first, then filename order (D1).
  const byName = [...paths].sort((a, b) => path.basename(a).localeCompare(path.basename(b)) || a.localeCompare(b));
  const ordered = primary === undefined
    ? byName
    : [paths.find((p) => path.resolve(p) === path.resolve(primary)), ...byName.filter((p) => path.resolve(p) !== path.resolve(primary))];

  const samples = [];
  for (const p of ordered) samples.push(await extractSample(p));

  const conflicts = [];
  const spec = {
    slideSize: vote('slideSize', samples.map((s) => s.slideSize), conflicts),
    palette: {},
    fonts: { major: {}, minor: {} },
    observed: null,
    layouts: [],
    assets: [],
    conflicts,
    warnings: [],
  };
  for (const k of PALETTE_KEYS) spec.palette[k] = vote(`palette.${k}`, samples.map((s) => s.palette[k]), conflicts);
  for (const role of ['major', 'minor']) {
    for (const k of ['latin', 'ea']) spec.fonts[role][k] = vote(`fonts.${role}.${k}`, samples.map((s) => s.fonts[role][k]), conflicts);
  }
  spec.observed = observeSlides(samples);
  spec.layouts = mergeLayouts(samples, conflicts);

  // assets actually referenced by the merged spec
  const used = new Set();
  for (const l of spec.layouts) {
    if (l.background.type === 'image') used.add(l.background.value);
    for (const d of l.decorations) if (d.asset) used.add(d.asset);
  }
  const assetsDir = path.join(out, 'assets');
  rmSync(assetsDir, { recursive: true, force: true });
  mkdirSync(assetsDir, { recursive: true });
  for (const file of used) {
    const s = samples.find((x) => x.media.has(file));
    if (!s) continue;
    const m = s.media.get(file);
    writeFileSync(path.join(assetsDir, file), m.data);
    spec.assets.push({ file, source: m.source });
  }
  spec.warnings = [...new Set(samples.flatMap((s) => s.warnings))];
  if (samples.length === 1) spec.warnings.unshift('single sample: values taken as-is, no majority vote');

  writeFileSync(path.join(out, 'template-spec.json'), JSON.stringify(spec, null, 2) + '\n');
  console.log(`wrote ${path.join(out, 'template-spec.json')} (${spec.layouts.length} layouts, ${spec.assets.length} assets, ${conflicts.length} conflicts, ${spec.warnings.length} warnings)`);
  if (conflicts.length) {
    for (const c of conflicts) console.error(`conflict: ${c.path} -> ${JSON.stringify(c.chosen)} (${c.reason})`);
    process.exitCode = 3;
  }
}

main().catch((e) => {
  if (e instanceof InputError) { console.error(`error: ${e.message}`); process.exitCode = 2; }
  else { console.error(e); process.exitCode = 1; }
});
