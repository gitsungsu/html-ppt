// Build an HTML deck from a template spec, a Markdown/text content file, an image folder and CSV/XLSX data files.
// usage: node build-deck.mjs <templateDir>/template-spec.json <content.md|content.txt> --assets <imageDir> [--data <csv-or-xlsx>]... --out <deckDir> [--no-chrome]
// exit codes: 0 ok, 2 input error (missing spec/content/image/data file, bad arguments)
import { readFile, writeFile, mkdir, copyFile, stat, readdir } from 'node:fs/promises';
import { dirname, join, basename, extname, resolve, relative, isAbsolute } from 'node:path';
import { loadTable, InputError } from './lib/tables.mjs';

const USAGE = 'usage: node build-deck.mjs <template-spec.json> <content.md|content.txt> --assets <imageDir> [--data <csv-or-xlsx>]... --out <deckDir> [--no-chrome]';
const CANVAS_W = 1280;
const CANVAS_H = 720;
const FALLBACK = { dk: '#000000', lt: '#FFFFFF' };

// ---------- arguments ----------

function parseArgs(argv) {
  const pos = [];
  const opts = { data: [] };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--assets' || a === '--out' || a === '--data') {
      const v = argv[++i];
      if (v === undefined) throw new InputError(`${a} needs a value\n${USAGE}`);
      if (a === '--data') opts.data.push(v); else opts[a.slice(2)] = v;
    } else if (a === '--no-chrome') opts.noChrome = true;
    else if (a.startsWith('--')) throw new InputError(`unknown option ${a}\n${USAGE}`);
    else pos.push(a);
  }
  if (pos.length !== 2 || !opts.out) throw new InputError(USAGE);
  return { specPath: pos[0], contentPath: pos[1], ...opts };
}

// ---------- small helpers ----------

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
// Escaped text with inline emphasis: **bold** -> <b>, ~~light~~ -> weight 400. One pass, so markers do not nest.
const fmt = (s) => esc(s).replace(/\*\*([^*\n]+?)\*\*|~~([^~\n]+?)~~/g, (_, b, l) => (b !== undefined ? `<b>${b}</b>` : `<span style="font-weight:400">${l}</span>`));
const px = (n) => `${Math.round(n * 100) / 100}px`;
const exists = async (p) => { try { await stat(p); return true; } catch { return false; } };

const hex = (v) => {
  if (typeof v !== 'string') return null;
  const m = /^#?([0-9a-fA-F]{6})$/.exec(v.trim());
  return m ? `#${m[1].toUpperCase()}` : null;
};

const lum = (h) => {
  const c = [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16) / 255).map((v) => (v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4));
  return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2];
};
const contrast = (a, b) => {
  const [hi, lo] = [lum(a), lum(b)].sort((x, y) => y - x);
  return (hi + 0.05) / (lo + 0.05);
};
// The darker or lighter palette text colour, whichever reads better on a background colour.
const textOn = (bg, dk, lt) => (bg ? (contrast(bg, dk) >= contrast(bg, lt) ? dk : lt) : dk);

const fontStack = (f) => {
  const names = [f?.latin, f?.ea].filter((n) => typeof n === 'string' && n && !n.startsWith('+'));
  return [...new Set(names)].map((n) => `'${n.replace(/['"\\]/g, '')}'`).concat('sans-serif').join(',');
};

// Rough text measuring: CJK/fullwidth characters are 1em wide, others 0.58em.
const emWidth = (s) => [...s].reduce((n, c) => n + (c.codePointAt(0) >= 0x1100 ? 1 : 0.58), 0);
const linesFor = (s, boxPx, fontPx) => Math.max(1, Math.ceil((emWidth(s) * fontPx) / Math.max(1, boxPx)));

// ---------- content parsing ----------

function parseContent(text) {
  const slides = [];
  let cur = [];
  const flush = () => { if (cur.some((l) => l.trim())) slides.push(cur); cur = []; };
  for (const line of text.replace(/^\uFEFF/, '').split(/\r?\n/)) {
    if (line.trim() === '---') flush(); else cur.push(line);
  }
  flush();
  return slides.map((lines) => {
    const s = { layout: null, title: null, items: [], plain: [], images: [], tables: [] };
    let first = true;
    for (const raw of lines) {
      const line = raw.trim();
      if (!line) continue;
      if (first) {
        first = false;
        const m = /^<!--\s*layout:\s*(.+?)\s*-->$/i.exec(line);
        if (m) { s.layout = m[1]; continue; }
      }
      let m;
      if ((m = /^!\[[^\]]*\]\(([^)]+)\)$/.exec(line))) s.images.push(m[1].trim());
      else if ((m = /^\[table:\s*(.+?)\s*\]$/i.exec(line))) s.tables.push(m[1]);
      else if ((m = /^#{1,6}\s+(.*)$/.exec(line))) { if (s.title === null) s.title = m[1].trim(); else { s.items.push(m[1].trim()); s.plain.push(true); } }
      else if ((m = /^[-*]\s+(.*)$/.exec(line))) { s.items.push(m[1].trim()); s.plain.push(false); }
      else { s.items.push(line); s.plain.push(!/^\d{1,3}\.\s/.test(line)); } // plain = no bullet or number-dot marker; "2025. 11" is not one (AC10)
    }
    return s;
  });
}

// ---------- layout selection and placeholder slots ----------

function pickLayout(spec, slide, warnings, n) {
  const layouts = spec.layouts ?? [];
  if (slide.layout) {
    const found = layouts.find((l) => l.name === slide.layout) ?? layouts.find((l) => String(l.name).toLowerCase() === slide.layout.toLowerCase());
    if (found) return found;
    warnings.push(`slide ${n}: layout "${slide.layout}" not found in the spec; chose a layout automatically`);
  }
  if (slide.tables.length) return layouts.find((l) => l.kind === 'table') ?? layouts.find((l) => l.kind === 'content') ?? layouts[0];
  return layouts.find((l) => l.kind === 'content') ?? layouts[0];
}

const roleOfPlaceholder = (type) => {
  const t = String(type ?? '').toLowerCase();
  if (t === 'title' || t === 'ctrtitle') return 'title';
  if (t === 'pic' || t === 'picture') return 'image';
  if (['dt', 'ftr', 'sldnum', 'hdr', 'sldimg', 'date', 'footer', 'slidenumber'].includes(t)) return null;
  return 'body'; // body, subTitle, obj, tbl, chart, ... and placeholders without a type
};

// Default boxes (EMU, as fractions of the slide) for content that the layout has no placeholder for.
const DEFAULTS = {
  title: [0.05, 0.05, 0.9, 0.16],
  body: [0.05, 0.25, 0.9, 0.65],
};

function slotsOf(layout, size) {
  const slots = { title: null, body: [], image: [] };
  const list = [...(layout.placeholders ?? [])].sort((a, b) => (a.idx ?? 0) - (b.idx ?? 0));
  for (const p of list) {
    const role = roleOfPlaceholder(p.type);
    if (!role || ![p.x, p.y, p.w, p.h].every(Number.isFinite)) continue;
    const box = { x: p.x, y: p.y, w: p.w, h: p.h, type: p.type, textStyle: p.textStyle };
    if (role === 'title') slots.title ??= box; else slots[role].push(box);
  }
  const def = (k) => ({ x: DEFAULTS[k][0] * size.cx, y: DEFAULTS[k][1] * size.cy, w: DEFAULTS[k][2] * size.cx, h: DEFAULTS[k][3] * size.cy });
  slots.defaultTitle = def('title');
  slots.defaultBody = def('body');
  return slots;
}

// ---------- rendering ----------

const SIZES = [24, 22, 20, 18, 16];

function renderList(items, fs, color, fontFamily, extra = '') {
  const lis = items.map((t) => `<li style="margin:0 0 ${px(fs * 0.4)} 0">${fmt(t)}</li>`).join('');
  return `<ul data-ph="body" style="margin:0;padding:0 0 0 ${px(fs * 1.2)};list-style:disc;font-family:${fontFamily};font-size:${px(fs)};line-height:1.3;color:${color};${extra}">${lis}</ul>`;
}

const NUMERIC = /^[-+]?[$₩€¤£¥]?[\d,]*\.?\d+%?[$₩€¤£¥]?$/;

function tableMetrics(rows, boxW, fs) {
  const cols = Math.max(1, ...rows.map((r) => r.length));
  const weight = Array.from({ length: cols }, (_, c) => Math.max(2, ...rows.map((r) => emWidth(r[c] ?? ''))));
  const sum = weight.reduce((a, b) => a + b, 0);
  let h = 0;
  for (const r of rows) {
    const lines = Math.max(1, ...Array.from({ length: cols }, (_, c) => linesFor(r[c] ?? '', (boxW * weight[c]) / sum - 24, fs)));
    h += lines * fs * 1.3 + 16;
  }
  return h;
}

function renderTable(rows, fs, p) {
  const cols = Math.max(1, ...rows.map((r) => r.length));
  const border = `1px solid ${p.border}`;
  const cell = (tag, text, extra) => `<${tag} style="padding:8px 12px;border:${border};text-align:${NUMERIC.test(text.trim()) ? 'right' : 'left'};vertical-align:middle;${extra}">${esc(text)}</${tag}>`;
  const [head, ...rest] = rows;
  const tr = (r, tag, extra) => `<tr>${Array.from({ length: cols }, (_, c) => cell(tag, r[c] ?? '', extra)).join('')}</tr>`;
  const thExtra = `background-color:${p.headBg};color:${p.headText};font-weight:700;`;
  return `<table style="width:100%;border-collapse:collapse;table-layout:auto;font-family:${p.font};font-size:${px(fs)};line-height:1.3;color:${p.text};margin:0 0 ${px(fs * 0.6)} 0"><tbody>`
    + (head ? tr(head, 'th', thExtra) : '') + rest.map((r) => tr(r, 'td', '')).join('') + '</tbody></table>';
}

// The placeholder's textStyle (colour, bold, align) as CSS parts; null fields keep the defaults.
const ALIGN = { l: 'left', ctr: 'center', r: 'right' };
const JUSTIFY = { l: 'flex-start', ctr: 'center', r: 'flex-end' };
function textStyleOf(box) {
  const ts = box?.textStyle ?? {};
  return {
    color: hex(ts.color),
    weight: ts.bold === true ? 700 : ts.bold === false ? 400 : null,
    align: ALIGN[ts.align] ?? null,
    justify: JUSTIFY[ts.align] ?? null,
    sizePt: Number.isFinite(ts.size) && ts.size > 0 ? ts.size : null,
  };
}

function posStyle(box, k, extra = '') {
  return `position:absolute;left:${px(box.x * k)};top:${px(box.y * k)};width:${px(box.w * k)};height:${px(box.h * k)};${extra}`;
}

function renderSlide(slide, layout, ctx) {
  const { spec, k, frameW, frameH, palette, assetHref, warnings, n, tables, imageHref, noChrome } = ctx;
  const dk = hex(palette.dk1) ?? FALLBACK.dk;
  const lt = hex(palette.lt1) ?? FALLBACK.lt;
  const bg = layout.background ?? { type: 'none' };
  let frameBg = '#FFFFFF';
  let bgColor = '#FFFFFF';
  if (bg.type === 'color' && hex(bg.value)) { frameBg = bgColor = hex(bg.value); }
  else if (bg.type === 'image' && bg.value) {
    const href = assetHref(bg.value);
    if (href) { frameBg = `#FFFFFF url('${href}') center/100% 100% no-repeat`; bgColor = null; }
  }
  const textColor = bgColor ? textOn(bgColor, dk, lt) : dk;
  // Slide-level style learned by the extractor (spec.observed) wins over theme fonts; a placeholder's own colour wins over it.
  const obs = spec.observed ?? {};
  const headBg = hex(obs.colours?.tableHeader) ?? hex(palette.accent1) ?? hex(palette.dk2) ?? dk;
  const headText = hex(obs.colours?.tableHeaderText) ?? textOn(headBg, dk, lt);
  const major = fontStack(spec.fonts?.major);
  const minor = fontStack(spec.fonts?.minor);
  const hasFont = (f) => [f?.latin, f?.ea].some((n) => typeof n === 'string' && n && !n.startsWith('+'));
  // Per-layout learned title/body style (observed.layoutStyles) wins over the placeholder textStyle and the global observed values.
  const ls = obs.layoutStyles?.[layout.name] ?? {};
  const lsFont = (role, fallback) => (typeof ls[role]?.latin === 'string' && ls[role].latin ? fontStack({ latin: ls[role].latin, ea: obs.fonts?.[role === 'title' ? 'title' : 'body']?.ea }) : fallback);
  const lsBox = (role, box) => {
    const s = ls[role];
    if (!s || ![s.x, s.y, s.w, s.h].every(Number.isFinite)) return box;
    // Clamp the learned box into the 5% safe area (the spec keeps the learned value); boxes already inside are unchanged.
    const { cx, cy } = spec.slideSize;
    const l = Math.max(s.x, cx * 0.05), t = Math.max(s.y, cy * 0.05);
    const r = Math.min(s.x + s.w, cx * 0.95), b = Math.min(s.y + s.h, cy * 0.95);
    return { ...box, x: l, y: t, w: Math.max(0, r - l), h: Math.max(0, b - t) };
  };
  const lsText = (role, ts) => {
    const s = ls[role] ?? {};
    return {
      color: hex(s.colour) ?? ts.color,
      weight: s.bold === true ? 700 : s.bold === false ? 400 : ts.weight,
      align: ALIGN[s.align] ?? ts.align,
      justify: JUSTIFY[s.align] ?? ts.justify,
      sizePt: Number.isFinite(s.size) && s.size > 0 ? s.size : ts.sizePt,
    };
  };
  const titleFont = lsFont('title', hasFont(obs.fonts?.title) ? fontStack(obs.fonts.title) : major);
  const bodyFont = lsFont('body', hasFont(obs.fonts?.body) ? fontStack(obs.fonts.body) : minor);
  const obsTitleColor = hex(obs.colours?.title);
  const obsBodyColor = hex(obs.colours?.body);

  const out = [];
  // Layout decorations (master and layout shapes, already merged into layout.decorations by the extractor).
  for (const d of noChrome ? [] : layout.decorations ?? []) {
    if (![d.x, d.y, d.w, d.h].every(Number.isFinite)) continue;
    const box = { x: Math.max(0, d.x), y: Math.max(0, d.y) };
    box.w = Math.min(d.w, spec.slideSize.cx - box.x);
    box.h = Math.min(d.h, spec.slideSize.cy - box.y);
    if (box.w <= 0 || box.h <= 0) continue;
    const attrs = 'data-tpl="chrome" data-role="chrome"';
    if (d.kind === 'pic') {
      const href = d.asset ? assetHref(d.asset) : null;
      if (!href) continue;
      out.push(`<img ${attrs} src="${esc(href)}" alt="" style="${posStyle(box, k, 'object-fit:contain;')}">`);
      continue;
    }
    const fill = hex(d.fill);
    const style = posStyle(box, k, fill ? `background-color:${fill};` : '');
    if (d.text) {
      const fs = Math.max(16, Math.min(24, box.h * k * 0.5));
      const color = textOn(fill ?? bgColor, dk, lt);
      out.push(`<div ${attrs} style="${style}font-family:${minor};font-size:${px(fs)};line-height:1.2;color:${color};">${esc(d.text)}</div>`);
    } else out.push(`<div ${attrs} style="${style}"></div>`);
  }

  // AC10: the first plain body line goes to the learned subtitle box (no bullet); without one, a learned
  // text box style (layoutStyles.textbox) replaces the body style, its null fields falling back to it.
  const subIdx = ls.subtitle && [ls.subtitle.x, ls.subtitle.y, ls.subtitle.w, ls.subtitle.h].every(Number.isFinite) ? slide.plain.indexOf(true) : -1;
  const items = slide.items.filter((_, i) => i !== subIdx);
  const useTextbox = subIdx < 0 && !!ls.textbox;

  const slots = slotsOf(layout, spec.slideSize);
  if (slots.title === null && slide.title !== null && layout.placeholders?.length) warnings.push(`slide ${n}: layout "${layout.name}" has no title placeholder; used a default title box`);

  // Image slot(s): the layout's picture placeholder, else a second body placeholder, else the right part of the body box.
  let bodyBox = lsBox('body', slots.body[0] ?? slots.defaultBody);
  if (useTextbox && !tables.length) bodyBox = lsBox('textbox', bodyBox); // a table keeps the wider body box
  let imageBox = null;
  if (slide.images.length) {
    if (slots.image[0]) imageBox = slots.image[0];
    else if (slots.body[1]) imageBox = slots.body[1];
    else {
      const gap = 0.03 * bodyBox.w;
      const bodyW = (bodyBox.w - gap) * 0.55;
      imageBox = { x: bodyBox.x + bodyW + gap, y: bodyBox.y, w: bodyBox.w - bodyW - gap, h: bodyBox.h };
      bodyBox = { ...bodyBox, w: bodyW };
    }
  }

  if (slide.title !== null) {
    const box = lsBox('title', slots.title ?? slots.defaultTitle);
    const wPx = box.w * k;
    const hPx = box.h * k;
    const ts = lsText('title', textStyleOf(box));
    const fs = ts.sizePt !== null ? ts.sizePt * 12700 * k : ([40, 36, 32, 28, 24].find((s) => linesFor(slide.title, wPx, s) * s * 1.2 <= hPx) ?? 24);
    const tsAlign = ts.align ? `text-align:${ts.align};justify-content:${ts.justify};` : '';
    out.push(`<h1 data-role="title" data-ph="title" style="${posStyle(box, k, `margin:0;display:flex;align-items:center;${tsAlign}font-family:${titleFont};font-size:${px(fs)};font-weight:${ts.weight ?? 400};line-height:1.2;color:${ts.color ?? obsTitleColor ?? textColor};`)}"><span style="display:inline;">${fmt(slide.title)}</span></h1>`);
  }

  if (subIdx >= 0) {
    const ts = lsText('subtitle', textStyleOf(null));
    const fs = ts.sizePt !== null ? ts.sizePt * 12700 * k : 18;
    const align = ts.align ? `text-align:${ts.align};` : '';
    out.push(`<p data-role="subtitle" data-ph="subtitle" style="${posStyle(lsBox('subtitle', ls.subtitle), k,`margin:0;white-space:nowrap;${align}font-family:${lsFont('subtitle', bodyFont)};font-size:${px(fs)};font-weight:${ts.weight ?? 400};line-height:1.3;color:${ts.color ?? obsBodyColor ?? textColor};`)}">${fmt(slide.items[subIdx])}</p>`);
  }

  if (items.length || tables.length) {
    const wPx = bodyBox.w * k;
    const hPx = bodyBox.h * k;
    const rowsAll = tables.map((t) => t.rows);
    const need = (fs) => {
      const listH = items.reduce((h, t) => h + linesFor(t, wPx - fs * 1.2, fs) * fs * 1.3 + fs * 0.4, 0);
      const tfs = Math.min(fs, 20);
      return listH + rowsAll.reduce((h, r) => h + tableMetrics(r, wPx, tfs) + tfs * 0.6, 0);
    };
    const bodyTs = lsText('body', textStyleOf(bodyBox));
    const bs = useTextbox ? lsText('textbox', bodyTs) : bodyTs;
    const listFont = useTextbox ? lsFont('textbox', bodyFont) : bodyFont;
    const fs = bs.sizePt !== null ? bs.sizePt * 12700 * k : (SIZES.find((s) => need(s) <= hPx) ?? 16);
    if (need(fs) > hPx) warnings.push(`slide ${n}: body content may not fit its box even at ${px(fs)}`);
    const bsExtra = `${bs.weight ? `font-weight:${bs.weight};` : ''}${bs.align ? `text-align:${bs.align};` : ''}`;
    const plain = String(bodyBox.type ?? '').toLowerCase() === 'subtitle';
    const inner = (items.length
      ? (plain
        ? `<p data-ph="body" style="margin:0;font-family:${listFont};font-size:${px(fs)};line-height:1.3;color:${bs.color ?? obsBodyColor ?? textColor};${bsExtra}">${items.map(fmt).join('<br>')}</p>`
        : renderList(items, fs, bs.color ?? obsBodyColor ?? textColor, listFont, bsExtra))
      : '')
      + tables.map((t) => renderTable(t.rows, Math.min(fs, 20), {
        font: bodyFont, text: obsBodyColor ?? textColor, border: bgColor && lum(bgColor) < 0.4 ? '#7F7F7F' : '#BFBFBF',
        headBg, headText,
      })).join('');
    out.push(`<div data-role="body" style="${posStyle(bodyBox, k)}">${inner}</div>`);
  }

  if (imageBox) {
    const n2 = slide.images.length;
    const gap = n2 > 1 ? 12 / k : 0;
    const w = (imageBox.w - gap * (n2 - 1)) / n2;
    slide.images.forEach((name, i) => {
      const box = { x: imageBox.x + i * (w + gap), y: imageBox.y, w, h: imageBox.h };
      out.push(`<img data-role="image" src="${esc(imageHref(name))}" alt="${esc(basename(name))}" style="${posStyle(box, k, 'object-fit:contain;')}">`);
    });
  }

  const left = (CANVAS_W - frameW) / 2;
  const top = (CANVAS_H - frameH) / 2;
  return `<section class="slide" data-layout="${esc(layout.name ?? '')}">\n`
    + `<div class="frame" style="position:absolute;left:${px(left)};top:${px(top)};width:${px(frameW)};height:${px(frameH)};${noChrome ? '' : `background:${frameBg};`}">\n${out.join('\n')}\n</div>\n</section>`;
}

// ---------- main ----------

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const specPath = resolve(args.specPath);
  let spec;
  try { spec = JSON.parse((await readFile(specPath, 'utf8')).replace(/^\uFEFF/, '')); } catch (e) { throw new InputError(`cannot read spec ${args.specPath}: ${e.code === 'ENOENT' ? 'file not found' : e.message}`); }
  if (!spec.slideSize?.cx || !spec.slideSize?.cy || !Array.isArray(spec.layouts) || !spec.layouts.length) throw new InputError('spec needs slideSize {cx,cy} and at least one layout');
  let contentText;
  try { contentText = await readFile(args.contentPath, 'utf8'); } catch { throw new InputError(`content file not found: ${args.contentPath}`); }
  const slides = parseContent(contentText);
  if (!slides.length) throw new InputError('content has no slides');

  const warnings = [];
  const errors = [];
  const templateAssets = join(dirname(specPath), 'assets');
  const assetsDir = args.assets ? resolve(args.assets) : null;
  const outDir = resolve(args.out);
  const outAssets = join(outDir, 'assets');
  const copies = new Map(); // output file name -> source path

  // Data files, matched by base name (with or without extension).
  const dataFiles = args.data.map((p) => ({ path: resolve(p), name: basename(p), stem: basename(p, extname(p)), used: false }));
  const tableCache = new Map();
  const findData = (ref) => {
    const r = basename(ref).toLowerCase();
    const exact = dataFiles.filter((d) => d.name.toLowerCase() === r);
    const pool = exact.length ? exact : dataFiles.filter((d) => d.stem.toLowerCase() === r);
    if (pool.length > 1) warnings.push(`table "${ref}" matches ${pool.length} data files (${pool.map((d) => d.name).join(', ')}); used ${pool[0].name}`);
    return pool[0] ?? null;
  };

  // Images from the --assets folder (exact name, then case-insensitive).
  let assetNames = null;
  const findImage = async (name) => {
    if (!assetsDir) return null;
    const direct = resolve(assetsDir, name);
    const rel = relative(assetsDir, direct);
    if (rel.startsWith('..') || isAbsolute(rel)) return null;
    if (await exists(direct)) return { src: direct, out: rel.split(/[\\/]/).join('_') };
    assetNames ??= await readdir(assetsDir).catch(() => []);
    const hit = assetNames.find((f) => f.toLowerCase() === name.toLowerCase());
    return hit ? { src: join(assetsDir, hit), out: hit } : null;
  };
  const imageOut = new Map();
  for (const s of slides) {
    for (const name of s.images) {
      if (imageOut.has(name)) continue;
      const hit = await findImage(name);
      if (!hit) { errors.push(`image not found in --assets: ${name}`); continue; }
      imageOut.set(name, hit.out);
      copies.set(hit.out, hit.src);
    }
  }

  // Tables per slide.
  const slideTables = [];
  for (const s of slides) {
    const list = [];
    for (const ref of s.tables) {
      const d = findData(ref);
      if (!d) { errors.push(`data file not given for [table: ${ref}] (pass it with --data)`); continue; }
      d.used = true;
      if (!(await exists(d.path))) { errors.push(`data file not found: ${d.path}`); continue; }
      if (!tableCache.has(d.path)) {
        try { tableCache.set(d.path, await loadTable(d.path)); } catch (e) { if (e instanceof InputError) { errors.push(e.message); continue; } throw e; }
      }
      const rows = tableCache.get(d.path);
      if (!rows.length) { warnings.push(`data file ${d.name} has no rows`); continue; }
      list.push({ name: d.name, rows });
    }
    slideTables.push(list);
  }
  for (const d of dataFiles) {
    if (!d.used) warnings.push(`unused data file: ${d.name}`);
    if (!(await exists(d.path)) && !errors.some((e) => e.includes(d.path))) errors.push(`data file not found: ${d.path}`);
  }
  if (errors.length) throw new InputError(errors.join('\n'));

  // Template assets (logos, background pictures). Missing or EMF/WMF files are skipped with a warning.
  const tplOut = new Map();
  const assetHref = (file) => {
    const name = basename(String(file));
    if (tplOut.has(name)) return tplOut.get(name);
    let href = null;
    if (/\.(emf|wmf)$/i.test(name)) warnings.push(`template picture ${name} is EMF/WMF and is not drawn`);
    else {
      const src = join(templateAssets, name);
      const outName = `tpl-${name}`;
      copies.set(outName, src);
      href = `assets/${encodeURI(outName)}`;
    }
    tplOut.set(name, href);
    return href;
  };
  const imageHref = (name) => `assets/${encodeURI(imageOut.get(name))}`;

  const { cx, cy } = spec.slideSize;
  const k = Math.min(CANVAS_W / cx, CANVAS_H / cy);
  const ctxBase = { spec, k, frameW: cx * k, frameH: cy * k, palette: spec.palette ?? {}, assetHref, warnings, imageHref, noChrome: !!args.noChrome };

  const html = [];
  slides.forEach((s, i) => {
    const layout = pickLayout(spec, s, warnings, i + 1);
    html.push(renderSlide(s, layout, { ...ctxBase, n: i + 1, tables: slideTables[i] }));
  });

  await mkdir(outAssets, { recursive: true });
  const copied = [];
  for (const [outName, src] of copies) {
    if (!(await exists(src))) {
      warnings.push(`asset file missing, not copied: ${src}`);
      continue;
    }
    await copyFile(src, join(outAssets, outName));
    copied.push(outName);
  }

  const doc = `<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${esc(basename(args.contentPath, extname(args.contentPath)))}</title>
<style>
html,body{margin:0;padding:0;background:#FFFFFF}
section.slide{position:relative;display:block;width:${CANVAS_W}px;height:${CANVAS_H}px;background:#FFFFFF;box-sizing:border-box}
img{display:block}
</style>
</head>
<body>
${html.join('\n')}
</body>
</html>
`;
  await writeFile(join(outDir, 'deck.html'), doc);
  await writeFile(join(outDir, 'build-report.json'), JSON.stringify({
    slides: slides.length,
    deck: 'deck.html',
    assets: copied,
    warnings,
    specWarnings: spec.warnings ?? [],
  }, null, 2));
  console.log(`wrote ${join(outDir, 'deck.html')} (${slides.length} slides, ${warnings.length} warnings)`);
}

main().catch((e) => {
  if (e instanceof InputError) {
    console.error(e.message);
    process.exit(2);
  }
  console.error(e.stack ?? e);
  process.exit(1);
});
