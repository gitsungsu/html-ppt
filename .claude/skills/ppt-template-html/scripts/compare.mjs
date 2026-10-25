// Compare each layout of a sample PPTX (rendered by PowerPoint) with the same layout generated as HTML.
// usage: node compare.mjs --spec <template-spec.json> --sample <sample.pptx> --filler <filler.md> --out <dir>
// Steps: make-layout-slides.ps1 (original PNGs), build-deck.mjs (one generated deck per layout), then per layout
//   colour    top 8 colours by pixel share, 8x8 minimum-cost matching on Delta E 76, mean value; pass <= 15 (D2), human-check 12-18
//   structure boxes matched by data-role, IoU per pair; pass when nothing is missing and mean IoU >= 0.8 (D2), human-check 0.6-0.8
//   layout    pass = colour pass AND structure pass; picture-background layouts go to the human-check list
// Both images are letterboxed to 1280x720 on white (D3). Pixels are read in headless Chromium (canvas), no image packages.
// Outputs: <dir>/report.json, <dir>/report.md, <dir>/original/*.png, <dir>/original-lb/*.png, <dir>/generated-png/*.png
// Exit codes: 0 ran (see the report for layout verdicts), 2 usage or input error, 4 PowerPoint not available (comparison skipped, D4).
import { createRequire } from 'node:module';
import { mkdirSync, readFileSync, writeFileSync, existsSync } from 'node:fs';
import { resolve, join, dirname } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { spawnSync } from 'node:child_process';

const here = dirname(fileURLToPath(import.meta.url));
const usage = 'usage: node compare.mjs --spec <template-spec.json> --sample <sample.pptx> --filler <filler.md> --out <dir>';
const args = process.argv.slice(2);
const opt = (name) => (args.includes(name) ? args[args.indexOf(name) + 1] : undefined);
const [specArg, sampleArg, fillerArg, outArg] = ['--spec', '--sample', '--filler', '--out'].map(opt);
if (!specArg || !sampleArg || !fillerArg || !outArg) {
  console.error(usage);
  process.exit(2);
}
const specFile = resolve(specArg);
const sampleFile = resolve(sampleArg);
const fillerFile = resolve(fillerArg);
const out = resolve(outArg);
for (const f of [specFile, sampleFile, fillerFile]) {
  if (!existsSync(f)) {
    console.error(`input not found: ${f}\n${usage}`);
    process.exit(2);
  }
}

const COLOUR_PASS = 15;
const COLOUR_BAND = [12, 18];
const IOU_PASS = 0.8;
const IOU_BAND = [0.6, 0.8];
const W = 1280;
const H = 720;

const spec = JSON.parse(readFileSync(specFile, 'utf8').replace(/^﻿/, ''));
const layouts = spec.layouts || [];
if (!layouts.length || !spec.slideSize) {
  console.error('spec needs slideSize and layouts[]');
  process.exit(2);
}
// Same file-name rule as make-layout-slides.ps1.
const safeName = (name) => name.replace(/[\\/:*?"<>|]/g, '_');

// One filler slide: comment lines dropped, text after a "---" line ignored.
const fillerText = (() => {
  const keep = [];
  for (const line of readFileSync(fillerFile, 'utf8').replace(/^﻿/, '').split(/\r?\n/)) {
    const t = line.trim();
    if (/^---+$/.test(t)) { if (keep.length) break; else continue; }
    if (/^<!--.*-->$/.test(t)) continue;
    keep.push(line);
  }
  return keep.join('\n').trim();
})();

mkdirSync(join(out, 'original'), { recursive: true });
mkdirSync(join(out, 'original-lb'), { recursive: true });
mkdirSync(join(out, 'generated-png'), { recursive: true });

// (a) original PNGs
const ps = spawnSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', join(here, 'make-layout-slides.ps1'),
  '-Sample', sampleFile, '-Spec', specFile, '-Filler', fillerFile, '-Out', out], { encoding: 'utf8' });
if (ps.stdout) console.log(ps.stdout.trim());
if (ps.status === 4) {
  console.error('PowerPoint is not available: comparison skipped (AC5 only).');
  process.exit(4);
}
if (ps.status !== 0) {
  console.error(`make-layout-slides.ps1 failed (exit ${ps.status}): ${(ps.stderr || '').trim()}`);
  process.exit(2);
}

// (b) generated decks, one per layout
const buildErrors = {};
for (const l of layouts) {
  const dir = join(out, 'generated', safeName(l.name));
  mkdirSync(dir, { recursive: true });
  const content = join(dir, 'content.md');
  writeFileSync(content, `<!-- layout: ${l.name} -->\n${fillerText}\n`, 'utf8');
  const r = spawnSync(process.execPath, [join(here, 'build-deck.mjs'), specFile, content, '--out', dir], { encoding: 'utf8' });
  if (r.status !== 0 && r.status !== 3) buildErrors[l.name] = `build-deck exit ${r.status}: ${(r.stderr || r.stdout || '').trim().split('\n')[0]}`;
  else if (!existsSync(join(dir, 'deck.html'))) buildErrors[l.name] = 'build-deck wrote no deck.html';
}

// ---- pixel work, in the page ----
// Draws a PNG letterboxed on white into a 1280x720 canvas and returns the top 8 colours (4 bits per channel bins).
const analyzeInPage = async ({ dataUrl, W, H }) => {
  const img = new Image();
  await new Promise((res, rej) => { img.onload = res; img.onerror = () => rej(new Error('image load failed')); img.src = dataUrl; });
  const cv = document.createElement('canvas');
  cv.width = W;
  cv.height = H;
  const ctx = cv.getContext('2d', { willReadFrequently: true });
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, W, H);
  const s = Math.min(W / img.naturalWidth, H / img.naturalHeight);
  const w = img.naturalWidth * s;
  const h = img.naturalHeight * s;
  ctx.drawImage(img, (W - w) / 2, (H - h) / 2, w, h);
  const d = ctx.getImageData(0, 0, W, H).data;
  const count = new Float64Array(4096);
  const sum = new Float64Array(4096 * 3);
  for (let i = 0; i < d.length; i += 4) {
    const k = ((d[i] >> 4) << 8) | ((d[i + 1] >> 4) << 4) | (d[i + 2] >> 4);
    count[k]++;
    sum[k * 3] += d[i];
    sum[k * 3 + 1] += d[i + 1];
    sum[k * 3 + 2] += d[i + 2];
  }
  const bins = [];
  for (let k = 0; k < 4096; k++) if (count[k]) bins.push(k);
  bins.sort((a, b) => count[b] - count[a]);
  const total = W * H;
  const colors = bins.slice(0, 8).map((k) => ({
    rgb: [0, 1, 2].map((c) => Math.round(sum[k * 3 + c] / count[k])),
    share: count[k] / total,
  }));
  return { colors, letterboxed: cv.toDataURL('image/png') };
};

const toLab = ([r, g, b]) => {
  const lin = (c) => ((c /= 255) <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
  const [R, G, B] = [lin(r), lin(g), lin(b)];
  const x = (0.4124564 * R + 0.3575761 * G + 0.1804375 * B) / 0.95047;
  const y = 0.2126729 * R + 0.7151522 * G + 0.072175 * B;
  const z = (0.0193339 * R + 0.119192 * G + 0.9503041 * B) / 1.08883;
  const f = (t) => (t > 216 / 24389 ? Math.cbrt(t) : (24389 / 27 * t + 16) / 116);
  return [116 * f(y) - 16, 500 * (f(x) - f(y)), 200 * (f(y) - f(z))];
};
const dE76 = (a, b) => Math.hypot(a[0] - b[0], a[1] - b[1], a[2] - b[2]);

// Minimum-cost matching of the smaller colour list into the larger one (at most 8 x 8: bitmask DP). Returns the mean pair cost.
const matchColours = (A, B) => {
  const [S, L] = A.length <= B.length ? [A, B] : [B, A];
  if (!S.length) return null;
  const labS = S.map((c) => toLab(c.rgb));
  const labL = L.map((c) => toLab(c.rgb));
  const full = 1 << L.length;
  let dp = new Array(full).fill(Infinity);
  dp[0] = 0;
  for (let i = 0; i < S.length; i++) {
    const next = new Array(full).fill(Infinity);
    for (let mask = 0; mask < full; mask++) {
      if (dp[mask] === Infinity) continue;
      for (let j = 0; j < L.length; j++) {
        if (mask & (1 << j)) continue;
        const v = dp[mask] + dE76(labS[i], labL[j]);
        if (v < next[mask | (1 << j)]) next[mask | (1 << j)] = v;
      }
    }
    dp = next;
  }
  return Math.min(...dp) / S.length;
};

// ---- structure ----
// Placeholder type to the data-role a generated element carries (null = not rendered, so not compared).
const roleOfPlaceholder = (type) => {
  const t = String(type || '').toLowerCase();
  if (/date|^dt$|footer|^ftr$|sld.?num|slide.?num|^sldnum$/.test(t)) return null;
  if (/title/.test(t) && !/sub/.test(t)) return 'title';
  if (/pic|image|media|clipart/.test(t)) return 'image';
  if (/body|sub|obj|text|content|chart|tbl|table/.test(t)) return 'body';
  return null;
};
const cx = spec.slideSize.cx;
const cy = spec.slideSize.cy;
const k = Math.min(W / cx, H / cy);
const offX = (W - cx * k) / 2;
const offY = (H - cy * k) / 2;
const toBox = (role, e) => ({ role, x: offX + e.x * k, y: offY + e.y * k, w: e.w * k, h: e.h * k });
const originalBoxes = (l) => [
  ...(l.placeholders || []).map((p) => [roleOfPlaceholder(p.type), p]).filter(([r]) => r).map(([r, p]) => toBox(r, p)),
  ...(l.decorations || []).map((d) => toBox('chrome', d)),
].filter((b) => b.w > 0 && b.h > 0);

const iou = (a, b) => {
  const w = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
  const h = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
  if (w <= 0 || h <= 0) return 0;
  const inter = w * h;
  return inter / (a.w * a.h + b.w * b.h - inter);
};
// Pair original and generated boxes of the same role, best IoU first, one to one.
const matchBoxes = (orig, gen) => {
  const pairs = [];
  const usedO = new Set();
  const usedG = new Set();
  const cand = [];
  orig.forEach((o, i) => gen.forEach((g, j) => { if (o.role === g.role) cand.push({ i, j, v: iou(o, g) }); }));
  cand.sort((a, b) => b.v - a.v);
  for (const c of cand) {
    if (usedO.has(c.i) || usedG.has(c.j)) continue;
    usedO.add(c.i);
    usedG.add(c.j);
    pairs.push({ role: orig[c.i].role, iou: c.v });
  }
  const missing = orig.filter((_, i) => !usedO.has(i)).map((o) => o.role);
  return { pairs, missing };
};

// ---- run ----
let chromium;
try {
  ({ chromium } = createRequire(import.meta.url)('playwright'));
} catch {
  console.error('playwright is not installed: cd scripts && npm install && npx playwright install chromium');
  process.exit(2);
}
let browser;
try {
  browser = await chromium.launch({ args: ['--disable-lcd-text', '--font-render-hinting=none'] });
} catch (e) {
  console.error(`cannot start Chromium: ${String(e.message).split('\n')[0]}\nrun once: npx playwright install chromium`);
  process.exit(2);
}

const results = [];
try {
  const tool = await browser.newPage({ viewport: { width: W, height: H } });
  await tool.goto('about:blank');
  const analyze = (file) => tool.evaluate(analyzeInPage, { dataUrl: `data:image/png;base64,${readFileSync(file).toString('base64')}`, W, H });
  const saveDataUrl = (file, url) => writeFileSync(file, Buffer.from(url.split(',')[1], 'base64'));

  for (const l of layouts) {
    const safe = safeName(l.name);
    const res = { name: l.name, original: `original/${safe}.png`, generated: `generated-png/${safe}.png`, humanCheck: [] };
    results.push(res);
    if (buildErrors[l.name]) {
      res.error = buildErrors[l.name];
      res.pass = false;
      continue;
    }
    const origFile = join(out, 'original', `${safe}.png`);
    if (!existsSync(origFile)) {
      res.error = 'no original PNG (layout not found in the sample)';
      res.pass = false;
      continue;
    }
    try {
      // (c) generated screenshot and boxes
      const page = await browser.newPage({ viewport: { width: W, height: H } });
      const deckUrl = pathToFileURL(join(out, 'generated', safe, 'deck.html')).href;
      await page.goto(deckUrl, { waitUntil: 'networkidle', timeout: 60000 });
      await page.evaluate(() => document.fonts.ready);
      await page.waitForTimeout(300);
      const slide = await page.$('section.slide');
      if (!slide) throw new Error('generated deck has no section.slide');
      await slide.evaluate((el) => el.scrollIntoView({ block: 'start' }));
      await page.waitForTimeout(100);
      const genFile = join(out, 'generated-png', `${safe}.png`);
      await page.screenshot({ path: genFile });
      const genBoxes = await slide.evaluate((sec) => {
        const r = sec.getBoundingClientRect();
        const frame = sec.querySelector('.frame') || sec;
        return [...frame.querySelectorAll('[data-role]')].map((e) => {
          const b = e.getBoundingClientRect();
          return { role: e.getAttribute('data-role'), x: b.left - r.left, y: b.top - r.top, w: b.width, h: b.height };
        }).filter((b) => b.w > 0 && b.h > 0);
      });
      await page.close();

      // (d) colour
      const o = await analyze(origFile);
      saveDataUrl(join(out, 'original-lb', `${safe}.png`), o.letterboxed);
      const g = await analyze(genFile);
      const mean = matchColours(o.colors, g.colors);
      res.colour = { meanDeltaE: mean, pass: mean !== null && mean <= COLOUR_PASS, original: o.colors, generated: g.colors };
      if (mean !== null && mean >= COLOUR_BAND[0] && mean <= COLOUR_BAND[1]) res.humanCheck.push(`colour DeltaE ${mean.toFixed(1)} is in the ${COLOUR_BAND.join('-')} band`);

      // (e) structure
      const orig = originalBoxes(l);
      const { pairs, missing } = matchBoxes(orig, genBoxes);
      const meanIoU = pairs.length ? pairs.reduce((s, p) => s + p.iou, 0) / pairs.length : null;
      const structPass = missing.length === 0 && (meanIoU === null || meanIoU >= IOU_PASS);
      res.structure = { originalBoxes: orig.length, generatedBoxes: genBoxes.length, missing, meanIoU, pairs, pass: structPass };
      if (meanIoU !== null && meanIoU >= IOU_BAND[0] && meanIoU < IOU_BAND[1]) res.humanCheck.push(`mean IoU ${meanIoU.toFixed(2)} is in the ${IOU_BAND.join('-')} band`);

      // (f) verdict
      if (l.background && l.background.type === 'image') res.humanCheck.push('picture background');
      res.pass = res.colour.pass && res.structure.pass;
    } catch (e) {
      res.error = String(e.message).split('\n')[0];
      res.pass = false;
    }
  }
} finally {
  await browser.close();
}

// ---- reports ----
const humanCheckList = results.filter((r) => r.humanCheck.length).map((r) => ({ layout: r.name, reasons: r.humanCheck }));
const summary = {
  layouts: results.length,
  pass: results.filter((r) => r.pass).length,
  fail: results.filter((r) => !r.pass).length,
  humanCheck: humanCheckList.length,
  thresholds: { colourDeltaE: COLOUR_PASS, colourBand: COLOUR_BAND, iou: IOU_PASS, iouBand: IOU_BAND },
};
writeFileSync(join(out, 'report.json'), JSON.stringify({ spec: specFile, sample: sampleFile, filler: fillerFile, summary, layouts: results, humanCheckList }, null, 2), 'utf8');

const f1 = (v) => (v === null || v === undefined ? '-' : v.toFixed(1));
const f2 = (v) => (v === null || v === undefined ? '-' : v.toFixed(2));
const md = [
  '# Layout comparison',
  '',
  `Sample: ${sampleFile}`,
  `Spec: ${specFile}`,
  `Thresholds: colour DeltaE76 mean <= ${COLOUR_PASS} (human-check ${COLOUR_BAND.join('-')}), mean IoU >= ${IOU_PASS} with no missing box (human-check ${IOU_BAND.join('-')})`,
  '',
  `Layouts: ${summary.layouts}, pass: ${summary.pass}, fail: ${summary.fail}, human-check: ${summary.humanCheck}`,
  '',
  '| Layout | Colour DeltaE | Colour | Mean IoU | Missing boxes | Structure | Layout | Human check |',
  '|---|---|---|---|---|---|---|---|',
  ...results.map((r) => (r.error
    ? `| ${r.name} | - | - | - | - | - | FAIL | error: ${r.error} |`
    : `| ${r.name} | ${f1(r.colour.meanDeltaE)} | ${r.colour.pass ? 'pass' : 'fail'} | ${f2(r.structure.meanIoU)} | ${r.structure.missing.length ? r.structure.missing.join(', ') : '-'} | ${r.structure.pass ? 'pass' : 'fail'} | ${r.pass ? 'PASS' : 'FAIL'} | ${r.humanCheck.length ? r.humanCheck.join('; ') : '-'} |`)),
  '',
  '## Human-check list',
  '',
  ...(humanCheckList.length ? humanCheckList.map((h) => `- ${h.layout}: ${h.reasons.join('; ')}`) : ['(none)']),
  '',
  'Image pairs: original/<layout>.png (PowerPoint), original-lb/<layout>.png (letterboxed 1280x720), generated-png/<layout>.png (HTML).',
  '',
];
writeFileSync(join(out, 'report.md'), md.join('\n'), 'utf8');
console.log(`compared ${summary.layouts} layout(s): ${summary.pass} pass, ${summary.fail} fail, ${summary.humanCheck} for human check`);
console.log(`report: ${join(out, 'report.md')}`);
