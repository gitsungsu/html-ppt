// Render-check a PPT-style HTML deck: slide count, console and network errors, overflow, the optional presenter (?present),
// and the rules of the PPT guide that a page can measure (references/ppt-conversion-guide.md, section 23):
// 16:9 canvas, background contrast (text at least 4.5:1), text inside the 5% safe area, no overlapping or clipped text,
// text not below 16px (body text at 24px is checked by eye).
// usage: node check-starter.mjs <file-or-url> <outDir> [--mode slides|sections] [--slides N]
//   slides    html-ppt format: every section.slide (1280x720) is a slide; the stacked page is checked slide by slide
//   sections  Marp decks (one page, one visible slide): ArrowRight moves on, but may first reveal a fragment, so presses repeat until the next slide is active
// Playwright is installed in this folder's node_modules: cd scripts && npm install && npx playwright install chromium
import { createRequire } from 'node:module';
import { mkdirSync } from 'node:fs';
import { resolve, join } from 'node:path';
import { pathToFileURL } from 'node:url';

const { chromium } = createRequire(import.meta.url)('playwright');

const args = process.argv.slice(2);
const opt = (name, def) => (args.includes(name) ? args[args.indexOf(name) + 1] : def);
// Positionals are the arguments that are neither options nor option values.
const [target, outDir] = args.filter((a, i) => !a.startsWith('--') && !args[i - 1]?.startsWith('--'));
const mode = opt('--mode', 'slides');
const slidesArg = opt('--slides', '0');
const usage = 'usage: node check-starter.mjs <file-or-url> <outDir> [--mode slides|sections] [--slides N]';

if (!target || !outDir) {
  console.error(usage);
  process.exit(2);
}
if (!['slides', 'sections'].includes(mode)) {
  console.error(`unknown --mode "${mode}"\n${usage}`);
  process.exit(2);
}
if (!/^\d+$/.test(String(slidesArg))) {
  console.error(`--slides must be a whole number, got "${slidesArg}"\n${usage}`);
  process.exit(2);
}
const expected = Number(slidesArg);
const url = /^https?:/.test(target) ? target : pathToFileURL(resolve(target)).href;
mkdirSync(outDir, { recursive: true });

const problems = [];
const notes = new Set();
const report = (text) => problems.push(text);

const CAPTION_FLOOR = 16; // px = 12pt, the smallest text size; guide section 11 asks for 19px captions, relaxed here
const BODY_FLOOR = 24;    // px = 18pt, the smallest body size in guide section 11; 19-23px is noted, not failed
const pad = (n) => String(n).padStart(2, '0');
const shots = [];

// WCAG contrast of two computed colours ("rgb(r, g, b)" or "rgba(r, g, b, a)"); alpha is ignored.
const rgbOf = (s) => (s.match(/[\d.]+/g) || []).slice(0, 3).map(Number);
const luminance = ([r, g, b]) => {
  const lin = (c) => ((c /= 255) <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
  return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b);
};
const contrast = (a, b) => {
  const [hi, lo] = [luminance(a), luminance(b)].sort((x, y) => y - x);
  return (hi + 0.05) / (lo + 0.05);
};

// Runs in the page on a slide's root element. Returns the canvas box, the text boxes relative to it, and the layout faults.
const measure = (root) => {
  const c = root.getBoundingClientRect();
  // The colour a text box sits on: the nearest ancestor with a background (white if none).
  const bgOf = (el) => {
    for (let n = el; n && n !== document.documentElement; n = n.parentElement) {
      const bg = getComputedStyle(n).backgroundColor;
      if (bg !== 'rgba(0, 0, 0, 0)' && bg !== 'transparent') return bg;
    }
    return 'rgb(255, 255, 255)';
  };
  const boxes = [];
  for (const el of root.querySelectorAll('*')) {
    if (el.namespaceURI !== 'http://www.w3.org/1999/xhtml') continue; // HTML text only; chart text inside SVG is not measured
    if (![...el.childNodes].some((n) => n.nodeType === 3 && n.textContent.trim())) continue;
    const cs = getComputedStyle(el);
    if (cs.display === 'none' || cs.visibility === 'hidden') continue;
    const r = el.getBoundingClientRect();
    if (r.width === 0 || r.height === 0) continue;
    boxes.push({ el, label: el.textContent.trim().replace(/\s+/g, ' ').slice(0, 18), size: parseFloat(cs.fontSize), color: /, 0\)$|^transparent$/.test(cs.color) ? null : cs.color, bg: bgOf(el), l: r.left, t: r.top, r: r.right, b: r.bottom });
  }
  const overlaps = [];
  for (let i = 0; i < boxes.length; i++) {
    for (let j = i + 1; j < boxes.length; j++) {
      const a = boxes[i];
      const b = boxes[j];
      if (a.el.contains(b.el) || b.el.contains(a.el)) continue; // a text box inside another one is not an overlap
      const w = Math.min(a.r, b.r) - Math.max(a.l, b.l);
      const h = Math.min(a.b, b.b) - Math.max(a.t, b.t);
      if (w <= 0 || h <= 0) continue;
      const smaller = Math.min((a.r - a.l) * (a.b - a.t), (b.r - b.l) * (b.b - b.t));
      if (w * h > 0.2 * smaller) overlaps.push(`"${a.label}" / "${b.label}"`);
    }
  }
  const overflow = [];
  for (const d of root.querySelectorAll('*')) {
    const b = d.getBoundingClientRect();
    if (b.width === 0 || b.height === 0) continue;
    if (b.right > c.right + 2 || b.bottom > c.bottom + 2 || b.left < c.left - 2 || b.top < c.top - 2) {
      const cls = d.className && d.className.baseVal === undefined ? '.' + String(d.className).split(' ')[0] : '';
      overflow.push(`${d.tagName.toLowerCase()}${cls}`);
    }
  }
  // Boxes that hide their overflow (overflow: hidden or clip) and hold more content than they show: the text is clipped.
  const clipped = [];
  for (const d of root.querySelectorAll('*')) {
    const cs = getComputedStyle(d);
    if (!/hidden|clip/.test(cs.overflowX + ' ' + cs.overflowY) || d.clientHeight === 0) continue;
    if (d.scrollHeight > d.clientHeight + 2 || d.scrollWidth > d.clientWidth + 2) {
      const cls = typeof d.className === 'string' && d.className ? '.' + d.className.split(' ')[0] : '';
      clipped.push(`${d.tagName.toLowerCase()}${cls}`);
    }
  }
  return {
    canvas: { w: c.width, h: c.height },
    backgroundImage: getComputedStyle(root).backgroundImage,
    texts: boxes.map((t) => ({ label: t.label, size: t.size, color: t.color, bg: t.bg, l: t.l - c.left, t: t.t - c.top, r: t.r - c.left, b: t.b - c.top })),
    overlaps,
    overflow: overflow.slice(0, 5),
    clipped: clipped.slice(0, 5),
    columns: [...root.querySelectorAll('*')].filter((d) => !['auto', '1'].includes(String(getComputedStyle(d).columnCount))).length,
  };
};

// Runs in the page on a slide's root element. Lists colours the html-ppt converter cannot read (oklch and other
// modern colour functions). Tailwind v4 palette classes compute to oklch(); the converter then drops those colours.
const colourIssues = (root) => {
  const modern = /\b(oklch|oklab|lab|lch|color|color-mix)\(/;
  const found = new Set();
  const note = (el, label, value) => {
    if (value && modern.test(value)) {
      const cls = (el.getAttribute('class') || '').split(' ')[0];
      found.add(`${el.tagName.toLowerCase()}${cls ? '.' + cls : ''} ${label} ${value}`);
    }
  };
  for (const el of [root, ...root.querySelectorAll('*')]) {
    const cs = getComputedStyle(el);
    if (!el.getClientRects().length || cs.visibility === 'hidden') continue; // no boxes: display none, or inside one
    if (cs.backgroundColor !== 'rgba(0, 0, 0, 0)') note(el, 'background', cs.backgroundColor);
    if (cs.backgroundImage !== 'none') note(el, 'background-image', cs.backgroundImage);
    const ownText = [...el.childNodes].some((n) => n.nodeType === 3 && n.textContent.trim());
    if (ownText) note(el, 'text', cs.color);
    for (const side of ['Top', 'Right', 'Bottom', 'Left']) {
      if (parseFloat(cs[`border${side}Width`]) > 0) note(el, 'border', cs[`border${side}Color`]);
    }
    if (cs.boxShadow !== 'none') note(el, 'shadow', cs.boxShadow);
    if (parseFloat(cs.outlineWidth) > 0) note(el, 'outline', cs.outlineColor);
    if (ownText && cs.textDecorationLine !== 'none') note(el, 'text-decoration', cs.textDecorationColor);
    for (const pseudo of ['::before', '::after']) {
      const ps = getComputedStyle(el, pseudo);
      if (ps.content === 'none' || ps.content === 'normal') continue;
      if (ps.backgroundColor !== 'rgba(0, 0, 0, 0)') note(el, `${pseudo} background`, ps.backgroundColor);
      if (ps.backgroundImage !== 'none') note(el, `${pseudo} background-image`, ps.backgroundImage);
    }
    if (el instanceof SVGElement) {
      if (cs.fill !== 'none') note(el, 'fill', cs.fill);
      if (cs.stroke !== 'none') note(el, 'stroke', cs.stroke);
    }
  }
  return [...found].slice(0, 5);
};

let browser;
try {
  browser = await chromium.launch();
} catch (e) {
  console.error(`cannot start Chromium: ${String(e.message).split('\n')[0]}\nrun once: npx playwright install chromium`);
  process.exit(1);
}
try {
  const page = await browser.newPage({ viewport: { width: 1280, height: 720 } });
  page.on('console', (m) => { if (m.type() === 'error') report(`console error: ${m.text()}`); });
  page.on('pageerror', (e) => report(`page error: ${e.message}`));
  page.on('requestfailed', (r) => report(`request failed: ${r.url()}`));
  page.on('response', (r) => { if (r.status() >= 400) report(`HTTP ${r.status()}: ${r.url()}`); });

  // The guide's per-slide rules, applied to one slide: root is the slide element.
  const guide = async (no, root) => {
    const { canvas, texts, overlaps, overflow, clipped, backgroundImage, columns } = await root.evaluate(measure);
    const k = canvas.w / 1280; // the px floors are for the 1280x720 canvas; scale them for other 16:9 sizes
    if (Math.abs(canvas.w / canvas.h - 16 / 9) > 0.01 * (16 / 9)) {
      report(`slide ${no}: canvas is ${Math.round(canvas.w)}x${Math.round(canvas.h)}, not 16:9`);
    }
    for (const t of texts) {
      const floor = CAPTION_FLOOR * k;
      if (t.size < floor) report(`slide ${no}: text below ${Math.round(floor)}px (${t.size}px): "${t.label}"`);
      // Transparent text (for example gradient text painted by background-clip) has no colour to check.
      if (!t.color) notes.add(`slide ${no}: transparent or gradient text is not contrast-checked`);
      const ratio = t.color ? contrast(rgbOf(t.color), rgbOf(t.bg)) : Infinity;
      if (ratio < 4.5) report(`slide ${no}: "${t.label}" contrast ${ratio.toFixed(2)}:1 is below 4.5:1`);
      if (t.l < 0.05 * canvas.w - 1 || t.r > 0.95 * canvas.w + 1 || t.t < 0.05 * canvas.h - 1 || t.b > 0.95 * canvas.h + 1) {
        report(`slide ${no}: "${t.label}" is outside the 5% safe area`);
      }
    }
    const captions = texts.filter((t) => t.size >= CAPTION_FLOOR * k && t.size < BODY_FLOOR * k).length;
    if (captions) notes.add(`slide ${no}: ${captions} text box(es) at ${CAPTION_FLOOR}-23px (below the 24px body size, guide section 11)`);
    if (backgroundImage !== 'none') notes.add(`slide ${no}: background is an image or gradient; the converter may draw it as a picture`);
    if (columns) notes.add(`slide ${no}: ${columns} multi-column element(s); the converter splits them into boxes, so check the column breaks in the render`);
    for (const pair of overlaps) report(`slide ${no}: overlapping text ${pair}`);
    for (const o of overflow) report(`slide ${no} overflow: ${o}`);
    for (const c of clipped) report(`slide ${no}: text clipped inside ${c}`);
    for (const c of await root.evaluate(colourIssues)) report(`slide ${no}: colour the converter cannot read (use hex, rgb or rgba): ${c}`);
  };

  await page.goto(url, { waitUntil: 'networkidle', timeout: 60000 });
  await page.evaluate(() => document.fonts.ready);
  await page.waitForTimeout(1200); // Tailwind, Lucide, ECharts and Mermaid draw after load
  // Mermaid draws its SVG asynchronously: wait until every diagram has one before measuring.
  await page.waitForFunction(() => [...document.querySelectorAll('pre.mermaid')].every((p) => p.querySelector('svg')), null, { timeout: 15000 })
    .catch(() => report('a Mermaid diagram was not drawn within 15 s'));
  const fences = await page.evaluate(() => document.querySelectorAll('code[class*="language-mermaid"]').length);
  if (fences) report(`${fences} Mermaid fence(s) left unconverted: run the injector`);
  if (mode === 'sections') {
    // Marp hides a fragment until it is revealed; show them all so every fragment is measured and screenshotted.
    await page.addStyleTag({ content: '[data-marpit-fragment] { visibility: visible !important; }' });
  }

  // Index of the slide Marp's viewer shows, among the Marp slide SVGs (-1 when none).
  const activeSlide = () => page.evaluate(() => [...document.querySelectorAll('svg[data-marpit-svg]')]
    .findIndex((s) => s.classList.contains('bespoke-marp-active')));
  const slides = await page.$$(mode === 'sections' ? 'section' : 'section.slide');
  if (expected && slides.length !== expected) report(`slide count ${slides.length}, expected ${expected}`);
  console.log(`slides found: ${slides.length}`);
  if (slides.length === 0) report(`no slides found for mode "${mode}" (check --mode and the file)`);

  for (let i = 0; i < slides.length; i++) {
    let root = slides[i];
    if (mode === 'sections') {
      // Marp's viewer shows one slide at a time (svg.bespoke-marp-active). ArrowRight may first reveal a fragment
      // (a list item marked data-marpit-fragment), so press until the wanted slide is the active one.
      for (let p = 0; i > 0 && p < 30 && (await activeSlide()) !== i; p++) {
        await page.keyboard.press('ArrowRight');
        await page.waitForTimeout(150);
      }
      if (i > 0 && (await activeSlide()) !== i) {
        report(`slide ${i + 1}: the viewer did not reach this slide after 30 ArrowRight presses`);
        break;
      }
      await page.waitForTimeout(350);
      root = await page.$('svg.bespoke-marp-active');
      if (!root) {
        report(`slide ${i + 1}: no active slide found`);
        continue;
      }
    }
    // Each slide is exactly the viewport's height: scroll it to the top and take a viewport screenshot.
    // (Element screenshots of the stacked page came out misaligned.)
    await root.evaluate((el) => el.scrollIntoView({ block: 'start' }));
    await page.waitForTimeout(200);
    const file = join(outDir, `slide-${pad(i + 1)}.png`);
    await page.screenshot({ path: file });
    shots.push(file);
    await guide(i + 1, root);
  }

  if (mode === 'slides' && slides.length > 1) {
    // Presenter (optional): ?present shows one slide at a time. The stacked page above is what the converter reads.
    const presentUrl = new URL(url);
    presentUrl.searchParams.set('present', '');
    await page.goto(presentUrl.href, { waitUntil: 'networkidle', timeout: 60000 });
    await page.waitForTimeout(800);
    const visible = () => page.evaluate(() => [...document.querySelectorAll('section.slide')]
      .map((s, i) => (getComputedStyle(s).visibility !== 'hidden' ? i : -1))
      .filter((i) => i >= 0));
    if (!(await page.evaluate(() => document.body.classList.contains('present')))) {
      notes.add('presenter: this deck has no ?present support (optional, not checked)');
    } else {
      let shown = await visible();
      if (shown.length !== 1 || shown[0] !== 0) report(`presenter: ?present should show only slide 1, shows ${JSON.stringify(shown)}`);
      const steps = Math.min(2, slides.length - 1);
      for (let s = 0; s < steps; s++) {
        await page.keyboard.press('ArrowRight');
        await page.waitForTimeout(150);
      }
      shown = await visible();
      if (shown.length !== 1 || shown[0] !== steps) report(`presenter: after ${steps} ArrowRight, slide ${steps + 1} should show, shows ${JSON.stringify(shown)}`);
      else console.log(`presenter: ?present steps with ArrowRight (${steps} step(s) checked)`);
    }
  }
} catch (e) {
  report(`run error: ${String(e.message).split('\n')[0]}`);
} finally {
  await browser.close();
}

for (const s of shots) console.log(`screenshot: ${s}`);
for (const n of notes) console.log(`note: ${n}`);
if (problems.length) {
  console.log(`FAILED (${problems.length}):`);
  for (const p of problems) console.log(`  - ${p}`);
  process.exit(1);
}
console.log('PASSED: no console errors or failed requests; overflow, clipped text, 16:9, background contrast, 5% safe area, text size and overlap checked');
