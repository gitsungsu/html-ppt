// Render-check a generated HTML deck. Started as a copy of ../../ppt-style-html/scripts/check-starter.mjs (same rules),
// plus two layers: (1) data-tpl attribution (template decoration "chrome" vs "content"), (2) a background layer taken
// from the browser's real paint order (document.elementsFromPoint) instead of the ancestor walk.
// usage: node check-deck.mjs <deck.html> <outDir> [--slides N] [--text-size warn|fail]
//   --text-size warn: text-size findings on content print result "warn" and do not count as content fails (default fail)
// stdout: one line per finding, nothing for passes:
//   <kind> | slide <n> | <rule> | <result> | <selector> | "<excerpt>"
//   kind: content | chrome   result: fail | warn | human-check | reported  slide 0 = deck-level   selector indices are 0-based
//   "note: ..." lines are notes, not findings.
// exit: 0 no content fail, 1 at least one content fail, 2 usage or input error. <outDir>/report.json repeats the findings.
// Playwright lives in this folder's node_modules: cd scripts && npm install && npx playwright install chromium
import { createRequire } from 'node:module';
import { mkdirSync, existsSync, writeFileSync } from 'node:fs';
import { resolve, join } from 'node:path';
import { pathToFileURL } from 'node:url';

const { chromium } = createRequire(import.meta.url)('playwright');

const args = process.argv.slice(2);
const opt = (name, def) => (args.includes(name) ? args[args.indexOf(name) + 1] : def);
const usage = 'usage: node check-deck.mjs <deck.html> <outDir> [--slides N] [--text-size warn|fail]';
const fail2 = (msg) => {
  console.error(`${msg}\n${usage}`);
  process.exit(2);
};
const unknown = args.find((a) => a.startsWith('--') && a !== '--slides' && a !== '--text-size');
if (unknown) fail2(`unknown option "${unknown}"`);
// Positionals are the arguments that are neither options nor option values.
const positional = args.filter((a, i) => !a.startsWith('--') && !args[i - 1]?.startsWith('--'));
const [target, outDir] = positional;
const slidesArg = opt('--slides', '0');
if (!target || !outDir || positional.length !== 2) fail2('expected <deck.html> and <outDir>');
if (!/^\d+$/.test(String(slidesArg))) fail2(`--slides must be a whole number, got "${slidesArg}"`);
const textSizeArg = opt('--text-size', 'fail');
if (textSizeArg !== 'warn' && textSizeArg !== 'fail') fail2(`--text-size must be warn or fail, got "${textSizeArg}"`);
const textSizeResult = textSizeArg === 'warn' ? 'warn' : undefined; // warn: text-size findings on content are printed but not counted
const isUrl = /^https?:/.test(target);
if (!isUrl && !existsSync(resolve(target))) fail2(`file not found: ${target}`);
const expected = Number(slidesArg);
const url = isUrl ? target : pathToFileURL(resolve(target)).href;
mkdirSync(outDir, { recursive: true });

const findings = [];
const notes = new Set();
const clean = (s) => String(s).replace(/\s+/g, ' ').replace(/\|/g, '/').replace(/"/g, "'");
// detail keeps the full message for report.json; the stdout line only carries the short excerpt.
const add = (kind, slide, rule, result, selector, excerpt, detail) => findings.push({ kind, slide, rule, result, selector, excerpt: clean(excerpt ?? ''), detail });
// Deck-level problem (slide 0): always a content failure.
const deck = (rule, message) => add('content', 0, rule, 'fail', '-', message, message);
// A rule broken by one element: chrome failures are only reported (plan D5), content failures count.
const sel = (slide, ref) => (ref.idx < 0 ? `section.slide[${slide - 1}]` : `section.slide[${slide - 1}] *[${ref.idx}] ${ref.tag}`);
const onEl = (slide, rule, ref, detail, result) => add(ref.kind, slide, rule, result ?? (ref.kind === 'chrome' ? 'reported' : 'fail'), sel(slide, ref), ref.text, detail);

const CAPTION_FLOOR = 16; // px = 12pt, the smallest text size; guide section 11 asks for 19px captions, relaxed here
const BODY_FLOOR = 24;    // px = 18pt, the smallest body size in guide section 11; 19-23px is noted, not failed
const KEEP = 5;           // overflow and clipped findings kept per slide
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
// Every element is described by ref(): its index among root.querySelectorAll('*'), tag, kind (chrome when inside [data-tpl]) and text.
const measure = (root) => {
  const c = root.getBoundingClientRect();
  const all = [...root.querySelectorAll('*')];
  const idxOf = new Map(all.map((e, i) => [e, i]));
  const ref = (el) => ({
    idx: el === root ? -1 : idxOf.get(el),
    tag: el.tagName.toLowerCase(),
    desc: el.tagName.toLowerCase() + (typeof el.className === 'string' && el.className ? '.' + el.className.split(' ')[0] : ''),
    kind: el.closest('[data-tpl]') ? 'chrome' : 'content',
    text: (el.textContent || '').trim().replace(/\s+/g, ' ').slice(0, 24),
  });

  // Alpha of a computed colour: 0..1, or null when the string is not rgb()/rgba() (for example oklch()).
  const alphaOf = (s) => {
    if (s === 'transparent') return 0;
    const m = s.match(/^rgba?\(([^)]*)\)$/);
    if (!m) return null;
    const p = m[1].split(/[,/ ]+/).filter(Boolean);
    if (p.length < 4) return 1;
    return p[3].endsWith('%') ? parseFloat(p[3]) / 100 : parseFloat(p[3]);
  };
  const slideBg = alphaOf(getComputedStyle(root).backgroundColor) > 0 ? getComputedStyle(root).backgroundColor : 'rgb(255, 255, 255)';
  // The layer a text element E sits on at one point: the first element of document.elementsFromPoint (paint order,
  // top to bottom), from E onward and E included, that paints a background.
  //   undefined  E (or a descendant) is not at this point: drop the point
  //   null       nothing paints: the slide background applies
  //   {human}    picture, gradient, translucent or replaced element: a human decides
  //   {color}    opaque solid colour
  const layerAt = (E, x, y) => {
    const list = document.elementsFromPoint(x, y);
    let i = list.indexOf(E);
    if (i < 0) i = list.findIndex((n) => E.contains(n));
    if (i < 0) return undefined;
    for (const n of list.slice(i)) {
      const cs = getComputedStyle(n);
      const a = alphaOf(cs.backgroundColor);
      const replaced = ['img', 'svg', 'canvas', 'video'].includes(n.localName);
      if (!(a > 0 || a === null) && cs.backgroundImage === 'none' && !replaced) continue;
      if (replaced || cs.backgroundImage !== 'none' || a === null || a < 1) return { human: true };
      for (let m = n; m; m = m.parentElement) if (parseFloat(getComputedStyle(m).opacity) !== 1) return { human: true };
      return { color: cs.backgroundColor };
    }
    return null;
  };
  // Sample points of E: the four corners (inset) and the centre of every rect of E's own text nodes.
  const range = document.createRange();
  const pointsOf = (el) => {
    const pts = [];
    for (const n of el.childNodes) {
      if (n.nodeType !== 3 || !n.textContent.trim()) continue;
      range.selectNodeContents(n);
      for (const r of range.getClientRects()) {
        if (r.width <= 0 || r.height <= 0) continue;
        const dx = Math.min(2, r.width / 4);
        const dy = Math.min(2, r.height / 4);
        pts.push([r.left + dx, r.top + dy], [r.right - dx, r.top + dy], [r.left + dx, r.bottom - dy], [r.right - dx, r.bottom - dy], [(r.left + r.right) / 2, (r.top + r.bottom) / 2]);
      }
    }
    return pts;
  };
  // Opaque solid colours under E, and whether any point needs a human.
  const backdrop = (el) => {
    const bgs = new Set();
    let human = false;
    let kept = 0;
    for (const [x, y] of pointsOf(el)) {
      if (x < 0 || y < 0 || x >= window.innerWidth || y >= window.innerHeight) continue;
      if (x < c.left || x > c.right || y < c.top || y > c.bottom) continue;
      const layer = layerAt(el, x, y);
      if (layer === undefined) continue;
      kept++;
      if (layer === null) bgs.add(slideBg);
      else if (layer.human) human = true;
      else bgs.add(layer.color);
    }
    if (!kept) bgs.add(slideBg);
    return { bgs: [...bgs], human };
  };

  const boxes = [];
  for (const el of all) {
    if (el.namespaceURI !== 'http://www.w3.org/1999/xhtml') continue; // HTML text only; chart text inside SVG is not measured
    if (![...el.childNodes].some((n) => n.nodeType === 3 && n.textContent.trim())) continue;
    const cs = getComputedStyle(el);
    if (cs.display === 'none' || cs.visibility === 'hidden') continue;
    const r = el.getBoundingClientRect();
    if (r.width === 0 || r.height === 0) continue;
    boxes.push({ el, ref: ref(el), size: parseFloat(cs.fontSize), color: /^rgba\(.*,\s*0\)$|^transparent$/.test(cs.color) ? null : cs.color, ...backdrop(el), l: r.left, t: r.top, r: r.right, b: r.bottom });
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
      if (w * h > 0.2 * smaller) overlaps.push({ a: a.ref, b: b.ref });
    }
  }
  const overflow = [];
  for (const d of all) {
    const b = d.getBoundingClientRect();
    if (b.width === 0 || b.height === 0) continue;
    if (b.right > c.right + 2 || b.bottom > c.bottom + 2 || b.left < c.left - 2 || b.top < c.top - 2) overflow.push(ref(d));
  }
  // Boxes that hide their overflow (overflow: hidden or clip) and hold more content than they show: the text is clipped.
  const clipped = [];
  for (const d of all) {
    const cs = getComputedStyle(d);
    if (!/hidden|clip/.test(cs.overflowX + ' ' + cs.overflowY) || d.clientHeight === 0) continue;
    if (d.scrollHeight > d.clientHeight + 2 || d.scrollWidth > d.clientWidth + 2) clipped.push(ref(d));
  }
  return {
    canvas: { w: c.width, h: c.height },
    backgroundImage: getComputedStyle(root).backgroundImage,
    texts: boxes.map((t) => ({ ref: t.ref, size: t.size, color: t.color, bgs: t.bgs, human: t.human, l: t.l - c.left, t: t.t - c.top, r: t.r - c.left, b: t.b - c.top })),
    overlaps,
    overflow,
    clipped,
    columns: all.filter((d) => !['auto', '1'].includes(String(getComputedStyle(d).columnCount))).length,
  };
};

// Runs in the page on a slide's root element. Lists colours the html-ppt converter cannot read (oklch and other
// modern colour functions). Tailwind v4 palette classes compute to oklch(); the converter then drops those colours.
const colourIssues = (root) => {
  const all = [...root.querySelectorAll('*')];
  const modern = /\b(oklch|oklab|lab|lch|color|color-mix)\(/;
  const found = new Map();
  const note = (el, label, value) => {
    if (value && modern.test(value)) {
      const cls = (el.getAttribute('class') || '').split(' ')[0];
      const key = `${el.tagName.toLowerCase()}${cls ? '.' + cls : ''} ${label} ${value}`;
      if (found.has(key)) return;
      found.set(key, {
        key,
        idx: el === root ? -1 : all.indexOf(el),
        tag: el.tagName.toLowerCase(),
        kind: el.closest('[data-tpl]') ? 'chrome' : 'content',
        text: (el.textContent || '').trim().replace(/\s+/g, ' ').slice(0, 24),
      });
    }
  };
  for (const el of [root, ...all]) {
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
  return [...found.values()].slice(0, 5);
};

let browser;
try {
  browser = await chromium.launch();
} catch (e) {
  console.error(`cannot start Chromium: ${String(e.message).split('\n')[0]}\nrun once: npx playwright install chromium`);
  process.exit(2);
}
let slideCount = 0;
try {
  const page = await browser.newPage({ viewport: { width: 1280, height: 720 } });
  page.on('console', (m) => { if (m.type() === 'error') deck('console', `console error: ${m.text()}`); });
  page.on('pageerror', (e) => deck('page-error', `page error: ${e.message}`));
  page.on('requestfailed', (r) => deck('request-failed', `request failed: ${r.url()}`));
  page.on('response', (r) => { if (r.status() >= 400) deck('http', `HTTP ${r.status()}: ${r.url()}`); });

  // The guide's per-slide rules, applied to one slide: root is the slide element.
  const guide = async (no, root) => {
    const { canvas, texts, overlaps, overflow, clipped, backgroundImage, columns } = await root.evaluate(measure);
    const k = canvas.w / 1280; // the px floors are for the 1280x720 canvas; scale them for other 16:9 sizes
    const slideRef = { idx: -1, tag: 'section', kind: 'content', text: '' };
    if (Math.abs(canvas.w / canvas.h - 16 / 9) > 0.01 * (16 / 9)) {
      onEl(no, 'canvas', { ...slideRef, text: `${Math.round(canvas.w)}x${Math.round(canvas.h)}` }, `slide ${no}: canvas is ${Math.round(canvas.w)}x${Math.round(canvas.h)}, not 16:9`);
    }
    for (const t of texts) {
      const floor = CAPTION_FLOOR * k;
      if (t.size < floor) onEl(no, 'text-size', t.ref, `slide ${no}: text below ${Math.round(floor)}px (${t.size}px): "${t.ref.text}"`, t.ref.kind === 'content' ? textSizeResult : undefined);
      // Transparent text (for example gradient text painted by background-clip) has no colour to check.
      if (!t.color) notes.add(`slide ${no}: transparent or gradient text is not contrast-checked`);
      // Picture, gradient, translucent or replaced layer under any sample point: a human decides.
      if (t.human) onEl(no, 'contrast', t.ref, `slide ${no}: "${t.ref.text}" sits on a picture, gradient, translucent or replaced layer: check by eye`, 'human-check');
      // Opaque solid points: the lowest contrast counts.
      const ratio = t.color && t.bgs.length ? Math.min(...t.bgs.map((bg) => contrast(rgbOf(t.color), rgbOf(bg)))) : Infinity;
      if (ratio < 4.5) onEl(no, 'contrast', t.ref, `slide ${no}: "${t.ref.text}" contrast ${ratio.toFixed(2)}:1 is below 4.5:1`);
      if (t.l < 0.05 * canvas.w - 1 || t.r > 0.95 * canvas.w + 1 || t.t < 0.05 * canvas.h - 1 || t.b > 0.95 * canvas.h + 1) {
        onEl(no, 'safe-area', t.ref, `slide ${no}: "${t.ref.text}" is outside the 5% safe area`);
      }
    }
    const captions = texts.filter((t) => t.size >= CAPTION_FLOOR * k && t.size < BODY_FLOOR * k).length;
    if (captions) notes.add(`slide ${no}: ${captions} text box(es) at ${CAPTION_FLOOR}-23px (below the 24px body size, guide section 11)`);
    if (backgroundImage !== 'none') notes.add(`slide ${no}: background is an image or gradient; the converter may draw it as a picture`);
    if (columns) notes.add(`slide ${no}: ${columns} multi-column element(s); the converter splits them into boxes, so check the column breaks in the render`);
    for (const { a, b } of overlaps) {
      // One side content makes it a content failure; two chrome boxes are only reported.
      const main = a.kind === 'content' ? a : b.kind === 'content' ? b : a;
      const other = main === a ? b : a;
      onEl(no, 'overlap', { ...main, kind: a.kind === 'content' || b.kind === 'content' ? 'content' : 'chrome', text: `${main.text} / ${other.text}` }, `slide ${no}: overlapping text "${a.text}" / "${b.text}"`);
    }
    for (const o of overflow.slice(0, KEEP)) onEl(no, 'overflow', o, `slide ${no} overflow: ${o.desc}`);
    if (overflow.length > KEEP) notes.add(`slide ${no}: ${overflow.length - KEEP} more overflow finding(s) not listed`);
    for (const c of clipped.slice(0, KEEP)) onEl(no, 'clipped', c, `slide ${no}: text clipped inside ${c.desc}`);
    if (clipped.length > KEEP) notes.add(`slide ${no}: ${clipped.length - KEEP} more clipped finding(s) not listed`);
    for (const c of await root.evaluate(colourIssues)) onEl(no, 'colour', c, `slide ${no}: colour the converter cannot read (use hex, rgb or rgba): ${c.key}`);
  };

  await page.goto(url, { waitUntil: 'networkidle', timeout: 60000 });
  await page.evaluate(() => document.fonts.ready);
  await page.waitForTimeout(1200); // Tailwind, Lucide, ECharts and Mermaid draw after load
  // Mermaid draws its SVG asynchronously: wait until every diagram has one before measuring.
  await page.waitForFunction(() => [...document.querySelectorAll('pre.mermaid')].every((p) => p.querySelector('svg')), null, { timeout: 15000 })
    .catch(() => deck('mermaid', 'a Mermaid diagram was not drawn within 15 s'));
  const fences = await page.evaluate(() => document.querySelectorAll('code[class*="language-mermaid"]').length);
  if (fences) deck('mermaid-fence', `${fences} Mermaid fence(s) left unconverted: run the injector`);
  // Background hit-testing must see every element, also the ones a deck made click-through.
  await page.addStyleTag({ content: '*{pointer-events:auto !important}' });

  const slides = await page.$$('section.slide');
  slideCount = slides.length;
  if (expected && slides.length !== expected) deck('slide-count', `slide count ${slides.length}, expected ${expected}`);
  if (slides.length === 0) deck('no-slides', 'no slides found (check the file)');

  for (let i = 0; i < slides.length; i++) {
    const root = slides[i];
    // Each slide is exactly the viewport's height: scroll it to the top and take a viewport screenshot.
    // (Element screenshots of the stacked page came out misaligned.)
    await root.evaluate((el) => el.scrollIntoView({ block: 'start' }));
    await page.waitForTimeout(200);
    const file = join(outDir, `slide-${pad(i + 1)}.png`);
    await page.screenshot({ path: file });
    shots.push(file);
    await guide(i + 1, root);
  }

  if (slides.length > 1) {
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
      if (shown.length !== 1 || shown[0] !== 0) deck('presenter', `presenter: ?present should show only slide 1, shows ${JSON.stringify(shown)}`);
      const steps = Math.min(2, slides.length - 1);
      for (let s = 0; s < steps; s++) {
        await page.keyboard.press('ArrowRight');
        await page.waitForTimeout(150);
      }
      shown = await visible();
      if (shown.length !== 1 || shown[0] !== steps) deck('presenter', `presenter: after ${steps} ArrowRight, slide ${steps + 1} should show, shows ${JSON.stringify(shown)}`);
    }
  }
} catch (e) {
  deck('run-error', `run error: ${String(e.message).split('\n')[0]}`);
} finally {
  await browser.close();
}

findings.sort((a, b) => a.slide - b.slide); // stable: deck-level first, then slide by slide in discovery order
for (const f of findings) console.log(`${f.kind} | slide ${f.slide} | ${f.rule} | ${f.result} | ${f.selector} | "${f.excerpt}"`);
console.log(`slides found: ${slideCount}`);
for (const n of notes) console.log(`note: ${n}`);
const contentFails = findings.filter((f) => f.kind === 'content' && f.result === 'fail').length;
const exitCode = contentFails ? 1 : 0;
writeFileSync(join(outDir, 'report.json'), JSON.stringify({ file: target, expectedSlides: expected, slidesFound: slideCount, contentFails, exitCode, findings, notes: [...notes], screenshots: shots }, null, 2));
process.exit(exitCode);
