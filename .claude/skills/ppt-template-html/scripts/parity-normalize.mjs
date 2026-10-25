// Normalise and compare checker logs (check-deck.mjs and the ppt-style-html check-starter.mjs).
// usage:
//   node parity-normalize.mjs <log> --kind deck|starter [--detail] [--slides N]
//   node parity-normalize.mjs --compare <a.json> <b.json> <expected.json> --key slides10|slides9
//   node parity-normalize.mjs --compare <a.json> <expected.json> --key slides10|slides9
//   node parity-normalize.mjs --compare-detail <actual.json> <expected.json>
// Exit codes: 0 equal / printed, 1 sets differ, 2 usage or input error.
import { readFileSync } from 'node:fs';

const usage = `usage:
  node parity-normalize.mjs <log> --kind deck|starter [--detail] [--slides N]
  node parity-normalize.mjs --compare <a.json> <b.json> <expected.json> --key slides10|slides9
  node parity-normalize.mjs --compare <a.json> <expected.json> --key slides10|slides9
  node parity-normalize.mjs --compare-detail <actual.json> <expected.json>`;
const die = (msg, code = 2) => {
  console.error(`${msg}\n${usage}`);
  process.exit(code);
};

// check-starter problem lines ("  - <text>") -> (slide, rule). Slide 0 = deck-wide findings.
// Rule names are the shared vocabulary that check-deck.mjs also prints in its <rule> column.
const STARTER_DECK_WIDE = [
  [/^console error: /, 'console'],
  [/^page error: /, 'page-error'],
  [/^request failed: /, 'request-failed'],
  [/^HTTP \d+: /, 'http'],
  [/^a Mermaid diagram was not drawn/, 'mermaid'],
  [/^\d+ Mermaid fence\(s\) left unconverted/, 'mermaid-fence'],
  [/^slide count \d+, expected \d+/, 'slide-count'],
  [/^no slides found/, 'no-slides'],
  [/^run error: /, 'run-error'],
  [/^presenter: /, 'presenter'],
];
const STARTER_PER_SLIDE = [
  [/^slide (\d+): canvas is /, 'canvas'],
  [/^slide (\d+): text below /, 'text-size'],
  [/^slide (\d+): ".*" contrast [\d.]+:1 is below 4\.5:1$/, 'contrast'],
  [/^slide (\d+): ".*" is outside the 5% safe area$/, 'safe-area'],
  [/^slide (\d+): overlapping text /, 'overlap'],
  [/^slide (\d+) overflow: /, 'overflow'],
  [/^slide (\d+): text clipped inside /, 'clipped'],
  [/^slide (\d+): colour the converter cannot read/, 'colour'],
  [/^slide (\d+): the viewer did not reach/, 'viewer'],
  [/^slide (\d+): no active slide found/, 'viewer'],
];

const parseStarter = (text) => {
  const out = [];
  for (const line of text.split(/\r?\n/)) {
    const m = line.match(/^\s+- (.*)$/);
    if (!m) continue;
    const body = m[1];
    let hit = null;
    for (const [re, rule] of STARTER_DECK_WIDE) if (re.test(body)) hit = { slide: 0, rule };
    if (!hit) {
      for (const [re, rule] of STARTER_PER_SLIDE) {
        const s = body.match(re);
        if (s) { hit = { slide: Number(s[1]), rule }; break; }
      }
    }
    if (!hit) {
      console.error(`warning: unmapped check-starter line: ${body}`);
      continue;
    }
    out.push(hit);
  }
  return out;
};

// check-deck lines: <kind> | slide <n> | <rule> | <result> | <selector> | "<excerpt>"; split on the first five "|" only.
const parseDeck = (text) => {
  const out = [];
  let found = null;
  for (const line of text.split(/\r?\n/)) {
    const f = line.match(/^slides found: (\d+)/);
    if (f) found = Number(f[1]);
    const parts = [];
    let rest = line;
    for (let i = 0; i < 5; i++) {
      const at = rest.indexOf('|');
      if (at < 0) break;
      parts.push(rest.slice(0, at).trim());
      rest = rest.slice(at + 1);
    }
    if (parts.length < 4) continue;
    const slide = parts[1].match(/^slide (\d+)$/);
    if (!slide || !/^\S+$/.test(parts[0]) || !parts[2]) continue;
    out.push({ slide: Number(slide[1]), kind: parts[0], rule: parts[2], result: parts[3] });
  }
  return { rows: out, found };
};

const uniqueBy = (items, keyOf) => {
  const seen = new Map();
  for (const it of items) if (!seen.has(keyOf(it))) seen.set(keyOf(it), it);
  return [...seen.values()];
};
const bySlideRule = (a, b) => a.slide - b.slide || (a.rule < b.rule ? -1 : a.rule > b.rule ? 1 : 0);

const args = process.argv.slice(2);
const opt = (name) => (args.includes(name) ? args[args.indexOf(name) + 1] : undefined);
const readJson = (file) => {
  try {
    return JSON.parse(readFileSync(file, 'utf8').replace(/^﻿/, ''));
  } catch (e) {
    return die(`cannot read JSON ${file}: ${e.message}`);
  }
};
// Files after a flag, up to the next "--" argument.
const filesAfter = (flag) => {
  const out = [];
  for (let i = args.indexOf(flag) + 1; i < args.length && !args[i].startsWith('--'); i++) out.push(args[i]);
  return out;
};
const printDiff = (labelA, setA, labelB, setB) => {
  for (const k of [...setA].filter((x) => !setB.has(x)).sort()) console.log(`only in ${labelA}: ${k}`);
  for (const k of [...setB].filter((x) => !setA.has(x)).sort()) console.log(`only in ${labelB}: ${k}`);
};
const pairKey = (p) => `${p.slide}|${p.rule}`;

if (args.includes('--compare-detail')) {
  const files = filesAfter('--compare-detail');
  if (files.length !== 2) die('--compare-detail takes <actual.json> <expected.json>');
  const key = (p) => `${p.slide}|${p.kind}|${p.rule}|${p.result}`;
  const actual = new Set(readJson(files[0]).map(key));
  const expectedJson = readJson(files[1]);
  if (!Array.isArray(expectedJson)) die('expected file for --compare-detail must be an array');
  const expected = new Set(expectedJson.map(key));
  const same = actual.size === expected.size && [...actual].every((k) => expected.has(k));
  if (same) console.log(`detail sets equal (${actual.size} entries)`);
  else {
    console.log('detail sets differ (slide|kind|rule|result):');
    printDiff('actual', actual, 'expected', expected);
  }
  process.exit(same ? 0 : 1);
}

if (args.includes('--compare')) {
  const files = filesAfter('--compare');
  if (files.length !== 2 && files.length !== 3) die('--compare takes two or three JSON files');
  const key = opt('--key');
  const expectedFile = files[files.length - 1];
  let expected = readJson(expectedFile);
  if (!Array.isArray(expected)) {
    if (!key) die(`${expectedFile} is keyed; --key slides10|slides9 is required`);
    if (!Array.isArray(expected[key])) die(`${expectedFile} has no array under key "${key}"`);
    expected = expected[key];
  }
  const sets = files.slice(0, -1).map((f) => new Set(readJson(f).map(pairKey)));
  const exp = new Set(expected.map(pairKey));
  let ok = true;
  files.slice(0, -1).forEach((f, i) => {
    const same = sets[i].size === exp.size && [...sets[i]].every((k) => exp.has(k));
    if (!same) {
      ok = false;
      console.log(`${f} differs from expected (slide|rule):`);
      printDiff(f, sets[i], 'expected', exp);
    }
  });
  if (sets.length === 2) {
    const same = sets[0].size === sets[1].size && [...sets[0]].every((k) => sets[1].has(k));
    if (!same) {
      ok = false;
      console.log('the two checker results differ (slide|rule):');
      printDiff(files[0], sets[0], files[1], sets[1]);
    }
  }
  if (ok) console.log(`pair sets equal (${exp.size} pairs)`);
  process.exit(ok ? 0 : 1);
}

// Log normalisation mode.
const logFile = args.find((a, i) => !a.startsWith('--') && !['--kind', '--slides', '--key'].includes(args[i - 1]));
const kind = opt('--kind');
if (!logFile) die('missing <log> file');
if (!['deck', 'starter'].includes(kind)) die('--kind must be deck or starter');
const slidesArg = opt('--slides');
if (slidesArg !== undefined && !/^\d+$/.test(slidesArg)) die('--slides must be a whole number');
const detail = args.includes('--detail');
if (detail && kind !== 'deck') die('--detail works with --kind deck only');
let text;
try {
  text = readFileSync(logFile, 'utf8');
} catch (e) {
  die(`cannot read ${logFile}: ${e.message}`);
}

if (kind === 'starter') {
  console.log(JSON.stringify(uniqueBy(parseStarter(text), pairKey).sort(bySlideRule), null, 2));
} else {
  const { rows, found } = parseDeck(text);
  if (!detail) {
    const fails = rows.filter((r) => r.result === 'fail').map(({ slide, rule }) => ({ slide, rule }));
    console.log(JSON.stringify(uniqueBy(fails, pairKey).sort(bySlideRule), null, 2));
  } else {
    // N: --slides, else the "slides found: N" line, else the highest slide number seen.
    const n = slidesArg !== undefined ? Number(slidesArg) : found ?? Math.max(0, ...rows.map((r) => r.slide));
    const out = uniqueBy(rows, (r) => `${r.slide}|${r.kind}|${r.rule}|${r.result}`);
    for (let s = 1; s <= n; s++) {
      const hasContent = rows.some((r) => r.slide === s && r.kind === 'content' && ['fail', 'human-check'].includes(r.result));
      if (!hasContent) out.push({ slide: s, kind: 'content', rule: 'any', result: 'pass' });
    }
    out.sort((a, b) => a.slide - b.slide || (a.kind < b.kind ? -1 : a.kind > b.kind ? 1 : 0) || (a.rule < b.rule ? -1 : 1));
    console.log(JSON.stringify(out, null, 2));
  }
}
