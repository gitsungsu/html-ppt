// Table readers for build-deck.mjs: CSV (RFC 4180) and the first sheet of an XLSX workbook.
// Every reader returns string[][] (rows of display strings). XLSX numbers are shown as percent, currency or
// grouped numbers when their number format says so (plan D6); dates and every other format stay raw.
import { readFile } from 'node:fs/promises';
import { extname } from 'node:path';
import JSZip from 'jszip';

export class InputError extends Error {}

// XLSX limits (plan AC7): file 50MB, 2000 entries, 200MB uncompressed in total.
const MAX_FILE = 50 * 1024 * 1024;
const MAX_ENTRIES = 2000;
const MAX_UNCOMPRESSED = 200 * 1024 * 1024;

// ---------- CSV ----------

export function parseCsv(text) {
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
  const rows = [];
  let row = [];
  let field = '';
  let quoted = false;
  let touched = false; // the current record has at least one character or separator
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; } else quoted = false;
      } else field += ch;
      continue;
    }
    if (ch === '"') { quoted = true; touched = true; }
    else if (ch === ',') { row.push(field); field = ''; touched = true; }
    else if (ch === '\r' || ch === '\n') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      if (touched || field) { row.push(field); rows.push(row); }
      row = []; field = ''; touched = false;
    } else { field += ch; touched = true; }
  }
  if (touched || field) { row.push(field); rows.push(row); }
  return rows;
}

// ---------- XLSX ----------

const unescapeXml = (s) => s
  .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
  .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)))
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

const attrsOf = (tag) => {
  const out = {};
  for (const m of tag.matchAll(/([\w:.-]+)\s*=\s*"([^"]*)"|([\w:.-]+)\s*=\s*'([^']*)'/g)) out[m[1] ?? m[3]] = unescapeXml(m[2] ?? m[4]);
  return out;
};

// Text of a shared string item or inline string: the <t> runs, without phonetic (<rPh>) runs.
const textOf = (xml) => [...xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '').matchAll(/<t\b[^>]*?(?:\/>|>([\s\S]*?)<\/t>)/g)]
  .map((m) => unescapeXml(m[1] ?? '')).join('');

const colIndex = (letters) => letters.toUpperCase().split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;

const group = (intPart) => intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');

// Built-in number formats (ECMA-376 18.8.30) that the plan formats: 3,4 grouped; 5-8 currency; 9,10 percent.
const BUILTIN = {
  3: { kind: 'number', grouped: true, min: 0, max: 0 },
  4: { kind: 'number', grouped: true, min: 2, max: 2 },
  5: { kind: 'currency', grouped: true, min: 0, max: 0, symbol: '$', prefix: true },
  6: { kind: 'currency', grouped: true, min: 0, max: 0, symbol: '$', prefix: true },
  7: { kind: 'currency', grouped: true, min: 2, max: 2, symbol: '$', prefix: true },
  8: { kind: 'currency', grouped: true, min: 2, max: 2, symbol: '$', prefix: true },
  9: { kind: 'percent', grouped: false, min: 0, max: 0 },
  10: { kind: 'percent', grouped: false, min: 2, max: 2 },
};

const SYMBOLS = '$₩€¤£¥';

// Classify a custom format code: returns a format description, or null when the value stays raw.
function parseFormatCode(code) {
  let section = '';
  let inQuote = false;
  for (let i = 0; i < code.length; i++) { // first section: up to an unquoted ';'
    const ch = code[i];
    if (ch === '"') inQuote = !inQuote;
    if (ch === ';' && !inQuote) break;
    section += ch;
  }
  if (/^\s*general\s*$/i.test(section)) return null;
  // Replace literal parts with markers so that letters inside them do not look like date or exponent codes.
  let symbol = null;
  let symbolAt = -1;
  let stripped = '';
  for (let i = 0; i < section.length; i++) {
    const ch = section[i];
    if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      const lit = section.slice(i + 1, end < 0 ? section.length : end);
      if (symbol === null && lit && [...lit].every((c) => SYMBOLS.includes(c))) { symbol = lit; symbolAt = stripped.length; }
      stripped += '\u0001';
      i = end < 0 ? section.length : end;
    } else if (ch === '[') {
      const end = section.indexOf(']', i);
      const inner = section.slice(i + 1, end < 0 ? section.length : end);
      const cur = /^\$([^-\]]*)/.exec(inner);
      if (cur && cur[1] && symbol === null) { symbol = cur[1]; symbolAt = stripped.length; stripped += '\u0001'; }
      else if (/^(h+|m+|s+)$/i.test(inner)) stripped += 'h'; // elapsed time [h] [mm] [ss]
      i = end < 0 ? section.length : end;
    } else if (ch === '\\' || ch === '_' || ch === '*') {
      i++; // escaped, padding or fill character
      stripped += '\u0001';
    } else if (SYMBOLS.includes(ch)) {
      if (symbol === null) { symbol = ch; symbolAt = stripped.length; }
      stripped += '\u0001';
    } else stripped += ch;
  }
  if (/[ymdhs]/i.test(stripped) || /e[+-]/i.test(stripped) || /\//.test(stripped)) return null; // date, time, scientific, fraction
  const firstDigit = stripped.search(/[0#?]/);
  if (firstDigit < 0) return null;
  const numeric = /[0#?,.]+/.exec(stripped.slice(firstDigit))[0];
  if (/,$/.test(numeric) && !/\./.test(numeric)) return null; // trailing comma scales by 1000: left raw
  const [intCode, decCode = ''] = numeric.split('.');
  const percent = stripped.includes('%');
  const grouped = /,/.test(intCode);
  const min = (decCode.match(/0/g) ?? []).length;
  const max = (decCode.match(/[0#?]/g) ?? []).length;
  if (percent) return { kind: 'percent', grouped, min, max };
  if (symbol) return { kind: 'currency', grouped, min, max, symbol, prefix: symbolAt <= firstDigit };
  if (grouped) return { kind: 'number', grouped, min, max };
  return null;
}

function formatNumber(raw, fmt) {
  const n = Number(raw);
  if (!fmt || !Number.isFinite(n)) return raw;
  const value = fmt.kind === 'percent' ? n * 100 : n;
  let fixed = Math.abs(value).toFixed(fmt.max);
  if (fmt.max > fmt.min && fixed.includes('.')) { // optional decimals: drop trailing zeros down to the minimum
    let [i, d] = fixed.split('.');
    while (d.length > fmt.min && d.endsWith('0')) d = d.slice(0, -1);
    fixed = d ? `${i}.${d}` : i;
  }
  let [intPart, decPart] = fixed.split('.');
  if (fmt.grouped) intPart = group(intPart);
  let out = decPart ? `${intPart}.${decPart}` : intPart;
  if (fmt.kind === 'percent') out += '%';
  else if (fmt.kind === 'currency') out = fmt.prefix ? fmt.symbol + out : `${out}${fmt.symbol}`;
  const negative = value < 0 && Number(fixed.replace(/,/g, '')) !== 0;
  return negative ? `-${out}` : out;
}

export async function parseXlsx(buffer) {
  if (buffer.length > MAX_FILE) throw new InputError(`xlsx larger than ${MAX_FILE} bytes`);
  let zip;
  try { zip = await JSZip.loadAsync(buffer); } catch (e) { throw new InputError(`not a readable xlsx: ${e.message}`); }
  const names = Object.keys(zip.files);
  if (names.length > MAX_ENTRIES) throw new InputError(`xlsx has more than ${MAX_ENTRIES} entries`);
  let total = 0;
  for (const f of Object.values(zip.files)) total += f._data?.uncompressedSize ?? 0;
  if (total > MAX_UNCOMPRESSED) throw new InputError(`xlsx uncompressed size exceeds ${MAX_UNCOMPRESSED} bytes`);

  const read = async (path) => {
    const f = zip.file(path);
    return f ? f.async('string') : null;
  };

  // First sheet: xl/workbook.xml -> r:id -> xl/_rels/workbook.xml.rels -> target.
  const workbook = await read('xl/workbook.xml');
  if (!workbook) throw new InputError('xlsx has no xl/workbook.xml');
  const sheetTag = /<sheet\b[^>]*>/.exec(workbook);
  if (!sheetTag) throw new InputError('xlsx has no sheet');
  const sheetAttrs = attrsOf(sheetTag[0]);
  const rid = sheetAttrs['r:id'] ?? Object.entries(sheetAttrs).find(([k]) => k.endsWith(':id'))?.[1];
  const rels = (await read('xl/_rels/workbook.xml.rels')) ?? '';
  let target = null;
  for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    const a = attrsOf(m[0]);
    if (a.Id === rid) { target = a.Target; break; }
  }
  if (!target) throw new InputError('xlsx first sheet relationship not found');
  const sheetPath = target.startsWith('/') ? target.slice(1) : `xl/${target}`.replace(/\/\.\//g, '/');
  const sheet = await read(sheetPath);
  if (!sheet) throw new InputError(`xlsx sheet part missing: ${sheetPath}`);

  const sstXml = await read('xl/sharedStrings.xml');
  const shared = sstXml ? [...sstXml.matchAll(/<si\b[^>]*?(?:\/>|>([\s\S]*?)<\/si>)/g)].map((m) => textOf(m[1] ?? '')) : [];

  // Number formats: custom codes by id, then the xf index (cell style) -> format.
  const stylesXml = await read('xl/styles.xml');
  const custom = new Map();
  const xfFormat = [];
  if (stylesXml) {
    for (const m of stylesXml.matchAll(/<numFmt\b[^>]*>/g)) {
      const a = attrsOf(m[0]);
      if (a.numFmtId !== undefined) custom.set(Number(a.numFmtId), a.formatCode ?? '');
    }
    const xfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml);
    if (xfs) {
      for (const m of xfs[1].matchAll(/<xf\b[^>]*>/g)) {
        const id = Number(attrsOf(m[0]).numFmtId ?? 0);
        xfFormat.push(custom.has(id) ? parseFormatCode(custom.get(id)) : (BUILTIN[id] ?? null));
      }
    }
  }

  const grid = [];
  let rowNo = -1;
  for (const rm of sheet.matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
    const rAttr = attrsOf(rm[1]).r;
    rowNo = rAttr ? Number(rAttr) - 1 : rowNo + 1;
    let colNo = -1;
    for (const cm of (rm[2] ?? '').matchAll(/<c\b([^>]*?)\/>|<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
      const a = attrsOf(cm[1] ?? cm[2]);
      const ref = /^([A-Za-z]+)(\d+)$/.exec(a.r ?? '');
      colNo = ref ? colIndex(ref[1]) : colNo + 1;
      const body = cm[3] ?? '';
      const v = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(body)?.[1];
      let text = '';
      if (a.t === 's') text = shared[Number(unescapeXml(v ?? ''))] ?? '';
      else if (a.t === 'inlineStr') text = textOf(/<is\b[^>]*>([\s\S]*?)<\/is>/.exec(body)?.[1] ?? '');
      else if (a.t === 'b') text = v === '1' ? 'TRUE' : 'FALSE';
      else if (v !== undefined) {
        text = unescapeXml(v);
        if (a.t === undefined || a.t === 'n') text = formatNumber(text, xfFormat[Number(a.s ?? 0)]);
      }
      (grid[rowNo] ??= [])[colNo] = text;
    }
  }
  const width = Math.max(0, ...grid.map((r) => (r ? r.length : 0)));
  const rows = [];
  for (const r of grid) if (r) rows.push(Array.from({ length: width }, (_, i) => r[i] ?? ''));
  return rows.filter((r) => r.some((c) => c !== ''));
}

// Load a table file by extension. Throws InputError for a missing or unsupported file.
export async function loadTable(path) {
  const ext = extname(path).toLowerCase();
  let data;
  try { data = await readFile(path); } catch { throw new InputError(`data file not found: ${path}`); }
  if (ext === '.csv') return parseCsv(data.toString('utf8'));
  if (ext === '.xlsx') return parseXlsx(data);
  throw new InputError(`unsupported data file type (use .csv or .xlsx): ${path}`);
}
