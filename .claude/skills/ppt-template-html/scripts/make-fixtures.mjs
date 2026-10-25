// Builds the non-COM fixtures into ../fixtures. Run after make-standard-office.ps1.
// Usage: node make-fixtures.mjs
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const here = path.dirname(fileURLToPath(import.meta.url));
const skill = path.resolve(here, '..');
const root = path.resolve(skill, '../../..');
const fx = path.join(skill, 'fixtures');
const testDir = path.join(root, 'test');

fs.mkdirSync(path.join(fx, 'oversize'), { recursive: true });

const write = (rel, data) => {
  const p = path.join(fx, rel);
  fs.writeFileSync(p, data);
  console.log(`wrote ${rel} (${fs.statSync(p).size} bytes)`);
};

// 1. standard-office derived samples
const office = path.join(fx, 'standard-office.pptx');
if (!fs.existsSync(office)) {
  console.error('standard-office.pptx missing: run make-standard-office.ps1 first');
  process.exit(2);
}
const officeBytes = fs.readFileSync(office);
write('standard-office-copy.pptx', officeBytes);

const zip = await JSZip.loadAsync(officeBytes);
const themeFile = zip.file('ppt/theme/theme1.xml');
if (!themeFile) {
  console.error('ppt/theme/theme1.xml not found in standard-office.pptx');
  process.exit(2);
}
const themeXml = await themeFile.async('string');
const accentRe = /(<a:accent1>\s*<a:srgbClr val=")([0-9A-Fa-f]{6})(")/;
const m = accentRe.exec(themeXml);
if (!m) {
  console.error('accent1 srgbClr not found in theme1.xml');
  process.exit(2);
}
const newAccent = m[2].toUpperCase() === 'C0392B' ? '1E8449' : 'C0392B';
zip.file('ppt/theme/theme1.xml', themeXml.replace(accentRe, `$1${newAccent}$3`));
write('standard-office-accent.pptx', await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));

// 2. byte copies of the sample masters
for (const [src, dst] of [
  ['template-4x3.pptx', 'template-4x3-copy.pptx'],
  ['template-1610.pptx', 'template-1610-copy.pptx'],
  ['template-master.pptx', 'template-master-copy.pptx'],
]) {
  write(dst, fs.readFileSync(path.join(testDir, src)));
}

// 3. data.csv and data.xlsx: same 3x4 table, XLSX first sheet with a percent and a currency cell
write(
  'data.csv',
  ['Item,Share,Revenue,Count', 'Alpha,25.0%,"$1,200.00",1500', 'Beta,40.0%,"$3,450.50",820', ''].join('\n'),
);

const xlsx = new JSZip();
xlsx.file(
  '[Content_Types].xml',
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
    '</Types>',
);
xlsx.file(
  '_rels/.rels',
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '</Relationships>',
);
xlsx.file(
  'xl/workbook.xml',
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
);
xlsx.file(
  'xl/_rels/workbook.xml.rels',
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' +
    '</Relationships>',
);
// cellXfs: 0 general, 1 percent (0.0%), 2 currency ($#,##0.00)
xlsx.file(
  'xl/styles.xml',
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="2"><numFmt numFmtId="164" formatCode="0.0%"/><numFmt numFmtId="165" formatCode="&quot;$&quot;#,##0.00"/></numFmts>' +
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="3">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '</cellXfs></styleSheet>',
);
const strings = ['Item', 'Share', 'Revenue', 'Count', 'Alpha', 'Beta'];
xlsx.file(
  'xl/sharedStrings.xml',
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">` +
    strings.map((s) => `<si><t>${s}</t></si>`).join('') +
    '</sst>',
);
const s = (ref, i) => `<c r="${ref}" t="s"><v>${i}</v></c>`;
const n = (ref, v, style = 0) => `<c r="${ref}"${style ? ` s="${style}"` : ''}><v>${v}</v></c>`;
xlsx.file(
  'xl/worksheets/sheet1.xml',
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
    `<row r="1">${s('A1', 0)}${s('B1', 1)}${s('C1', 2)}${s('D1', 3)}</row>` +
    `<row r="2">${s('A2', 4)}${n('B2', 0.25, 1)}${n('C2', 1200, 2)}${n('D2', 1500)}</row>` +
    `<row r="3">${s('A3', 5)}${n('B3', 0.4, 1)}${n('C3', 3450.5, 2)}${n('D3', 820)}</row>` +
    '</sheetData></worksheet>',
);
write('data.xlsx', await xlsx.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));

// 4. oversize fixtures (input-safety limits: 50MB file, 2000 entries, 200MB uncompressed total)
const MB = 1024 * 1024;
const bigZip = new JSZip();
bigZip.file('big.bin', Buffer.alloc(51 * MB), { compression: 'STORE' });
write('oversize/too-big.pptx', await bigZip.generateAsync({ type: 'nodebuffer', compression: 'STORE' }));

const manyZip = new JSZip();
for (let i = 0; i < 2001; i++) manyZip.file(`f/${i}.txt`, 'x');
write('oversize/too-many.pptx', await manyZip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));

const bombZip = new JSZip();
bombZip.file('zeros.bin', Buffer.alloc(201 * MB), { compression: 'DEFLATE' });
write('oversize/bomb.pptx', await bombZip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
