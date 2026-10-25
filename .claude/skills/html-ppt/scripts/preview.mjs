#!/usr/bin/env node
// Render each slide of an HTML deck to PNG so the draft can be reviewed/edited.
// usage: node preview.mjs <deck.html> [outDir=./preview] [slideNumber]
import { chromium } from 'playwright';
import { pathToFileURL } from 'node:url';
import { resolve, join } from 'node:path';
import { mkdirSync } from 'node:fs';

const [, , input, outDir = 'preview', only] = process.argv;
if (!input) { console.error('usage: node preview.mjs <deck.html> [outDir] [slideNumber]'); process.exit(1); }
mkdirSync(outDir, { recursive: true });

const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1280, height: 720 }, deviceScaleFactor: 1 });
await page.goto(pathToFileURL(resolve(input)).href, { waitUntil: 'networkidle' });
await page.evaluate(() => document.fonts.ready);
await page.addStyleTag({ content: 'section.slide{margin:0 !important;transform:none !important;box-shadow:none !important}' });
const els = await page.$$('section.slide');
for (let i = 0; i < els.length; i++) {
  if (only && +only !== i + 1) continue;
  const f = join(outDir, `slide-${String(i + 1).padStart(2, '0')}.png`);
  await els[i].screenshot({ path: f });
  console.log(f);
}
await browser.close();
