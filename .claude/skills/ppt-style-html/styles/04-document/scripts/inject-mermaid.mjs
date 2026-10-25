// Marp turns ```mermaid fences into plain code blocks. This turns them into Mermaid diagrams in the built HTML.
// Labels are SVG text at 24px (htmlLabels is a root option in Mermaid 11; the flowchart option is ignored there).
// The injected script carries the marker below, so a rerun recognises it without looking at the deck text.
// The palette and font match styles/03-developer/index.html; keep them in step.
// usage: node scripts/inject-mermaid.mjs dist/deck.html
import { readFileSync, writeFileSync } from 'node:fs';

const [file] = process.argv.slice(2);
if (!file) {
  console.error('usage: node scripts/inject-mermaid.mjs <built-deck.html>');
  process.exit(2);
}

const MARKER = '<script type="module" data-mermaid>';
let html = readFileSync(file, 'utf8');
const pattern = /<pre[^>]*><code class="language-mermaid">([\s\S]*?)<\/code><\/pre>/g;
const count = [...html.matchAll(pattern)].length;
const converted = html.includes('<pre class="mermaid">');
if (count === 0 && !converted) {
  if (html.includes('language-mermaid')) {
    console.error('a mermaid fence is in the output but no block matched: check the Marp output format');
    process.exit(1);
  }
  console.log('no mermaid blocks: nothing to inject');
  process.exit(0);
}

const hasScript = html.includes(MARKER);
if (count === 0 && hasScript) {
  console.log('already converted: no mermaid fences left, script present');
  process.exit(0);
}
if (!html.includes('</body>')) {
  console.error('no </body> tag: refusing to write a deck without the Mermaid script');
  process.exit(1);
}

if (count > 0) html = html.replace(pattern, '<pre class="mermaid">$1</pre>');
if (!hasScript) {
  html = html.replace('</body>', `${MARKER}
import mermaid from 'https://cdn.jsdelivr.net/npm/mermaid@11.17.2/dist/mermaid.esm.min.mjs';
mermaid.initialize({ startOnLoad: false, htmlLabels: false, theme: 'base', themeVariables: { fontSize: '24px', fontFamily: 'Pretendard, "Malgun Gothic", sans-serif', primaryColor: '#e0f2fe', primaryBorderColor: '#0369a1', primaryTextColor: '#111111', lineColor: '#555555' } });
await document.fonts.load('24px Pretendard');
await mermaid.run();
</script>
</body>`);
}
writeFileSync(file, html);
console.log(`mermaid blocks: ${count}${hasScript ? '' : ', script added'}`);
