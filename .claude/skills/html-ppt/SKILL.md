---
name: html-ppt
description: Draft a presentation as an HTML deck (1280x720 slides), preview and iteratively revise it, then export a native, editable PowerPoint (.pptx). Use when the user asks for a PPT/슬라이드/발표자료 시안, wants to edit or tweak an existing HTML deck, or wants the final .pptx generated from it.
---

# html-ppt: HTML 시안 → 수정 → 네이티브 PPTX

HTML is the source of truth. Slides are drafted and revised as HTML, previewed as PNG, and finally converted to a real `.pptx` whose text, shapes and tables are editable in PowerPoint (not screenshots).

Scripts live in `${CLAUDE_SKILL_DIR}/scripts` (Node ≥18, `playwright` + `pptxgenjs`). First run in a fresh checkout: `cd scripts && npm i && npx playwright install chromium`.

## Workflow

1. **Brief** — if missing, settle topic, audience, slide count, tone, language. Don't over-ask; pick sensible defaults.
2. **Draft** — copy `${CLAUDE_SKILL_DIR}/templates/deck.html` to the project (e.g. `deck.html`) and rewrite the slides. Use `frontend-design` thinking for a distinctive look; one idea per slide.
3. **Preview** — `node scripts/preview.mjs deck.html preview/` → `preview/slide-NN.png`. Read the PNGs and check overflow, overlap, contrast, density. Pass a slide number as 3rd arg to render one slide.
4. **Revise** — edit the HTML (Edit tool, small diffs), re-preview only the changed slide. Repeat until the user approves. Keep the user's edits; never regenerate the whole deck for a small request.
5. **Export** — `node scripts/build_pptx.mjs deck.html out.pptx [--template master.pptx] [--hard-wrap] [--font-map "A=B"] [--no-embed-fonts] [--no-svg] [--settle ms]`. Needs `npm i` in `scripts/` (playwright, pptxgenjs, jszip, wawoff2, opentype.js). Output reports slide/text/shape/picture counts.
6. **Verify (Windows + PowerPoint)** — `powershell -File scripts/render_pptx.ps1 out.pptx pptx-check/` exports each slide to PNG at its own aspect ratio; compare with the previews and fix HTML (not the pptx) if something drifted, then re-export.

## Authoring contract (what converts cleanly)

- Each slide is `<section class="slide">` at exactly **1280×720 px**; slides are stacked in the page. Keep the `section.slide` CSS from the template.
- **Text**: any element whose children are text/inline tags (`span b strong em a br`) becomes ONE text box with runs (font, size, weight, italic, underline, color, letter-spacing, alignment, line-height, uppercase). `ul/ol > li` get `•` / `1.` prefixes. Horizontal margins on inline elements (e.g. `small { margin-left: 6px }`) become spaces of about that width.
- **Boxes**: `background-color`, uniform `border`, `border-radius` (incl. circles), per-side borders (e.g. table row lines) become native shapes. Tables, flex and grid work because geometry comes from the rendered layout.
- **Pictures**: `<img>`, `<canvas>`, `<video>`, `background-image: url()`, anything marked `data-ppt="raster"`, rasterized at 2x. An SVG with `foreignObject` is split: its vector part stays a vector, and the HTML inside is converted to native text and shapes. An SVG with external images or unsupported filter primitives becomes a PNG (`--no-svg` forces PNG). Other inline `<svg>` that is not decomposed (see Charts) is embedded as a **vector SVG**; `feDropShadow`/`flood-opacity` are rewritten into primitives PowerPoint draws correctly.
- **Charts (SVG → native shapes)**: an inline `<svg>` built from `rect`, `circle`, `ellipse`, `line`, `polyline`, `polygon`, `path` and `text` is split into separate objects: each shape or path becomes a freeform (curves kept exact) with its own fill, stroke and gradient; `text` becomes a text box aligned by `text-anchor`; `stroke-dasharray` is cut into one freeform per dash, so every dash is editable. An SVG that uses `clipPath`, `mask`, `pattern`, `filter`, `image`, `use`, `symbol`, `marker`, `switch` or `textPath` stays a vector picture.
- **Gradients**: `linear-gradient` / `radial-gradient` (with alpha stops) on any box, on the slide background, and as live text fill (`background-clip:text`) become native gradient fills; `inset` box-shadow becomes a native inner shadow. Text on a faded (`opacity<1`) filled box is blended against the slide/gradient backdrop, so it fades together with the box. Gradient text on an inline `<span>` (`background-clip:text`) is applied per run, with the gradient stops mapped onto the span's own width (CSS paints over the span, PowerPoint over the text box).
- **Shapes**: per-corner `border-radius` (custom geometry), `-webkit-text-stroke` (text outline), `filter: drop-shadow()` (native shadow); a child flush with a rounded `overflow:hidden` parent inherits that corner radius, and pictures that poke out of the rounded corner are cut with the clip applied.
- **Transforms**: `rotate()` / `scale()` / `scaleX`/`scaleY` / `skew()` / `translate()` on any element, including groups with block children, are applied to every shape and text box inside (children stay editable). Skewed boxes become polygons; stretched single-line text is warped to the stretched size (PowerPoint "plain" text warp). Skewed text is converted to outlines so its slant is exact, not approximated as italic.
- **Filters**: `blur()`, `grayscale/sepia/saturate/hue-rotate/brightness/contrast/invert/opacity()` and `drop-shadow()` stay native (colours of the whole subtree are recomputed, text included). `filter: url(#id)` is native when the referenced SVG `<filter>` is a colour pipeline (`feColorMatrix` matrix/saturate/hueRotate/luminanceToAlpha, `feComponentTransfer` identity/linear/gamma/table/discrete per channel, any chain of them, linearRGB honoured), a single `feGaussianBlur`, or a drop shadow (`feDropShadow`, or the spelled-out blur–offset–flood–composite–merge chain). `backdrop-filter` (frosted glass): the blurred backdrop is a picture under the element, while the panel and its text stay native.
- **Clipping**: `overflow:hidden` containers stay native; only the child that actually overflows is rasterized, with the clip applied.
- **Blend modes**: `mix-blend-mode` on a plain solid box (no text, no image, no gradient) is computed for all separable modes (`multiply`, `screen`, `overlay`, `darken`, `lighten`, `color-dodge`, `color-burn`, `hard-light`, `soft-light`, `difference`, `exclusion`) and the non-separable ones (`hue`, `saturation`, `color`, `luminosity`). The box is drawn as its colour blended with the slide background, and each solid box beneath it gets an opaque overlap rectangle in the blended colour (lowest first). Exact for opaque rectangular boxes; verified against the browser on slides 8, 12 and 13.
- **Fonts**: the browser's real font is detected (an uninstalled first family falls through to the one actually drawn). `@font-face` web fonts are **embedded in the pptx**: TTF, WOFF and WOFF2 directly, CFF-based OpenType (`.otf`) converted to TrueType outlines by the pure-JS `otf2ttf.mjs` (no Python needed). Verified with the real Pretendard (WOFF, WOFF2, CFF `.otf`; regular and bold). Fonts whose licence forbids embedding (OS/2 fsType restricted) are not embedded; their text is converted to **vector outlines** instead, so it looks the same without the font being shipped. `--no-embed-fonts` disables embedding; `--font-map "Pretendard=Malgun Gothic"` substitutes a font.
- **Text outlines**: text that PowerPoint cannot draw as live text (blurred text, skewed text, text in licence-restricted fonts) becomes native freeform shapes built from the real glyph outlines (cubic and quadratic curves kept exact, `opentype.js`). PowerPoint ignores blur on text runs and on text boxes (tested), and has no text skew, so these stay vector shapes: they look right and scale cleanly, but are not editable as text.
- **Still pictures (PowerPoint has no equivalent)**: `feTurbulence` and other procedural/generative filter primitives (the noise is drawn in the browser and placed as a picture), spatial filter primitives (offset/morphology/displacement/lighting) outside the shadow chain, `backdrop-filter` (the blurred backdrop is a picture), `mix-blend-mode` on text, images, gradients or pictures, `<img>`/`<canvas>`/`<video>`, and SVG with external images. Drop-shadow softness is matched closely (measured within about 7/255 at the shadow edge) but is an approximation of the browser's Gaussian.
- **Line breaks**: `<br>` and the lines of a `<pre>` block are emitted as explicit lines (one paragraph per line, indentation kept), so code panels keep their spacing.
- **Text baseline**: each text box is placed so that PowerPoint's baseline lands on the browser's baseline (font ascent measured per run, PowerPoint's exact line spacing modelled). Measured on `ultra-complex.html`: titles and body text agree within 2px.
- **Runtime content**: the converter waits `--settle` ms (default 2000) after load before reading the DOM, so animated readouts (fps, canvases, charts) have settled.
- **Wrapping**: default is soft wrap (natural to edit). Add `--hard-wrap` to insert line breaks exactly where the browser wrapped, so lines match the preview.
- Regression test: `test/complex.html` in the project exercises all of the above; rebuild (`--hard-wrap`) and compare with `render_pptx.ps1` after changing the converter. Template regression: `test/tpl-test.html` with `--template test/template-master.pptx`.
- Wrapping can differ by a few px between browser and PowerPoint; leave ~5% spare width/height in tight text boxes.

## Templates (slide master)

`--template master.pptx` makes the output adopt that file's slide master, layouts, theme and slide size. The master's background and decorations (bars, footers) show behind the slide content, and the layouts are the ones PowerPoint offers in its Layout menu.

- **Layout per slide**: `data-layout="Title Only"` on `section.slide` picks a layout by name (exact match first, then substring). Without it, or when no name matches, the slide uses `Title Only`, then `Blank`, then the first layout in the file.
- **Title placeholder**: `data-ph="title"` on the element that holds the slide title. When the chosen layout has a title placeholder, the text takes its position and size from the layout; otherwise it keeps its HTML position.
- **Body and subtitle placeholders**: `data-ph="body"` on a container (`ul`/`ol`, or a `div` of `p`/`li` lines) fills the layout's first `body`/`obj` placeholder, one paragraph per line, with the layout's own bullets (no `•` or numbers are added) and each run's bold, colour and font kept. `data-ph="subtitle"` on one paragraph fills the layout's `subTitle` placeholder as one paragraph without bullets. Both take the layout's position and size. When the layout has no such placeholder, the element stays a text box at its HTML position, with list markers as before.
- **Background**: in template mode the HTML slide background (colour, gradient, background layer) is not written, so the master's background shows.
- **Size**: the template's slide size is used. Content (positions, font sizes, line widths) is scaled by `min(width / 13.333 in, height / 7.5 in)` and centred on the slide. Tested with a 16:9 and a 4:3 sample master; other aspect ratios use the same formula but are not tested.
- **Overlap**: the converter does not move content away from the master's decorations. Content drawn over a footer or bar covers it, as it would in PowerPoint.

```html
<section class="slide" data-layout="Title Only">
  <h1 data-ph="title">Quarterly review</h1>
  ...
</section>
```

## Rules of thumb

- Prefer a real `<table>` over pictures of tables. Inline SVG charts become native shapes and text, so keep them to the decomposable subset (see Charts).
- Do not use `position:absolute` text stacks unless needed; flow layout converts more reliably.
- Always preview before export, and verify after export when PowerPoint is available. Report honestly what did not carry over.

## Accepted limits (decided by the user)

These are PowerPoint format constraints, verified by experiment, and were accepted as-is:

- **Blurred and skewed text** stays as vector outline shapes (PowerPoint ignores blur on text runs and on text boxes, and cannot skew text). The shapes look right but are not editable as text. Option kept: outlines.
- **`feTurbulence` (noise)** and other generative filter output are placed as pictures.
- **`backdrop-filter`**: the blurred backdrop is a picture under a native panel.
- **Blend modes on text, images, gradients or picture backdrops** are pictures rendered with their backdrop.
- **Templates**: verified with the sample masters in `test/` (16:9 and 4:3) only; no customer template has been tested. Content that runs over a master's footer or bar covers it (no automatic avoidance).

Everything else in this document is native and verified against the browser preview on `test/complex.html` (13 slides).
