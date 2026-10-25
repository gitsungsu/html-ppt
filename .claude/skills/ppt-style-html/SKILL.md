---
name: ppt-style-html
description: Turn content into a PPT-style HTML presentation that converts to PowerPoint with native text and shapes (styles 1–3; the Marp style 4 is not converted; 16:9 slides, white or dark background, no animation by default). Picks a style (Standard, Data Story, Developer, or Document) from the purpose, audience and content, or blends them. Use when the user asks for an HTML slide deck, PPT-style HTML, 발표 슬라이드 HTML, or a web presentation built from notes or a document.
---

# ppt-style-html: content → PowerPoint-ready HTML slides

The HTML is an intermediate format for PowerPoint. Layout, text, tables, charts and icons should survive `html-ppt` conversion as native objects, so the starters use the html-ppt slide format: stacked `section.slide` elements at 1280×720, read in full by the converter. This skill writes no `.pptx` itself; the conversion step uses the separate `html-ppt` skill.

## Styles

| # | Style | Stack | Starter | PowerPoint result (html-ppt, 2026-10-22) |
| --- | --- | --- | --- | --- |
| 1 | Standard | HTML + Tailwind + Lucide icons | `styles/01-standard/index.html` | 5 slides, 26 text boxes, 34 shapes, 0 pictures. Icons become shapes; colours and icons verified in the PowerPoint render. |
| 2 | Data Story | HTML + Tailwind + ECharts (SVG renderer) + tables | `styles/02-data-story/index.html` | 5 slides, 40 text boxes, 25 shapes, 0 pictures. The chart (two bars) is native shapes and text; its labels sit about 10px higher than in the browser. |
| 3 | Developer | HTML + Tailwind + code blocks + Mermaid (SVG, labels as text) | `styles/03-developer/index.html` | 5 slides, 19 text boxes, 5 shapes, 1 vector SVG. The code block is native text; the flow diagram is one vector image (Mermaid draws arrowheads as markers). |
| 4 | Document | Marp + custom CSS + Mermaid | `styles/04-document/deck.md` | Not converted by html-ppt (Marp's markup is not the html-ppt slide format). Marp's own editable PPTX needs LibreOffice, which was not available, so it is not verified. Use styles 1–3 when the deck must become PowerPoint. |

The Executive preset is style 1 with KPI cards and inline SVG figures (guide section 4.2); there is no separate starter. Reveal.js, Slidev and GSAP are not used: their slides are not in the html-ppt format, so the converter cannot read them.

## Choosing a style

Fill in the content profile in `references/analysis-rubric.md`, then apply these rules in order; the first match wins:

1. **Interactive data** (hover, drill-down or live updates needed, or two or more charts) for a **report** or **persuade** purpose → **2 Data Story**. This comes first, so a code-heavy deck that needs charts becomes style 2.
2. **Code-heavy** (two or more code blocks longer than 8 lines, or a developer audience with a live walkthrough) → **3 Developer**.
3. **Document** (purpose **share**, mostly text and tables, written in Markdown and read as a document) → **4 Document**. If the deck must become PowerPoint, use style 1, 2 or 3 instead.
4. **Default** (business talk, mixed content, custom layout) → **1 Standard**.

If no signal is strong, choose **1** and say which rule matched.

Blends (at most two per deck): **1 + ECharts (from 2)** when one chart needs hover detail and the purpose is not report or persuade; **1 or 2 + Mermaid (from 3)** when a flow is simpler as text. Do not blend **3** with **4**, and do not blend **3** with ECharts (use **2**). No animation library is blended in. A new style is allowed only when none of the four fits; it still follows the rules below.

## Workflow

1. **Brief**: topic, audience, purpose, slide count, language, and any style the user named. A named style is used as `rule: user choice`; the blend limits still apply.
2. **Analyze**: the content profile, and the story structure (guide section 20: opening, context, problem, insight, evidence, solution, impact, next step, closing).
3. **Choose** the style (above).
4. **Write**: copy the starter, rewrite the source text into slide messages (guide section 19), keep one message per slide, and fill in the style profile. Map each slide to a structure from guide section 12 (title, section divider, title + content, two column, three column, KPI, comparison, timeline, process, chart, diagram, full image).
5. **Check**: `node scripts/check-starter.mjs <deck.html> <out> --slides N` (Marp: `--mode sections`). Read the screenshots.
6. **Convert** (html-ppt skill, path shown as `<html-ppt>`):
   `node <html-ppt>/scripts/build_pptx.mjs deck.html out.pptx --hard-wrap`
   `powershell -File <html-ppt>/scripts/render_pptx.ps1 out.pptx renders/`
   The converter's summary does not report lost colours. Open the PNGs in `renders/` and check every slide by eye: colours and icons, text clipped by its card, mid-word or awkward line breaks, labels at the edge, and whether chart and table text is still readable. `--hard-wrap` makes the line breaks match the browser (without it, PowerPoint breaks Korean mid-word).
7. **Report**: the style and rule, the check result, the conversion counts, the render check, and what did not convert.

## Authoring rules for PowerPoint conversion

These come from the html-ppt contract and from the conversions above.

- **Slides**: each slide is `<section class="slide">` at exactly 1280×720 px, stacked in the page. Speaker notes go in `<aside class="notes">` (the converter reads them into the PowerPoint notes; that code path is not checked in PowerPoint).
- **Colours: hex, rgb() or rgba().** Tailwind v4 palette classes (`bg-sky-300`, `text-slate-500`) compute to `oklch()`, which the converter cannot read: the backgrounds, bars and icons disappear, and the summary still looks normal. Write `bg-[#7dd3fc]` or `style="color: #555555"` instead. The checker fails any `oklch`, `oklab`, `lab`, `lch`, `color()` or `color-mix()` colour in text, backgrounds, borders, shadows, outlines and pseudo-elements.
- **Backgrounds**: white is the default. A dark theme is allowed when every text box has at least 4.5:1 contrast with the colour it sits on (checked). A gradient or image background is noted, because the converter may draw it as a picture.
- **Sizes**: body text at least 24px (18pt); captions and labels at least 16px (12pt; the checker's floor, and 16–23px is noted); titles 48–72px. Keep content inside the 5% safe area (64px at the sides, 36px at the top and bottom; the starters use 96px and 64px).
- **Layout**: flex, grid and tables convert from their rendered geometry. Keep boxes square where the browser shows square corners: the converter kept rounded bottom corners on a bar that was rounded only at the top.
- **Icons**: Lucide is loaded at run time and converted to shapes (verified).
- **Charts**: ECharts with `renderer: 'svg'` and `animation: false` converts to shapes and text (verified). Charts built from `div` bars also convert (verified).
- **Tables**: native `<table>` converts to text boxes and rules (verified).
- **Code**: `<pre>` text with coloured spans converts as text lines (verified); use a monospace font such as Consolas.
- **Diagrams**: Mermaid with `htmlLabels: false` and `themeVariables.fontSize: '24px'` draws labels as SVG text; the whole diagram becomes one vector image (not editable shapes).
- **Fonts**: Pretendard is loaded with `@font-face` from the jsdelivr CDN (OFL); the converter embeds it (verified: regular and bold; the SemiBold 600 face that some labels use is not checked). Other `@font-face` fonts are embedded when their licence allows.
- **Animation**: none. The optional presenter (`?present`) shows one slide at a time with arrow keys; the converter reads the default stacked page, so the presenter does not change the export.
- **Avoid**: web UI (navigation bars, dropdowns, modals, hover-only information), GSAP and other motion libraries, 3D, parallax, blend modes, backdrop filters, clip paths and masks on content, and text used as pictures.

## Verified conversion (2026-10-22)

Each starter was converted with `build_pptx.mjs --hard-wrap` and rendered with `render_pptx.ps1`. The renders were checked for colours and icons (for example, the 01 bars keep their blue ramp, and the 03 card text keeps its accent colour). Earlier runs with Tailwind v4 palette classes showed the cards, bars and icons missing from the render, with the summary unchanged.

## Check script

`scripts/check-starter.mjs` renders each slide in headless Chromium and applies the guide's measurable rules (section 23):

- 16:9 canvas (aspect ratio within 1%; the pixel size is not tested);
- background contrast: every text box has at least 4.5:1 against the colour it sits on; a gradient or image background is noted;
- text inside the 5% safe area;
- no overlapping text boxes (an overlap above 20% of the smaller box fails; HTML text only, not chart or SVG text);
- text at least 16px (16–23px is noted; body text at 24px is checked by eye);
- text clipped inside a box that hides its overflow;
- no colours the converter cannot read (oklch and related);
- overflow outside the slide;
- console errors, page errors and failed requests;
- slide count (`--slides N`);
- the presenter, when the deck has one: `?present` shows one slide, and ArrowRight steps through slides (a deck without `?present` is noted, not failed);
- multi-column text (`column-count`) is noted: the converter splits it into boxes, so the column breaks must be checked in the render.

It cannot check the other items of section 23 (images, chart legibility, story and content quality, line breaks and clipping as they appear in PowerPoint, the PowerPoint render itself); check those on the screenshots and the render.

```bash
node scripts/check-starter.mjs styles/01-standard/index.html out/01 --slides 5
node scripts/check-starter.mjs styles/04-document/dist/deck.html out/04 --mode sections --slides 5
```

Setup, once per machine: `cd scripts && npm install && npx playwright install chromium`, and for style 4 also `cd styles/04-document && npm install`. The checker needs the CDNs for Tailwind, Lucide, ECharts, Mermaid and Pretendard.

## Guide checklist (section 23)

- **Slide**: 16:9 and background contrast (checked); content inside the safe area (checked); no overlaps (checked).
- **Content**: a clear title and one message per slide, with no long paragraphs (manual).
- **Typography**: body at least 24px (checked for text below 16px; 24px is manual); one font family (Pretendard).
- **Images**: PNG, JPEG or SVG with `object-fit` (no images in the starters; manual when used).
- **Charts**: legend and numbers readable, not overly complex (manual; the starters use one series).
- **Conversion**: no animation needed for meaning (by design; the checker does not test motion); HTML renders as static slides; nothing lost in the PowerPoint render (manual, using the render).

Other guide sections the skill follows: 1 and 24 (defaults: 16:9, white, animation off), 4 (presets, mapped above), 5 (avoided stacks), 6 (animation policy), 7 and 8 (fixed canvas and safe area), 9 (one message), 10 (fonts), 11 (sizes), 12 (structures), 13 (charts: bar, line, donut, etc., simple and without 3D), 14 (tables: header emphasis, few rows and columns), 15 (images), 16 (colours: #111111 text, #555555 secondary, one accent hue), 17 (default theme: style 1), 18 (web UI avoided), 19 (rewrite, do not copy), 20 (story structure), 21 (simplest stack first), 22 (conversion first), 25 (final test: could this be used as is in a business talk after conversion?).

## Style profile

Every deck starts with a comment that records the choice:

```html
<!-- style-profile
purpose: persuade
audience: executives, live talk
primary: 1 Standard (HTML + Tailwind + Lucide)
blend: none
rule: 4 (default); no chart or code
-->
```

In Markdown decks (style 4), put the same lines as YAML comments (`# key: value`) in the frontmatter.

## Pinned versions

- CDN (exact, checked with HTTP 200): Tailwind `@tailwindcss/browser` 4.1.11, Lucide 0.460.0, ECharts 5.5.1, Mermaid 11.17.2, Pretendard 1.3.9 (woff2).
- npm (exact, at least two weeks old): `playwright` 1.64.0 (in `scripts/`), `@marp-team/marp-cli` 4.5.1 (in `styles/04-document/`).
- When a pin changes, re-run the check on every starter and the conversion with the render check.

## Files

- `styles/01-standard/index.html`, `styles/02-data-story/index.html`, `styles/03-developer/index.html`: html-ppt format, one file each (CDN libraries, presenter included).
- `styles/04-document/`: `deck.md`, `theme.css`, `package.json`, `scripts/inject-mermaid.mjs` (build: `npm run build` → `dist/deck.html`).
- `scripts/check-starter.mjs`, `scripts/package.json`: the checker and its Playwright dependency.
- `references/analysis-rubric.md`: the content profile and the decision rules.
- `references/ppt-conversion-guide.md`: the full guide, copied unchanged from the project root, with a header on how this skill applies it.

## Not verified

- Conversion of style 4 to PowerPoint (needs LibreOffice for Marp's editable export).
- Conversion on a machine without the `html-ppt` skill (the converter is a separate skill).
- Mermaid diagrams are vector images, not editable shapes (the arrowheads are markers).
- Line breaks and clipping in PowerPoint: checked by eye on the renders only; the checker measures the HTML.
- The speaker-note export and the fonts on a machine where Pretendard is not embedded.
