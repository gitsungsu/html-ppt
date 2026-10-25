---
name: ppt-template-html
description: Learn the common template (layouts, fonts, colours, background/logo/decorations) from two or more sample PPTX files, then turn Markdown/text, images and CSV/Excel into HTML slides (1280x720 section.slide) that follow that template, check them, and compare each layout against the PowerPoint original. Use when the user says "PPT 양식 학습", "PPT 템플릿으로 HTML 슬라이드", "샘플 PPT 양식대로 슬라이드 만들어줘", "양식 추출", "template to HTML slides", "learn a PowerPoint template", "build slides in the style of these decks", or "compare HTML slides with the PPTX layout". HTML only: no PPTX output, no charts.
---

# ppt-template-html

양식(template) 학습 -> 콘텐츠 준비 -> HTML 생성 -> 검증 -> 비교. The template spec file (`template-spec.json`) is the single source of truth; the generator reads only that spec.

`SKILL=.claude/skills/ppt-template-html`. Run every command from the project root in Git Bash, except the two PowerShell scripts (PowerPoint COM, Windows only).

## Workflow

1. **Learn the template.** Give two or more sample PPTX files. The extractor reads the package XML (no PowerPoint needed): theme colours (12 hex), fonts (latin/ea), per layout the effective background and decorations (master logo inherited unless `showMasterSp="0"`), placeholder kinds, indexes and inherited positions. Scalars are merged by majority. A tie takes the `--primary` sample (else the first file name in sort order) and is written to `conflicts[]`.
   **Layout art vs slide content:** a layout picture is dropped when (a) the slides of that layout place a picture at the same position and size (a copy of slide content, e.g. the contents-slide numbers), or (b) a later opaque (JPEG) full-slide picture of the same layout covers it. The notes say which pictures were dropped. Repeated pictures are re-added by the 60% rule.
   **Single-sample mode** (spec change: AC2 now allows 1 sample): with one sample every value is taken as-is, `conflicts` is `[]`, and `warnings[]` gets `single sample: values taken as-is, no majority vote`. Zero samples still exits 2.
   **Slide-level learning** (AC8, for decks whose design lives on the slides rather than in the theme): the extractor also reads every slide. It writes `observed` = `{fonts:{title:{latin,ea}, body:{latin,ea}}, colours:{title, body, accent}, slides:{<layout>: count}, notes:[...]}`, using character-weighted modes of the slide text runs (title = runs in `title`/`ctrTitle` placeholders, body = all other runs; runs without an explicit font or colour do not vote). Fonts vote by family (trailing weight words such as `ExtraBold` are dropped); accent is the most used colour whose max minus min channel is above 40. A picture (same media file, top-level on the slide, not inside a group) that appears on at least 60% of a layout's slides is added to that layout's `decorations` (position from its first occurrence); if it covers at least 90% of the slide it becomes the layout `background` (`type: "image"`) instead. Theme `palette` and `fonts` are left unchanged. Precedence in `build-deck.mjs`: a placeholder's own `textStyle.color` > `observed.colours.title|body` > the palette text colour; `observed.fonts.title|body` > theme `fonts.major|minor` (decoration text keeps the theme font).
   **Subtitle and text box learning** (AC10): the extractor also reads top-level non-placeholder text boxes (`p:sp` with text; group children skipped) with the same fields plus a character count. On each slide, the nearest box whose top is at or below the title's bottom and whose horizontal range overlaps the title's is a subtitle candidate; the largest other box below the title is the text box. They go to `observed.layoutStyles[<layout>].subtitle` and `.textbox` (median position, weighted mode for the rest; omitted without candidates). In `build-deck.mjs`, if the layout has `subtitle`, the slide's first plain body line (not `- `, `* ` or `1. `-style) is drawn at the subtitle position with its font, size, colour, bold and alignment and no bullet; the other lines stay in the list. Otherwise a `textbox` style sets the list's font, size and colour (and its position when the slide has no table, so a table keeps the wider body box). Null fields fall back to the body style.
   **Per-layout learned styles** (AC9): for each top-level title (`title`/`ctrTitle`) and body (`body`/`obj`/`subTitle`) placeholder on a slide, the extractor reads the slide's own `a:xfrm` (if any), the first paragraph's `algn` and the first text run's `sz`, `b`, colour and latin font. Per layout name it writes `observed.layoutStyles[<layout>] = {title:{x,y,w,h,align,size,bold,colour,latin}, body:{...}}` (only roles that occur): position and size are the median of the slides that set an `xfrm`, the other fields the character-weighted mode; missing values are `null`. It also writes `observed.colours.tableHeader` / `tableHeaderText`: the most common first-row cell fill and first-run text colour of the first `a:tbl` found on the slides (`null` if there is no table). Precedence in `build-deck.mjs` for title/body position, size, align, bold, colour and font: `observed.layoutStyles[<layout>]` > the placeholder `textStyle` > global `observed` > default; a `null` field falls through. Learned boxes (title, body, subtitle, textbox) are clamped into the 5% safe area at build time; the spec keeps the learned value. Table headers use `tableHeader` (background) and `tableHeaderText` (text) when set, else the theme accent and a contrasting text colour.
2. **Prepare content.** Markdown or `.txt` (same rules). One `---` line separates slides. A slide may start with `<!-- layout: NAME -->` to pick a layout. A picture is `![](file.png)` (looked up in `--assets`). A table slide has a `[table: file.csv]` line naming a file passed with `--data` (CSV, or the first sheet of XLSX). In titles, subtitles and body lines, `**text**` is bold and `~~text~~` is light (weight 400, e.g. inside a bold title); markers apply on one line, do not nest, and are not read in table cells.
3. **Build.** `build-deck.mjs` writes `<deckDir>/deck.html` and copies media to `<deckDir>/assets/`. Every slide is a `section.slide` (1280x720, white) holding a centred `.frame` (scaled to the sample slide size, letterboxed with white; the frame has no `data-tpl`). Decorations carry `data-tpl="chrome" data-role="chrome"`; placeholders carry `data-role="title|body|image"`. Colours are hex only.
4. **Check.** `check-deck.mjs` renders the deck in Chromium and applies the ppt-style-html rules (16:9, contrast 4.5:1, 5% safe area, overlap, text size, clipping, unreadable colours such as oklch, overflow, slide count, console and request errors). Findings are split into content (counted) and chrome (reported only, plan D5).
5. **Compare** (needs PowerPoint). `compare.mjs` renders every layout of the sample in PowerPoint and as generated HTML with the same filler text, then judges colour and structure per layout.

## Commands

```
node $SKILL/scripts/extract-template.mjs <sample.pptx> <sample.pptx> [more...] --out <templateDir> [--primary <sample.pptx>]
node $SKILL/scripts/build-deck.mjs <templateDir>/template-spec.json <content.md> --assets <imageDir> [--data <csv-or-xlsx>]... --out <deckDir> [--no-chrome]
node $SKILL/scripts/check-deck.mjs <deck.html> <outDir> [--slides N] [--text-size warn|fail]
node $SKILL/scripts/parity-normalize.mjs <log> --kind deck|starter [--detail] [--slides N]
node $SKILL/scripts/parity-normalize.mjs --compare <a.json> <b.json> <expected.json> --key slides10|slides9
node $SKILL/scripts/parity-normalize.mjs --compare <a.json> <expected.json> --key slides10|slides9
node $SKILL/scripts/parity-normalize.mjs --compare-detail <actual.json> <chrome-expected.json>
node $SKILL/scripts/compare.mjs --spec <template-spec.json> --sample <sample.pptx> --filler $SKILL/fixtures/filler.md --out <dir>
powershell -File $SKILL/fixtures/make-standard-office.ps1      # test fixture, PowerPoint COM
node $SKILL/scripts/make-fixtures.mjs                          # other test fixtures
powershell -File $SKILL/scripts/make-layout-slides.ps1 -Sample <abs.pptx> -Spec <abs spec.json> -Filler <abs filler.md> -Out <abs dir>   # called by compare.mjs
```

`--assets` and `--data` are required whenever `content.md` references such files. `--primary` must be one of the sample files. Use `--no-chrome` for PPTX export, since the master layouts carry the decorations. `check-deck --text-size warn` keeps the template's original body size (about 14px) and reports text under 16px as `warn`, which does not count toward the exit code; the default `fail` is unchanged.

`parity-normalize.mjs` details:
- `--kind deck` keeps only `result=fail` lines, as unique `{slide, rule}`. `--kind starter` maps each `check-starter` problem line to `(slide, rule)` by message prefix; console, page error, request failed, HTTP, slide count and run errors map to slide 0.
- `--detail` (deck only) prints `{slide, kind, rule, result}` for every line, plus `{slide, kind:"content", rule:"any", result:"pass"}` for each slide `1..N` without a content fail or human-check line. N comes from `--slides`, else the `slides found: N` line, else the highest slide seen.
- Rule names (shared by both checkers): `canvas`, `text-size`, `contrast`, `safe-area`, `overlap`, `overflow`, `clipped`, `colour`, `slide-count`, `console`, `page-error`, `request-failed`, `http`, `mermaid`, `mermaid-fence`, `no-slides`, `run-error`, `presenter`.
- Log lines are split on the first five `|` only; the rest is the excerpt.

`compare.mjs` writes `<dir>/report.json`, `<dir>/report.md`, and the image pairs `original/`, `original-lb/` (letterboxed 1280x720), `generated-png/`. Per layout:
- Colour: top 8 colours by pixel share (4-bit bins), minimum-cost matching on Delta E 76, mean value. Pass at <= 15; human-check band 12-18.
- Structure: boxes of the original (spec placeholders and decorations, EMU converted to letterbox coordinates) against generated elements with `data-role` inside `.frame`, paired by role, IoU per pair. Pass when no box is missing and mean IoU >= 0.8; human-check band 0.6-0.8.
- Layout pass = colour pass AND structure pass. Picture-background layouts always go to the human-check list. Thresholds (D2) are starting values to calibrate with test pairs.

## Exit codes

| Code | Meaning |
|---|---|
| 0 | OK. `check-deck`: no content fail. `parity-normalize`: sets equal. |
| 1 | `check-deck`: at least one content fail. `parity-normalize --compare*`: sets differ. |
| 2 | Input or usage error (no sample, oversize file, bad arguments, missing file). |
| 3 | `extract-template`: tie conflict resolved by `--primary` / file-name order, recorded in `conflicts[]`. |
| 4 | `compare.mjs` / `make-layout-slides.ps1`: PowerPoint COM not available, comparison skipped (D4). |

Limits for PPTX and XLSX input (AC7): 50 MB file, 2000 entries, 200 MB uncompressed total; above that the tool exits 2.

## Copied checker and its two intentional differences

`scripts/check-deck.mjs` started as a copy of `.claude/skills/ppt-style-html/scripts/check-starter.mjs`. SHA-256 of that original at copy time:

```
7170353eaccef8f14c36afcac9cc332cbf47c18788ea37b2bf5b01290ebb576d
```

Verify with `powershell -Command "(Get-FileHash -Algorithm SHA256 .claude/skills/ppt-style-html/scripts/check-starter.mjs).Hash"` (PowerShell prints upper case). If the hash changed, re-sync the rules in `check-deck.mjs` and rerun the parity test.

All `check-starter` rules are kept. The only intentional differences:

1. **Background walk is paint-order based.** `check-starter` walks up the ancestors for the first background colour; `check-deck` samples points inside the text's glyph rectangles and asks `document.elementsFromPoint` which element actually paints under the text (the text element itself first). The lowest contrast among solid points is reported. The `.frame` counts as an ordinary background layer.
2. **Picture, gradient, translucent and replaced (`img`/`svg`/`canvas`/`video`) backgrounds are human-check**, not a solved colour. `check-starter` treats a translucent background as that colour. The parity fixture contains neither case.
3. **Transparent-colour check is narrower.** `check-deck` skips the contrast check for a text colour only when it is `rgba(..., 0)` or `transparent`. `check-starter` also skips black `rgb(0, 0, 0)` text (its regex `/, 0\)$/` matches the tail of `rgb(0, 0, 0)`), a bug inherited from the starter and not changed there.

Extra layer: `closest('[data-tpl]')` decides content versus chrome (decoration). Chrome failures are `reported`, never counted (D5).

## Known limits

- Dates come out as raw values. Number formats other than percent, currency and thousands separators are raw (XLSX first sheet only).
- Groups, SmartArt, animations: not read. EMF/WMF pictures are not drawn, only listed in `warnings[]`; this path is unsupported and untested (no fixture).
- Text on a picture background: contrast is a human check.
- `::before` / `::after` backgrounds are not measured. Strips thinner than the sample-point spacing can be missed.
- Media (pictures, logos) live in `assets/` beside `deck.html`; move the folder together with the HTML.
- Non-16:9 samples are centred in a scaled frame with white side bars (D3).
- Comparison fills title/body placeholders only; date, footer, slide-number and picture placeholders are left empty in the PowerPoint render.

## Setup

- **P1 (this skill).** `(cd $SKILL/scripts && npm install && npx playwright install chromium)`. Pins: `playwright` 1.64.0, `jszip` 3.10.2.
- **P2 (PowerPoint COM).** Needed only for `make-standard-office.ps1` and `make-layout-slides.ps1` (fixtures and `compare.mjs`). Without PowerPoint everything else still works.
- **P3 (parity test, optional).** The original checker needs its own install: `(cd .claude/skills/ppt-style-html/scripts && npm install && npx playwright install chromium)`; check with `[ -d .claude/skills/ppt-style-html/scripts/node_modules/playwright ]`. Without it, skip the `check-starter` run and compare `check-deck` output against `fixtures/parity-expected.json` only.

## Test shortcuts

```
node $SKILL/scripts/check-deck.mjs $SKILL/fixtures/chrome-deck.html out/chrome --slides 3 > out/chrome.log
node $SKILL/scripts/parity-normalize.mjs out/chrome.log --kind deck --detail > out/chrome.json
node $SKILL/scripts/parity-normalize.mjs --compare-detail out/chrome.json $SKILL/fixtures/chrome-expected.json
node $SKILL/scripts/check-deck.mjs $SKILL/fixtures/parity-deck.html out/parity --slides 10 > out/parity-deck.log
node $SKILL/scripts/parity-normalize.mjs out/parity-deck.log --kind deck > out/parity-deck.json
node $SKILL/scripts/parity-normalize.mjs --compare out/parity-deck.json $SKILL/fixtures/parity-expected.json --key slides10
```
