# Style selection rubric

Run this before choosing a style. Record the result as a style profile in the deck (format in SKILL.md, "Style profile"), so the choice can be explained and revised.

## 1. Purpose and audience

Primary purpose (pick one):
- **inform**: explain a topic or a status
- **persuade**: pitch, decision request, product launch
- **report**: results, KPIs, financials
- **teach**: training, seminar, how-to
- **demo**: live product or code walkthrough
- **share**: draft or planning document read asynchronously

Audience and setting:
- audience: executives, clients, team, developers, public
- setting: live projected talk, self-paced reading on screen, shared link, PDF/print
- time per slide: under 30 s (skim), 1-2 min (talk), several minutes (read)

## 2. Content profile

Measure from the input before deciding:
- `slide_count`
- `density`: words per slide (sparse ≤ 40, medium 40-100, dense > 100)
- `structure`: sequence, comparison, hierarchy, process/flow, timeline, data story, code walkthrough
- `charts`: number of charts the deck needs (one per figure, not per series)
- `diagrams`: flows, sequences, architecture diagrams
- `code_blocks`: blocks longer than 8 lines
- `interactive`: does the reader need hover detail, drill-down or live updates? yes/no
- `motion_matters`: does the entrance or transition carry the message? yes/no (default no: animation is off unless the user asks)
- `key_slides`: which slides carry the main message (these get the strongest emphasis)

## 3. Primary style: decision rules

Apply in order; the first match sets the primary style. All four styles are HTML in the html-ppt slide format, except style 4 (see its note).

1. **interactive data** (`interactive` = yes, or `charts` ≥ 2) with purpose **report** or **persuade**: **2 Data Story** (HTML, ECharts drawn as SVG, tables). Checked first, so a code-heavy deck that needs charts becomes style 2.
2. **code-heavy** (`code_blocks` ≥ 2, or a developer audience with a live walkthrough): **3 Developer** (HTML, code blocks as text, Mermaid flow as SVG).
3. **document** (purpose **share**, mostly text and tables, the author writes Markdown, the deck is read as a document): **4 Document** (Marp). Style 4 is not converted by html-ppt; if the deck must become PowerPoint, write it in style 1, 2 or 3 instead.
4. **default** (business talk, mixed content, custom layout wanted): **1 Standard** (HTML, Tailwind, Lucide icons). For executive reports with KPIs or comparisons, use style 1 with KPI cards and inline SVG figures (guide section 4.2).

## 4. Blending

A blend adds one library from another style to the primary. Allowed blends:
- **1 + ECharts (from 2)**: one chart needs hover detail, the purpose is not report or persuade (otherwise rule 2 applies), and the rest of the deck is a standard layout.
- **1 or 2 + Mermaid (from 3)**: a flow or sequence diagram is simpler to write as text than to draw by hand.

Limits:
- At most two blended libraries per deck.
- Do not blend 3 with 4: they use different authoring models (HTML slides or Markdown).
- Do not blend 3 with ECharts; if charts matter in a code deck, use 2 instead.
- No motion library is blended in: animation is off by default (guide section 6). A fade may be added only when the user asks for it.

## 5. A new style

Use a new combination only when the content needs something none of the four provides, for example a print-first poster layout or a fixed brand template. Record the stack and the reason in the style profile. The new style still follows the html-ppt authoring contract in SKILL.md (1280×720 `section.slide` slides, hex colours, native text and shapes, no animation by default).

## 6. Worked examples

| Input | Signals | Primary | Blend |
| --- | --- | --- | --- |
| Quarterly results, 10 slides, 3 interactive charts, executives | interactive, charts ≥ 2, report | 2 | none |
| Architecture walkthrough, 12 slides, 2 code blocks, 1 sequence diagram, developers | code_blocks ≥ 2 | 3 | none (the sequence diagram is part of style 3) |
| Internal planning memo, 8 slides, text and a flowchart, shared as a document | share, document | 4 | none |
| Product pitch, 8 slides, no data, brand look | default | 1 | none (no animation unless the user asks) |

## 7. Style profile

Record the choice in the deck with the style-profile comment described in SKILL.md ("Style profile").
