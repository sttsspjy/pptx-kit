---
name: generate
description: Builds a finished, editable PowerPoint deck from report content (a Mirae slide-content export, a structured content.json, or plain markdown) in the design language of a pptx-kit design pack — plan, render with PowerPoint, review every slide image, fix, report.
argument-hint: "<packDir> <contentDir> [outDir]"
disable-model-invocation: true
---

# Build a new deck in an extracted design

Arguments: `$ARGUMENTS` → `<packDir> <contentDir> [outDir]` (default outDir: `./deck-<content folder name>`). Ask for missing paths. Plugin root: `${CLAUDE_PLUGIN_ROOT}`. Everything runs locally; nothing is uploaded.

Make a finished, send-ready PowerPoint deck from new report content, in the design language of a reference deck that was extracted into a design pack. The result must look like it was made by the same team with the same template: same structure, chrome, grid, type, colours, components and writing style — and contain only the new content.

## Read once (cost rule)

Read everything below **once, up front, in one parallel batch**, and do not read these files again during the run (re-reading them doubled the cost of a run in measurement; you already have them in context):

- `${CLAUDE_PLUGIN_ROOT}/reference/pack-format.md`, `${CLAUDE_PLUGIN_ROOT}/reference/pptxgenjs.md`
- pack: `DESIGN.md` (fully), `catalog.json`, `tokens.json`, `pack.json`, every `lib/*.mjs`
- content: `content.md`, `content.json` (if present)
- `python "${CLAUDE_PLUGIN_ROOT}/tools/imginfo.py" "<contentDir>"` — pixel size of every content picture

Later, read a layout's `layout.mjs` (for its `schema`) and `SPEC.md` only for layouts you actually use, once each. View a layout's `preview.png` or its source render `src/png/sNN.png` when choosing between layouts.

## Content formats (contentDir)

1. **Mirae slide content** — `content.md` (header + report markdown) and optionally `content.json` with `"schema": "mirae-slide-content/1"`: `meta {project, title, exportedAt, language, sessions[]}`, `report` (markdown), `quotes [{text, speaker?, session, t?}]`, `media [{path, kind, caption?, session?, t?}]`; pictures under `media/`. The report's headings are the sections and messages; quotes and media are evidence to place next to the findings they support (match by session and topic, never by guess).
2. **Structured content** — `content.json` with `meta`, `sections`, `messages` (each message: id, section, kind hint, headline, insight, bullets, quotes, images, table / data).
3. **Plain markdown** — only `content.md`: `#` = title, `##` = sections, `###` / paragraphs / lists = messages; images referenced in the markdown.

Paths inside content are relative to the content folder. Text language = `meta.language` or the content's language.

## Output — `<out>/`

- `deck.json` — `{ "pack": "<pack, relative to out>", "content": "<content, relative to out>", "slides": [ {"layout": id, "data": {…}} | {"module": "custom/<name>.mjs", "data": {…}} ] }`. Content picture paths in data stay relative to the content folder.
- `custom/*.mjs` — only for slides no layout fits; built from `ctx.lib` components and `ctx.tokens`, inside the grid in `DESIGN.md`, same module shape as a layout.
- `decks/<n>/deck.pptx` + renders via `node "${CLAUDE_PLUGIN_ROOT}/tools/render-deck.mjs" --spec "<out>/deck.json" --out "<out>/decks/<n>"` (n = 1, 2, … per render), and a final copy `<out>/deck.pptx`.
- `REPORT.md` — per slide: content message id(s), layout or custom module, why it fits, anything trimmed; then **Content inconsistencies** (flagged, not changed).

## Process

1. Preflight: `node "${CLAUDE_PLUGIN_ROOT}/tools/doctor.mjs" --pack "<pack>"`. Missing fonts or PowerPoint → stop and tell the user the fix (font sources are in `pack.json`).
2. Plan the deck: map content to the source's structure (cover, contents, section and subsection dividers, content slides, insert/example slides, closing). Choose a layout per message by its message shape using `catalog.json` (`useWhen` / `avoidWhen` / capacities) and the reference render. Never force content into a layout built for a different shape (e.g. quotes into a table, a chart into a table, three items into a two-column layout). Quantitative comparisons and distributions become charts in the deck's chart style (`lib/components.mjs`), unless the reference deck uses a table for that shape.
3. Write data that satisfies each layout's `schema`. Content beyond capacity → split into two slides or shorten by deleting words, never by shrinking type.
4. Render the deck, then review every slide image at full size (Read `decks/<n>/png/sNN.png`) against the reference render of its layout and against `DESIGN.md`: zero overflow, no collisions (labels vs bars, text vs frames), no empty slots, correct chrome (breadcrumb text, tracker state, page number), pictures in proper frames with correct fit, emphasis used as the source uses it, consistent spacing, colours from `tokens.json` only.
5. Fix and re-render until every slide passes; batch related fixes, then re-check only the changed slides and their neighbours. Copy the final `deck.pptx` to `<out>/deck.pptx`. Then write `REPORT.md`.

## Rules

- **Fidelity: never change a number, label, name or quote from the content.** If the content looks inconsistent (totals that do not add up, a label that contradicts a number, the same figure given twice differently), keep it as given and flag it in `REPORT.md` → Content inconsistencies.
- Do not invent findings, numbers, quotes, names or pictures. Picture slots without a matching picture → choose a layout without that slot.
- Quotes stay verbatim with their speaker / session attribution in the deck's attribution format; a quote too long for a slot gets a layout with more room or its own slide — never a rewrite.
- Headlines and bullets may be shortened only by deleting words; never reword a claim into a stronger or different one.
- Text as given (any language); names and acronyms as spelled.
- Text is native and editable; no slide or text rendered as an image.
- Finish with: slide count, overflow count (must be 0), the REPORT.md path, the deck path, and any flagged inconsistencies.
