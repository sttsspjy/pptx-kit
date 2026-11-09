---
name: consolidate
description: Builds or rebuilds the pack-level layer of a pptx-kit design pack from its verified layouts — final tokens.json, lib/components.mjs (incl. charts in the deck's style), a rendered component demo, catalog.json, DESIGN.md and pack.json font sources. Runs at the end of extract; re-run alone after layouts change.
argument-hint: "<packDir>"
disable-model-invocation: true
---

# Consolidate slide templates into a design pack

Arguments: `$ARGUMENTS` → `<packDir>`. Plugin root: `${CLAUDE_PLUGIN_ROOT}`. Read `${CLAUDE_PLUGIN_ROOT}/reference/pack-format.md` and `${CLAUDE_PLUGIN_ROOT}/reference/pptxgenjs.md` first.

A set of slide templates was rebuilt one slide at a time from a source deck (`<pack>/layouts/*`: `layout.mjs`, `SPEC.md`, `sample.json`, `stress-*.json`). Each was verified against its source render, so the layouts themselves stay as they are. Your job is the pack-level layer that lets a later model build NEW decks in this design: a design system document, a catalogue, and shared components for slides no layout covers.

## Output (inside `<pack>`)

1. `tokens.json` — final: colour roles (hex), type scale (role → face, size, weight, colour, line spacing), spacing scale, grid (canvas, margins, content box, header bands), radii, strokes, shadows. Keep every key existing layouts read (grep them); add, don't rename.
2. `lib/components.mjs` (new file; do not edit other existing `lib/` files, layouts import them; when it already exists from an earlier run, extend it without breaking its exports) — components for composing new content slides inside the deck's grid, lifted from the verified layouts' code where it exists, each taking `(slide, ctx, box, data)` and returning the height used: headline + insight block (via existing chrome), bullet list, quote card, picture / phone-screenshot frame with marks, panel / card, numbered finding, dark header bar + body, table in the deck's table style, arrow / connector, and **charts in this deck's visual language** (bar, stacked 100 % bar, line, donut; native PptxGenJS charts styled with the deck's colours, faces and sizes — if the source deck has no charts, derive the style from its palette, type roles and table styling; keep them quiet: no 3D, no gridline clutter, data labels in the body face; chart colours only from `tokens.json`).
3. `lib/demo.mjs` + renders — a layout-shaped module that places every component once; copy it into `layouts/_demo/layout.mjs` (with a `sample.json`) and render it: `node "${CLAUDE_PLUGIN_ROOT}/tools/render.mjs" --pack "<pack>" --layout _demo --out "<pack>/_work/demo/<n>"`; view `render.png` and iterate until each component looks native to the deck. Several demo slides are fine (`_demo2`, …) when one canvas cannot hold every component.
4. `catalog.json` — `{ "canvas": {w, h, unit: "pt"}, "layouts": [ … ] }`, one entry per layout (not `_demo*`): `id, name, kind (cover | contents | section | subsection | content | insert | closing), purpose, useWhen, avoidWhen, capacities (items, lines, chars), imageSlots (count, aspect, fit), dataFields (top-level keys), sourceSlide`. Read each layout's `schema` and `SPEC.md`; write capacities from the schema.
5. `DESIGN.md` — the design system for a model building new decks in this style, in measured numbers, not adjectives: deck structure and slide-kind sequence (look at the contact sheets `src/sheet-*.png` for the whole deck); chrome rules per slide kind (breadcrumb text pattern, tracker states, page number, insert/example slides); grid and spacing; type roles; colour roles and the keyword-emphasis rule; writing rules (headline length and line count, insight line, bullet depth, quote text + attribution format, number formats, tone); image rules (which pictures go in which frame, fit, marks); how to choose a layout; how to compose a new slide from `lib/components.mjs` when no layout fits, staying on the grid; chart rules (palette lock: chart colours only from the listed roles); minimum type sizes; known limits. Reference source renders only as `src/png/sNN.png` of archetype slides (other slide PNGs are pruned from shared packs) and contact sheets.
6. `pack.json` — keep its keys; for every entry in `fonts`, fill `source` (official download page if you know it, else `"ask the deck owner"`) and `licence` (e.g. `"SIL OFL"`, `"proprietary — bundled with Windows"`, `"unknown"`). Never bundle font files unless the licence is free (OFL or similar) — then they may go in `<pack>/fonts/` with the licence text.

## Rules

- Do not modify `layouts/*/layout.mjs` or existing `lib/*.mjs` (except extending `lib/components.mjs` / `lib/demo.mjs` you own).
- Measured values come from the layouts' SPEC.md files, facts (`facts/sNN.json`) and source renders; resolve disagreements against the source.
- Paths in pack code are pack-relative complete string literals (`'facts/media/<file>'`, no computed file names — prune keeps only referenced media); pack code imports no npm packages (`ctx.deps`).
- Finish: `node "${CLAUDE_PLUGIN_ROOT}/tools/pack.mjs" check "<pack>"` and fix every error. Then reply with: component list, demo render verdict, anything unresolved.
