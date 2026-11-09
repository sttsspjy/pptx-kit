---
name: review
description: Reviews and finishes a deck made by pptx-kit generate — checks fidelity to the content, visual form, message, slots and craft slide by slide against the design pack, fixes in the deck sources, re-renders, and writes REVIEW.md.
argument-hint: "<packDir> <contentDir> <deckDir>"
disable-model-invocation: true
---

# Review and finish a generated deck

Arguments: `$ARGUMENTS` → `<packDir> <contentDir> <deckDir>` (deckDir = the generate output folder with `deck.json`). Ask for missing paths. Plugin root: `${CLAUDE_PLUGIN_ROOT}`.

A model built a deck from new report content in the design language of an extracted design pack. Bring it to send-ready: find every defect, fix it in the deck's source files, re-render, verify. The result must look like the same team made it with the same template as the reference deck, and contain only the new content.

## Inputs — read once, up front, in one parallel batch; do not re-read later

- Design pack `<pack>` (read-only): `DESIGN.md` (fully), `catalog.json`, `tokens.json`, `lib/*.mjs`; `layouts/<id>/` (`layout.mjs` + `schema`, `SPEC.md`) only for layouts the deck uses. Reference renders of the source deck: `src/png/sNN.png`, layout previews `layouts/<id>/preview.png`.
- Content `<content>/content.md`, `content.json` (if present), `media/` — the only source of facts, numbers, quotes and pictures.
- Deck dir `<out>/`: `deck.json`, `custom/*.mjs`, `REPORT.md`, latest render `decks/<n>/png/sNN.png` + `sheet.png` (highest n).
- `${CLAUDE_PLUGIN_ROOT}/reference/pack-format.md` (CLIs, deck spec).

## Process

1. View `sheet.png`, then every slide at full size next to the reference render of its layout (or the closest source slide).
2. List defects per slide before editing. Check:
   - **Fidelity** — every number, label, quote and name matches the content exactly. Never change a number, even when the content looks inconsistent; flag it in `REVIEW.md` instead. Nothing invented.
   - **Form** — the visual form fits the message the way the reference deck would show it: comparisons and distributions as charts in the deck's chart style, not tables, unless the reference uses a table for that shape. No content forced into a layout made for a different shape.
   - **Message** — headline and insight state what the slide shows, written the way the reference deck writes them.
   - **Slots** — no empty picture or text slots; a slot without matching content → a layout or module without that slot. Use an available content picture only where it truly belongs.
   - **Craft** — zero overflow, no collisions (labels vs bars, text vs frames), minimum type size and spacing per `DESIGN.md`, colours from `tokens.json` only, consistent chrome (breadcrumb, page number, tracker), pictures framed and fitted as in the reference.
3. Fix in `deck.json` and `custom/*.mjs` (new modules from `lib` components and `tokens` only). Batch related fixes, then render the whole deck once per batch: `node "${CLAUDE_PLUGIN_ROOT}/tools/render-deck.mjs" --spec "<out>/deck.json" --out "<out>/decks/<n+1>"`.
4. Re-check only the slides you changed plus their neighbours. Repeat until no defect is left. Copy the final `deck.pptx` to `<out>/deck.pptx`.
5. Write `<out>/REVIEW.md`: per slide — defects found, fix made, anything left with its reason; plus content inconsistencies for the author.

## Rules

- Prefer the smallest change that removes the defect; do not rebuild slides that pass.
- Content beyond a slot's capacity → split or shorten by deleting words, never by shrinking type below `DESIGN.md` minimums. Quotes stay verbatim.
- Text as given; names and acronyms as spelled. Text stays native and editable.
- Never edit the pack.
- Finish with: defects found / fixed / left, the deck path, the REVIEW.md path.
