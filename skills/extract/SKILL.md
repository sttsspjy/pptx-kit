---
name: extract
description: Extracts a team's design system from an original PowerPoint deck into a shareable pptx-kit design pack (source intake, archetype pick, parallel per-slide rebuild verified against PowerPoint renders, then consolidation into DESIGN.md, catalog.json, tokens.json and shared components). Long-running and usage-heavy; run only when the user asks to extract or build a pack.
argument-hint: "<source.pptx> <packDir> [--yes] [--max N]"
disable-model-invocation: true
---

# Extract a design pack from a source deck

Arguments: `$ARGUMENTS` → `<source.pptx> <packDir>`; `--yes` skips the archetype confirmation; `--max N` caps archetypes (default 25).

Plugin root: `${CLAUDE_PLUGIN_ROOT}`. Plugin files you open with Read (another SKILL.md, agents/*.md) still show the unresolved plugin-root placeholder; it stands for that same path. Read `${CLAUDE_PLUGIN_ROOT}/reference/pack-format.md` first (pack layout, CLIs). Everything runs locally: PowerPoint, Node, Python and this Claude Code session; nothing is uploaded anywhere.

Ask for the two paths if missing. Keep `<packDir>` on an ASCII path when possible. Use absolute paths in every command and in every task you give a subagent.

## 1. Preflight

`node "${CLAUDE_PLUGIN_ROOT}/tools/doctor.mjs"` — stop on any failed check except `ascii working dir` and explain the fix (see the setup skill: `/pptx-kit:setup`).

## 2. Intake

`node "${CLAUDE_PLUGIN_ROOT}/tools/intake.mjs" --src "<source.pptx>" --pack "<packDir>"`

- Missing fonts → stop. Tell the user which faces are missing; they install them (per-user install is enough) and you rerun. Only pass `--allow-missing-fonts` if the user explicitly accepts substitute-font renders (the pack will then be measured against wrong renders).
- Output: `facts/sNN.json`, `facts/media/`, `facts/deck-stats.json`, `src/png/sNN.png`, `src/sheet-N.png`, `pack.json`.

## 3. Archetype pick

1. Read `facts/deck-stats.json` and view every contact sheet `src/sheet-*.png` (full deck).
2. Classify every slide by visual structure and slide kind: cover, contents, section divider, subsection / interstitial, content variants (by body pattern: quote cards, table, screenshot + numbered findings, comparison columns, flow, matrix, image grid, charts drawn from shapes …), insert / example slides, closing. View a slide at full size (`src/png/sNN.png`) when a thumbnail is ambiguous.
3. Pick ONE representative per distinct structure — the instance with the most content (it defines capacity), clean of one-off decoration. Cover every slide kind the deck uses. Typical result: 10–25 archetypes; never more than `--max` (default 25) without the user's consent.
4. Pick the **seed**: a typical content slide with the full chrome (breadcrumb / header, tracker, page number). The seed's rebuilder owns `tokens.json` and `lib/` (chrome + shared helpers); all others only read them.
5. Write `<packDir>/_work/archetypes.json`:
   `{ "seed": 11, "archetypes": [ { "slide": 11, "id": "s11-quote-columns", "kind": "content", "why": "...", "alsoLike": [12, 14, 40] } ] }` (`id` = `sNN-<short-purpose>`, NN zero-padded to 2 as in the PNG names).
6. Unless `--yes`: show the table (slide, id, kind, why, how many slides it stands for) and the estimate — each archetype is one long subagent run (roughly 20–60 min wall time in parallel batches, and a large share of a subscription's usage window for a full deck). Ask the user to confirm or edit the list. Do not start rebuilding before the answer.

## 4. Rebuild (subagents)

Each archetype is rebuilt by the `pptx-kit:slide-rebuilder` agent. Task text per agent:

> Pack: `<abs packDir>`. Rebuild source slide sNN as layout `<id>` (kind: <kind>; similar slides: <alsoLike>). You [own | do not own] `tokens.json` and `lib/`. Follow your method; finish with the report it asks for.

1. Seed first, alone. When it finishes, check `tokens.json` and `lib/*.mjs` exist and the seed renders (`render.mjs --compare`).
2. Then all other archetypes, at most 4 agents running at once (PowerPoint renders queue on one lock; more agents only wait). Launch the next as one finishes.
3. If the `pptx-kit:slide-rebuilder` agent type is unavailable (pptx-kit opened as a plain folder), use a general-purpose agent and prefix the task with: "Read `${CLAUDE_PLUGIN_ROOT}/agents/slide-rebuilder.md` and follow it as your instructions; its unresolved plugin-root placeholder stands for `${CLAUDE_PLUGIN_ROOT}`."

## 5. Verify

For every layout, sequentially: `node "${CLAUDE_PLUGIN_ROOT}/tools/render.mjs" --pack "<packDir>" --layout <id> --out "<packDir>/_work/verify/<id>" --compare "<packDir>/src/png/sNN.png"`. Also render `stress-max.json` and `stress-min.json` with `--data` (overflow only). Build a table: id, ssim, overflow (sample / max / min), errors.

- Fails = errors, any overflow, or ssim < 0.95 → send that layout back to a rebuilder once, with the failing JSON lines: "Continue from the existing files in `layouts/<id>/`; fix: …". Re-verify. Report what still fails; do not loop further.
- View 3–4 `compare.png` files yourself (lowest ssim first) — a high score can hide a wrong line break.

## 6. Consolidate

Read `${CLAUDE_PLUGIN_ROOT}/skills/consolidate/SKILL.md` and follow it for `<packDir>` (it writes tokens.json final, lib/components.mjs, lib/demo.mjs, catalog.json, DESIGN.md, pack.json fonts).

## 7. Finish

1. `node "${CLAUDE_PLUGIN_ROOT}/tools/pack.mjs" prune "<packDir>"` (dry run) — show removed MB; then `node "${CLAUDE_PLUGIN_ROOT}/tools/pack.mjs" prune "<packDir>" --apply`.
2. `node "${CLAUDE_PLUGIN_ROOT}/tools/pack.mjs" check "<packDir>"` — fix every error.
3. Report: archetypes and their final ssim / overflow, components added, unresolved items, pack size, fonts the receiver must install (with sources). Tell the user the pack folder is the shareable unit (copy the folder) and that the next step is `/pptx-kit:generate <packDir> <contentDir>`.
