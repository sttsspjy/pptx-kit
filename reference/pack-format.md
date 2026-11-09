# Pack format and tool contracts

A **pack** is one folder that holds a team's extracted design. It is self-contained: copy the folder to share it. Every path inside a pack is relative to the pack root (renders run with cwd = pack root).

```
<pack>/
  pack.json           name, source {file, slides, importedAt}, canvas {w, h} pt, fonts [{face, chars, source, licence}]
  DESIGN.md           the design system in measured numbers (written by consolidate)
  catalog.json        { canvas, layouts: [{ id, name, kind, purpose, useWhen, avoidWhen, capacities, imageSlots, dataFields, sourceSlide }] }
  tokens.json         colour roles, type scale, spacing, grid, radii, strokes, shadows
  lib/*.mjs           shared code: chrome, components (merged into ctx.lib)
  layouts/<id>/       layout.mjs, sample.json, stress-max.json, stress-min.json, SPEC.md, preview.png
  facts/sNN.json      resolved shape facts of the archetype slides; facts/media/ = pictures the pack uses
  facts/deck-stats.json
  src/png/sNN.png     PowerPoint renders of the archetype source slides (2 px per pt) — the truth for appearance
  src/sheet-N.png     contact sheets of the whole source deck
  assets/             optional extra pictures the pack draws (logos, icons) that are not in facts/media
  _work/              extraction scratch (source.pptx, archetypes.json, verify/) — removed by `pack.mjs prune --apply`, never shared
```

`<id>` = `sNN-<short-purpose>` (e.g. `s11-quote-columns`), NN = source slide number. Ids starting with `_` (e.g. `_demo`) are not layouts.

## Layout module

`export default { id, name, purpose, useWhen, avoidWhen, schema, render(slide, data, ctx) }`

- `schema`: JSON Schema of `data` with `maxLength` / `minItems` / `maxItems` = verified capacities.
- `ctx = { pptx, tokens, lib, packDir, assetsDir, mediaDir, canvas, deps, imageSize, xmlPatch }`
  - `lib` = every export of every `lib/*.mjs` merged (a default export is keyed by file name).
  - `canvas` = `{ w, h }` in pt (pack.json, default 960×540 = PptxGenJS LAYOUT_WIDE 13.333×7.5 in).
  - `imageSize(path)` → `{ width, height }` px of an image (pack-relative or content-relative path).
  - `xmlPatch(baseName, fn)` → unique object name; pass it as `objectName`; after writing, `fn(elementXml)` returns the replacement XML of that `<p:sp>`/`<p:pic>`/`<p:graphicFrame>`. Use it for what PptxGenJS 4.0.1 cannot emit (separate latin/east-asian faces, `fontAlgn`, rounded picture clip, bullet font).
  - `deps` = `{ PptxGenJS, JSZip, imageSize }`. Pack code never imports npm packages itself (a pack must run from any install).
- Image paths in code and data: pack files as complete string literals `facts/media/<file>` / `assets/<file>` (prune keeps only files it finds referenced); content pictures relative to the content folder (deck spec `content`). Never absolute paths in a pack.
- A custom module (generation) has the same shape and lives in the deck folder (`custom/*.mjs`); it reaches pack code only through `ctx.lib` / `ctx.tokens`.

## CLIs (`K` = the plugin root; run with `node` / `python`)

All print ONE JSON line on stdout. Relative arguments resolve against the current directory.

| Command | Does |
|---|---|
| `node K/tools/doctor.mjs [--install] [--pack <pack>]` | Environment check; `--install` installs missing Node/Python packages; `--pack` checks the pack's fonts |
| `node K/tools/intake.mjs --src <deck.pptx> --pack <pack> [--allow-missing-fonts]` | Source intake: facts, media, deck-stats, PowerPoint PNGs, contact sheets, pack.json. Refuses when a deck font is missing |
| `node K/tools/render.mjs --pack <pack> --layout <id> [--data <json>] --out <dir> [--compare <src png>]` | Render one layout → `slide.pptx`, `render.png`; with `--compare` also `compare.png` (2×2: source / render / diff heatmap + numbered boxes / 50 % blend) and `ssim`, `meanAbsDiff`, `diffBoxes` [{x,y,w,h,score}] in pt; `overflow` [{name, text, boxH, textH}] from PowerPoint's own text measurement |
| `node K/tools/render-deck.mjs --spec <deck.json> --out <dir>` | Render a deck → `deck.pptx`, `png/sNN.png`, `sheet.png`; per-slide `overflow` and `errors` |
| `node K/tools/pack.mjs check <pack>` | Self-containment check (loads every layout, missing referenced files, absolute paths, catalog ↔ layouts, fonts) |
| `node K/tools/pack.mjs prune <pack> [--apply] [--keep-src-all]` | Shrink a finished pack for sharing (dry run without `--apply`) |
| `python K/tools/crop.py <png> <x> <y> <w> <h> [--scale S] --out <file>` | Crop (px) + zoom a region to view it |
| `python K/tools/sheet.py <pngDir> <out.png> <cols> <start> <end>` | Contact sheet |
| `python K/tools/imginfo.py <file or dir>...` | Pixel size + aspect of pictures |
| `python K/tools/fonts.py --faces "<face>"...` or `--pack <pack>` | Installed-font check |

Deck spec (`deck.json`): `{ "pack": "<pack>", "content": "<content dir>", "slides": [ {"layout": "<id>", "data": {…}} | {"module": "custom/<name>.mjs", "data": {…}} ] }`. `pack`, `content`, `module` and a string `data` (JSON file) are relative to the spec file. Image paths in data resolve against the pack root, then `content`, then the spec folder.

PowerPoint renders are serialized across processes by a lock (`%TEMP%\pptx-kit-pp.lock`); parallel callers queue. A render takes ~2–5 s per slide. Non-ASCII paths are staged through an ASCII temp folder automatically.

## Slide PNG coordinates

Renders are 2 px per pt: a 960×540 pt slide is 1920×1080 px. `crop.py` takes px; facts, SPEC.md and code use pt.
