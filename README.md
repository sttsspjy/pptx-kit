# pptx-kit

A Claude Code plugin that learns a team's PowerPoint design from one of their original decks and then builds new decks in that design.

1. **extract** — rebuilds the deck's distinct slide types as parametric, editable templates, each verified against PowerPoint's own render of the source slide, and writes a shareable **design pack** (design system document, layout catalogue, tokens, components, charts in the deck's style).
2. **generate** — turns report content (a Mirae slide-content export, a structured `content.json`, or plain markdown) into a finished deck in that design.
3. **review** — checks a generated deck slide by slide against the pack and the content, fixes it, and writes `REVIEW.md`.

All model work runs in your own Claude Code session on your own subscription. There is no API key, no SDK and no network call in the plugin; PowerPoint, Node and Python run locally. The only downloads are the npm / pip packages that `setup` installs.

## Requirements

- Windows 10/11 with **desktop PowerPoint** (Microsoft 365 or Office 2016+). PowerPoint is the render oracle: every slide is measured with PowerPoint's own text layout and exported by PowerPoint.
- Node.js 20+.
- Python 3.10+ with pillow, numpy, python-pptx, scikit-image, scipy, lxml (`requirements.txt`; `setup` installs them).
- The fonts of the deck. A pack lists its fonts with a download source in `pack.json`; install them (per-user install is enough) before extracting or generating, or renders use a substitute font.

## Install

Pick one:

- **Install from this folder (persistent)** — in your shell:
  1. `claude plugin marketplace add C:\path\to\pptx-kit`
  2. `claude plugin install pptx-kit@pptx-kit-local`
  3. Start Claude Code and run `/pptx-kit:setup`.
- **One session** — `claude --plugin-dir C:\path\to\pptx-kit`, then `/pptx-kit:setup`.
- **Every session without a marketplace** — move the folder under `%USERPROFILE%\.claude\skills\` and restart Claude Code.

A local-folder marketplace loads the plugin in place and does not install Node packages for you: `/pptx-kit:setup` runs `npm ci` in the plugin folder the first time.

## Use

| Command | What it does |
|---|---|
| `/pptx-kit:setup [packDir]` | Checks Windows, Node, Python packages, PowerPoint COM and (with a pack) its fonts; installs missing packages |
| `/pptx-kit:extract <source.pptx> <packDir>` | Builds a design pack: intake → archetype pick (you confirm the list) → parallel per-slide rebuild subagents → verification → consolidate → prune |
| `/pptx-kit:consolidate <packDir>` | Re-runs only the pack-level layer (DESIGN.md, catalog.json, tokens.json, components, chart style) |
| `/pptx-kit:generate <packDir> <contentDir> [outDir]` | Builds `deck.pptx` + `REPORT.md` from the content in the pack's design |
| `/pptx-kit:review <packDir> <contentDir> <deckDir>` | Finds and fixes defects in a generated deck, writes `REVIEW.md` |

Extraction is the expensive step: each slide type is one long subagent run (a full deck typically has 10–25 types). It is done once per team deck; the pack is then reused for every new deck. Generation reads the pack once up front and renders the deck with PowerPoint; a 15–25 slide deck takes roughly 20–40 minutes.

### Content folder (generate)

- `content.md` — required unless `content.json` carries everything: header + report in markdown.
- `content.json` — optional. Mirae export: `{"schema": "mirae-slide-content/1", "meta": {project, title, exportedAt, language, sessions[]}, "report", "quotes": [{text, speaker?, session, t?}], "media": [{path, kind, caption?, session?, t?}]}`. Structured content with `meta`, `sections`, `messages` is also accepted.
- `media/` — pictures referenced by the content (paths relative to the content folder).

Numbers, labels and quotes are taken exactly as given; inconsistencies are flagged in `REPORT.md`, never changed.

### Design pack

A pack is one folder; copy it to share it. It contains `pack.json`, `DESIGN.md`, `catalog.json`, `tokens.json`, `lib/`, `layouts/<id>/` (template code, sample data, stress data, SPEC.md, preview), the source renders and facts of the archetype slides and the pictures the templates use. Font files are not included unless their licence allows it. Format details: [reference/pack-format.md](reference/pack-format.md).

A pack contains renders and pictures of the source deck: share it only with people allowed to see that deck.

## Privacy

Everything runs on your PC inside your Claude Code session: the source deck, the pack, the content and the output never leave the machine except as part of your normal Claude Code conversation with the model.

## Tools

`tools/` holds the CLIs the skills call (each prints one JSON line): `intake.mjs`, `render.mjs`, `render-deck.mjs`, `pack.mjs` (check / prune), `doctor.mjs`, `facts.py`, `compare.py`, `crop.py`, `sheet.py`, `imginfo.py`, `fonts.py`, `pp-export.ps1`. Smoke test: `node test/make-sample-deck.mjs work/sample.pptx` then `node tools/intake.mjs --src work/sample.pptx --pack work/pack`.

## Licence

No licence file yet — add one before distributing outside the team.
