---
name: setup
description: Checks and prepares this PC for pptx-kit (Windows, Node 20+, Python 3 packages, desktop PowerPoint via COM, bundled Node packages, fonts of a design pack). Use before the first extract/generate, when a pptx-kit tool fails to start, or to check a shared pack's fonts.
argument-hint: "[packDir]"
---

# pptx-kit setup (doctor)

Arguments: `$ARGUMENTS` (optional pack folder whose fonts to check).

1. Run `node "${CLAUDE_PLUGIN_ROOT}/tools/doctor.mjs" --install` (add `--pack "<packDir>"` when a pack folder was given). `--install` runs `npm ci` in the plugin root when the bundled Node packages are missing and `pip install --user -r requirements.txt` when Python packages are missing — these package downloads are the only network use of pptx-kit.
2. Show the result as a short table: check, status, detail.
3. For every failed check, give its `fix` as a numbered step. Specifics:
   - **powerpoint**: desktop PowerPoint (Microsoft 365 / Office 2016+) must be installed and activated; PowerPoint for the web has no COM. pptx-kit uses PowerPoint as the render oracle, so there is no workaround.
   - **pack fonts**: each missing face with its `source` from the pack's `pack.json`. Installing per user is enough (right-click the font file → *Install*). Renders made without the font use a substitute and are wrong — install before generating.
   - **python**: if Python is installed but not on PATH, set `PPTX_KIT_PYTHON` to the full path of `python.exe`.
   - **ascii working dir**: a warning only; tools stage non-ASCII paths through `%TEMP%`, but ASCII folders avoid surprises.
4. When everything passes, say so in one line and name the next step: `/pptx-kit:extract <deck.pptx> <packDir>` to build a pack, or `/pptx-kit:generate <packDir> <contentDir>` to make a deck.
