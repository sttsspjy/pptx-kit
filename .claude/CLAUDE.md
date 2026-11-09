# pptx-kit (this folder is a Claude Code plugin)

Best use: start Claude Code with the plugin loaded — `claude --plugin-dir .` from this folder — then use `/pptx-kit:setup`, `/pptx-kit:extract`, `/pptx-kit:consolidate`, `/pptx-kit:generate`, `/pptx-kit:review`. See README.md.

When this folder is opened without the plugin loaded (no `/pptx-kit:*` commands available):

- The skills are `skills/<name>/SKILL.md` (setup, extract, consolidate, generate, review). When the user asks for one of these jobs, Read that SKILL.md and follow it.
- In those files the plugin-root placeholder (`CLAUDE_PLUGIN_ROOT` in `${…}` form) means the absolute path of this folder.
- The per-slide rebuilder used by extract is `agents/slide-rebuilder.md`; spawn a general-purpose subagent told to Read and follow it.
- Run `node tools/doctor.mjs --install` before the first job.

Editing the plugin itself:

- `tools/*` CLIs print exactly one JSON line on stdout; keep that contract (skills parse it).
- Pack paths are relative to the pack root; tools take explicit paths; no hard-coded machine paths, no API keys, no network calls.
- Read JSON with UTF-8 explicitly in Python. PowerPoint renders are serialized by `%TEMP%\pptx-kit-pp.lock`; never call PowerPoint `Quit()` while a presentation is open.
- PptxGenJS 4.0.1 traps are listed in `reference/pptxgenjs.md`.
- Validate after changes: `claude plugin validate .claude-plugin/plugin.json`.
