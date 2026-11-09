// node intake.mjs --src <deck.pptx> --pack <packDir> [--allow-missing-fonts] [--per-sheet 20]
// Source intake for extraction:
//   <pack>/_work/source.pptx            copy of the source deck (removed by `pack.mjs prune --apply`)
//   <pack>/facts/sNN.json, facts/media/  every shape resolved (facts.py), media deduped by content
//   <pack>/facts/deck-stats.json         colours, faces, sizes, common geometry, canvas
//   <pack>/src/png/sNN.png               PowerPoint export of every slide, 2 px per pt (the truth for appearance)
//   <pack>/src/sheet-N.png               contact sheets, --per-sheet slides each (5 columns)
//   <pack>/pack.json                     name, source, canvas, fonts (merged into an existing pack.json)
// Refuses to export when a deck face is not installed (renders would use a substitute font) unless --allow-missing-fonts.
// stdout: one JSON line { ok, pack, slides, canvas, media, sheets, fonts, missingFonts, errors, ms }.
import fs from 'node:fs';
import path from 'node:path';
import { parseArgs, resolvePath, readJson, ppExport, runPy, rel } from './render-lib.mjs';

const t0 = Date.now();
const a = parseArgs(process.argv.slice(2));
const res = { ok: false, pack: null, slides: 0, canvas: null, media: 0, sheets: [], fonts: [], missingFonts: [], errors: [] };

try {
  if (!a.src || !a.pack) throw new Error('usage: --src <deck.pptx> --pack <packDir> [--allow-missing-fonts] [--per-sheet 20]');
  const src = resolvePath(a.src);
  if (!fs.existsSync(src)) throw new Error(`source not found: ${a.src}`);
  const pack = resolvePath(a.pack);
  res.pack = rel(pack);
  const work = path.join(pack, '_work');
  fs.mkdirSync(work, { recursive: true });
  const deck = path.join(work, 'source.pptx');
  fs.copyFileSync(src, deck);

  const factsDir = path.join(pack, 'facts');
  const f = JSON.parse(runPy('facts.py', [deck, factsDir]).split('\n').pop());
  res.slides = f.slides; res.media = f.media;
  const statsFile = path.join(factsDir, 'deck-stats.json');
  const stats = readJson(statsFile);
  res.canvas = stats.canvas;

  const fc = JSON.parse(runPy('fonts.py', ['--stats', statsFile]));
  const chars = Object.fromEntries((stats.faces || []).map((x) => [x.face, x.chars]));
  res.fonts = fc.faces.map((x) => ({ face: x.face, chars: chars[x.face] || 0, installed: x.installed }));
  res.missingFonts = fc.missing;
  if (fc.missing.length && !a['allow-missing-fonts']) {
    throw new Error(`fonts not installed: ${fc.missing.join(', ')} — install them (per-user install is enough), then rerun; or pass --allow-missing-fonts to accept substitute-font renders`);
  }

  const pngDir = path.join(pack, 'src', 'png');
  fs.rmSync(pngDir, { recursive: true, force: true });
  ppExport(deck, pngDir, false, { measure: false });

  const per = Math.max(1, parseInt(a['per-sheet'] || '20', 10));
  for (let s = 1, n = 1; s <= res.slides; s += per, n++) {
    const out = path.join(pack, 'src', `sheet-${n}.png`);
    runPy('sheet.py', [pngDir, out, '5', String(s), String(Math.min(res.slides, s + per - 1))]);
    res.sheets.push(rel(out));
  }

  const pj = path.join(pack, 'pack.json');
  const prev = fs.existsSync(pj) ? readJson(pj) : {};
  const prevFonts = Object.fromEntries((prev.fonts || []).map((x) => [x.face, x]));
  const meta = {
    schema: 'pptx-kit-pack/1',
    name: prev.name || path.basename(pack),
    ...prev,
    source: { file: path.basename(src), slides: res.slides, importedAt: new Date().toISOString() },
    canvas: res.canvas,
    fonts: res.fonts.map((x) => ({ face: x.face, chars: x.chars, source: prevFonts[x.face]?.source || null, licence: prevFonts[x.face]?.licence || null })),
  };
  fs.writeFileSync(pj, JSON.stringify(meta, null, 2) + '\n', 'utf8');
} catch (e) {
  res.errors.push(String(e && e.message ? e.message : e));
}
res.ok = res.errors.length === 0;
res.ms = Date.now() - t0;
process.stdout.write(JSON.stringify(res) + '\n');
process.exitCode = res.ok ? 0 : 1;
