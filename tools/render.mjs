// node render.mjs --pack <packDir> --layout <id> [--data <file.json>] --out <dir> [--compare <src png>]
// Writes <out>/slide.pptx, <out>/render.png (2 px per pt), with --compare <out>/compare.png.
// Relative paths: args against cwd; image paths in data against the pack root, then the data file's dir, then cwd.
// stdout: one JSON line { ok, render, compare, ssim, meanAbsDiff, diffBoxes, overflow, errors, ms }.
import fs from 'node:fs';
import path from 'node:path';
import { parseArgs, resolvePath, readJson, loadPack, loadModule, newPres, renderSlide, writePptx,
  designedTables, ppExport, tableOverflow, runPy, rel } from './render-lib.mjs';

const t0 = Date.now();
const a = parseArgs(process.argv.slice(2));
const res = { ok: false, render: null, compare: null, ssim: null, meanAbsDiff: null, diffBoxes: [], overflow: [], errors: [] };

try {
  if (!a.pack || !a.layout || !a.out) throw new Error('usage: --pack <dir> --layout <id> [--data <json>] --out <dir> [--compare <png>]');
  const pack = resolvePath(a.pack);
  const layoutDir = path.join(pack, 'layouts', a.layout);
  const layoutFile = path.join(layoutDir, 'layout.mjs');
  if (!fs.existsSync(layoutFile)) throw new Error(`layout not found: ${rel(layoutFile)}`);
  const dataFile = a.data ? resolvePath(a.data) : path.join(layoutDir, 'sample.json');
  if (!fs.existsSync(dataFile)) throw new Error(`data not found: ${rel(dataFile)}`);
  const out = resolvePath(a.out);
  fs.mkdirSync(out, { recursive: true });
  for (const f of ['render.png', 'compare.png']) fs.rmSync(path.join(out, f), { force: true });

  const data = readJson(dataFile);
  const packCtx = await loadPack(pack);
  const mod = await loadModule(layoutFile);
  const pptx = newPres(packCtx.canvas);
  res.errors.push(...(await renderSlide(pptx, mod, data, packCtx, { base: pack, extraBases: [path.dirname(dataFile), process.cwd()] })));
  const pptxFile = path.join(out, 'slide.pptx');
  await writePptx(pptx, pptxFile, pack);

  const com = ppExport(pptxFile, out, true);
  const s1 = com.slides[0] || { i: 1, overflow: [], tables: [] };
  res.overflow = [...s1.overflow, ...tableOverflow(s1, await designedTables(pptxFile))];
  res.render = rel(path.join(out, 'render.png'));

  if (a.compare) {
    const src = resolvePath(a.compare);
    if (!fs.existsSync(src)) throw new Error(`compare source not found: ${a.compare}`);
    const cmpFile = path.join(out, 'compare.png');
    const c = JSON.parse(runPy('compare.py', [src, path.join(out, 'render.png'), '--out', cmpFile]));
    Object.assign(res, { compare: rel(cmpFile), ssim: c.ssim, meanAbsDiff: c.meanAbsDiff, diffBoxes: c.diffBoxes });
  }
} catch (e) {
  res.errors.push(String(e && e.message ? e.message : e));
}
res.ok = res.errors.length === 0;
res.ms = Date.now() - t0;
process.stdout.write(JSON.stringify(res) + '\n');
process.exitCode = res.ok ? 0 : 1;
