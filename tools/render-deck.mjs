// node render-deck.mjs --spec <deck.json> --out <dir>
// spec { "pack": "<packDir>", "content"?: "<contentDir>", "slides": [ {"layout":"<id>","data":{...}} | {"module":"<path.mjs>","data":{...}} ] }
// Paths in spec (pack, content, module, data file): absolute, else relative to the spec file's dir. "data" may be a JSON file path.
// Image paths in data: absolute, else first existing under [pack root, content dir, spec dir, cwd].
// Writes <out>/deck.pptx, <out>/png/sNN.png, <out>/sheet.png. stdout: one JSON line { ok, pptx, sheet, slides:[{i, overflow, errors}], errors, ms }.
import fs from 'node:fs';
import path from 'node:path';
import { parseArgs, resolvePath, readJson, loadPack, loadModule, newPres, renderSlide, writePptx,
  designedTables, ppExport, tableOverflow, runPy, rel } from './render-lib.mjs';

const t0 = Date.now();
const a = parseArgs(process.argv.slice(2));
const res = { ok: false, pptx: null, sheet: null, slides: [], errors: [] };

try {
  if (!a.spec || !a.out) throw new Error('usage: --spec <deck.json> --out <dir>');
  const specFile = resolvePath(a.spec);
  const specDir = path.dirname(specFile);
  const spec = readJson(specFile);
  if (!Array.isArray(spec.slides) || !spec.slides.length) throw new Error('spec.slides is empty');
  const fromSpec = (p) => resolvePath(p, specDir);
  const pack = spec.pack ? fromSpec(spec.pack) : null;
  if (pack && !fs.existsSync(pack)) throw new Error(`pack not found: ${pack}`);
  const content = spec.content ? fromSpec(spec.content) : null;
  const packCtx = await loadPack(pack);
  const base = pack || specDir;
  const extraBases = [content, specDir, process.cwd()].filter(Boolean);
  const out = resolvePath(a.out);
  const pngDir = path.join(out, 'png');
  fs.rmSync(pngDir, { recursive: true, force: true });
  fs.mkdirSync(pngDir, { recursive: true });

  const pptx = newPres(packCtx.canvas);
  for (let i = 0; i < spec.slides.length; i++) {
    const s = spec.slides[i];
    const entry = { i: i + 1, overflow: [], errors: [] };
    res.slides.push(entry);
    try {
      let file;
      if (s.layout) {
        if (!pack) throw new Error('spec.pack required for "layout" slides');
        file = path.join(pack, 'layouts', s.layout, 'layout.mjs');
      } else if (s.module) file = fromSpec(s.module);
      else throw new Error('slide needs "layout" or "module"');
      if (!fs.existsSync(file)) throw new Error(`not found: ${rel(file)}`);
      const data = typeof s.data === 'string' ? readJson(fromSpec(s.data)) : (s.data ?? {});
      const mod = await loadModule(file);
      entry.errors.push(...(await renderSlide(pptx, mod, data, packCtx, { base, extraBases, prefix: `s${String(i + 1).padStart(2, '0')}-` })));
    } catch (e) {
      pptx.addSlide();
      entry.errors.push(String(e && e.message ? e.message : e));
    }
  }
  const pptxFile = path.join(out, 'deck.pptx');
  await writePptx(pptx, pptxFile, base);
  res.pptx = rel(pptxFile);

  const com = ppExport(pptxFile, pngDir, false);
  const designed = await designedTables(pptxFile);
  for (const cs of com.slides) {
    const entry = res.slides[cs.i - 1];
    if (entry) entry.overflow = [...cs.overflow, ...tableOverflow(cs, designed)];
  }
  const sheet = path.join(out, 'sheet.png');
  runPy('sheet.py', [pngDir, sheet, '4', '1', String(spec.slides.length)]);
  res.sheet = rel(sheet);
} catch (e) {
  res.errors.push(String(e && e.message ? e.message : e));
}
res.ok = res.errors.length === 0 && res.slides.every((s) => s.errors.length === 0);
res.ms = Date.now() - t0;
process.stdout.write(JSON.stringify(res) + '\n');
process.exitCode = res.ok ? 0 : 1;
