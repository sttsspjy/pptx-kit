// node pack.mjs check <packDir>
//   Self-containment + shape check. Errors: missing tokens.json / layouts, layout without render()/schema/sample.json,
//   referenced pack file (facts/media/…, assets/…, src/png/…) missing, absolute paths in lib/layouts, catalog ids without a layout.
//   Warnings: missing DESIGN.md / catalog.json / pack.json / SPEC.md / preview, layouts not in catalog, fonts not installed here.
// node pack.mjs prune <packDir> [--apply] [--keep-src-all]
//   Shrinks a finished pack to what generation needs. Removes _work/, decks/, every out/ dir (layouts/<id>/out/final/render.png,
//   else the newest render.png under out/, is kept as layouts/<id>/preview.png), facts/media files nothing references,
//   facts/sNN.json and src/png/sNN.png of slides no layout came from (all src/png kept with --keep-src-all). Dry run without --apply.
// stdout: one JSON line.
import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { parseArgs, resolvePath, readJson, runPy, rel } from './render-lib.mjs';

const [cmd, packArg] = process.argv.slice(2);
const a = parseArgs(process.argv.slice(4));
const res = { ok: false, cmd, pack: null, errors: [], warnings: [] };

const walk = (d, out = []) => {
  if (!fs.existsSync(d)) return out;
  for (const e of fs.readdirSync(d, { withFileTypes: true })) {
    const p = path.join(d, e.name);
    if (e.isDirectory()) walk(p, out); else out.push(p);
  }
  return out;
};
const size = (files) => files.reduce((s, f) => s + (fs.existsSync(f) ? fs.statSync(f).size : 0), 0);
const MB = (b) => Math.round((b / 1048576) * 10) / 10;

// Text files whose string literals may reference pack files (outputs under out/ excluded).
function sourceTexts(pack) {
  const files = [
    ...walk(path.join(pack, 'lib')),
    ...walk(path.join(pack, 'layouts')).filter((f) => !f.split(path.sep).includes('out')),
    path.join(pack, 'tokens.json'), path.join(pack, 'catalog.json'), path.join(pack, 'DESIGN.md'),
  ];
  return files.filter((f) => fs.existsSync(f) && /\.(mjs|js|json|md)$/.test(f)).map((f) => ({ f, t: fs.readFileSync(f, 'utf8') }));
}
const REF = /(?:^|['"`(\s])((?:facts\/media|assets|src\/png)\/[^'"`)\s]+\.[A-Za-z0-9]+)/g;

function references(texts) {
  const refs = new Map();
  for (const { f, t } of texts) for (const m of t.matchAll(REF)) {
    if (/NN|XX|[*{<]/.test(m[1])) continue; // placeholders such as src/png/sNN.png
    if (!refs.has(m[1])) refs.set(m[1], new Set());
    refs.get(m[1]).add(f);
  }
  return refs;
}

function layoutIds(pack) {
  const d = path.join(pack, 'layouts');
  return fs.existsSync(d) ? fs.readdirSync(d, { withFileTypes: true }).filter((e) => e.isDirectory()).map((e) => e.name) : [];
}

function sourceSlides(pack) {
  const set = new Set();
  for (const id of layoutIds(pack)) { const m = id.match(/^s(\d+)-/); if (m) set.add(+m[1]); }
  const cat = path.join(pack, 'catalog.json');
  if (fs.existsSync(cat)) {
    const c = readJson(cat);
    for (const e of Array.isArray(c) ? c : c.layouts || []) {
      const m = String(e.sourceSlide ?? '').match(/(\d+)/);
      if (m) set.add(+m[1]);
    }
  }
  return set;
}

function catalogIds(pack) {
  const cat = path.join(pack, 'catalog.json');
  if (!fs.existsSync(cat)) return null;
  const c = readJson(cat);
  return (Array.isArray(c) ? c : c.layouts || []).map((e) => e.id).filter(Boolean);
}

async function check(pack) {
  const ids = layoutIds(pack).filter((id) => !id.startsWith('_'));
  if (!fs.existsSync(path.join(pack, 'tokens.json'))) res.errors.push('tokens.json missing');
  if (!ids.length) res.errors.push('no layouts/');
  for (const f of ['DESIGN.md', 'catalog.json', 'pack.json']) if (!fs.existsSync(path.join(pack, f))) res.warnings.push(`${f} missing`);
  res.layouts = [];
  for (const id of ids) {
    const d = path.join(pack, 'layouts', id);
    const row = { id };
    try {
      const mod = await import(pathToFileURL(path.join(d, 'layout.mjs')).href);
      const m = mod.default || mod;
      if (typeof m.render !== 'function') res.errors.push(`${id}: no render()`);
      if (!m.schema) res.errors.push(`${id}: no schema`);
      if (m.id && m.id !== id) res.warnings.push(`${id}: module id "${m.id}" differs from folder`);
      for (const k of ['purpose', 'useWhen', 'avoidWhen']) if (!m[k]) res.warnings.push(`${id}: no ${k}`);
    } catch (e) {
      res.errors.push(`${id}: layout.mjs does not load: ${String(e.message || e).slice(0, 200)}`);
    }
    if (!fs.existsSync(path.join(d, 'sample.json'))) res.errors.push(`${id}: sample.json missing`);
    if (!fs.existsSync(path.join(d, 'SPEC.md'))) res.warnings.push(`${id}: SPEC.md missing`);
    row.preview = fs.existsSync(path.join(d, 'preview.png'));
    res.layouts.push(row);
  }
  const cat = catalogIds(pack);
  if (cat) {
    for (const id of cat) if (!ids.includes(id)) res.errors.push(`catalog.json: "${id}" has no layout folder`);
    for (const id of ids) if (!cat.includes(id)) res.warnings.push(`layout ${id} not in catalog.json`);
  }
  const texts = sourceTexts(pack);
  for (const [ref, from] of references(texts)) {
    if (!fs.existsSync(path.join(pack, ref))) res.errors.push(`missing pack file ${ref} (referenced by ${[...from].map(rel).slice(0, 3).join(', ')})`);
  }
  for (const { f, t } of texts) {
    if (!/\.(mjs|js)$/.test(f)) continue;
    const m = t.match(/['"`]([A-Za-z]:[\\/][^'"`]*|\/(?:Users|home)\/[^'"`]*)['"`]/);
    if (m) res.errors.push(`absolute path in ${rel(f)}: ${m[1].slice(0, 80)} — packs must be relative to the pack root`);
  }
  try {
    const fc = JSON.parse(runPy('fonts.py', ['--pack', pack]));
    res.fonts = fc.faces.map((x) => ({ face: x.face, installed: x.installed }));
    if (fc.missing.length) res.warnings.push(`fonts not installed on this PC: ${fc.missing.join(', ')}`);
  } catch (e) {
    res.warnings.push(`font check failed: ${e.message}`);
  }
  const pj = path.join(pack, 'pack.json');
  if (fs.existsSync(pj)) {
    const fonts = readJson(pj).fonts || [];
    const noSrc = fonts.filter((x) => !x.source).map((x) => x.face);
    if (noSrc.length) res.warnings.push(`pack.json fonts without "source": ${noSrc.join(', ')}`);
  }
  res.sizeMB = MB(size(walk(pack)));
}

function prune(pack) {
  const apply = !!a.apply;
  const remove = [];
  const keepPreview = [];
  for (const d of ['_work', 'decks']) if (fs.existsSync(path.join(pack, d))) remove.push(path.join(pack, d));
  const outDirs = [];
  const findOut = (d) => {
    if (!fs.existsSync(d)) return;
    for (const e of fs.readdirSync(d, { withFileTypes: true })) {
      if (!e.isDirectory()) continue;
      const p = path.join(d, e.name);
      if (e.name === 'out') outDirs.push(p); else findOut(p);
    }
  };
  findOut(path.join(pack, 'layouts'));
  findOut(path.join(pack, 'lib'));
  for (const o of outDirs) {
    const layoutDir = path.dirname(o);
    if (path.dirname(layoutDir) === path.join(pack, 'layouts') && !path.basename(layoutDir).startsWith('_')) {
      const fin = path.join(o, 'final', 'render.png');
      const renders = walk(o).filter((f) => path.basename(f) === 'render.png' && !/stress/i.test(f));
      const pick = fs.existsSync(fin) ? fin : renders.sort((x, y) => fs.statSync(y).mtimeMs - fs.statSync(x).mtimeMs)[0];
      if (pick) keepPreview.push([pick, path.join(layoutDir, 'preview.png')]);
    }
    remove.push(o);
  }
  const refs = references(sourceTexts(pack));
  for (const f of walk(path.join(pack, 'facts', 'media'))) {
    if (!refs.has(`facts/media/${path.basename(f)}`)) remove.push(f);
  }
  const keep = sourceSlides(pack);
  for (const r of refs.keys()) { const m = r.match(/^src\/png\/s(\d+)\.png$/); if (m) keep.add(+m[1]); }
  for (const f of walk(path.join(pack, 'facts')).filter((f) => /[\\/]s\d+\.json$/.test(f))) {
    if (!keep.has(+path.basename(f).match(/\d+/)[0])) remove.push(f);
  }
  if (!a['keep-src-all']) {
    for (const f of walk(path.join(pack, 'src', 'png'))) {
      const m = path.basename(f).match(/^s(\d+)\.png$/);
      if (m && !keep.has(+m[1])) remove.push(f);
    }
  }
  const before = size(walk(pack));
  const removedBytes = remove.reduce((s, p) => s + (fs.statSync(p).isDirectory() ? size(walk(p)) : fs.statSync(p).size), 0);
  if (apply) {
    for (const [from, to] of keepPreview) fs.copyFileSync(from, to);
    for (const p of remove) fs.rmSync(p, { recursive: true, force: true });
  }
  Object.assign(res, {
    applied: apply, previews: keepPreview.length, removed: remove.length,
    removedList: remove.slice(0, 40).map(rel), beforeMB: MB(before), afterMB: MB(before - removedBytes), keptSlides: [...keep].sort((x, y) => x - y),
  });
}

try {
  if (!['check', 'prune'].includes(cmd) || !packArg) throw new Error('usage: pack.mjs check|prune <packDir> [--apply] [--keep-src-all]');
  const pack = resolvePath(packArg);
  if (!fs.existsSync(pack)) throw new Error(`pack not found: ${packArg}`);
  res.pack = rel(pack);
  if (cmd === 'check') await check(pack); else prune(pack);
} catch (e) {
  res.errors.push(String(e && e.message ? e.message : e));
}
res.ok = res.errors.length === 0;
process.stdout.write(JSON.stringify(res) + '\n');
process.exitCode = res.ok ? 0 : 1;
