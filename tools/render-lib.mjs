// Shared by render.mjs / render-deck.mjs / intake.mjs: pack loading, slide building, PowerPoint export, overflow, compare.
// Path base = the pack root: layouts and lib reference pack files as "facts/media/x.png", "assets/x.png" (cwd = pack root while rendering).
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';
import PptxGenJS from 'pptxgenjs';
import JSZip from 'jszip';
import imageSizeMod from 'image-size';

console.log = (...a) => console.error(...a); // stdout carries exactly one JSON line

export const TOOLS = path.dirname(fileURLToPath(import.meta.url));
const imageSize = imageSizeMod.imageSize || imageSizeMod.default || imageSizeMod;

// Arg value -> absolute path (cwd-relative unless absolute).
export function resolvePath(p, base = process.cwd()) {
  if (!p || p === true) return p;
  return path.isAbsolute(p) ? p : path.resolve(base, p);
}

export function parseArgs(argv) {
  const out = {};
  for (let i = 0; i < argv.length; i++) {
    if (!argv[i].startsWith('--')) continue;
    const k = argv[i].slice(2);
    const v = argv[i + 1] && !argv[i + 1].startsWith('--') ? argv[++i] : true;
    out[k] = v;
  }
  return out;
}

export function readJson(p) {
  return JSON.parse(fs.readFileSync(p, 'utf8').replace(/^\uFEFF/, ''));
}

// pack.json canvas (pt) or the 960x540 default (PptxGenJS LAYOUT_WIDE).
export function packCanvas(packDir) {
  const f = packDir && path.join(packDir, 'pack.json');
  if (f && fs.existsSync(f)) {
    const c = readJson(f).canvas;
    if (c && c.w > 0 && c.h > 0) return { w: c.w, h: c.h };
  }
  return { w: 960, h: 540 };
}

// ctx = { pptx, tokens, lib, packDir, assetsDir, mediaDir, canvas, deps, imageSize, xmlPatch } for a pack.
export async function loadPack(packDir) {
  const ctx = { tokens: {}, lib: {}, packDir: packDir || null, canvas: packCanvas(packDir) };
  if (packDir) {
    const tokPath = path.join(packDir, 'tokens.json');
    if (fs.existsSync(tokPath)) ctx.tokens = readJson(tokPath);
    const libDir = path.join(packDir, 'lib');
    if (fs.existsSync(libDir)) {
      for (const f of fs.readdirSync(libDir).filter((f) => f.endsWith('.mjs')).sort()) {
        const mod = await import(pathToFileURL(path.join(libDir, f)).href);
        for (const [k, v] of Object.entries(mod)) {
          if (k === 'default') ctx.lib[path.basename(f, '.mjs')] = v;
          else ctx.lib[k] = v;
        }
      }
    }
    ctx.assetsDir = path.join(packDir, 'assets');
    ctx.mediaDir = path.join(packDir, 'facts', 'media');
  }
  return ctx;
}

export async function loadModule(file) {
  const mod = await import(pathToFileURL(file).href);
  const m = mod.default || mod;
  if (typeof m.render !== 'function') throw new Error(`${file}: default export has no render()`);
  return m;
}

export function newPres(canvas = { w: 960, h: 540 }) {
  const pptx = new PptxGenJS();
  if (canvas.w === 960 && canvas.h === 540) pptx.layout = 'LAYOUT_WIDE';
  else {
    pptx.defineLayout({ name: 'PACK', width: canvas.w / 72, height: canvas.h / 72 });
    pptx.layout = 'PACK';
  }
  pptx.__xmlPatches = new Map();
  pptx.__xmlSeq = 0;
  return pptx;
}

// Registers fn(elementXml) -> elementXml for the <p:sp>/<p:pic>/<p:graphicFrame> named `base#n`; returns that name (use it as objectName).
// Applied by writePptx after PptxGenJS writes (for what 4.0.1 cannot emit: separate latin/ea faces, fontAlgn, roundRect picture clip, bullet fonts).
export function xmlPatch(pptx, base, fn) {
  if (!pptx.__xmlPatches) { pptx.__xmlPatches = new Map(); pptx.__xmlSeq = 0; }
  const name = `${base}#${++pptx.__xmlSeq}`;
  pptx.__xmlPatches.set(name, fn);
  return name;
}

function applyPatch(xml, name, fn) {
  const esc = name.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const at = xml.indexOf(`name="${esc}"`);
  if (at < 0) return xml;
  const cands = ['<p:sp>', '<p:pic>', '<p:graphicFrame>', '<p:cxnSp>'].map((t) => [t, xml.lastIndexOf(t, at)]);
  const [tag, start] = cands.reduce((a, b) => (b[1] > a[1] ? b : a));
  if (start < 0) return xml;
  const close = tag.replace('<', '</');
  const end = xml.indexOf(close, at) + close.length;
  return xml.slice(0, start) + fn(xml.slice(start, end)) + xml.slice(end);
}

// Relative image path -> first existing of [base dirs]; absolute / data: untouched.
function resolveImage(p, bases) {
  if (!p || typeof p !== 'string' || p.startsWith('data:') || path.isAbsolute(p)) return p;
  for (const b of bases) {
    const c = path.resolve(b, p);
    if (fs.existsSync(c)) return c;
  }
  throw new Error(`image not found: ${p} (searched ${bases.join(' ; ')})`);
}

// Proxy: every object gets a meaningful objectName when the layout did not set one; relative image paths resolve against `bases`.
function snippet(text) {
  let s = '';
  if (typeof text === 'string') s = text;
  else if (Array.isArray(text)) s = text.map((t) => (typeof t === 'string' ? t : t?.text ?? '')).join('');
  return s.replace(/\s+/g, ' ').trim().slice(0, 24);
}
export function namedSlide(slide, prefix, bases) {
  const n = { text: 0, shape: 0, image: 0, table: 0, chart: 0 };
  const name = (kind, extra) => `${prefix}${kind}${++n[kind]}${extra ? ':' + extra : ''}`;
  const proxy = new Proxy(slide, {
    get(t, k) {
      const v = t[k];
      if (typeof v !== 'function') return v;
      if (k === 'addText') return (text, o = {}) => { t.addText(text, { objectName: name('text', snippet(text)), ...o }); return proxy; };
      if (k === 'addShape') return (s, o = {}) => { t.addShape(s, { objectName: name('shape', typeof s === 'string' ? s : ''), ...o }); return proxy; };
      if (k === 'addImage') return (o = {}) => {
        const p = o.path ? resolveImage(o.path, bases) : o.path;
        t.addImage({ objectName: name('image', path.basename(o.path || '')), ...o, ...(p ? { path: p } : {}) });
        return proxy;
      };
      if (k === 'addTable') return (rows, o = {}) => { t.addTable(rows, { objectName: name('table'), ...o }); return proxy; };
      if (k === 'addChart') return (type, d, o = {}) => { t.addChart(type, d, { objectName: name('chart'), ...o }); return proxy; };
      return v.bind(t);
    },
  });
  return proxy;
}

// Renders one slide object into pptx; returns error strings. base = pack root (cwd during render); extraBases = more dirs for image lookup.
export async function renderSlide(pptx, mod, data, packCtx, { base, extraBases = [], prefix = '' } = {}) {
  const root = base || packCtx.packDir || process.cwd();
  const bases = [root, ...extraBases];
  const slide = namedSlide(pptx.addSlide(), prefix, bases);
  const ctx = {
    pptx, ...packCtx,
    deps: { PptxGenJS, JSZip, imageSize },
    imageSize: (p) => imageSize(fs.readFileSync(resolveImage(p, bases))),
    xmlPatch: (b, fn) => xmlPatch(pptx, b, fn),
  };
  const cwd = process.cwd();
  process.chdir(root);
  try {
    await mod.render(slide, data, ctx);
    return [];
  } catch (e) {
    return [`render ${mod.id || ''}: ${e && e.stack ? e.stack.split('\n').slice(0, 4).join(' | ') : String(e)}`];
  } finally {
    process.chdir(cwd);
  }
}

export async function writePptx(pptx, file, base = process.cwd()) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const cwd = process.cwd();
  process.chdir(base);
  try {
    let buf = await pptx.write({ outputType: 'nodebuffer' });
    if (pptx.__xmlPatches && pptx.__xmlPatches.size) {
      const zip = await JSZip.loadAsync(buf);
      for (const f of Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))) {
        let xml = await zip.file(f).async('string');
        for (const [name, patch] of pptx.__xmlPatches) xml = applyPatch(xml, name, patch);
        zip.file(f, xml);
      }
      buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    }
    fs.writeFileSync(file, buf);
  } finally {
    process.chdir(cwd);
  }
}

// Designed table row heights (pt) per slide from the written pptx: [{slide, name, rows:[pt]}].
export async function designedTables(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const out = [];
  for (const k of Object.keys(zip.files)) {
    const m = k.match(/^ppt\/slides\/slide(\d+)\.xml$/);
    if (!m) continue;
    const xml = await zip.file(k).async('string');
    for (const gf of xml.split('<p:graphicFrame>').slice(1)) {
      if (!gf.includes('<a:tbl>')) continue;
      const name = (gf.match(/<p:cNvPr [^>]*name="([^"]*)"/) || [])[1] || '';
      const rows = [...gf.matchAll(/<a:tr h="(\d+)"/g)].map((r) => Math.round((+r[1] / 12700) * 100) / 100);
      out.push({ slide: +m[1], name: decodeXml(name), rows });
    }
  }
  return out;
}
function decodeXml(s) {
  return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
}

const nonAscii = (p) => /[^\x20-\x7E]/.test(p);
function asciiTmp() {
  for (const d of [path.join(os.tmpdir(), 'pptx-kit'), 'C:\\pptx-kit-tmp']) {
    if (nonAscii(d)) continue;
    fs.mkdirSync(d, { recursive: true });
    return fs.mkdtempSync(path.join(d, 'x-'));
  }
  throw new Error('no ASCII temp dir for PowerPoint staging');
}

// PowerPoint export (+ overflow unless measure=false). single -> <outDir>/render.png; else <outDir>/sNN.png.
// Non-ASCII paths are staged through an ASCII temp dir. Returns { slides:[{i, overflow, tables}], w, h, ms } or throws.
export function ppExport(pptxFile, outDir, single, { measure = true } = {}) {
  fs.mkdirSync(outDir, { recursive: true });
  const stage = nonAscii(pptxFile) || nonAscii(outDir) ? asciiTmp() : null;
  const inFile = stage ? path.join(stage, 'in.pptx') : pptxFile;
  const pngDir = stage ? path.join(stage, 'png') : outDir;
  if (stage) fs.copyFileSync(pptxFile, inFile);
  const json = path.join(stage || outDir, `.pp-${process.pid}.json`);
  const args = ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', path.join(TOOLS, 'pp-export.ps1'),
    '-Pptx', inFile, '-OutDir', pngDir, '-Json', json];
  if (single) args.push('-Single');
  if (!measure) args.push('-NoMeasure');
  try {
    const r = spawnSync('powershell.exe', args, { encoding: 'utf8', timeout: 1800000, windowsHide: true });
    if (r.status !== 0 || !fs.existsSync(json)) {
      throw new Error(`pp-export failed (${r.status}): ${(r.stderr || r.stdout || r.error || '').toString().trim().slice(0, 600)}`);
    }
    const res = readJson(json);
    if (stage) for (const f of fs.readdirSync(pngDir)) fs.copyFileSync(path.join(pngDir, f), path.join(outDir, f));
    return res;
  } finally {
    fs.rmSync(json, { force: true });
    if (stage) fs.rmSync(stage, { recursive: true, force: true });
  }
}

// Merge COM table heights with designed heights: rows that grew > 1 pt are overflow.
export function tableOverflow(comSlide, designed) {
  const out = [];
  for (const t of comSlide.tables || []) {
    const d = designed.find((x) => x.slide === comSlide.i && x.name === t.name);
    if (!d) continue;
    t.rows.forEach((h, r) => {
      const want = d.rows[r];
      if (want != null && h - want > 1) out.push({ name: `${t.name} row ${r + 1}`, text: t.text[r] || '', boxH: want, textH: h });
    });
  }
  return out;
}

// Python launcher: $PPTX_KIT_PYTHON, else `python`, else `py -3` (first that runs Python 3).
let PY = null;
export function python() {
  if (PY) return PY;
  const cands = [];
  if (process.env.PPTX_KIT_PYTHON) cands.push([process.env.PPTX_KIT_PYTHON, []]);
  cands.push(['python', []], ['py', ['-3']], ['python3', []]);
  for (const [cmd, pre] of cands) {
    const r = spawnSync(cmd, [...pre, '-c', 'import sys; print(sys.version_info[0])'], { encoding: 'utf8', windowsHide: true });
    if (r.status === 0 && r.stdout.trim() === '3') return (PY = [cmd, pre]);
  }
  throw new Error('Python 3 not found (set PPTX_KIT_PYTHON)');
}

export function runPy(script, args, timeout = 1800000) {
  const [cmd, pre] = python();
  const r = spawnSync(cmd, [...pre, path.join(TOOLS, script), ...args],
    { encoding: 'utf8', timeout, windowsHide: true, maxBuffer: 64 << 20, env: { ...process.env, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1' } });
  if (r.status !== 0) throw new Error(`${script} failed: ${(r.stderr || r.stdout || String(r.error || '')).trim().slice(-600)}`);
  return r.stdout.trim();
}

// Output path for JSON: relative to cwd when inside it, forward slashes.
export function rel(p) {
  if (!p) return p;
  const r = path.relative(process.cwd(), p);
  return (r.startsWith('..') || path.isAbsolute(r) ? p : r).split(path.sep).join('/');
}
