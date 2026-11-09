// node doctor.mjs [--install] [--pack <packDir>]
// Environment check for pptx-kit: Windows, Node >= 20, Node packages (plugin root node_modules), Python 3 + packages,
// PowerPoint COM, ASCII working dir, and (with --pack) the pack's fonts.
// --install: `npm ci` in the plugin root when packages are missing; `pip install --user -r requirements.txt` when Python packages are missing.
// stdout: one JSON line { ok, checks: [{name, ok, detail, fix?}] }. No imports of the Node packages it checks.
import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const TOOLS = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.dirname(TOOLS);
const args = process.argv.slice(2);
const install = args.includes('--install');
const packArg = args.includes('--pack') ? args[args.indexOf('--pack') + 1] : null;
const checks = [];
const add = (name, ok, detail, fix) => checks.push({ name, ok, detail, ...(ok || !fix ? {} : { fix }) });
const run = (cmd, a, o = {}) => spawnSync(cmd, a, { encoding: 'utf8', windowsHide: true, shell: false, ...o });

add('windows', process.platform === 'win32', process.platform, 'pptx-kit needs Windows with desktop PowerPoint');

const major = +process.versions.node.split('.')[0];
add('node', major >= 20, process.versions.node, 'install Node.js 20 or newer');

const PKGS = ['pptxgenjs', 'jszip', 'image-size'];
const missingPkgs = () => PKGS.filter((p) => !fs.existsSync(path.join(ROOT, 'node_modules', p, 'package.json')));
let miss = missingPkgs();
if (miss.length && install) {
  const npm = process.platform === 'win32' ? 'npm.cmd' : 'npm';
  const hasLock = fs.existsSync(path.join(ROOT, 'package-lock.json'));
  const r = run(npm, [hasLock ? 'ci' : 'install', '--no-audit', '--no-fund', '--omit=dev'], { cwd: ROOT, shell: true });
  if (r.status !== 0) add('npm install', false, (r.stderr || r.stdout || '').trim().slice(-400), `run npm ci in ${ROOT}`);
  miss = missingPkgs();
}
add('node packages', !miss.length, miss.length ? `missing: ${miss.join(', ')}` : PKGS.join(', '), `run the setup skill with install, or npm ci in ${ROOT}`);

let py = null;
const cands = [];
if (process.env.PPTX_KIT_PYTHON) cands.push([process.env.PPTX_KIT_PYTHON, []]);
cands.push(['python', []], ['py', ['-3']], ['python3', []]);
for (const [cmd, pre] of cands) {
  const r = run(cmd, [...pre, '-c', 'import sys; print("%d.%d" % sys.version_info[:2])']);
  if (r.status === 0 && r.stdout.trim().startsWith('3.')) { py = { cmd, pre, version: r.stdout.trim() }; break; }
}
add('python', !!py, py ? `${[py.cmd, ...py.pre].join(' ')} ${py.version}` : 'not found', 'install Python 3.10+ (python.org), or set PPTX_KIT_PYTHON to its python.exe');

const MODS = [['PIL', 'pillow'], ['numpy', 'numpy'], ['pptx', 'python-pptx'], ['skimage', 'scikit-image'], ['scipy', 'scipy'], ['lxml', 'lxml']];
const missingMods = () => MODS.filter(([m]) => run(py.cmd, [...py.pre, '-c', `import ${m}`]).status !== 0).map(([, p]) => p);
if (py) {
  let mm = missingMods();
  if (mm.length && install) {
    const r = run(py.cmd, [...py.pre, '-m', 'pip', 'install', '--user', '-r', path.join(ROOT, 'requirements.txt')]);
    if (r.status !== 0) add('pip install', false, (r.stderr || r.stdout || '').trim().slice(-400), `${py.cmd} -m pip install --user -r "${path.join(ROOT, 'requirements.txt')}"`);
    mm = missingMods();
  }
  add('python packages', !mm.length, mm.length ? `missing: ${mm.join(', ')}` : MODS.map(([, p]) => p).join(', '),
    `${py.cmd} -m pip install --user -r "${path.join(ROOT, 'requirements.txt')}"`);
}

if (process.platform === 'win32') {
  const ps = [
    '$ErrorActionPreference = "Stop"',
    '$a = New-Object -ComObject PowerPoint.Application',
    '$v = $a.Version',
    'if ($a.Presentations.Count -eq 0) { $a.Quit() }',
    '[void][Runtime.InteropServices.Marshal]::ReleaseComObject($a)',
    'Write-Output $v',
  ].join('; ');
  const r = run('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', ps], { timeout: 120000 });
  const v = (r.stdout || '').trim();
  add('powerpoint', r.status === 0 && /^\d+/.test(v), r.status === 0 ? `PowerPoint ${v}` : (r.stderr || '').trim().slice(0, 300),
    'install desktop Microsoft PowerPoint (Microsoft 365 / Office); the web or Store-sandboxed app has no COM');
}

const cwd = process.cwd();
add('ascii working dir', !/[^\x20-\x7E]/.test(cwd), cwd, 'work in a folder whose path is ASCII (renders are staged through %TEMP% otherwise)');

if (packArg && py) {
  const r = run(py.cmd, [...py.pre, path.join(TOOLS, 'fonts.py'), '--pack', path.resolve(packArg)], { env: { ...process.env, PYTHONUTF8: '1' } });
  try {
    const fc = JSON.parse(r.stdout.trim());
    let src = {};
    const pj = path.join(path.resolve(packArg), 'pack.json');
    if (fs.existsSync(pj)) src = Object.fromEntries((JSON.parse(fs.readFileSync(pj, 'utf8').replace(/^﻿/, '')).fonts || []).map((f) => [f.face, f.source]));
    add('pack fonts', !fc.missing.length, fc.missing.length ? `missing: ${fc.missing.join(', ')}` : `${fc.faces.length} faces installed`,
      fc.missing.map((f) => `${f}: ${src[f] || 'ask the pack owner where to get it'}`).join(' | '));
  } catch {
    add('pack fonts', false, (r.stderr || r.stdout || '').trim().slice(-300));
  }
}

const ok = checks.every((c) => c.ok || c.name === 'ascii working dir');
process.stdout.write(JSON.stringify({ ok, root: ROOT.split(path.sep).join('/'), checks }) + '\n');
process.exitCode = ok ? 0 : 1;
