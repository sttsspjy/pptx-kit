// node test/make-sample-deck.mjs <out.pptx>
// Generic 6-slide test deck (cover, bullets, table, picture + caption, shapes + connectors, section divider) in Arial / Malgun Gothic.
// The picture is a generated gradient PNG written next to the deck. No external content.
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const req = createRequire(path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'package.json'));
const PptxGenJS = req('pptxgenjs');

function crc32(buf) {
  let c, crc = 0xffffffff;
  for (let n = 0; n < buf.length; n++) {
    c = (crc ^ buf[n]) & 0xff;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crc = (crc >>> 8) ^ c;
  }
  return (crc ^ 0xffffffff) >>> 0;
}
function png(w, h, px) {
  const raw = Buffer.alloc((w * 3 + 1) * h);
  for (let y = 0; y < h; y++) {
    raw[y * (w * 3 + 1)] = 0;
    for (let x = 0; x < w; x++) raw.set(px(x, y), y * (w * 3 + 1) + 1 + x * 3);
  }
  const chunk = (type, data) => {
    const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
    const td = Buffer.concat([Buffer.from(type), data]);
    const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(td));
    return Buffer.concat([len, td, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2;
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

const out = path.resolve(process.argv[2] || 'sample.pptx');
fs.mkdirSync(path.dirname(out), { recursive: true });
const img = path.join(path.dirname(out), 'sample-picture.png');
fs.writeFileSync(img, png(480, 320, (x, y) => [40 + Math.round((x / 480) * 160), 90 + Math.round((y / 320) * 120), 200 - Math.round((x / 480) * 80)]));

const C = { ink: '1F2937', muted: '6B7280', accent: '2563EB', soft: 'EFF6FF', line: 'D1D5DB' };
const F = { latin: 'Arial', ko: 'Malgun Gothic' };
const p = new PptxGenJS();
p.layout = 'LAYOUT_WIDE';
const chrome = (s, n, crumb) => {
  s.addText(crumb, { x: 0.5, y: 0.25, w: 8, h: 0.3, fontFace: F.ko, fontSize: 10, color: C.muted, margin: 0, valign: 'top' });
  s.addShape(p.ShapeType.line, { x: 0.5, y: 0.62, w: 12.33, h: 0, line: { color: C.line, width: 0.75 } });
  s.addText(String(n), { x: 12.33, y: 7.0, w: 0.5, h: 0.3, fontFace: F.latin, fontSize: 9, color: C.muted, align: 'right', margin: 0, valign: 'top' });
};

let s = p.addSlide();
s.background = { color: C.ink };
s.addText('Sample Research Report', { x: 0.8, y: 2.6, w: 11, h: 0.9, fontFace: F.latin, fontSize: 40, bold: true, color: 'FFFFFF', margin: 0, valign: 'top' });
s.addText('샘플 보고서 · 2026.11', { x: 0.8, y: 3.6, w: 11, h: 0.5, fontFace: F.ko, fontSize: 18, color: 'CBD5E1', margin: 0, valign: 'top' });
s.addShape(p.ShapeType.rect, { x: 0.8, y: 2.3, w: 1.2, h: 0.08, fill: { color: C.accent }, line: { color: C.accent, width: 0 } });

s = p.addSlide();
chrome(s, 2, 'Findings  >  Overview');
s.addText('Three things stood out in the sessions', { x: 0.5, y: 0.9, w: 12, h: 0.6, fontFace: F.latin, fontSize: 26, bold: true, color: C.ink, margin: 0, valign: 'top' });
s.addText([
  { text: '첫 번째 발견: 사용자는 설정 화면을 찾지 못했다', options: { bullet: true, breakLine: true } },
  { text: 'Second: onboarding copy was skipped by most participants', options: { bullet: true, breakLine: true } },
  { text: 'Third: the export button was mistaken for share', options: { bullet: true } },
], { x: 0.5, y: 1.8, w: 7.5, h: 3, fontFace: F.ko, fontSize: 16, color: C.ink, paraSpaceAfter: 10, valign: 'top', margin: 0 });
s.addShape(p.ShapeType.roundRect, { x: 8.6, y: 1.8, w: 4.2, h: 3, fill: { color: C.soft }, line: { color: C.accent, width: 1 }, rectRadius: 0.15 });
s.addText('“I looked everywhere for the settings.”\n— P3', { x: 8.9, y: 2.1, w: 3.6, h: 2.4, fontFace: F.latin, fontSize: 16, italic: true, color: C.ink, valign: 'middle', margin: 0 });

s = p.addSlide();
chrome(s, 3, 'Method  >  Participants');
s.addText('Participants', { x: 0.5, y: 0.9, w: 12, h: 0.6, fontFace: F.latin, fontSize: 26, bold: true, color: C.ink, margin: 0, valign: 'top' });
const head = (t) => ({ text: t, options: { bold: true, color: 'FFFFFF', fill: { color: C.ink } } });
s.addTable([
  [head('ID'), head('Age'), head('Segment'), head('Device')],
  ['P1', '28', '신규 사용자', 'Android'],
  ['P2', '35', 'Returning', 'iOS'],
  ['P3', '41', '신규 사용자', 'Android'],
  ['P4', '52', 'Returning', 'iOS'],
], { x: 0.5, y: 1.8, w: 12.33, colW: [1.5, 1.5, 5.33, 4], rowH: 0.5, fontFace: F.ko, fontSize: 14, color: C.ink, border: { type: 'solid', pt: 0.75, color: C.line }, valign: 'middle', margin: [0, 0.1, 0, 0.1] });

s = p.addSlide();
chrome(s, 4, 'Findings  >  Screen');
s.addText('The settings entry is below the fold', { x: 0.5, y: 0.9, w: 12, h: 0.6, fontFace: F.latin, fontSize: 26, bold: true, color: C.ink, margin: 0, valign: 'top' });
s.addImage({ path: img, x: 0.5, y: 1.8, w: 6, h: 4 });
s.addShape(p.ShapeType.rect, { x: 3.2, y: 4.6, w: 2.4, h: 0.8, fill: { color: 'FF0000', transparency: 100 }, line: { color: 'EF4444', width: 2 } });
s.addText('Settings sits at the bottom of a long list; 3 of 4 participants scrolled past it.', { x: 7, y: 1.8, w: 5.8, h: 1.5, fontFace: F.latin, fontSize: 16, color: C.ink, valign: 'top', margin: 0 });

s = p.addSlide();
chrome(s, 5, 'Findings  >  Flow');
s.addText('Where the flow breaks', { x: 0.5, y: 0.9, w: 12, h: 0.6, fontFace: F.latin, fontSize: 26, bold: true, color: C.ink, margin: 0, valign: 'top' });
['Open app', 'Find menu', 'Settings', 'Change'].forEach((t, i) => {
  const x = 0.5 + i * 3.2;
  s.addShape(p.ShapeType.ellipse, { x: x + 0.9, y: 2.2, w: 1, h: 1, fill: { color: i === 2 ? 'EF4444' : C.accent }, line: { color: 'FFFFFF', width: 0 } });
  s.addText(String(i + 1), { x: x + 0.9, y: 2.2, w: 1, h: 1, fontFace: F.latin, fontSize: 20, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle', margin: 0 });
  s.addText(t, { x, y: 3.4, w: 2.8, h: 0.5, fontFace: F.latin, fontSize: 14, color: C.ink, align: 'center', valign: 'top', margin: 0 });
  if (i < 3) s.addShape(p.ShapeType.line, { x: x + 2.0, y: 2.7, w: 2.1, h: 0, line: { color: C.line, width: 2, endArrowType: 'triangle' } });
});

s = p.addSlide();
s.background = { color: C.accent };
s.addText('02', { x: 0.8, y: 2.4, w: 3, h: 1, fontFace: F.latin, fontSize: 54, bold: true, color: 'FFFFFF', margin: 0, valign: 'top' });
s.addText('Recommendations', { x: 0.8, y: 3.5, w: 11, h: 0.8, fontFace: F.latin, fontSize: 32, color: 'FFFFFF', margin: 0, valign: 'top' });

await p.writeFile({ fileName: out });
process.stdout.write(JSON.stringify({ ok: true, out, picture: img }) + '\n');
