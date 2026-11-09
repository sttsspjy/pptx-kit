# PptxGenJS 4.0.1 quick reference

Do not read the library source. Working code = `<pack>/lib/*.mjs` and existing layouts.

- Units: x/y/w/h in inches (pt / 72) or `'NN%'`. Colours hex without `#`.
- `slide.addText(textOrRuns, opts)`; runs = `[{ text, options: { fontFace, fontSize, bold, italic, underline, color, highlight, charSpacing, breakLine, softBreakBefore, bullet, paraSpaceBefore, paraSpaceAfter, lineSpacing, lineSpacingMultiple, align, indentLevel } }]`. Box opts: `x y w h, margin, valign ('top'|'middle'|'bottom'), align, wrap, fill: {color, transparency}, line: {color, width, dashType}, shape, rectRadius, rotate, vert, shadow, objectName`. `breakLine: true` ends a paragraph; paragraph props go on the first run of the paragraph.
- `lineSpacing` = exact pt; `lineSpacingMultiple` = 0–9.99; `paraSpaceBefore/After` in pt; `charSpacing` in pt.
- `bullet: { characterCode: '<hex>', indent: <pt> }`; `bullet.indent` defaults to 27 pt. If a source bullet uses a symbol font (e.g. Wingdings ü = ✓), verify the glyph in a render or draw the mark as its own text box.
- `slide.addShape(ctx.pptx.ShapeType.rect | roundRect | ellipse | line | triangle | rightArrow | …, { x y w h, fill, line: {color, width, dashType, beginArrowType, endArrowType}, rectRadius, rotate, flipH, flipV, shadow: {type:'outer', color, opacity, blur, offset, angle} })`.
- `slide.addImage({ path, x y w h, sizing: { type: 'crop'|'cover'|'contain', x, y, w, h }, rounding, transparency })` — for source crops (`srcRect` %), compute the crop box from the picture's pixel size (`ctx.imageSize(path)`).
- `slide.addTable(rows, { x y w, colW: [..], rowH: [..], margin, border: {type, pt, color}, fill, fontFace, fontSize, valign })`; cell = `{ text | runs, options: { fill, color, bold, align, valign, colspan, rowspan, border: [t, r, b, l], margin } }`.
- `slide.addChart(ctx.pptx.ChartType.bar | line | doughnut | …, data, opts)` — native, editable charts; style every colour, face and size explicitly from tokens (`chartColors`, `catAxisLabelFontFace`, `dataLabelFontFace`, `valGridLine: { style: 'none' }`, …).

## Traps (measured)

- Text `margin` is pt ordered `[left, right, bottom, top]`.
- `valign` defaults to middle — always set it.
- `rectRadius` behaves as inches in 4.0.1 (types say 0–1) — verify with one render.
- Never `fit: 'shrink'`; no autofit.
- Table cell `margin` is ordered `[top, right, bottom, left]` and read as inches when the first value is < 1, else pt.
- PptxGenJS writes one face for latin and east-asian runs and cannot emit `fontAlgn`: when the source needs different faces or `fontAlgn`, patch the XML with `ctx.xmlPatch` (see `pack-format.md`).
- Set `objectName` on content boxes so overflow reports are readable (unnamed objects get `text3:<first words>`).
