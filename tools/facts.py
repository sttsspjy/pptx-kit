"""facts.py <deck.pptx> <outDir> [--slides 5,9]  ->  outDir/sNN.json, outDir/media/<sha1-12>.<ext>, outDir/deck-stats.json (full runs only)
Picture paths in facts are "facts/media/<file>", i.e. relative to the pack root when outDir = <pack>/facts.

Per slide: every shape flattened in z-order with absolute pt geometry, group path, resolved fills/lines/effects,
text frame + paragraphs + runs with inherited values resolved (shape lstStyle -> fontRef -> placeholder chain ->
master txStyles -> presentation defaultTextStyle -> theme fonts), pictures (deduped media), tables.
Numbers are pt rounded to 0.01. Unnormalized parts are kept as raw XML under `xml`.
"""
import sys, os, re, json, math, hashlib, colorsys, collections, copy
from lxml import etree
from pptx import Presentation

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'a14': 'http://schemas.microsoft.com/office/drawing/2010/main',
    'asvg': 'http://schemas.microsoft.com/office/drawing/2016/SVG/main',
}
A = '{%s}' % NS['a']; P = '{%s}' % NS['p']; R = '{%s}' % NS['r']
EMU = 12700.0


def q(tag):
    pre, local = tag.split(':')
    return '{%s}%s' % (NS[pre], local)


def ptv(v, default=None):
    if v is None:
        return default
    return round(int(v) / EMU, 2)


def r2(v):
    return round(v, 2)


def lname(el):
    return etree.QName(el).localname


def raw(el, limit=6000):
    s = etree.tostring(el, encoding='unicode')
    s = re.sub(r'\s+xmlns(:\w+)?="[^"]*"', '', s)
    s = re.sub(r'<a:extLst>.*?</a:extLst>', '', s, flags=re.S)
    return s if len(s) <= limit else s[:limit] + '...[trimmed]'


def clean(d):
    """Drop None / empty containers recursively."""
    if isinstance(d, dict):
        out = {}
        for k, v in d.items():
            v = clean(v)
            if v is None or v == {} or (v == [] and k not in ('runs', 'paragraphs', 'shapes', 'cells')):
                continue
            out[k] = v
        return out
    if isinstance(d, list):
        return [clean(x) for x in d]
    return d


# ---------------------------------------------------------------- colours
PRESET = {'black': '000000', 'white': 'FFFFFF', 'red': 'FF0000', 'green': '008000', 'blue': '0000FF', 'yellow': 'FFFF00',
          'gray': '808080', 'grey': '808080', 'darkGray': 'A9A9A9', 'lightGray': 'D3D3D3', 'orange': 'FFA500',
          'cyan': '00FFFF', 'magenta': 'FF00FF', 'navy': '000080', 'silver': 'C0C0C0', 'purple': '800080'}


def _lin(c):
    return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4


def _srgb(c):
    c = max(0.0, min(1.0, c))
    return c * 12.92 if c <= 0.0031308 else 1.055 * c ** (1 / 2.4) - 0.055


class Theme:
    def __init__(self, theme_el, clr_map):
        self.colors = {}
        cs = theme_el.find('.//' + q('a:clrScheme'))
        for c in cs:
            ch = c[0]
            if lname(ch) == 'srgbClr':
                self.colors[lname(c)] = ch.get('val').upper()
            elif lname(ch) == 'sysClr':
                self.colors[lname(c)] = (ch.get('lastClr') or '000000').upper()
        self.clr_map = clr_map
        fs = theme_el.find('.//' + q('a:fontScheme'))
        self.fonts = {}
        for kind, el in (('mj', fs.find(q('a:majorFont'))), ('mn', fs.find(q('a:minorFont')))):
            lat = el.find(q('a:latin')).get('typeface')
            ea = el.find(q('a:ea')).get('typeface') or ''
            hang = el.find(q('a:font') + '[@script="Hang"]')
            if not ea and hang is not None:
                ea = hang.get('typeface')
            cs_ = el.find(q('a:cs')).get('typeface') or lat
            self.fonts[kind] = {'lt': lat, 'ea': ea or lat, 'cs': cs_}
        fm = theme_el.find('.//' + q('a:fmtScheme'))
        self.fill_styles = list(fm.find(q('a:fillStyleLst')))
        self.bg_fill_styles = list(fm.find(q('a:bgFillStyleLst')))
        self.ln_styles = list(fm.find(q('a:lnStyleLst')))
        self.effect_styles = [e.find(q('a:effectLst')) for e in fm.find(q('a:effectStyleLst'))]

    def font(self, face):
        if face and face.startswith('+'):
            m = re.match(r'\+(mj|mn)-(lt|ea|cs)', face)
            if m:
                return self.fonts[m.group(1)][m.group(2)]
        return face

    def scheme(self, val):
        val = self.clr_map.get(val, val)
        return self.colors.get(val, '000000')


def color_of(el, theme, ph=None):
    """el = element containing a colour choice (solidFill, gs, buClr, ...) or the colour element itself.
    Returns {'hex','alpha'?,'scheme'?} or None."""
    if el is None:
        return None
    c = el
    if lname(c) not in ('srgbClr', 'schemeClr', 'prstClr', 'sysClr', 'scrgbClr', 'hslClr'):
        c = None
        for ch in el:
            if lname(ch) in ('srgbClr', 'schemeClr', 'prstClr', 'sysClr', 'scrgbClr', 'hslClr'):
                c = ch
                break
        if c is None:
            return None
    kind = lname(c)
    scheme = None
    alpha = 1.0
    if kind == 'srgbClr':
        hexv = c.get('val').upper()
    elif kind == 'schemeClr':
        scheme = c.get('val')
        if scheme == 'phClr':
            if ph is None:
                hexv = '000000'
            else:
                hexv = ph['hex']; alpha = ph.get('alpha', 1.0); scheme = ph.get('scheme')
        else:
            hexv = theme.scheme(scheme)
    elif kind == 'prstClr':
        hexv = PRESET.get(c.get('val'), '000000'); scheme = c.get('val')
    elif kind == 'sysClr':
        hexv = (c.get('lastClr') or '000000').upper()
    elif kind == 'scrgbClr':
        hexv = ''.join('%02X' % round(_srgb(int(c.get(k)) / 100000) * 255) for k in ('r', 'g', 'b'))
    else:
        rr, gg, bb = colorsys.hls_to_rgb(int(c.get('hue')) / 21600000, int(c.get('lum')) / 100000, int(c.get('sat')) / 100000)
        hexv = '%02X%02X%02X' % (round(rr * 255), round(gg * 255), round(bb * 255))
    rgb = [int(hexv[i:i + 2], 16) / 255 for i in (0, 2, 4)]
    mods = []
    for m in c:
        n = lname(m); v = int(m.get('val', '0')) / 100000
        mods.append('%s%d' % (n, round(v * 100)))
        if n in ('lumMod', 'lumOff', 'satMod', 'satOff', 'hueMod', 'hueOff', 'lum', 'sat', 'hue'):
            h, l, s = colorsys.rgb_to_hls(*rgb)
            if n == 'lumMod': l *= v
            elif n == 'lumOff': l += v
            elif n == 'lum': l = v
            elif n == 'satMod': s *= v
            elif n == 'satOff': s += v
            elif n == 'sat': s = v
            elif n == 'hueMod': h = (h * v) % 1
            elif n == 'hueOff': h = (h + int(m.get('val')) / 21600000) % 1
            elif n == 'hue': h = int(m.get('val')) / 21600000
            l = max(0, min(1, l)); s = max(0, min(1, s))
            rgb = list(colorsys.hls_to_rgb(h, l, s))
        elif n == 'shade':
            rgb = [_srgb(_lin(x) * v) for x in rgb]
        elif n == 'tint':
            rgb = [_srgb(_lin(x) * v + (1 - v)) for x in rgb]
        elif n == 'alpha':
            alpha = v
        elif n == 'alphaMod':
            alpha *= v
        elif n == 'alphaOff':
            alpha += v
        elif n == 'inv':
            rgb = [1 - x for x in rgb]
        elif n == 'gray':
            g = 0.3 * rgb[0] + 0.59 * rgb[1] + 0.11 * rgb[2]; rgb = [g, g, g]
        elif n == 'comp':
            h, l, s = colorsys.rgb_to_hls(*rgb); rgb = list(colorsys.hls_to_rgb((h + 0.5) % 1, l, s))
    out = {'hex': '%02X%02X%02X' % tuple(round(max(0, min(1, x)) * 255) for x in rgb)}
    if alpha < 0.999:
        out['alpha'] = r2(max(0, min(1, alpha)))
    if scheme:
        out['scheme'] = scheme + ((' ' + ' '.join(m for m in mods if not m.startswith('alpha'))) if any(not m.startswith('alpha') for m in mods) else '')
    return out


def fill_of(el, theme, ph=None, ctx=None):
    """el = a fill element (solidFill/noFill/gradFill/blipFill/pattFill/grpFill)."""
    n = lname(el)
    if n == 'noFill':
        return {'type': 'none'}
    if n == 'solidFill':
        c = color_of(el, theme, ph)
        return dict({'type': 'solid'}, **(c or {}))
    if n == 'gradFill':
        stops = []
        for gs in el.findall(q('a:gsLst') + '/' + q('a:gs')):
            c = color_of(gs, theme, ph) or {}
            stops.append(dict({'pos': r2(int(gs.get('pos')) / 1000)}, **c))
        g = {'type': 'gradient', 'stops': stops}
        lin = el.find(q('a:lin'))
        if lin is not None:
            g['angle'] = r2(int(lin.get('ang', '0')) / 60000)
        path = el.find(q('a:path'))
        if path is not None:
            g['path'] = path.get('path'); g['xml'] = raw(path)
        return g
    if n == 'blipFill':
        out = {'type': 'picture'}
        if ctx is not None:
            out.update(ctx.blip(el))
        return out
    if n == 'pattFill':
        return {'type': 'pattern', 'prst': el.get('prst'), 'fg': color_of(el.find(q('a:fgClr')), theme, ph),
                'bg': color_of(el.find(q('a:bgClr')), theme, ph)}
    if n == 'grpFill':
        return {'type': 'group'}
    return None


FILL_TAGS = ('noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill')


def find_fill(parent):
    if parent is None:
        return None
    for ch in parent:
        if lname(ch) in FILL_TAGS:
            return ch
    return None


def style_ref(style, tag, theme):
    if style is None:
        return None, None
    ref = style.find(q('a:' + tag))
    if ref is None:
        return None, None
    return ref.get('idx'), color_of(ref, theme)


def line_of(ln, style, theme, ctx):
    idx, phc = style_ref(style, 'lnRef', theme)
    base = None
    if idx not in (None, '0'):
        i = int(idx) - 1
        if 0 <= i < len(theme.ln_styles):
            base = theme.ln_styles[i]
    if ln is None and base is None:
        return None
    out = {}
    w = (ln.get('w') if ln is not None else None) or (base.get('w') if base is not None else None)
    f = find_fill(ln) if ln is not None else None
    src_ph = None
    if f is None and base is not None:
        f = find_fill(base); src_ph = phc
    if f is None:
        return None
    fl = fill_of(f, theme, src_ph, ctx)
    if fl is None or fl.get('type') == 'none':
        return {'type': 'none'}
    out.update(fl)
    out['w'] = ptv(w, 0.75) if w is not None else 0.75
    for src in (ln, base):
        if src is None:
            continue
        d = src.find(q('a:prstDash'))
        if d is not None and 'dash' not in out:
            out['dash'] = d.get('val')
        cd = src.find(q('a:custDash'))
        if cd is not None and 'dash' not in out:
            out['dash'] = 'custom'; out['dashXml'] = raw(cd)
        for k in ('cap', 'cmpd'):
            if src.get(k) and k not in out and not (k == 'cmpd' and src.get(k) == 'sng'):
                out[k] = src.get(k)
        for j in ('round', 'bevel', 'miter'):
            if src.find(q('a:' + j)) is not None and 'join' not in out:
                out['join'] = j
    if ln is not None:
        for end in ('headEnd', 'tailEnd'):
            e = ln.find(q('a:' + end))
            if e is not None and e.get('type') not in (None, 'none'):
                out[end] = {'type': e.get('type'), 'w': e.get('w', 'med'), 'len': e.get('len', 'med')}
    if out.get('dash') == 'solid':
        del out['dash']
    return out


def effects_of(sppr, style, theme):
    el = sppr.find(q('a:effectLst')) if sppr is not None else None
    src = 'explicit'
    if el is None:
        idx, phc = style_ref(style, 'effectRef', theme)
        if idx not in (None, '0') and 0 < int(idx) <= len(theme.effect_styles):
            el = theme.effect_styles[int(idx) - 1]; src = 'theme'
        else:
            return None
    else:
        phc = None
    if el is None or len(el) == 0:
        return None
    out = {'src': src}
    for e in el:
        n = lname(e)
        if n in ('outerShdw', 'innerShdw'):
            out[n] = {'blur': ptv(e.get('blurRad', '0')), 'dist': ptv(e.get('dist', '0')),
                      'dir': r2(int(e.get('dir', '0')) / 60000), 'algn': e.get('algn'),
                      'color': color_of(e, theme, phc)}
            if e.get('sx') or e.get('sy'):
                out[n]['sx'] = e.get('sx'); out[n]['sy'] = e.get('sy')
        else:
            out.setdefault('other', []).append(raw(e))
    eff_dag = sppr.find(q('a:effectDag')) if sppr is not None else None
    if eff_dag is not None:
        out.setdefault('other', []).append(raw(eff_dag))
    return out


# ---------------------------------------------------------------- geometry
RADIUS_PRST = {'roundRect': ('adj', 16667), 'round2SameRect': ('adj1', 16667), 'round1Rect': ('adj', 16667),
               'round2DiagRect': ('adj1', 16667), 'wedgeRoundRectCallout': ('adj3', 16667), 'snipRoundRect': ('adj1', 16667)}


def geom_of(sppr, w, h):
    if sppr is None:
        return None
    pg = sppr.find(q('a:prstGeom'))
    if pg is not None:
        g = {'prst': pg.get('prst')}
        adj = {}
        for gd in pg.findall(q('a:avLst') + '/' + q('a:gd')):
            m = re.match(r'val\s+(-?\d+)', gd.get('fmla', ''))
            if m:
                adj[gd.get('name')] = int(m.group(1))
        if adj:
            g['adj'] = adj
        if g['prst'] in RADIUS_PRST:
            k, d = RADIUS_PRST[g['prst']]
            v = adj.get(k, d)
            g['radius'] = r2(min(w, h) * min(v, 50000) / 100000)
            if g['prst'] == 'round2SameRect':
                g['radius2'] = r2(min(w, h) * min(adj.get('adj2', 0), 50000) / 100000)
        return g
    cg = sppr.find(q('a:custGeom'))
    if cg is not None:
        return {'cust': raw(cg, 20000)}
    return None


def cust_path_pt(sppr, x, y, w, h, flip):
    """custGeom paths -> absolute pt commands [[M,x,y],[L,x,y],[C,x1,y1,x2,y2,x,y],[Q,..],[A,raw],[Z]] (flip applied, rotation not)."""
    cg = sppr.find(q('a:custGeom')) if sppr is not None else None
    if cg is None:
        return None
    out = []
    for path in cg.findall(q('a:pathLst') + '/' + q('a:path')):
        pw = int(path.get('w', '0')) or 1; ph = int(path.get('h', '0')) or 1
        def P_(pt):
            try:
                px, py = int(pt.get('x')), int(pt.get('y'))
            except (TypeError, ValueError):
                return [None, None]
            fx = px / pw; fy = py / ph
            if 'h' in (flip or ''):
                fx = 1 - fx
            if 'v' in (flip or ''):
                fy = 1 - fy
            return [r2(x + fx * w), r2(y + fy * h)]
        cmds = []
        for c in path:
            n = lname(c)
            pts = [v for pt in c.findall(q('a:pt')) for v in P_(pt)]
            if n == 'moveTo': cmds.append(['M'] + pts)
            elif n == 'lnTo': cmds.append(['L'] + pts)
            elif n == 'cubicBezTo': cmds.append(['C'] + pts)
            elif n == 'quadBezTo': cmds.append(['Q'] + pts)
            elif n == 'close': cmds.append(['Z'])
            elif n == 'arcTo': cmds.append(['A', raw(c)])
        p = {'cmds': cmds}
        if path.get('fill') == 'none':
            p['fill'] = 'none'
        if path.get('stroke') == '0':
            p['stroke'] = False
        out.append(p)
    return out


class Xf:
    """Group child -> parent transform."""
    def __init__(self, xfrm):
        off = xfrm.find(q('a:off')); ext = xfrm.find(q('a:ext'))
        cho = xfrm.find(q('a:chOff')); che = xfrm.find(q('a:chExt'))
        self.x, self.y = int(off.get('x')), int(off.get('y'))
        self.w, self.h = int(ext.get('cx')), int(ext.get('cy'))
        self.cx, self.cy = int(cho.get('x')), int(cho.get('y'))
        cw, ch = int(che.get('cx')), int(che.get('cy'))
        self.sx = self.w / cw if cw else 1.0
        self.sy = self.h / ch if ch else 1.0
        self.rot = int(xfrm.get('rot', '0')) / 60000
        self.fh = xfrm.get('flipH') == '1'; self.fv = xfrm.get('flipV') == '1'

    def apply(self, b):
        x, y, w, h, rot, fh, fv = b
        x = self.x + (x - self.cx) * self.sx; y = self.y + (y - self.cy) * self.sy
        w *= self.sx; h *= self.sy
        gx, gy = self.x + self.w / 2, self.y + self.h / 2
        mx, my = x + w / 2, y + h / 2
        if self.fh:
            mx = 2 * gx - mx; fh = not fh; rot = -rot
        if self.fv:
            my = 2 * gy - my; fv = not fv; rot = -rot
        if self.rot:
            t = math.radians(self.rot)
            dx, dy = mx - gx, my - gy
            mx = gx + dx * math.cos(t) - dy * math.sin(t); my = gy + dx * math.sin(t) + dy * math.cos(t)
            rot += self.rot
        return (mx - w / 2, my - h / 2, w, h, rot % 360, fh, fv)


def xfrm_box(xfrm):
    off = xfrm.find(q('a:off')); ext = xfrm.find(q('a:ext'))
    if off is None or ext is None:
        return None
    return (int(off.get('x')), int(off.get('y')), int(ext.get('cx')), int(ext.get('cy')),
            int(xfrm.get('rot', '0')) / 60000, xfrm.get('flipH') == '1', xfrm.get('flipV') == '1')


# ---------------------------------------------------------------- text
PPR_CHILD_GROUPS = {
    'lnSpc': ('lnSpc',), 'spcBef': ('spcBef',), 'spcAft': ('spcAft',),
    'buClr': ('buClrTx', 'buClr'), 'buSz': ('buSzTx', 'buSzPct', 'buSzPts'), 'buFont': ('buFontTx', 'buFont'),
    'bu': ('buNone', 'buAutoNum', 'buChar', 'buBlip'),
}
RUN_FILL = FILL_TAGS


def spacing(el):
    if el is None:
        return None
    pc = el.find(q('a:spcPct')); pp = el.find(q('a:spcPts'))
    if pc is not None:
        return {'pct': r2(int(pc.get('val')) / 1000)}
    if pp is not None:
        return {'pt': r2(int(pp.get('val')) / 100)}
    return None


class TextResolver:
    """chain = list of lstStyle-like elements (each has lvlNpPr), highest priority first; font_ref = p:style fontRef."""
    def __init__(self, theme, chain, font_ref=None):
        self.theme = theme; self.chain = [c for c in chain if c is not None]
        self.font_ref = font_ref
        self.own_first = chain[0] is not None if chain else False

    def lvl_list(self, lvl):
        out = []
        for ci, c in enumerate(self.chain):
            e = c.find(q('a:lvl%dpPr' % (lvl + 1)))
            if e is not None:
                out.append(('shape' if ci == 0 and self.own_first else 'inh', e))
        return out

    def para(self, p):
        ppr = p.find(q('a:pPr'))
        lvl = int(ppr.get('lvl', '0')) if ppr is not None else 0
        tagged = ([('ppr', ppr)] if ppr is not None else []) + self.lvl_list(lvl)
        srcs = [e for _, e in tagged]
        def attr(k, d=None):
            for s in srcs:
                if s.get(k) is not None:
                    return s.get(k)
            return d
        def child(group):
            for s in srcs:
                for t in PPR_CHILD_GROUPS[group]:
                    e = s.find(q('a:' + t))
                    if e is not None:
                        return e
            return None
        out = {'lvl': lvl, 'align': attr('algn', 'l'), 'marL': ptv(attr('marL', '0')), 'indent': ptv(attr('indent', '0'))}
        if attr('marR') and attr('marR') != '0':
            out['marR'] = ptv(attr('marR'))
        if attr('fontAlgn') and attr('fontAlgn') not in ('auto', 'base'):
            out['fontAlgn'] = attr('fontAlgn')
        out['latinLnBrk'] = attr('latinLnBrk') == '1' or None
        out['lnSpc'] = spacing(child('lnSpc')) or {'pct': 100.0}
        out['spcBef'] = spacing(child('spcBef')) or {'pt': 0.0}
        out['spcAft'] = spacing(child('spcAft')) or {'pt': 0.0}
        bu = child('bu')
        if bu is not None and lname(bu) != 'buNone':
            b = {}
            if lname(bu) == 'buChar':
                b['char'] = bu.get('char')
            elif lname(bu) == 'buAutoNum':
                b['autonum'] = bu.get('type'); b['startAt'] = int(bu.get('startAt', '1'))
            else:
                b['blip'] = True
            f = child('buFont')
            if f is not None and lname(f) == 'buFont':
                b['font'] = self.theme.font(f.get('typeface'))
            c = child('buClr')
            if c is not None and lname(c) == 'buClr':
                b['color'] = color_of(c, self.theme)
            s = child('buSz')
            if s is not None and lname(s) == 'buSzPct':
                b['sizePct'] = r2(int(s.get('val')) / 1000)
            elif s is not None and lname(s) == 'buSzPts':
                b['sizePt'] = r2(int(s.get('val')) / 100)
            out['bullet'] = b
        tabs = child_tabs(srcs)
        if tabs:
            out['tabs'] = tabs
        return out, [(t, s.find(q('a:defRPr'))) for t, s in tagged]

    def run(self, rpr, defs):
        """rpr = explicit a:rPr (or None); defs = defRPr list (pPr-level first)."""
        own = [d for t, d in defs if t != 'inh']; inh = [d for t, d in defs if t == 'inh']
        fr = [('fontRef', self.font_ref)] if self.font_ref is not None else []
        srcs = [x for x in [rpr] + own + fr + inh if x is not None]
        def attr(k):
            for s in srcs:
                if isinstance(s, tuple):
                    continue
                if s.get(k) is not None:
                    return s.get(k)
            return None
        def face(tag):
            for s in srcs:
                if isinstance(s, tuple):
                    idx = s[1].get('idx')
                    if idx in ('major', 'minor'):
                        return self.theme.font('+%s-%s' % ('mj' if idx == 'major' else 'mn', {'latin': 'lt', 'ea': 'ea', 'cs': 'cs'}[tag]))
                    continue
                e = s.find(q('a:' + tag))
                if e is not None and e.get('typeface'):
                    return self.theme.font(e.get('typeface'))
            return None
        def fill():
            for s in srcs:
                if isinstance(s, tuple):
                    c = color_of(s[1], self.theme)
                    if c:
                        return dict({'type': 'solid'}, **c)
                    continue
                f = find_fill(s)
                if f is not None:
                    return fill_of(f, self.theme)
            return None
        sz = attr('sz')
        out = {'latin': face('latin'), 'ea': face('ea'), 'sz': r2(int(sz) / 100) if sz else 18.0,
               'b': attr('b') in ('1', 'true'), 'i': attr('i') in ('1', 'true') or None}
        u = attr('u')
        if u and u != 'none':
            out['u'] = u
        st = attr('strike')
        if st and st != 'noStrike':
            out['strike'] = st
        f = fill()
        if f:
            if f.get('type') == 'solid':
                out['color'] = f['hex']
                if 'alpha' in f:
                    out['alpha'] = f['alpha']
            else:
                out['fill'] = f
        else:
            out['color'] = self.theme.scheme('tx1')
        spc = attr('spc')
        if spc and spc != '0':
            out['spc'] = r2(int(spc) / 100)
        bl = attr('baseline')
        if bl and bl != '0':
            out['baseline'] = r2(int(bl) / 1000)
        cap = attr('cap')
        if cap and cap != 'none':
            out['cap'] = cap
        if rpr is not None:
            hl = rpr.find(q('a:highlight'))
            if hl is not None:
                c = color_of(hl, self.theme)
                if c and c.get('alpha', 1) > 0:
                    out['highlight'] = c['hex']
            ln = rpr.find(q('a:ln'))
            if ln is not None:
                lf = find_fill(ln)
                if lf is not None:
                    fl = fill_of(lf, self.theme)
                    if fl and fl.get('type') == 'solid' and fl.get('alpha', 1) > 0:
                        out['outline'] = {'hex': fl['hex'], 'w': ptv(ln.get('w', '0'))}
            sym = rpr.find(q('a:sym'))
            if sym is not None:
                out['sym'] = sym.get('typeface')
            if rpr.find(q('a:effectLst')) is not None and len(rpr.find(q('a:effectLst'))):
                out['effects'] = raw(rpr.find(q('a:effectLst')))
        return out

    def body(self, txbody):
        paras = []
        for p in txbody.findall(q('a:p')):
            pd, defs = self.para(p)
            runs = []
            for ch in p:
                n = lname(ch)
                if n in ('r', 'fld'):
                    t = ch.find(q('a:t'))
                    rd = self.run(ch.find(q('a:rPr')), defs)
                    rd = dict({'t': t.text if t is not None and t.text else ''}, **rd)
                    if n == 'fld':
                        rd['field'] = ch.get('type')
                    runs.append(rd)
                elif n == 'br':
                    rd = self.run(ch.find(q('a:rPr')), defs)
                    runs.append({'t': '\n', 'br': True, 'sz': rd['sz']})
            epr = p.find(q('a:endParaRPr'))
            pd['endSz'] = self.run(epr, defs)['sz']
            pd['text'] = ''.join(r['t'] for r in runs)
            pd['runs'] = runs
            paras.append(pd)
        return paras


def child_tabs(srcs):
    for s in srcs:
        tl = s.find(q('a:tabLst'))
        if tl is not None and len(tl):
            return [{'pos': ptv(t.get('pos')), 'algn': t.get('algn', 'l')} for t in tl]
    return None


def body_pr(bp_list):
    """bp_list = bodyPr elements, highest priority first."""
    def attr(k, d=None):
        for b in bp_list:
            if b is not None and b.get(k) is not None:
                return b.get(k)
        return d
    out = {'insets': [ptv(attr('lIns', '91440')), ptv(attr('tIns', '45720')), ptv(attr('rIns', '91440')), ptv(attr('bIns', '45720'))],
           'anchor': attr('anchor', 't'), 'wrap': attr('wrap', 'square')}
    if attr('anchorCtr') == '1':
        out['anchorCtr'] = True
    v = attr('vert', 'horz')
    if v != 'horz':
        out['vert'] = v
    if attr('rot') and attr('rot') != '0':
        out['rot'] = r2(int(attr('rot')) / 60000)
    if attr('numCol') and attr('numCol') != '1':
        out['numCol'] = int(attr('numCol')); out['spcCol'] = ptv(attr('spcCol', '0'))
    if attr('upright') == '1':
        out['upright'] = True
    fit = 'none'
    for b in bp_list:
        if b is None:
            continue
        for t in ('noAutofit', 'normAutofit', 'spAutoFit'):
            e = b.find(q('a:' + t))
            if e is not None:
                fit = t
                if t == 'normAutofit':
                    out['fontScale'] = r2(int(e.get('fontScale', '100000')) / 1000)
                    out['lnSpcReduction'] = r2(int(e.get('lnSpcReduction', '0')) / 1000)
                break
        if fit != 'none':
            break
    out['autofit'] = {'noAutofit': 'none', 'normAutofit': 'shrink', 'spAutoFit': 'resize', 'none': 'none'}[fit]
    return out


# ---------------------------------------------------------------- tables: built-in Medium Style 2 (default GUID)
BUILTIN_TABLE_STYLES = {
    '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}': {'name': 'Medium Style 2 - Accent 1', 'accent': 'accent1'},
}


def builtin_cell_fill(style_id, theme, flags, r, c, nrows, ncols):
    st = BUILTIN_TABLE_STYLES.get(style_id)
    if not st:
        return None, None
    acc = theme.scheme(st['accent'])
    def tint(v):
        rgb = [int(acc[i:i + 2], 16) / 255 for i in (0, 2, 4)]
        return '%02X%02X%02X' % tuple(round(_srgb(_lin(x) * v + (1 - v)) * 255) for x in rgb)
    lt1 = theme.scheme('lt1'); dk1 = theme.scheme('dk1')
    if flags.get('firstRow') and r == 0 or flags.get('lastRow') and r == nrows - 1 \
            or flags.get('firstCol') and c == 0 or flags.get('lastCol') and c == ncols - 1:
        return acc, lt1
    band = flags.get('bandRow') and ((r - (1 if flags.get('firstRow') else 0)) % 2 == 0)
    return (tint(0.4) if band else tint(0.2)), dk1


# ---------------------------------------------------------------- slide walker
class SlideFacts:
    def __init__(self, deck, slide, sidx, media_dir):
        self.deck = deck; self.slide = slide; self.sidx = sidx; self.media_dir = media_dir
        self.theme = deck.theme
        self.items = []
        self.layout = slide.slide_layout; self.master = self.layout.slide_master

    # media ------------------------------------------------------------
    def save_part(self, rid, part_owner):
        try:
            part = part_owner.related_part(rid)
        except KeyError:
            return None
        blob = part.blob
        h = hashlib.sha1(blob).hexdigest()[:12]
        ext = part.partname.ext.lower()
        fn = '%s.%s' % (h, ext)
        path = os.path.join(self.media_dir, fn)
        if not os.path.exists(path):
            with open(path, 'wb') as f:
                f.write(blob)
        self.deck.media_seen[fn] = self.deck.media_seen.get(fn, 0) + 1
        info = {'media': 'facts/media/' + fn}
        if ext in ('png', 'jpg', 'jpeg', 'gif', 'bmp', 'tif', 'tiff'):
            try:
                from PIL import Image
                with Image.open(path) as im:
                    info['px'] = [im.width, im.height]
            except Exception:
                pass
        return info

    def blip(self, blipfill, part_owner=None):
        owner = part_owner or self.slide.part
        out = {}
        b = blipfill.find(q('a:blip'))
        if b is not None:
            rid = b.get(R + 'embed')
            if rid:
                info = self.save_part(rid, owner)
                if info:
                    out.update(info)
            if b.get(R + 'link'):
                out['link'] = True
            svg = b.find('.//' + q('asvg:svgBlip'))
            if svg is not None and svg.get(R + 'embed'):
                si = self.save_part(svg.get(R + 'embed'), owner)
                if si:
                    out['svg'] = si['media']
            for e in b:
                n = lname(e)
                if n == 'alphaModFix':
                    out['alpha'] = r2(int(e.get('amt', '100000')) / 100000)
                elif n != 'extLst':
                    out.setdefault('blipXml', []).append(raw(e))
            ext = b.find(q('a:extLst'))
            if ext is not None:
                ip = ext.find('.//' + q('a14:imgProps'))
                if ip is not None:
                    out.setdefault('blipXml', []).append(raw(ip, 3000))
        sr = blipfill.find(q('a:srcRect'))
        if sr is not None and any(sr.get(k) for k in 'ltrb'):
            out['crop'] = {k: r2(int(sr.get(k, '0')) / 1000) for k in 'ltrb'}
        if blipfill.find(q('a:tile')) is not None:
            out['tile'] = raw(blipfill.find(q('a:tile')))
        st = blipfill.find(q('a:stretch') + '/' + q('a:fillRect'))
        if st is not None and any(st.get(k) for k in 'ltrb'):
            out['fillRect'] = {k: r2(int(st.get(k, '0')) / 1000) for k in 'ltrb'}
        return out

    # placeholders -----------------------------------------------------
    def ph_chain(self, ph):
        """Returns [layout_sp, master_sp] matching the placeholder."""
        idx = ph.get('idx'); typ = ph.get('type', 'body')
        def find(tree, by_idx):
            for sp in tree.iter(P + 'sp'):
                pp = sp.find('.//' + P + 'ph')
                if pp is None:
                    continue
                if by_idx and idx is not None and pp.get('idx') == idx:
                    return sp
                if not by_idx and pp.get('type', 'body') == typ:
                    return sp
            return None
        lay = find(self.layout._element, True)
        if lay is None:
            lay = find(self.layout._element, False)
        mas = find(self.master._element, False)
        return [x for x in (lay, mas) if x is not None], typ

    # walk ---------------------------------------------------------------
    def walk(self, tree, xfs, path, grp_fill):
        for el in tree:
            n = lname(el)
            if n in ('sp', 'cxnSp', 'pic', 'graphicFrame', 'grpSp'):
                self.shape(el, n, xfs, path, grp_fill)
            elif n == 'AlternateContent':
                fb = el.find('{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback')
                ch = el.find('{http://schemas.openxmlformats.org/markup-compatibility/2006}Choice')
                self.walk(ch if ch is not None else fb, xfs, path, grp_fill)

    def box(self, xfrm, xfs):
        b = xfrm_box(xfrm) if xfrm is not None else None
        if b is None:
            return None
        for xf in reversed(xfs):
            b = xf.apply(b)
        return b

    def shape(self, el, n, xfs, path, grp_fill):
        nv = el.find('.//' + P + 'cNvPr')
        item = {'i': len(self.items), 'id': int(nv.get('id')), 'name': nv.get('name')}
        if path:
            item['group'] = list(path)
        if nv.get('hidden') == '1':
            item['hidden'] = True
        if nv.get('descr'):
            item['descr'] = nv.get('descr')
        theme = self.theme
        style = el.find(P + 'style')
        if n == 'grpSp':
            gpr = el.find(P + 'grpSpPr')
            xfrm = gpr.find(q('a:xfrm'))
            b = self.box(xfrm, xfs)
            item['kind'] = 'group'
            self.put_box(item, b)
            gf = find_fill(gpr)
            item['fill'] = fill_of(gf, theme, None, self) if gf is not None else None
            self.items.append(item)
            self.walk(el, xfs + [Xf(xfrm)], path + [nv.get('name')], item.get('fill') or grp_fill)
            return
        if n == 'graphicFrame':
            xfrm = el.find(P + 'xfrm')
            b = self.box(xfrm, xfs)
            self.put_box(item, b)
            tbl = el.find('.//' + q('a:tbl'))
            if tbl is not None:
                item['kind'] = 'table'
                item['table'] = self.table(tbl)
            else:
                item['kind'] = 'graphic'
                gd = el.find('.//' + q('a:graphicData'))
                item['xml'] = {'graphicData': raw(gd, 3000)}
            self.items.append(item)
            return
        sppr = el.find(P + 'spPr')
        ph = el.find('.//' + P + 'nvPr/' + P + 'ph')
        inherit = []
        ph_type = None
        if ph is not None:
            inherit, ph_type = self.ph_chain(ph)
            item['placeholder'] = {'type': ph_type, 'idx': ph.get('idx')}
        xfrm = sppr.find(q('a:xfrm')) if sppr is not None else None
        if xfrm is None or xfrm.find(q('a:off')) is None:
            for isp in inherit:
                x2 = isp.find(P + 'spPr/' + q('a:xfrm'))
                if x2 is not None and x2.find(q('a:off')) is not None:
                    xfrm = x2; break
        b = self.box(xfrm, xfs)
        self.put_box(item, b)
        w = b[2] / EMU if b else 0; h = b[3] / EMU if b else 0
        geo_src = sppr if sppr is not None and (sppr.find(q('a:prstGeom')) is not None or sppr.find(q('a:custGeom')) is not None) else \
            next((isp.find(P + 'spPr') for isp in inherit if isp.find(P + 'spPr/' + q('a:prstGeom')) is not None), sppr)
        item['geom'] = geom_of(geo_src, w, h)
        if item['geom'] and 'cust' in item['geom'] and b:
            item['geom']['pathPt'] = cust_path_pt(geo_src, item['x'], item['y'], item['w'], item['h'], item.get('flip'))
        # fill
        if n == 'pic':
            bf = el.find(P + 'blipFill')
            item['kind'] = 'picture'
            item['pic'] = self.blip(bf)
            f = find_fill(sppr)
            if f is not None:
                item['fill'] = fill_of(f, theme, None, self)
        else:
            f = find_fill(sppr)
            if f is None:
                for isp in inherit:
                    f = find_fill(isp.find(P + 'spPr'))
                    if f is not None:
                        break
            if f is not None:
                fl = fill_of(f, theme, None, self)
                if fl and fl.get('type') == 'group':
                    fl = dict(grp_fill or {'type': 'none'}, fromGroup=True)
                item['fill'] = fl
            else:
                idx, phc = style_ref(style, 'fillRef', theme)
                if idx in (None, '0'):
                    item['fill'] = {'type': 'none'}
                else:
                    i = int(idx)
                    lst = theme.fill_styles if i < 1000 else theme.bg_fill_styles
                    j = (i - 1) if i < 1000 else (i - 1001)
                    item['fill'] = dict(fill_of(lst[j], theme, phc, self), src='theme') if 0 <= j < len(lst) else None
        ln = sppr.find(q('a:ln')) if sppr is not None else None
        item['line'] = line_of(ln, style, theme, self)
        if item['line'] and item['line'].get('type') == 'none':
            item['line'] = None
        item['effects'] = effects_of(sppr, style, theme)
        extra = {}
        if sppr is not None:
            for ch in sppr:
                if lname(ch) in ('scene3d', 'sp3d'):
                    extra[lname(ch)] = raw(ch)
        prst = (item.get('geom') or {}).get('prst')
        if n == 'cxnSp' or prst in ('line', 'straightConnector1', 'bentConnector2', 'bentConnector3', 'curvedConnector3'):
            item['kind'] = 'line'
            if b:
                item['ends'] = self.line_ends(b)
            for k in ('stCxn', 'endCxn'):
                c = el.find('.//' + q('a:' + k))
                if c is not None:
                    item.setdefault('connect', {})[k] = {'id': int(c.get('id')), 'site': int(c.get('idx'))}
        # text
        tx = el.find(P + 'txBody')
        if tx is not None:
            font_ref = style.find(q('a:fontRef')) if style is not None else None
            chain = [tx.find(q('a:lstStyle'))]
            bps = [tx.find(q('a:bodyPr'))]
            for isp in inherit:
                itx = isp.find(P + 'txBody')
                if itx is not None:
                    chain.append(itx.find(q('a:lstStyle'))); bps.append(itx.find(q('a:bodyPr')))
            if ph is not None:
                txs = self.master._element.find(P + 'txStyles')
                key = 'titleStyle' if ph_type in ('title', 'ctrTitle') else ('bodyStyle' if ph_type in ('body', 'obj', 'subTitle') else 'otherStyle')
                chain.append(txs.find(P + key))
            chain.append(self.deck.default_text)
            tr = TextResolver(theme, chain, font_ref)
            paras = tr.body(tx)
            has_text = any(p['text'].strip() for p in paras)
            if has_text or n == 'sp' and item.get('kind') is None:
                if has_text:
                    bpr = body_pr(bps)
                    bpr['paragraphs'] = paras
                    bpr['plain'] = '\n'.join(p['text'] for p in paras)
                    item['text'] = bpr
        if 'kind' not in item:
            item['kind'] = 'text' if item.get('text') else 'shape'
        if extra:
            item['xml'] = extra
        self.items.append(item)

    def put_box(self, item, b):
        if not b:
            return
        x, y, w, h, rot, fh, fv = b
        item['x'] = r2(x / EMU); item['y'] = r2(y / EMU); item['w'] = r2(w / EMU); item['h'] = r2(h / EMU)
        if abs(rot) > 0.001 and abs(rot - 360) > 0.001:
            item['rot'] = r2(rot)
        if fh or fv:
            item['flip'] = ('h' if fh else '') + ('v' if fv else '')

    def line_ends(self, b):
        x, y, w, h, rot, fh, fv = b
        x1, y1, x2, y2 = x, y, x + w, y + h
        if fh:
            x1, x2 = x2, x1
        if fv:
            y1, y2 = y2, y1
        if rot:
            cx, cy = x + w / 2, y + h / 2
            t = math.radians(rot)
            def rt(px, py):
                dx, dy = px - cx, py - cy
                return cx + dx * math.cos(t) - dy * math.sin(t), cy + dx * math.sin(t) + dy * math.cos(t)
            x1, y1 = rt(x1, y1); x2, y2 = rt(x2, y2)
        return {'x1': r2(x1 / EMU), 'y1': r2(y1 / EMU), 'x2': r2(x2 / EMU), 'y2': r2(y2 / EMU)}

    def table(self, tbl):
        theme = self.theme
        tpr = tbl.find(q('a:tblPr'))
        flags = {k: tpr.get(k) == '1' for k in ('firstRow', 'lastRow', 'firstCol', 'lastCol', 'bandRow', 'bandCol')} if tpr is not None else {}
        sid = tpr.find(q('a:tableStyleId')).text if tpr is not None and tpr.find(q('a:tableStyleId')) is not None else None
        cols = [ptv(g.get('w')) for g in tbl.findall(q('a:tblGrid') + '/' + q('a:gridCol'))]
        rows_el = tbl.findall(q('a:tr'))
        out = {'style': sid, 'styleName': (BUILTIN_TABLE_STYLES.get(sid) or {}).get('name'),
               'flags': {k: v for k, v in flags.items() if v}, 'cols': cols, 'rows': [ptv(t.get('h')) for t in rows_el], 'cells': []}
        tpr_fill = find_fill(tpr) if tpr is not None else None
        if tpr_fill is not None:
            out['fill'] = fill_of(tpr_fill, theme, None, self)
        for ri, tr_ in enumerate(rows_el):
            row = []
            for ci, tc in enumerate(tr_.findall(q('a:tc'))):
                cell = {}
                for k in ('gridSpan', 'rowSpan'):
                    if tc.get(k) and tc.get(k) != '1':
                        cell[k] = int(tc.get(k))
                for k in ('hMerge', 'vMerge'):
                    if tc.get(k) == '1':
                        cell[k] = True
                tcpr = tc.find(q('a:tcPr'))
                g = lambda k, d: tcpr.get(k, d) if tcpr is not None else d
                cell['mar'] = [ptv(g('marL', '91440')), ptv(g('marT', '45720')), ptv(g('marR', '91440')), ptv(g('marB', '45720'))]
                cell['anchor'] = g('anchor', 't')
                if g('vert', 'horz') != 'horz':
                    cell['vert'] = g('vert', 'horz')
                sfill, stext = builtin_cell_fill(sid, theme, flags, ri, ci, len(rows_el), len(cols))
                f = find_fill(tcpr)
                if f is not None:
                    cell['fill'] = fill_of(f, theme, None, self)
                elif sfill:
                    cell['fill'] = {'type': 'solid', 'hex': sfill, 'src': 'tableStyle'}
                bd = {}
                if tcpr is not None:
                    for side in ('lnL', 'lnR', 'lnT', 'lnB', 'lnTlToBr', 'lnBlToTr'):
                        e = tcpr.find(q('a:' + side))
                        if e is not None:
                            l = line_of(e, None, theme, self)
                            bd[side[2:]] = l if l and l.get('type') != 'none' else None
                cell['border'] = {k: v for k, v in bd.items() if v} or None
                txb = tc.find(q('a:txBody'))
                if txb is not None:
                    chain = [txb.find(q('a:lstStyle')), self.deck.default_text]
                    tr = TextResolver(theme, chain)
                    paras = tr.body(txb)
                    if stext:
                        for p in paras:
                            for r in p['runs']:
                                if r.get('color') == theme.scheme('tx1') and not self._explicit_color(tc):
                                    r['color'] = stext
                    cell['paragraphs'] = paras
                    cell['text'] = '\n'.join(p['text'] for p in paras)
                row.append(cell)
            out['cells'].append(row)
        return out

    @staticmethod
    def _explicit_color(tc):
        return tc.find('.//' + q('a:rPr') + '/' + q('a:solidFill')) is not None

    def background(self):
        for owner in (self.slide, self.layout, self.master):
            bg = owner._element.find(P + 'cSld/' + P + 'bg')
            if bg is None:
                continue
            bpr = bg.find(P + 'bgPr')
            if bpr is not None:
                f = find_fill(bpr)
                return dict(fill_of(f, self.theme, None, self) or {}, src=type(owner).__name__)
            ref = bg.find(P + 'bgRef')
            if ref is not None:
                i = int(ref.get('idx'))
                phc = color_of(ref, self.theme)
                if i >= 1001:
                    return dict(fill_of(self.theme.bg_fill_styles[i - 1001], self.theme, phc, self) or {}, src=type(owner).__name__)
                return dict(phc or {}, type='solid', src=type(owner).__name__)
        return {'type': 'solid', 'hex': 'FFFFFF'}

    def run(self):
        self.walk(self.slide._element.find(P + 'cSld/' + P + 'spTree'), [], [], None)
        return clean({'slide': self.sidx, 'size': [self.deck.w, self.deck.h], 'layout': self.layout.name,
                      'background': self.background(), 'shapes': self.items})


class Deck:
    def __init__(self, path):
        self.prs = Presentation(path)
        self.w = round(self.prs.slide_width / EMU, 2); self.h = round(self.prs.slide_height / EMU, 2)
        master = self.prs.slide_masters[0]
        cm = master._element.find(P + 'clrMap')
        clr_map = dict(cm.attrib) if cm is not None else {}
        theme_part = None
        for rel in master.part.rels.values():
            if rel.reltype.endswith('/theme'):
                theme_part = rel.target_part
        self.theme = Theme(etree.fromstring(theme_part.blob), clr_map)
        self.default_text = self.prs.part._element.find(P + 'defaultTextStyle')
        self.media_seen = {}


def stats(all_facts):
    fills = collections.Counter(); texts = collections.Counter(); lines = collections.Counter()
    fonts = collections.Counter(); fonts_chars = collections.Counter()
    xs = collections.Counter(); ys = collections.Counter(); ws = collections.Counter(); hs = collections.Counter()
    pics = {}; kinds = collections.Counter(); faces = collections.Counter()
    def runs_of(paras):
        for p in paras:
            for r in p.get('runs', []):
                if r.get('br'):
                    continue
                yield r
    for f in all_facts:
        key = 's%02d' % f['slide']
        pics[key] = 0
        for s in f['shapes']:
            kinds[s['kind']] += 1
            if s['kind'] == 'picture':
                pics[key] += 1
            if s['kind'] != 'group' and 'x' in s:
                xs[s['x']] += 1; ys[s['y']] += 1; ws[s['w']] += 1; hs[s['h']] += 1
            fl = s.get('fill') or {}
            if fl.get('type') == 'solid':
                fills[fl['hex'] + ('@%s' % fl['alpha'] if 'alpha' in fl else '')] += 1
            ln = s.get('line') or {}
            if ln.get('type') == 'solid':
                lines[ln['hex'] + ('@%s' % ln['alpha'] if 'alpha' in ln else '')] += 1
            paras = (s.get('text') or {}).get('paragraphs', [])
            cells = [c for row in (s.get('table') or {}).get('cells', []) for c in row]
            for c in cells:
                paras = paras + c.get('paragraphs', [])
                cf = c.get('fill') or {}
                if cf.get('type') == 'solid':
                    fills[cf['hex']] += 1
            for r in runs_of(paras):
                if not r['t'].strip():
                    continue
                texts[r.get('color', '?')] += len(r['t'])
                k = (r.get('ea') or r.get('latin'), r['sz'], bool(r.get('b')))
                for fk in ('latin', 'ea'):
                    if r.get(fk):
                        faces[r[fk]] += len(r['t'])
                fonts[k] += 1; fonts_chars[k] += len(r['t'])
    top = lambda c, n: [{'v': k, 'n': v} for k, v in c.most_common(n)]
    return {
        'colors': {'fill': top(fills, 40), 'text': [{'v': k, 'chars': v} for k, v in texts.most_common(40)], 'line': top(lines, 30)},
        'fonts': [{'face': k[0], 'sz': k[1], 'bold': k[2], 'runs': v, 'chars': fonts_chars[k]} for k, v in fonts.most_common(60)],
        'common': {'x': top(xs, 30), 'y': top(ys, 30), 'w': top(ws, 30), 'h': top(hs, 30)},
        'kinds': dict(kinds), 'picturesPerSlide': pics,
        'faces': [{'face': k, 'chars': v} for k, v in faces.most_common()],
    }


def _j(o):
    return json.dumps(o, ensure_ascii=False, separators=(',', ':'))


def dump_facts(f):
    """Valid JSON, one shape per line; table rows one per line."""
    head = {k: v for k, v in f.items() if k != 'shapes'}
    lines = ['{' + _j(head)[1:-1] + ',"shapes":[']
    for n, s in enumerate(f['shapes']):
        s = dict(sorted(s.items(), key=lambda kv: ['i', 'id', 'name', 'kind'].index(kv[0]) if kv[0] in ('i', 'id', 'name', 'kind') else 9))
        end = ',' if n < len(f['shapes']) - 1 else ''
        if s.get('kind') == 'table':
            t = s['table']
            s2 = {k: v for k, v in s.items() if k != 'table'}
            t2 = {k: v for k, v in t.items() if k != 'cells'}
            rows = ',\n  '.join(_j(r) for r in t['cells'])
            lines.append(_j(s2)[:-1] + ',"table":' + _j(t2)[:-1] + ',"cells":[\n  ' + rows + '\n]}}' + end)
        else:
            lines.append(_j(s) + end)
    lines.append(']}')
    return '\n'.join(lines) + '\n'


def main():
    deck_path, out = sys.argv[1], sys.argv[2]
    only = None
    if '--slides' in sys.argv:
        only = {int(x) for x in sys.argv[sys.argv.index('--slides') + 1].split(',')}
    media = os.path.join(out, 'media')
    os.makedirs(media, exist_ok=True)
    deck = Deck(deck_path)
    all_facts = []
    for i, s in enumerate(deck.prs.slides, 1):
        if only and i not in only:
            continue
        f = SlideFacts(deck, s, i, media).run()
        with open(os.path.join(out, 's%02d.json' % i), 'w', encoding='utf-8') as fh:
            fh.write(dump_facts(f))
        all_facts.append(f)
    if not only:
        st = stats(all_facts)
        st['media'] = {'files': len(deck.media_seen), 'refs': sum(deck.media_seen.values())}
        st['canvas'] = {'w': deck.w, 'h': deck.h}
        st['slides'] = len(all_facts)
        with open(os.path.join(out, 'deck-stats.json'), 'w', encoding='utf-8') as fh:
            json.dump(st, fh, ensure_ascii=False, indent=1)
    print(json.dumps({'slides': len(all_facts), 'media': len(deck.media_seen), 'shapes': sum(len(f['shapes']) for f in all_facts)}))


if __name__ == '__main__':
    main()
