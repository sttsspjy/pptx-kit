"""fonts.py (--faces "<face>" ... | --pack <packDir> | --stats <deck-stats.json>)
Checks which font faces are installed (system + per-user fonts, files and registry), matching any name-table name
(family, full name, typographic family, family + subfamily) in any language, ignoring case, spaces and hyphens.
--pack reads faces from pack.json "fonts", else tokens.json values under *face*/*font* keys plus `fontFace: '...'` in lib/ and layouts/.
Prints one JSON line { faces: [{face, installed, file}], missing: [face] }."""
import os, re, sys, json, glob, struct, argparse

SKIP = {'', '+mn-lt', '+mj-lt', '+mn-ea', '+mj-ea', '+mn-cs', '+mj-cs'}


def norm(s):
    return re.sub(r'[\s\-_]+', '', s).lower()


def font_files():
    dirs = [os.path.join(os.environ.get('WINDIR', r'C:\Windows'), 'Fonts'),
            os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Microsoft', 'Windows', 'Fonts')]
    files = set()
    for d in dirs:
        if os.path.isdir(d):
            for ext in ('ttf', 'otf', 'ttc', 'otc'):
                files.update(glob.glob(os.path.join(d, '*.' + ext)) + glob.glob(os.path.join(d, '*.' + ext.upper())))
    try:
        import winreg
        for hive, key in ((winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts'),
                          (winreg.HKEY_CURRENT_USER, r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts')):
            try:
                k = winreg.OpenKey(hive, key)
            except OSError:
                continue
            i = 0
            while True:
                try:
                    _, v, _ = winreg.EnumValue(k, i)
                except OSError:
                    break
                i += 1
                if isinstance(v, str) and os.path.isabs(v) and os.path.exists(v):
                    files.add(v)
    except ImportError:
        pass
    return sorted(files)


def names_in(path):
    """All name-table strings (IDs 1, 2, 4, 16, 17) of every face in a TTF/OTF/TTC; returns [set(names)] per face."""
    out = []
    with open(path, 'rb') as f:
        data = f.read()
    offsets = [0]
    if data[:4] == b'ttcf':
        n = struct.unpack('>I', data[8:12])[0]
        offsets = list(struct.unpack('>%dI' % n, data[12:12 + 4 * n]))
    for off in offsets:
        num = struct.unpack('>H', data[off + 4:off + 6])[0]
        name_off = None
        for t in range(num):
            rec = data[off + 12 + 16 * t: off + 28 + 16 * t]
            if rec[:4] == b'name':
                name_off = struct.unpack('>I', rec[8:12])[0]
        if name_off is None:
            continue
        _, count, str_off = struct.unpack('>HHH', data[name_off:name_off + 6])
        by_id = {}
        for r in range(count):
            pid, eid, lid, nid, ln, so = struct.unpack('>HHHHHH', data[name_off + 6 + 12 * r: name_off + 18 + 12 * r])
            if nid not in (1, 2, 4, 16, 17):
                continue
            raw = data[name_off + str_off + so: name_off + str_off + so + ln]
            try:
                s = raw.decode('utf-16-be') if pid in (0, 3) else raw.decode('mac_roman')
            except UnicodeDecodeError:
                continue
            by_id.setdefault((pid, lid), {})[nid] = s.strip()
        names = set()
        for d in by_id.values():
            for nid in (1, 4, 16):
                if d.get(nid):
                    names.add(d[nid])
            if d.get(1) and d.get(2):
                names.add(d[1] + ' ' + d[2])
            if d.get(16) and d.get(17):
                names.add(d[16] + ' ' + d[17])
        out.append(names)
    return out


def index():
    idx = {}
    for f in font_files():
        try:
            for names in names_in(f):
                for n in names:
                    idx.setdefault(norm(n), f)
        except Exception:
            continue
    return idx


def faces_of_pack(pack):
    pj = os.path.join(pack, 'pack.json')
    if os.path.exists(pj):
        with open(pj, encoding='utf-8') as fh:
            fonts = json.load(fh).get('fonts') or []
        if fonts:
            return [f['face'] if isinstance(f, dict) else f for f in fonts]
    faces = set()
    tj = os.path.join(pack, 'tokens.json')
    if os.path.exists(tj):
        with open(tj, encoding='utf-8') as fh:
            tok = json.load(fh)
        def walk(o, key=''):
            if isinstance(o, dict):
                for k, v in o.items():
                    walk(v, k)
            elif isinstance(o, list):
                for v in o:
                    walk(v, key)
            elif isinstance(o, str) and re.search(r'face|font', key, re.I):
                faces.add(o)
        walk(tok)
    for f in glob.glob(os.path.join(pack, 'lib', '*.mjs')) + glob.glob(os.path.join(pack, 'layouts', '*', '*.mjs')):
        with open(f, encoding='utf-8') as fh:
            faces.update(re.findall(r"fontFace\s*:\s*['\"]([^'\"]+)['\"]", fh.read()))
    return sorted(faces)


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    ap = argparse.ArgumentParser()
    ap.add_argument('--faces', nargs='*'); ap.add_argument('--pack'); ap.add_argument('--stats')
    a = ap.parse_args()
    faces = list(a.faces or [])
    if a.pack:
        faces += faces_of_pack(a.pack)
    if a.stats:
        with open(a.stats, encoding='utf-8') as fh:
            faces += [f['face'] for f in json.load(fh).get('faces', [])]
    faces = [f for f in dict.fromkeys(faces) if f not in SKIP]
    idx = index()
    res = []
    for face in faces:
        hit = idx.get(norm(face))
        res.append({'face': face, 'installed': bool(hit), 'file': hit})
    print(json.dumps({'faces': res, 'missing': [r['face'] for r in res if not r['installed']]}, ensure_ascii=False))


if __name__ == '__main__':
    main()
