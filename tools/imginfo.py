"""imginfo.py <file|dir> ... -> one JSON line [{path, w, h, aspect}] for every image (png/jpg/jpeg/gif/bmp/webp; dirs recursive).
Paths are printed as given (dir entries joined to the dir arg), forward slashes."""
import os, sys, json
from PIL import Image

EXT = ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp')
sys.stdout.reconfigure(encoding='utf-8')
out = []
def one(p):
    try:
        with Image.open(p) as im:
            w, h = im.size
        out.append({'path': p.replace('\', '/'), 'w': w, 'h': h, 'aspect': round(w / h, 4)})
    except Exception as e:
        out.append({'path': p.replace('\', '/'), 'error': str(e)[:120]})
for a in sys.argv[1:]:
    if os.path.isdir(a):
        for d, _, fs in os.walk(a):
            for f in sorted(fs):
                if f.lower().endswith(EXT):
                    one(os.path.join(d, f))
    else:
        one(a)
print(json.dumps(out, ensure_ascii=False))
