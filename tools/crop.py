"""crop.py <png> <x> <y> <w> <h> [--scale S] --out <file>   (px; region clamped to the image)"""
import sys, argparse
from PIL import Image

ap = argparse.ArgumentParser()
ap.add_argument('png'); ap.add_argument('x', type=float); ap.add_argument('y', type=float)
ap.add_argument('w', type=float); ap.add_argument('h', type=float)
ap.add_argument('--scale', type=float, default=1.0); ap.add_argument('--out', required=True)
a = ap.parse_args()
im = Image.open(a.png).convert('RGB')
x0 = max(0, int(round(a.x))); y0 = max(0, int(round(a.y)))
x1 = min(im.width, int(round(a.x + a.w))); y1 = min(im.height, int(round(a.y + a.h)))
if x1 <= x0 or y1 <= y0:
    sys.exit('empty crop region')
c = im.crop((x0, y0, x1, y1))
if a.scale != 1.0:
    c = c.resize((max(1, round(c.width * a.scale)), max(1, round(c.height * a.scale))),
                 Image.LANCZOS if a.scale < 1 else Image.NEAREST if a.scale >= 3 else Image.BICUBIC)
c.save(a.out)
print('{"out": "%s", "w": %d, "h": %d}' % (a.out.replace('\\', '/'), c.width, c.height))
