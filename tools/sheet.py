"""sheet.py <pngDir> <out.png> <cols> <start> <end> [--tile 480]
Contact sheet of <pngDir>/*.png in natural order (s2 < s10 < s100), items start..end (1-based, inclusive), labelled by file name.
Tile width --tile px; tile height follows the first image's aspect."""
import sys, glob, os, re, json, argparse
from PIL import Image, ImageDraw

ap = argparse.ArgumentParser()
ap.add_argument('src'); ap.add_argument('out'); ap.add_argument('cols', type=int)
ap.add_argument('start', type=int); ap.add_argument('end', type=int)
ap.add_argument('--tile', type=int, default=480)
a = ap.parse_args()
nat = lambda f: [int(t) if t.isdigit() else t for t in re.split(r'(\d+)', os.path.basename(f))]
files = sorted(glob.glob(os.path.join(a.src, '*.png')), key=nat)[a.start - 1:a.end]
if not files:
    sys.exit('no png files in range')
w0, h0 = Image.open(files[0]).size
tw = a.tile; th = round(tw * h0 / w0)
rows = (len(files) + a.cols - 1) // a.cols
sheet = Image.new('RGB', (a.cols * tw, rows * (th + 18)), 'white')
d = ImageDraw.Draw(sheet)
for i, f in enumerate(files):
    im = Image.open(f).convert('RGB').resize((tw, th), Image.LANCZOS)
    x, y = (i % a.cols) * tw, (i // a.cols) * (th + 18)
    sheet.paste(im, (x, y + 18)); d.text((x + 4, y + 3), os.path.basename(f), fill='black')
os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
sheet.save(a.out)
print(json.dumps({'out': a.out.replace(os.sep, '/'), 'files': len(files)}))
