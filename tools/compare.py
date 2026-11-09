"""compare.py <source.png> <render.png> --out <compare.png>
Both images scaled to the slide size in pt (source px / 2; 1 px = 1 pt). Prints one JSON line:
{ ssim (grayscale), meanAbsDiff (RGB, 0-255), diffBoxes: [{x,y,w,h,score}] top 5 regions in pt }.
compare.png = 2x2 grid (source px size) [source | render / diff heatmap + boxes | 50% blend], labelled.
score = summed per-pixel max-channel difference in the region / 255 (pixel-equivalents of full difference)."""
import sys, json, argparse, os
import numpy as np
from PIL import Image, ImageDraw, ImageFont
from skimage.metrics import structural_similarity
from skimage.measure import label, regionprops
from scipy import ndimage

W, H = 960, 540      # replaced in main() by the source size / 2
THRESH = 40          # max-channel abs diff counted as a real difference
MIN_AREA = 6         # px; smaller blobs are anti-aliasing noise


def load(p):
    im = Image.open(p).convert('RGB')
    if im.size != (W, H):
        im = im.resize((W, H), Image.LANCZOS)
    return im


def font(sz):
    for f in ('C:/Windows/Fonts/arialbd.ttf', 'C:/Windows/Fonts/arial.ttf'):
        if os.path.exists(f):
            return ImageFont.truetype(f, sz)
    return ImageFont.load_default()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('src'); ap.add_argument('render'); ap.add_argument('--out', required=True)
    a = ap.parse_args()
    global W, H
    sw, sh = Image.open(a.src).size
    W, H = max(1, round(sw / 2)), max(1, round(sh / 2))
    s_im, r_im = load(a.src), load(a.render)
    s = np.asarray(s_im).astype(np.float32); r = np.asarray(r_im).astype(np.float32)
    sg = np.asarray(s_im.convert('L')).astype(np.float64); rg = np.asarray(r_im.convert('L')).astype(np.float64)
    ssim = float(structural_similarity(sg, rg, data_range=255))
    mad = float(np.abs(s - r).mean())
    d = np.abs(s - r).max(axis=2)
    mask = d > THRESH
    mask = ndimage.binary_opening(mask, iterations=1) | (ndimage.uniform_filter(d, 3) > THRESH)
    merged = ndimage.binary_dilation(mask, iterations=4)
    lab = label(merged)
    boxes = []
    for rp in regionprops(lab):
        y0, x0, y1, x1 = rp.bbox
        region = (lab[y0:y1, x0:x1] == rp.label) & mask[y0:y1, x0:x1]
        area = int(region.sum())
        if area < MIN_AREA:
            continue
        score = float(d[y0:y1, x0:x1][region].sum() / 255.0)
        boxes.append({'x': x0, 'y': y0, 'w': x1 - x0, 'h': y1 - y0, 'score': round(score, 1)})
    boxes.sort(key=lambda b: -b['score'])
    top = boxes[:5]

    # grid
    grid = Image.new('RGB', (W * 2, H * 2), 'white')
    grid.paste(s_im, (0, 0)); grid.paste(r_im, (W, 0))
    base = np.asarray(s_im.convert('L')).astype(np.float32)
    heat = np.stack([base * 0.25 + 191] * 3, axis=2)          # faded source
    k = np.clip(d / 128.0, 0, 1)[..., None]
    heat = heat * (1 - k) + np.array([230, 0, 0], np.float32) * k
    heat_im = Image.fromarray(heat.clip(0, 255).astype(np.uint8))
    hd = ImageDraw.Draw(heat_im)
    f_box = font(14)
    for n, b in enumerate(top, 1):
        hd.rectangle([b['x'] - 1, b['y'] - 1, b['x'] + b['w'], b['y'] + b['h']], outline=(0, 90, 255), width=2)
        hd.text((b['x'] + 2, max(0, b['y'] - 16)), str(n), fill=(0, 90, 255), font=f_box)
    grid.paste(heat_im, (0, H))
    grid.paste(Image.blend(s_im, r_im, 0.5), (W, H))
    g = ImageDraw.Draw(grid)
    g.line([(W, 0), (W, H * 2)], fill=(120, 120, 120), width=2)
    g.line([(0, H), (W * 2, H)], fill=(120, 120, 120), width=2)
    f_lab = font(16)
    labels = [('SOURCE', 0, 0), ('RENDER', W, 0), ('DIFF  ssim %.4f  mad %.2f' % (ssim, mad), 0, H), ('BLEND 50%', W, H)]
    for text, x, y in labels:
        tw = g.textlength(text, font=f_lab)
        g.rectangle([x, y, x + tw + 12, y + 22], fill=(20, 20, 20))
        g.text((x + 6, y + 2), text, fill=(255, 255, 255), font=f_lab)
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    grid.save(a.out)
    print(json.dumps({'ssim': round(ssim, 4), 'meanAbsDiff': round(mad, 2), 'diffBoxes': top}))


if __name__ == '__main__':
    main()
