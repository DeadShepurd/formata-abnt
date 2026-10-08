# Gera ícone e splash do app (Formata ABNT) para o projeto Android
import os, glob
from PIL import Image, ImageDraw

BLUE = (35, 64, 196)
PAPER = (243, 245, 250)
WHITE = (255, 255, 255)
INK = (35, 64, 196)
RES = 'android/app/src/main/res'

def page(size, scale=1.0):
    """folha A4 com canto dobrado e linhas de texto (parágrafo recuado)"""
    im = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    h = size * 0.62 * scale
    w = h / 1.414
    x0 = (size - w) / 2; y0 = (size - h) / 2
    fold = w * 0.28
    d.polygon([(x0, y0), (x0 + w - fold, y0), (x0 + w, y0 + fold), (x0 + w, y0 + h), (x0, y0 + h)], fill=WHITE)
    d.polygon([(x0 + w - fold, y0), (x0 + w - fold, y0 + fold), (x0 + w, y0 + fold)], fill=(196, 206, 245))
    lw = max(2, round(h * 0.055))
    m = w * 0.16
    ys = [0.30, 0.43, 0.56, 0.69, 0.82]
    for i, f in enumerate(ys):
        y = y0 + h * f
        left = x0 + m + (w * 0.18 if i in (0, 3) else 0)   # recuo de primeira linha
        right = x0 + w - m - (w * 0.22 if i in (2, 4) else 0)
        d.rounded_rectangle([left, y - lw / 2, right, y + lw / 2], radius=lw / 2, fill=INK)
    return im

def legacy(size, round_=False):
    big = size * 4
    im = Image.new('RGBA', (big, big), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    if round_:
        d.ellipse([0, 0, big - 1, big - 1], fill=BLUE)
    else:
        d.rounded_rectangle([0, 0, big - 1, big - 1], radius=big * 0.22, fill=BLUE)
    im.alpha_composite(page(big, 1.05))
    return im.resize((size, size), Image.LANCZOS)

def foreground(size):
    big = size * 4
    im = page(big, 0.62)   # zona segura do ícone adaptável
    return im.resize((size, size), Image.LANCZOS)

dens = {'mdpi': 1, 'hdpi': 1.5, 'xhdpi': 2, 'xxhdpi': 3, 'xxxhdpi': 4}
for k, f in dens.items():
    legacy(round(48 * f)).save(f'{RES}/mipmap-{k}/ic_launcher.png')
    legacy(round(48 * f), True).save(f'{RES}/mipmap-{k}/ic_launcher_round.png')
    foreground(round(108 * f)).save(f'{RES}/mipmap-{k}/ic_launcher_foreground.png')

with open(f'{RES}/values/ic_launcher_background.xml', 'w') as fh:
    fh.write('<?xml version="1.0" encoding="utf-8"?>\n<resources>\n    <color name="ic_launcher_background">#2340C4</color>\n</resources>\n')

for p in glob.glob(f'{RES}/drawable*/splash.png'):
    w, h = Image.open(p).size
    im = Image.new('RGBA', (w, h), PAPER + (255,))
    s = round(min(w, h) * 0.28)
    ic = legacy(s)
    im.alpha_composite(ic, ((w - s) // 2, (h - s) // 2))
    im.convert('RGB').save(p)

legacy(512).save('resources/icon-512.png')
print('ícones e splash gerados')
