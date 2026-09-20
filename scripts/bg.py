# -*- coding: utf-8 -*-
"""背景图合成 —— 无版权、可商用、可无限复刻。

方案里所有背景都在这儿用 numpy 算出来（数学渐变 + 光晕 + 噪点 + 暗角），
不去网上找图。理由见 references/红线与边界.md：搬别人的素材是红线。

改色只需改 _project.json 里的 accent 和 sections，重跑本脚本即可。
尺寸 1920×1080 是权衡过的：铺满 13.33in 仍有 144dpi，
换 2560px 会让文件大三倍，而投影仪上看不出差别。

用法：  python3 bg.py [产出目录]
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cfg import CFG
import numpy as np
from PIL import Image, ImageFilter

OUT = (sys.argv[1].rstrip('/') + '/_bg/' if len(sys.argv) > 1
       else CFG['base'] + '_bg/')
os.makedirs(OUT, exist_ok=True)

W, H = 1920, 1080
rng = np.random.default_rng(1103)          # 固定种子：每次重跑颜色一致

yy, xx = np.mgrid[0:H, 0:W].astype(np.float32)
u, v = xx / W, yy / H


def hex2rgb(h):
    return np.array([int(h[i:i + 2], 16) for i in (0, 2, 4)], dtype=np.float32)


def bloom(u, v, cx, cy, rx, ry, soft=1.6):
    """高斯光晕"""
    d2 = ((u - cx) / rx) ** 2 + ((v - cy) / ry) ** 2
    return np.exp(-d2 / soft)


def grain(shape, amount):
    return rng.normal(0, amount, shape)[..., None]


def vignette(u, v, strength=0.55, power=2.2):
    d = np.sqrt((u - 0.5) ** 2 * 1.0 + (v - 0.5) ** 2 * 1.6) / 0.62
    return 1.0 - strength * np.clip(d, 0, 1) ** power


def finish(rgb, gr=0.0, vig=0.0, uv=None, blur=0.0):
    if blur:
        img = Image.fromarray(np.clip(rgb, 0, 255).astype(np.uint8))
        rgb = np.asarray(img.filter(ImageFilter.GaussianBlur(blur))).astype(np.float32)
    if vig and uv is not None:
        rgb = rgb * vignette(uv[0], uv[1], vig)[..., None]
    if gr:
        rgb = rgb + grain(rgb.shape[:2], gr) * 255
    return np.clip(rgb, 0, 255).astype(np.uint8)


def save(arr, name):
    p = OUT + name
    Image.fromarray(arr).save(p, 'JPEG', quality=82, optimize=True,
                              progressive=True, subsampling=2)
    print('  %-24s %5.0f KB' % (name, os.path.getsize(p) / 1024))


# ── 调色板：深底不能是纯黑（投影会发灰），要带一点冷蓝 ──
NIGHT = hex2rgb('0A0D13')
NIGHT2 = hex2rgb('141B26')
GOLD = hex2rgb('C9A227')
WARM = hex2rgb('E8D5A8')
MOON = hex2rgb('B9CBDD')
TEAL = hex2rgb('1E3A44')

print('生成背景 → %s' % OUT)

# ── A · 雾 ── 一角冷光，低对比，压得住任何文字。全案主力底 ──
base = NIGHT[None, None, :] + (NIGHT2 - NIGHT)[None, None, :] * (0.35 + 0.65 * v)[..., None]
glow = bloom(u, v, 0.30, 0.22, 0.62, 0.95, 1.9)
rgb = base + (MOON - base) * (glow * 0.34)[..., None]
rgb += (GOLD - rgb) * (bloom(u, v, 0.30, 0.22, 0.30, 0.45, 2.4) * 0.07)[..., None]
save(finish(rgb, gr=0.016, vig=0.42, uv=(u, v), blur=1.2), 'A-雾.jpg')

# ── B · 破晓 ── 底部一线暖光。封面用 ──
base = NIGHT[None, None, :] + (NIGHT2 - NIGHT)[None, None, :] * (0.15 + 0.85 * (1 - v))[..., None]
dawn = np.clip((v - 0.52) / 0.48, 0, 1) ** 1.7
rgb = base + (WARM - base) * (dawn * 0.30)[..., None]
rgb += (GOLD - rgb) * (bloom(u, v, 0.72, 1.05, 0.75, 0.42, 2.0) * 0.22)[..., None]
save(finish(rgb, gr=0.018, vig=0.30, uv=(u, v), blur=1.6), 'B-破晓.jpg')

# ── C · 水感 ── 极缓波纹。需要质地暗示的页（产品/成分）用 ──
base = NIGHT[None, None, :] + (TEAL - NIGHT)[None, None, :] * (0.30 + 0.70 * v)[..., None]
wave = (np.sin(u * 6.2 + v * 2.1) * 0.5 + np.sin(u * 2.4 - v * 4.7) * 0.5 + 1) / 2
rgb = base + (MOON - base) * (bloom(u, v, 0.62, 0.30, 0.85, 1.1, 2.2) * (0.22 + 0.12 * wave))[..., None]
save(finish(rgb, gr=0.014, vig=0.46, uv=(u, v), blur=2.4), 'C-水感.jpg')

# ── D · 静夜 ── 近乎平黑 + 一道金线。目录/主张页/收尾用，最克制 ──
base = NIGHT[None, None, :] + (NIGHT2 - NIGHT)[None, None, :] * (0.55 + 0.45 * (1 - v))[..., None]
rgb = base + (MOON - base) * (bloom(u, v, 0.5, 0.05, 1.5, 0.75, 2.6) * 0.16)[..., None]
band = np.exp(-(((v - 0.78) / 0.020) ** 2))
rgb = rgb + (GOLD - rgb) * (band * 0.10)[..., None]
rgb = rgb + (GOLD - rgb) * (bloom(u, v, 0.86, 0.80, 0.30, 0.36, 2.0) * 0.12)[..., None]
save(finish(rgb, gr=0.013, vig=0.50, uv=(u, v), blur=1.0), 'D-静夜.jpg')

# ── 章节页：深底 + 单侧色晕，每个章节一个色 ──
for num, s in sorted(CFG['sections'].items()):
    c = hex2rgb(s['color'])
    base = NIGHT[None, None, :] + (c * 0.16 - NIGHT)[None, None, :] * (0.20 + 0.80 * v)[..., None]
    rgb = base + (c - base) * (bloom(u, v, 0.88, 0.30, 0.62, 0.85, 2.2) * 0.42)[..., None]
    rgb += (GOLD - rgb) * (bloom(u, v, 0.88, 0.30, 0.22, 0.30, 2.6) * 0.06)[..., None]
    save(finish(rgb, gr=0.014, vig=0.44, uv=(u, v), blur=1.4), '章节%s-%s.jpg' % (num, s['color']))
