# -*- coding: utf-8 -*-
"""版面工具箱 —— 建 PPT 用的全部原语。

设计思路：**每个原语自己算高度、自己贴边**。调用方只管按顺序往下排，
不用担心某段字多了一行就把页脚顶出界。所有「不能再往下了」的边界判断
都收敛到 FOOT_ZONE 这一个常量上。

一套底贯穿全案：10 张背景复用 50 多页，靠裁切和明度做变化。
python-pptx 按内容 SHA1 去重，所以复用不额外占体积。

配色和章节从 _project.json 读（见 cfg.py），换方案不用改这个文件。
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cfg import CFG
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree
from PIL import Image, ImageDraw, ImageFont

if not CFG.get('base'):
    raise SystemExit(
        'kit.py 找不到 _project.json，不知道这份方案的产出目录在哪。\n\n'
        '  先建 <产出目录>/_project.json，或者设环境变量再跑：\n'
        '      PROPOSAL_PROJECT=<产出目录> python3 deck.py\n\n'
        '  字段说明见 SKILL.md「4.2 建项目配置」。')
if not CFG.get('sections'):
    raise SystemExit('_project.json 里没有 sections，章节页和目录页画不出来。')

BASE = CFG['base']
BG   = BASE + '_bg/'
PH   = BASE + '_bg/ph/'
ASSET = BASE + '素材/'
os.makedirs(PH, exist_ok=True)

W, H = 13.333, 7.5        # 16:9，英寸
M    = 0.85               # 页边距
CW   = W - 2 * M          # 正文栏宽 11.633

# ── 字号：分三档，别混 ──────────────────────────────────
#   展示带（≥20pt）大标题 / 数字 / 大字主张 / 封面主标题。
#      这是版面的一部分，动它等于换设计，不跟着下面两档一起抬。
#   正文带（12–16pt）正文 / 导语 / 页脚 / 表格 / 目录明细 / 键值行。
#      **12pt 是下限**——投影仪上的可读线，再小现场就没人看得见。
#   角标带（9.5–11pt）**只给字距拉开的全大写小标签**：页眉右侧的英文章名、
#      封面右下那行英文、页码、姓名卡上的身份。这些是装饰性的定位标记，
#      不是给人读的正文，所以可以小；**别把正文往这一档里放**。
# 加新原语时照这三档挑，别顺手写个 9pt 给正文用。

# ── 配色 ────────────────────────────────────────────────
GOLD  = CFG['accent']
onN_1, onN_2, onN_3 = 'F2F0EC', 'B8B4AC', '7C7A74'   # 深底上的三级文字
onN_ln = '2A303B'
INK, INK2, INK3 = '1A1A1E', '4A4A52', '8A8A92'       # 浅底上的三级文字
LINE_L = 'E3E0D9'                                     # 浅底分隔线

class _Sec(dict):
    """章节号对不上时给一句人话。

    deck 脚本是从模板复制来改的，删几页、改章节号是常事。裸的 KeyError: '02'
    会让人以为是工具箱坏了，其实只是脚本里还留着一处引用别的方案才有的章节。
    """
    def __missing__(self, k):
        raise SystemExit(
            '章节 %r 不在 _project.json 的 sections 里。现有的：%s\n'
            '  改 deck 脚本里这一处，或者把该章节补进配置。'
            % (k, '、'.join(sorted(self)) or '（一个都没有）'))


SEC_COLOR = _Sec((k, v['color']) for k, v in CFG['sections'].items())
SEC_NAME  = _Sec((k, (v['name'], v['en'])) for k, v in CFG['sections'].items())
SEC_KEYS  = sorted(CFG['sections'])                   # 目录和章节页按这个顺序

_F = CFG['fonts']
CN_SERIF, EN_SERIF = _F['cn_serif'], _F['en_serif']
CN_SANS,  EN_SANS  = _F['cn_sans'],  _F['en_sans']

# ── 姓名卡要往图上画字，需要字体文件（PPT 正文用的是字体名，不依赖这些路径）──
# macOS 的 .ttc 要指定 index 取具体字重。换平台时按候选项往下找，
# 都找不到就退回 PIL 内置字体——中文会变方块，但不至于整个脚本崩掉。
def _font_file(cands, index=0):
    for p in cands:
        if os.path.exists(p):
            return p, index
    return None, 0

PF, _PFI = _font_file([
    '/System/Library/AssetsV2/com_apple_MobileAsset_Font8/'
    '86ba2c91f017a3749571a82f2c6d890ac7ffb2fb.asset/AssetData/PingFang.ttc',
    '/System/Library/Fonts/PingFang.ttc',
    '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
    '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',
    'C:/Windows/Fonts/msyh.ttc',
], index=7)
ST, _STI = _font_file([
    '/System/Library/Fonts/Supplemental/Songti.ttc',
    '/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc',
    'C:/Windows/Fonts/simsun.ttc',
], index=6)


def _ttf(path, index, size):
    """拿不到字体文件就退回 PIL 内置，别让整份方案卡在一张姓名卡上。"""
    try:
        return ImageFont.truetype(path, size, index=index)
    except Exception:
        return ImageFont.load_default()

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
def _el(tag, **kw):
    e = etree.Element('{%s}%s' % (A, tag))
    for k, v in kw.items(): e.set(k, str(v))
    return e

# ══ 排版度量 ═══════════════════════════════════════════
# 原语自己算高度、自己贴边，调用方就不用每页手数行数了。
import math
_WIDE = set('—…→←↑↓·　「」【】（）《》')
def _cw(ch, pt, bold):
    if ch in _WIDE: return pt
    o = ord(ch)
    if 0x2E80 <= o <= 0x9FFF or 0xFF00 <= o <= 0xFFEF: return pt * (1.02 if bold else 1.0)
    if ch == ' ': return pt * 0.30
    return pt * (0.56 if bold else 0.53)

def _wid(text, pt, bold=False):
    return sum(_cw(c, pt, bold) for c in text)

def nlines(text, pt, w_in, bold=False):
    """估算这段字在给定宽度下占几行"""
    return max(1, int(math.ceil(_wid(text, pt, bold) / max(1e-6, w_in * 72.0))))

LH_MIN = 1.42   # 自检渲染器 render.py 按这个系数估行高。宁可估高：
                # 估高了只是多留一点白，估低了下场就是渲染器报溢出。
def th(text, pt, w_in, bold=False, lh=1.45):
    """这段字在给定宽度下需要的高度（英寸）。

    两个坑，都踩过：
    1. 显式换行 \\n 必须单独算——它一定会断行，跟宽度无关。漏掉的话
       所有带 \\n 的原语都少算一行，文字从框里溢出去。
    2. 行距不能比 LH_MIN 更紧。文字框留够了，渲染器才认。
    """
    lh = max(lh, LH_MIN)
    return sum(nlines(seg, pt, w_in, bold) for seg in text.split('\n')) * pt * lh / 72.0

# ══ 文字 ═══════════════════════════════════════════════
def F(run, size=None, bold=None, color=None, cn=CN_SANS, en=EN_SANS, spc=None, italic=None):
    f = run.font
    if size is not None: f.size = Pt(size)
    if bold is not None: f.bold = bold
    if italic is not None: f.italic = italic
    if color is not None: f.color.rgb = RGBColor.from_string(color)
    rPr = run._r.get_or_add_rPr()
    if spc is not None: rPr.set('spc', str(int(round(spc * 100))))
    for tag, name in (('latin', en), ('ea', cn)):
        e = rPr.find(qn('a:' + tag))
        if e is None: e = _el(tag, typeface=name); rPr.append(e)
        else: e.set('typeface', name)

def tb(sl, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    b = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    t = b.text_frame
    t.word_wrap = True
    t.vertical_anchor = anchor
    t.margin_left = t.margin_right = t.margin_top = t.margin_bottom = 0
    return t

def P(t, first=False, align=PP_ALIGN.LEFT, before=0, after=0, line=None):
    p = t.paragraphs[0] if first else t.add_paragraph()
    p.alignment = align
    if before: p.space_before = Pt(before)
    if after:  p.space_after  = Pt(after)
    if line:   p.line_spacing = line
    return p

def R(p, text, *a, **kw):
    """加一段文字。文本里的 \\n 要变成真正的换行（<a:br/>）——
    run.text 里的 \\n 会被 python-pptx 吃掉，多行文案会挤成一团。"""
    segs = text.split('\n')
    last = None
    for i, seg in enumerate(segs):
        if i:
            p._p.append(_el('br'))
        if seg or len(segs) == 1:
            last = p.add_run(); last.text = seg; F(last, *a, **kw)
    return last

# ══ 形状 ═══════════════════════════════════════════════
def rect(sl, x, y, w, h, fill=None, line=None, lw=0.75, radius=None, shape=MSO_SHAPE.RECTANGLE):
    s = sl.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill:
        s.fill.solid(); s.fill.fore_color.rgb = RGBColor.from_string(fill)
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = RGBColor.from_string(line); s.line.width = Pt(lw)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    if radius is not None:
        try: s.adjustments[0] = radius
        except Exception: pass
    return s

def alpha(s, v):
    """给已填充的形状加透明度 v∈[0,1]"""
    c = s.fill._xPr.find(qn('a:solidFill')).find(qn('a:srgbClr'))
    c.append(_el('alpha', val=str(int(v * 100000))))

def line_h(sl, x, y, w, color, pt=0.75):
    c = sl.shapes.add_connector(1, Inches(x), Inches(y), Inches(x + w), Inches(y))
    c.line.color.rgb = RGBColor.from_string(color); c.line.width = Pt(pt)
    return c

# ══ 画布 ═══════════════════════════════════════════════
def deck():
    p = Presentation()
    p.slide_width, p.slide_height = Inches(W), Inches(H)
    return p

def _need_bg(bgfile):
    """背景没生成时给一句人话。

    不拦的话 python-pptx 会抛 FileNotFoundError，指向一张 .jpg——
    看起来像图片坏了，实际是「还没跑 bg.py」。顺序在 SKILL.md 4.3 / 4.5，
    但漏跑一次就要查半天。这里把话说明白。
    """
    if os.path.exists(BG + bgfile):
        return
    raise SystemExit(
        '背景图 %s 不存在——**先跑 bg.py 生成背景家族，再建 deck**：\n'
        '    python3 "$SKILL/scripts/bg.py" %s\n'
        '（deck 脚本只引用背景，不负责生成。顺序反了就是这个错。）'
        % (BG + bgfile, BASE))


def blank(prs, bgfile=None, scrim=0.0, solid='0A0D13'):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    if bgfile:
        _need_bg(bgfile)
        s.shapes.add_picture(BG + bgfile, 0, 0, width=prs.slide_width, height=prs.slide_height)
    else:
        f = s.background.fill; f.solid()
        f.fore_color.rgb = RGBColor.from_string(solid)
    if scrim:
        alpha(rect(s, 0, 0, W, H, fill='000000'), scrim)
    return s

# ══ 固定件 ═════════════════════════════════════════════
def eyebrow(s, left, right='', y=0.52):
    if left:
        t = tb(s, M, y, 8, 0.30); p = P(t, first=True)
        R(p, left, 11, True, GOLD, spc=1.6)
    if right:
        t = tb(s, M + CW - 5, y, 5, 0.30); p = P(t, first=True, align=PP_ALIGN.RIGHT)
        R(p, right, 9.5, False, onN_3, spc=1.8)

def title(s, text, y=0.94, size=30):
    h = th(text, size, CW, lh=1.30) + 0.10
    t = tb(s, M, y, CW, h); p = P(t, first=True, line=1.14)
    R(p, text, size, False, onN_1, cn=CN_SERIF, en=EN_SERIF)
    return y + h

def lead(s, text, y, w=None):
    w = w or CW
    h = th(text, 15, w) + 0.06
    t = tb(s, M, y, w, h); p = P(t, first=True)
    R(p, text, 15, False, onN_2)
    return y + h

FOOT_ZONE = H - 0.34          # 页脚**文字**允许的下边界（和页码同一行，左右分列，不打架）
BAR_ZONE  = H - 0.72          # 整幅宽的**色块**允许的下边界
# 为什么要两个：页码占 y ∈ [H-0.62, H-0.34]。满宽的色块（note_bar、cards_row
# 的卡片）如果只躲到 FOOT_ZONE，右下角会正好压住页码 —— 页脚文字没事是因为
# 它左对齐、页码右对齐，色块没有这个退路。这个坑出过一次，别再合并回去。
def foot(s, text, y=None, w=None):
    """页脚小字：按字数自适应高度，贴底对齐，绝不越界"""
    w = w or CW
    h = th(text, 10.5, w) + 0.06
    yy = (FOOT_ZONE - h) if y is None else min(y, FOOT_ZONE - h)
    t = tb(s, M, yy, w, h); p = P(t, first=True)
    R(p, text, 10.5, False, onN_3)
    return yy

def rule(s, x, y, w, color=GOLD, pt=1.2):
    return line_h(s, x, y, w, color, pt)

def card(s, x, y, w, h, fill=None, alpha_v=None, line=None, radius=0.055, pad=0.26):
    """半透明承载面板：深底上垫一层，保正文可读"""
    sh = rect(s, x, y, w, h,
              fill=fill or '161C26', line=line, radius=radius,
              shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    if alpha_v is not None: alpha(sh, alpha_v)
    return sh

# ══ 内容块 ═════════════════════════════════════════════
def bloc(s, x, y, w, items, gap=0.30, num_color=GOLD, head_size=14.5, body_size=12.5):
    """编号条目：一条一行，左编号右说明"""
    cy = y
    iw = w - 0.52
    for i, (head, body) in enumerate(items, 1):
        hh = th(head, head_size, iw, bold=True) + 0.05
        t = tb(s, x, cy, 0.40, hh); p = P(t, first=True)
        R(p, '%02d' % i, 11.5, True, num_color, spc=1.0)
        t = tb(s, x + 0.52, cy - 0.015, iw, hh); p = P(t, first=True)
        R(p, head, head_size, True, onN_1)
        cy += hh + 0.02
        if body:
            bh = th(body, body_size, iw) + 0.06
            t = tb(s, x + 0.52, cy, iw, bh); p = P(t, first=True)
            R(p, body, body_size, False, onN_2)
            cy += bh
        cy += gap
    return cy

def cards_row(s, y, h, items, gap=0.26, highlight=None, x=None, w=None, head_pt=15):
    """横排卡片。items=[(标题, 正文)]，highlight=高亮的下标集合。
    给定 h 不够时会自动长高；整体贴到页脚区之上。"""
    x = M if x is None else x
    w = CW if w is None else w
    n = len(items)
    cw = (w - (n - 1) * gap) / n
    hi = highlight or set()
    iw = cw - 0.52
    # 卡片至少要能装下「标题 + 最长正文」
    need = max(th(bd, 12, iw, lh=1.5) + 0.74 + 0.30 for _, bd in items)
    h = max(h, need + 0.06)
    y = min(y, BAR_ZONE - 0.10 - h)
    for i, (head, body) in enumerate(items):
        cx = x + i * (cw + gap)
        on = i in hi
        # 高亮不用整块亮色（太吵），改成「暖底 + 顶部金线 + 金色标题」
        card(s, cx, y, cw, h, fill=('241D10' if on else '161C26'),
             alpha_v=None if on else 0.72)
        if on:
            rect(s, cx, y, cw, 0.036, fill=GOLD)
        t = tb(s, cx + 0.26, y + 0.26, iw, head_pt * 1.5 / 72 + 0.06)
        p = P(t, first=True)
        R(p, head, head_pt, True, GOLD if on else onN_1)
        t = tb(s, cx + 0.26, y + 0.72, iw, max(0.30, h - 0.98))
        p = P(t, first=True, line=1.5)
        R(p, body, 12, False, onN_1 if on else onN_2)
    return y + h

def kv_rows(s, x, y, w, rows, k_w=2.5, gap=0.42, k_size=13, v_size=12.5):
    """键值行：左粗右细，行间细线"""
    cy = y
    for k, v in rows:
        h = max(th(k, k_size, k_w, bold=True), th(v, v_size, w - k_w)) + 0.06
        t = tb(s, x, cy, k_w, h); p = P(t, first=True)
        R(p, k, k_size, True, onN_1)
        t = tb(s, x + k_w, cy, w - k_w, h); p = P(t, first=True)
        R(p, v, v_size, False, onN_2)
        cy += h
        line_h(s, x, cy + 0.09, w, onN_ln, 0.5)
        cy += max(0.08, gap - 0.30)
    return cy

def stat(s, x, y, w, big, unit, label, big_size=54, color=None):
    bh = big_size * 1.42 / 72.0
    t = tb(s, x, y, w, bh + 0.06); p = P(t, first=True)
    R(p, big, big_size, False, color or onN_1, cn=CN_SERIF, en=EN_SERIF)
    if unit:
        R(p, ' ' + unit, 14, False, onN_2)
    lh_ = th(label, 12.5, w) + 0.06
    t = tb(s, x, y + bh + 0.12, w, lh_); p = P(t, first=True)
    R(p, label, 12.5, False, onN_2)
    return y + bh + 0.12 + lh_

def quote(s, text, y, size=40, sub=None, w=None):
    """大字主张：按换行估算高度，不再固定 1.13 英寸"""
    w = w or CW
    h = th(text, size, w, bold=False, lh=1.45) + 0.14
    t = tb(s, M, y, w, h); p = P(t, first=True, line=1.30)
    R(p, text, size, False, onN_1, cn=CN_SERIF, en=EN_SERIF)
    cy = y + h + 0.10
    if sub:
        sh_ = th(sub, 14, w) + 0.06
        t = tb(s, M, cy, w, sh_); p = P(t, first=True)
        R(p, sub, 14, False, onN_2)
        cy += sh_ + 0.12
    return cy

def shot(s, x, y, w, path, panel=True, pad=0.10):
    """截图条：把一张真截图按栏宽等比放上去，底下垫一层承载面板。

    高度**由图片自身的比例算，不写死**——截图是宽是扁由内容决定，
    写死高度就会拉伸变形。所以裁图的时候要按目标栏宽算好比例：
    比如打算放在 11.6 英寸宽的栏里、想要 1.7 英寸高，就裁成 6.8:1。
    """
    iw, ih = Image.open(path).size
    h = w * ih / iw
    if panel:
        card(s, x - pad, y - pad, w + 2 * pad, h + 2 * pad,
             fill='0E131B', alpha_v=0.55)
    s.shapes.add_picture(path, Inches(x), Inches(y), width=Inches(w))
    return y + h + (2 * pad if panel else 0)


def note_bar(s, text, y, w=None, color=GOLD):
    """提示条：这一页最该被记住的那一句。高度随字数长，整体不越界。

    y **接上一个原语的返回值**（`y = kv_rows(...)`，然后传 y + 0.30），
    别写死偏移 —— 写死的那个 y 是上一个原语**之前**的位置，提示条会正好
    按在正文中间，而且一声不响：越界检测只管页面边界，不管压没压住字。
    下面那个 min() 也帮不上忙，它只在「排到底了」的时候往上收。
    """
    w = w or CW
    th_ = th(text, 13.5, w - 0.60, bold=True) + 0.30
    y = min(y, BAR_ZONE - 0.10 - th_)
    sh = rect(s, M, y, w, th_, fill='1C1710', line=None,
              radius=0.10, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, M, y, 0.045, th_, fill=color)
    t = tb(s, M + 0.30, y + 0.14, w - 0.60, th_ - 0.24); p = P(t, first=True, line=1.35)
    R(p, text, 13.5, True, 'E4C97A')
    return y + th_

# ══ 表格 ═══════════════════════════════════════════════
NO_STYLE = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'

def table(s, x, y, w, head, rows, col_w=None, head_h=0.44, row_h=0.46,
          hi=None, size=12, head_size=11.5):
    """深底上的极简表：只有横线，无竖线无底色"""
    nr, nc = len(rows) + 1, len(head)
    shp = s.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(head_h + len(rows) * row_h))
    t = shp.table
    t.first_row = False; t.horz_banding = False
    t._tbl.find(qn('a:tblPr')).find(qn('a:tableStyleId')).text = NO_STYLE
    if col_w:
        tot = sum(col_w)
        for i, cwv in enumerate(col_w): t.columns[i].width = Emu(int(w * cwv / tot * 914400))
    t.rows[0].height = Inches(head_h)
    for i in range(len(rows)): t.rows[i + 1].height = Inches(row_h)
    hi = hi or set()

    def put(cell, txt, sz, bold, color, align=PP_ALIGN.LEFT):
        tf = cell.text_frame; tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.1)
        tf.margin_top = tf.margin_bottom = Inches(0.03)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = align
        R(p, txt, sz, bold, color)

    for c, htxt in enumerate(head):
        put(t.cell(0, c), htxt, head_size, True, onN_3)
    for r, row in enumerate(rows):
        on = r in hi
        for c, v in enumerate(row):
            put(t.cell(r + 1, c), v, size, on, onN_1 if on else onN_2,
                PP_ALIGN.LEFT)
    # 只留横线
    for r in range(nr):
        for c in range(nc):
            tc = t.cell(r, c)._tc
            tcPr = tc.get_or_add_tcPr()
            # tcPr 子元素有固定顺序：lnL,lnR,lnT,lnB 必须排在 fill 之前
            for i, tag in enumerate(('lnL', 'lnR', 'lnT')):
                e = _el(tag, w='0'); e.append(_el('noFill'))
                tcPr.insert(i, e)
            ln = _el('lnB', w='6350' if r else '9525')
            fill = _el('solidFill')
            fill.append(_el('srgbClr', val=onN_ln if r else '3A424F'))
            ln.append(fill); tcPr.insert(3, ln)
    return y + head_h + len(rows) * row_h

# ══ 人像占位卡（不抓真人照片，做成设计过的姓名卡）═══════
_MONO = {}
def _monogram(path, ch, name, role, note, w_in, h_in, accent):
    px = 640
    im = Image.new('RGB', (px, int(px * h_in / w_in)), '#141A23')
    d = ImageDraw.Draw(im, 'RGBA')
    for i in range(im.height):                      # 竖向渐变
        t = i / im.height
        d.line([(0, i), (px, i)], fill=(int(18 + 14 * t), int(23 + 18 * t), int(30 + 22 * t)))
    fch = _ttf(ST, _STI, int(px * 0.42))
    fnm = _ttf(PF, _PFI, int(px * 0.062))
    frl = _ttf(PF, _PFI, int(px * 0.040))
    fnt = _ttf(PF, _PFI, int(px * 0.034))
    d.text((px * 0.5, im.height * 0.42), ch, font=fch, fill=accent, anchor='mm')
    d.line([(px * 0.4, im.height * 0.63), (px * 0.6, im.height * 0.63)],
           fill=(58, 66, 79), width=2)
    d.text((px * 0.5, im.height * 0.70), name, font=fnm, fill=(238, 236, 231), anchor='mm')
    d.text((px * 0.5, im.height * 0.785), role, font=frl, fill=(150, 148, 142), anchor='mm')
    if note:
        d.text((px * 0.5, im.height * 0.855), note, font=fnt, fill=(120, 118, 113), anchor='mm')
    im.save(path, 'JPEG', quality=88)

def _fit(src, dst, w_in, h_in):
    """等比缩放 + 居中裁切，填满卡片，不留白边"""
    im = Image.open(src).convert('RGB')
    tw, th_ = 900, max(1, int(round(900 * h_in / w_in)))
    sc = max(tw / im.width, th_ / im.height)
    im = im.resize((max(tw, int(im.width * sc + .5)), max(th_, int(im.height * sc + .5))), Image.LANCZOS)
    l, t = (im.width - tw) // 2, (im.height - th_) // 2
    im = im.crop((l, t, l + tw, t + th_))
    # 底部压一层渐变暗角。真图上的姓名要是直接白字打上去，
    # 遇到浅色照片就糊了；垫一层渐变，任何底图都压得住。
    g = Image.new('L', (1, th_), 0)
    for i in range(th_):
        k = (i / th_ - 0.52) / 0.48
        g.putpixel((0, i), max(0, min(255, int(232 * k * k))))
    im = Image.composite(Image.new('RGB', im.size, '#0A0D11'), im, g.resize(im.size))
    im.save(dst, 'JPEG', quality=88, optimize=True)

def portrait(s, x, y, w, h, name, role, note='', accent=None, slot=None):
    """人物卡。素材/ 里有 slot 对应的图就用真图（居中裁切填满），没有就出姓名卡。"""
    real = None
    if slot:
        for ext in ('.jpg', '.jpeg', '.png', '.webp', '.JPG', '.PNG'):
            p = ASSET + slot + ext
            if os.path.exists(p): real = p; break
    key = PH + 'p_%s_%s.jpg' % (name, str(round(w * 10)))
    if real:
        _fit(real, key, w, h)
    elif not os.path.exists(key):
        _monogram(key, name[0], name, role, note, w, h, accent or '#8FA3B8')
    pic = s.shapes.add_picture(key, Inches(x), Inches(y), width=Inches(w), height=Inches(h))
    if real:
        # 真图没有姓氏字，姓名得单独补出来——否则换上真图之后，
        # 整页反而找不到这个人叫什么了。字号随卡片宽度走。
        t = tb(s, x + 0.20, y + h - 0.60, w - 0.40, 0.46)
        p = P(t, first=True)
        R(p, name, 13.5 if w >= 3.2 else 12, True, onN_1)
        R(p, '　' + role, 10, False, onN_2)
    return pic

# ══ 页面原型 ═══════════════════════════════════════════
def cover(prs, title_txt, sub, en, date, bg=None):
    """封面。品牌名从 _project.json 的 brand / sub_brand 来——
    不要在这儿写死客户名，那是每份方案都要换的东西。"""
    s = blank(prs, bg or CFG['cover_bg'])
    t = tb(s, M, 0.78, 8, 0.30); p = P(t, first=True)
    R(p, CFG['brand'], 11, False, GOLD, spc=2.4)
    if CFG.get('sub_brand'):
        R(p, '　·　', 11, False, onN_3)
        R(p, CFG['sub_brand'], 11, False, onN_2, spc=2.4)
    # 盒子高度和下面那道金线都跟着标题**量出来的**高度走。原来是写死 1.5 英寸
    # + 金线钉在 4.22——标题一折成两行就压在金线上。单行标题算出来还是 1.5，
    # 所以观感没变。
    h = th(title_txt, 76, CW, lh=1.26)
    t = tb(s, M, 2.62, CW, h); p = P(t, first=True, line=1.14)
    R(p, title_txt, 76, False, onN_1, cn=CN_SERIF, en=EN_SERIF)
    rule(s, M, 2.62 + h + 0.10, 1.55)
    t = tb(s, M, 2.62 + h + 0.40, 10, 0.40); p = P(t, first=True)
    R(p, sub, 16, False, onN_2)
    t = tb(s, M, 6.42, 7, 0.32); p = P(t, first=True)
    R(p, date, 11, False, onN_3, spc=1.4)
    t = tb(s, M + CW - 5, 6.42, 5, 0.32); p = P(t, first=True, align=PP_ALIGN.RIGHT)
    R(p, en, 9.5, False, onN_3, spc=2.0)
    return s

def toc(prs, items):
    """items=[(编号, 中文, 英文, 本章回答什么, 是否重心)]"""
    s = blank(prs, CFG['quiet_bg'])
    t = tb(s, M, 0.78, 8, 0.52); p = P(t, first=True)
    R(p, '目录', 22, False, onN_1, cn=CN_SERIF, en=EN_SERIF)
    R(p, '　CONTENTS', 10, False, onN_3, spc=2.2)
    y = 1.72
    for num, zh, en, what, key in items:
        c = SEC_COLOR[num] if key else onN_3
        t = tb(s, M, y, 0.7, 0.32); p = P(t, first=True)
        R(p, num, 14, True, c, spc=1.0)
        t = tb(s, M + 0.80, y - 0.05, 4.6, 0.36); p = P(t, first=True)
        R(p, zh, 18, False, onN_1 if key else onN_2, cn=CN_SERIF, en=EN_SERIF)
        R(p, '　' + en, 10, False, onN_3, spc=1.0)
        t = tb(s, M + 5.7, y + 0.04, CW - 5.7, 0.34); p = P(t, first=True)
        R(p, what, 12, False, onN_2 if key else onN_3)
        line_h(s, M, y + 0.46, CW, onN_ln, 0.5)
        y += 0.78
    return s

def divider(prs, num):
    zh, en = SEC_NAME[num]
    s = blank(prs, '章节%s-%s.jpg' % (num, SEC_COLOR[num]))
    rect(s, M, 2.62, 0.05, 1.62, fill=SEC_COLOR[num])
    t = tb(s, M + 0.34, 2.60, 9, 0.30); p = P(t, first=True)
    R(p, num, 12, True, SEC_COLOR[num], spc=1.6)
    t = tb(s, M + 0.34, 2.90, 10.5, 0.98); p = P(t, first=True)
    R(p, zh, 46, False, onN_1, cn=CN_SERIF, en=EN_SERIF)
    t = tb(s, M + 0.34, 3.98, 9, 0.30); p = P(t, first=True)
    R(p, en.upper(), 11, False, onN_3, spc=2.6)
    return s

def page(prs, sec, title_txt, lead_txt='', bg=None, scrim=0.42,
         title_size=30, eyebrow_right=None, foot_txt=None):
    """内容页骨架：背景 + 页眉 + 标题 + 导语，返回 (slide, y)"""
    s = blank(prs, bg or CFG['content_bg'], scrim=scrim)
    _, en = SEC_NAME[sec]
    eyebrow(s, '%s　%s' % (sec, SEC_NAME[sec][0]), eyebrow_right or en.upper())
    y = title(s, title_txt, 0.94, title_size)
    if lead_txt:
        y = lead(s, lead_txt, y + 0.08)
    if foot_txt:
        foot(s, foot_txt)
    return s, y + 0.16

def end(prs, line, sub='谢谢你的时间。', contact=None):
    """收尾页。line 是这页唯一那句大字，得为这份方案单独想——
    抄上一份的收尾句是最容易露怯的地方。

    contact 有就照原样打上，没有才退回占位符。**别把占位符写死**——
    联系方式是每份方案都不同、而且一定要填的东西，写死了就等着
    「〔提案方〕〔联系人〕」原样交到客户手上。用户给过就一定要用上。
    """
    s = blank(prs, CFG['quiet_bg'])
    GAP, SUB_H, CONTACT_Y = 0.14, 0.38, 6.50
    h = th(line, 60, CW, lh=1.26) + 0.10        # 按字数算，别写死
    # 副题**不能写死坐标**。h 是量出来的，line 一旦带 \n 变成两行，框就撑到
    # 6.09，而写死的 5.02 正好落在第二行上——出过一次，自检不报（越界检测
    # 只管页面边界，不管压没压住字），是翻图才看出来的。
    # 所以副题跟在框后；整块太靠下时往上收，落款位置不动。
    top = min(3.62, CONTACT_Y - GAP - SUB_H - GAP - h)
    t = tb(s, M, top, CW, h); p = P(t, first=True, line=1.14)
    R(p, line, 60, False, onN_1, cn=CN_SERIF, en=EN_SERIF)
    t = tb(s, M, top + h + GAP, CW, SUB_H); p = P(t, first=True)
    R(p, sub, 16, False, onN_2)
    t = tb(s, M, CONTACT_Y, CW, 0.32); p = P(t, first=True)
    R(p, contact or '〔提案方〕　　〔联系人〕　　〔邮箱 / 电话〕',
      11, False, onN_3, spc=0.8)
    return s


def page_number(prs, skip=2):
    """右下角页码。跳过前 skip 页（封面和目录不编号）。
    自检渲染器会找这个位置，所以别手写页码。"""
    for i, s in enumerate(prs.slides, 1):
        if i <= skip:
            continue
        t = tb(s, W - M - 1.0, H - 0.62, 1.0, 0.28)
        p = P(t, first=True, align=PP_ALIGN.RIGHT)
        R(p, '%02d' % i, 9.5, False, onN_3, spc=1.2)
