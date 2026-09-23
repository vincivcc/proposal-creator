# -*- coding: utf-8 -*-
"""通用 pptx 渲染器：把任意 .pptx 渲染成 PNG（自检 + 与参考方案比对）
用法: python3 _render_pptx.py <file.pptx> <outdir> [每页宽px] [页码...]
"""
import sys, os, io
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.enum.text import MSO_ANCHOR as MA, PP_ALIGN as PA

PF = ('/System/Library/AssetsV2/com_apple_MobileAsset_Font8/'
      '86ba2c91f017a3749571a82f2c6d890ac7ffb2fb.asset/AssetData/PingFang.ttc')
ST = '/System/Library/Fonts/Supplemental/Songti.ttc'
HN = '/System/Library/Fonts/HelveticaNeue.ttc'
GE = '/System/Library/Fonts/Supplemental/Georgia.ttf'
GB = '/System/Library/Fonts/Supplemental/Georgia Bold.ttf'
HI = '/System/Library/Fonts/Hiragino Sans GB.ttc'

FONTMAP = {
    'pingfang sc': (PF, 3), 'pingfang sc medium': (PF, 7), 'pingfang sc semibold': (PF, 11),
    'songti sc': (ST, 6), 'stsong': (ST, 4), 'songti tc': (ST, 7),
    'hiragino sans gb': (HI, 0),
    'georgia': (GE, 0), 'helvetica neue': (HN, 0), 'arial': (HN, 0),
    'optimael text': (GE, 0), 'optima lt std': (GE, 0), 'optima': (GE, 0),
    'mhei prc medium': (PF, 7), 'mhei prc': (PF, 3),
}
WARN = []          # 渲染时收集的问题：文字溢出 / 形状越界 / 文字压叠 / 色块压字
_LAYER = []        # 本页**绘制顺序**（px）。顺序就是 z 序，先画的在下。
                   # ('墨', x0,y0,x1,y1, 文字) = 一块字真实落到的范围
                   # ('块', x0,y0,x1,y1)      = 一块实色/半透明填充
_CUR = [0]         # 当前页号
_FC = ['']         # 当前页首个文字框内容，便于定位

def warn(kind, detail):
    WARN.append('P%02d  %s  %s' % (_CUR[0], kind, detail))

_fc = {}
def getfont(name, bold, px):
    key = ((name or '').lower().strip(), bool(bold), int(px))
    if key in _fc: return _fc[key]
    path, idx = FONTMAP.get(key[0]) or ((PF, 11) if bold else (PF, 3))
    if bold and path == GE: path, idx = GB, 0
    try:    f = ImageFont.truetype(path, max(6, int(px)), index=idx)
    except Exception: f = ImageFont.truetype(PF, max(6, int(px)), index=3)
    _fc[key] = f
    return f

WIDE = set('—…→←↑↓·　')
def is_cjk(ch):
    o = ord(ch)
    return (0x2E80 <= o <= 0x303F or 0x3400 <= o <= 0x4DBF or 0x4E00 <= o <= 0x9FFF
            or 0xF900 <= o <= 0xFAFF or 0xFF00 <= o <= 0xFFEF or ch in WIDE)

def atoms(t):
    out, buf = [], ''
    for ch in t:
        if is_cjk(ch):
            if buf: out.append(buf); buf = ''
            out.append(ch)
        elif ch == ' ':
            if buf: out.append(buf); buf = ''
            out.append(' ')
        else: buf += ch
    if buf: out.append(buf)
    return out

def segs(t):
    out, cur, kind = [], '', None
    for ch in t:
        k = 'cjk' if is_cjk(ch) else 'lat'
        if k != kind and cur: out.append((cur, kind)); cur = ''
        kind = k; cur += ch
    if cur: out.append((cur, kind))
    return out

LH = 1.42

def layout_par(par, maxw, PPT):
    """PPT = 每 pt 多少 px。返回 (行列表, 每行高px)
    注意：python-pptx 的 run.font.name 只映射 <a:latin>，中文得自己读 <a:ea>。"""
    items = []
    runs = []
    for ch in par._p:                       # 按 XML 顺序走，才能捞到 <a:br/>
        if ch.tag == qn('a:br'):
            runs.append(None)               # None = 强制换行
        elif ch.tag == qn('a:r'):
            runs.append(ch)
    for ch in runs:
        if ch is None:
            items.append(('\n', None, None, 0.0))
            continue
        r = next((x for x in par.runs if x._r is ch), None)
        if r is None or not r.text: continue
        sz = (r.font.size.pt if r.font.size else 14) * PPT
        spc = 0.0
        rPr = r._r.find(qn('a:rPr'))
        if rPr is not None and rPr.get('spc'): spc = float(rPr.get('spc')) / 100.0 * PPT
        col = (110, 110, 115)
        try:
            if r.font.color and r.font.color.rgb is not None:
                h = str(r.font.color.rgb); col = tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
        except Exception: pass
        bold = bool(r.font.bold)
        latin = r.font.name
        ea_el = rPr.find(qn('a:ea')) if rPr is not None else None
        ea = ea_el.get('typeface') if ea_el is not None else None
        fl = getfont(latin, bold, sz)
        # 没有 <a:ea>（英文模板常见）时，中文回退到 PingFang，否则会渲染成方块
        fe = getfont(ea, bold, sz) if ea else getfont('PingFang SC', bold, sz)
        for at in atoms(r.text):
            # atoms() 会把拉丁词合成多字符块，取首字符判断即可（CJK 永远单字成块）
            items.append((at, fe if is_cjk(at[0]) else fl, col, spc))
    if not items: return [], []
    lines, cur, curw = [], [], 0.0
    for at, f, col, spc in items:
        if f is None:                       # 强制换行
            lines.append(cur); cur, curw = [], 0.0
            continue
        w = f.getlength(at) + spc * len(at)
        if curw + w > maxw and cur:
            lines.append(cur); cur, curw = [], 0.0
            if at == ' ': continue
        cur.append((at, f, col, spc)); curw += w
    if cur: lines.append(cur)
    hs, prev = [], 14.0
    for ln in lines:
        szs = [a[1].size for a in ln if a[1] is not None]
        if szs: prev = max(szs)
        hs.append(prev * LH)
    return lines, hs

def render_text(dr, sh, PPT, PPI, ox=0.0, oy=0.0, box=None):
    """box=(x,y,w,h) 英寸；表格单元格没有 left/top，得由外面算好传进来"""
    tf = sh.text_frame
    bx, by, bw, bh = box or ((sh.left or 0) / 914400, (sh.top or 0) / 914400,
                             (sh.width or 0) / 914400, (sh.height or 0) / 914400)
    x0 = bx * PPI + ox
    y0 = by * PPI + oy
    w_in, h_in = bw, bh
    g = lambda v, d: (v.inches if v is not None else d)
    ml, mr = g(tf.margin_left, .1), g(tf.margin_right, .1)
    mt, mb = g(tf.margin_top, .05), g(tf.margin_bottom, .05)
    tx, ty = x0 + ml * PPI, y0 + mt * PPI
    ww = max(8.0, (w_in - ml - mr) * PPI)
    hh = (h_in - mt - mb) * PPI
    blocks = []
    for par in tf.paragraphs:
        lines, hs = layout_par(par, ww, PPT)
        sb = (par.space_before.pt if par.space_before else 0) * PPT
        sa = (par.space_after.pt if par.space_after else 0) * PPT
        blocks.append((par, lines, hs, sb, sa))
    total = sum(sb + sum(hs) + sa for _, _, hs, sb, sa in blocks)
    # 溢出检测：排下来比框还高，或横向超出框宽
    if total > hh + 1.5 and tf.text.strip():
        over = sum(len(l) for _, ls, _, _, _ in blocks for l in ls)
        warn('文字溢出', '框高 %.2fin 需 %.2fin（%d 字）「%s」'
             % (h_in, mt + total / PPI + mb, over, tf.text.strip().replace('\n', ' ')[:18]))
    for par, lines, _, _, _ in blocks:
        lw = sum(a[1].getlength(a[0]) + a[3] * len(a[0])
                 for ln in lines for a in ln if a[1] is not None) if lines else 0
        if lines and len(lines) == 1 and lw > ww + 2:
            warn('超宽未换行', '需 %.2fin 格宽 %.2fin「%s」' % (lw / PPI, ww / PPI, tf.text.strip()[:18]))
            break
    a = tf.vertical_anchor
    cy = ty + (hh - total) / 2 if a == MA.MIDDLE else (ty + hh - total if a == MA.BOTTOM else ty)
    ix0 = iy0 = 1e9; ix1 = iy1 = -1e9        # 这块字实际占到哪儿
    for par, lines, hs, sb, sa in blocks:
        cy += sb
        for ln, lh in zip(lines, hs):
            lw = sum(x[1].getlength(x[0]) + x[3] * len(x[0])
                     for x in ln if x[1] is not None)
            al = par.alignment
            cx = tx + (ww - lw) / 2 if al == PA.CENTER else (tx + ww - lw if al == PA.RIGHT else tx)
            base = cy + lh * 0.79
            if ln and lw > 0:                # 记的是落字范围，不是框——框留白大得多
                ix0 = min(ix0, cx); ix1 = max(ix1, cx + lw)
                iy0 = min(iy0, cy); iy1 = max(iy1, cy + lh)
            for at, f, col, spc in ln:
                if f is None: continue
                dr.text((cx, base), at, font=f, fill=col, anchor='ls')
                cx += f.getlength(at) + spc * len(at)
            cy += lh
        cy += sa
    if ix1 > ix0:
        _LAYER.append(('墨', ix0, iy0, ix1, iy1,
                       tf.text.strip().replace('\n', ' ')[:18]))

def solid_of(sh):
    """取形状的实色填充。**描边也算**——线是拿描边画的，不是拿填充画的。

    kit.line_h() / kit.rule() 用 add_connector 画横线，颜色写在 a:ln/a:solidFill，
    spPr 底下根本没有 a:solidFill。早先这里只看 spPr，于是**预览里所有横线
    都消失了**（封面那道金线、kv_rows 的行间线、目录的分隔线），而 pptx 里
    线是在的。预览缺一根线和设计缺一根线，看起来一模一样——所以这个 fallback
    不能省，否则「预览看着挺好」就成了一句空话。
    """
    try:
        # **不要走 sh.fill**：Connector 没有 .fill（会抛 AttributeError，
        # 被下面的 except 吃掉，于是线一声不响地消失）。直接读 XML。
        xpr = sh._element.find(qn('p:spPr'))
        if xpr is None: return None, 0
        el = xpr.find(qn('a:solidFill'))
        if el is None:                       # 退到描边
            ln = xpr.find(qn('a:ln'))
            el = ln.find(qn('a:solidFill')) if ln is not None else None
        if el is None: return None, 0
        c = el.find(qn('a:srgbClr'))
        if c is None: return None, 0
        h = c.get('val'); a = c.find(qn('a:alpha'))
        return tuple(int(h[i:i+2], 16) for i in (0, 2, 4)), \
               (int(a.get('val'))/100000.0 if a is not None else 1.0)
    except Exception:
        return None, 0

def group_xf(sh):
    """组的子坐标系 -> 页面坐标 的换算参数"""
    try:
        g = sh._element.find(qn('p:grpSpPr')).find(qn('a:xfrm'))
        off, ext = g.find(qn('a:off')), g.find(qn('a:ext'))
        co, ce = g.find(qn('a:chOff')), g.find(qn('a:chExt'))
        sx = ext.cx / ce.cx if ce.cx else 1.0
        sy = ext.cy / ce.cy if ce.cy else 1.0
        return sx, sy, off.x - co.x*sx, off.y - co.y*sy
    except Exception:
        return 1.0, 1.0, 0.0, 0.0

def _ln_color(tcPr, tag):
    """取单元格某条边框的颜色；w=0 或 noFill 视为无边框"""
    if tcPr is None: return None
    ln = tcPr.find(qn('a:' + tag))
    if ln is None: return None
    if ln.find(qn('a:noFill')) is not None: return None
    if ln.get('w') in ('0', None) and ln.find(qn('a:solidFill')) is None: return None
    c = ln.find(qn('a:solidFill') + '/' + qn('a:srgbClr'))
    return c.get('val') if c is not None else None

def render_table(dr, sh, PPT, PPI, img, ox=0.0, oy=0.0, sc=1.0):
    """画表格：底色 + 文字 + 横线。表格是 GraphicFrame，没有 left/top 之外的可绘制属性。"""
    t = sh.table
    x = (sh.left or 0) / 914400 * PPI * sc + ox
    y = (sh.top or 0) / 914400 * PPI * sc + oy
    ys, cy = [], y
    for row in t.rows:
        ys.append(cy); cy += row.height / 914400 * PPI * sc
    xs, cx = [], x
    for col in t.columns:
        xs.append(cx); cx += col.width / 914400 * PPI * sc
    for j, row in enumerate(t.rows):
        for i, cell in enumerate(row.cells):
            bx, by = xs[i], ys[j]
            bw, bh = t.columns[i].width / 914400, row.height / 914400
            if j >= len(t.rows) - 1: bh = None        # 末行吃掉剩余高度
            c, al = solid_of(cell)
            if c and al > 0.02:
                ov = Image.new('RGBA', img.size, (0, 0, 0, 0))
                ImageDraw.Draw(ov).rectangle([bx, by, bx + bw * PPI * sc, by + (bh or 1) * PPI * sc],
                                             fill=c + (int(al * 255),))
                b = img.convert('RGBA'); b.alpha_composite(ov)
                img.paste(b.convert('RGB'), (0, 0))
                dr = ImageDraw.Draw(img)
            render_text(dr, cell, PPT, PPI, ox, oy, box=(bx / (PPI * sc), by / (PPI * sc), bw, bh or 0.5))
            tcPr = cell._tc.find(qn('a:tcPr'))
            for tag, seg in (('lnB', [(bx, by + (bh or 0) * PPI * sc),
                                      (bx + bw * PPI * sc, by + (bh or 0) * PPI * sc)]),):
                h = _ln_color(tcPr, tag)
                if h:
                    dr.line(seg, fill=tuple(int(h[k:k+2], 16) for k in (0, 2, 4)),
                            width=max(1, int(0.75 * PPI / 72)))

def draw_shapes(dr, shapes, PPT, PPI, img, ox=0.0, oy=0.0, sc=1.0, depth=0):
    if depth > 6: return
    for sh in shapes:
        try: st = sh.shape_type
        except Exception: st = None
        cls = sh.__class__.__name__
        if cls == 'Picture':
            try:
                im = Image.open(io.BytesIO(sh.image.blob)).convert('RGBA')
                x = (sh.left or 0)/914400*PPI*sc + ox
                y = (sh.top or 0)/914400*PPI*sc + oy
                tw = max(1, int(round((sh.width or 0)/914400*PPI*sc)))
                th = max(1, int(round((sh.height or 0)/914400*PPI*sc)))
                im = im.resize((tw, th), Image.LANCZOS)
                base = img.convert('RGBA')
                base.alpha_composite(im, (int(x), int(y)))
                img.paste(base.convert('RGB'), (0, 0))
                dr = ImageDraw.Draw(img)
            except Exception: pass
            continue
        if cls == 'GraphicFrame' or (st is not None and st == 19):
            try:
                if sh.has_table:
                    render_table(dr, sh, PPT, PPI, img, ox, oy, sc)
                    dr = ImageDraw.Draw(img)
            except Exception: pass
            continue
        if st == 6 or cls == 'GroupShape':
            sx, sy, tx, ty = group_xf(sh)
            box = sh._element.find(qn('p:grpSpPr')).find(qn('a:xfrm'))
            co = box.find(qn('a:chOff')); ce = box.find(qn('a:chExt'))
            nox = ox + (tx - co.x*sx)*PPI/914400*sc*sx if False else ox + tx/914400*PPI*sc
            noy = oy + ty/914400*PPI*sc
            try: draw_shapes(dr, sh.shapes, PPT, PPI, img, nox, noy, sc*sx, depth+1)
            except Exception: pass
            continue
        if cls in ('AutoShape', 'Connector', 'Line') or (st is not None and st in (1, 5, 9)):
            c, al = solid_of(sh)
            if c and al > 0.02:
                x = (sh.left or 0)/914400*PPI*sc + ox
                y = (sh.top or 0)/914400*PPI*sc + oy
                w = (sh.width or 0)/914400*PPI*sc
                h = (sh.height or 0)/914400*PPI*sc
                _LAYER.append(('块', x, y, x + w, y + h))
                if al < 0.999:
                    ov = Image.new('RGBA', img.size, (0, 0, 0, 0))
                    ImageDraw.Draw(ov).rectangle([x, y, x+w, y+h], fill=c+(int(al*255),))
                    b = img.convert('RGBA'); b.alpha_composite(ov)
                    img.paste(b.convert('RGB'), (0, 0)); dr = ImageDraw.Draw(img)
                else:
                    dr.rectangle([x, y, x+w, y+h], fill=c)
        if sh.has_text_frame and sh.text_frame.text.strip():
            render_text(dr, sh, PPT, PPI, ox, oy)

def check_overlap(PPI):
    """同一页里**后画的东西压住先画的字** = 压叠。

    越界检测是逐形状的：框没出页面、字没撑破框，两块各自合格，
    压在一起却没人管。**它看不见形状之间的事。**
    P18 的副题压在大字第二行上、P13 的提示条按在正文上，两次都是这条漏的。

    两个坑：
    1. **比落字范围，不比框。** 框的留白比字大得多，拿框比会误报
       （封面的日期和右下英文同高，其实左右分列）。
    2. **色块单独算一条。** P13 那块提示条的底色压住了正文，而它自己那行字
       恰好落在两行正文的空隙里——只比「字 vs 字」是看不见的。
       绘制顺序就是 z 序，所以只看「排在后面的」。
    """
    T = 0.04 * PPI
    for i, a in enumerate(_LAYER):
        if a[0] != '墨': continue
        for b in _LAYER[i + 1:]:
            dx = min(a[3], b[3]) - max(a[1], b[1])
            dy = min(a[4], b[4]) - max(a[2], b[2])
            if dx <= T or dy <= T: continue
            if b[0] == '墨':
                warn('文字压叠', '「%s」↔「%s」 交叠 %.2f×%.2fin'
                     % (a[5], b[5], dx / PPI, dy / PPI))
            else:
                warn('色块压字', '「%s」被后画的色块盖住 %.2f×%.2fin'
                     % (a[5], dx / PPI, dy / PPI))

def render(path, outdir, width=1600, pages=None):
    os.makedirs(outdir, exist_ok=True)
    WARN.clear()
    prs = Presentation(path)
    SW, SH = prs.slide_width, prs.slide_height
    PPI = width / (SW/914400)
    PPT = PPI / 72.0
    made = []
    for i, s in enumerate(prs.slides, 1):
        if pages and i not in pages: continue
        _CUR[0] = i
        for sh in s.shapes:
            if sh.left is None: continue
            if (sh.left < -9144 or sh.top < -9144
                    or sh.left + (sh.width or 0) > SW + 9144
                    or sh.top + (sh.height or 0) > SH + 9144):
                warn('形状越界', '%s at (%.2f, %.2f) %.2fx%.2fin'
                     % (sh.shape_type, sh.left/914400, sh.top/914400,
                        (sh.width or 0)/914400, (sh.height or 0)/914400))
        Hp = int(round(width * SH / SW))
        img = Image.new('RGB', (width, Hp), (255, 255, 255))
        dr = ImageDraw.Draw(img)
        bg = s._element.find(qn('p:cSld')+'/'+qn('p:bg'))
        if bg is not None:
            sf = bg.find('.//'+qn('a:solidFill'))
            if sf is not None:
                c = sf.find(qn('a:srgbClr'))
                if c is not None:
                    h = c.get('val')
                    dr.rectangle([0, 0, width, Hp],
                                 fill=tuple(int(h[j:j+2], 16) for j in (0, 2, 4)))
        _LAYER.clear()                       # 逐页清：压叠只在同一页内算
        draw_shapes(dr, s.shapes, PPT, PPI, img)
        check_overlap(PPI)
        out = os.path.join(outdir, 'r%02d.png' % i)
        img.save(out)
        made.append(out)
    return made

if __name__ == '__main__':
    f, od = sys.argv[1], sys.argv[2]
    wd = int(sys.argv[3]) if len(sys.argv) > 3 else 1600
    pg = [int(x) for x in sys.argv[4:]] or None
    m = render(f, od, wd, pg)
    print('渲染 %d 页 -> %s' % (len(m), od))
    if WARN:
        print('\n⚠️  %d 处问题：' % len(WARN))
        for w in WARN: print('   ' + w)
    else:
        print('✅ 无溢出、无越界')
