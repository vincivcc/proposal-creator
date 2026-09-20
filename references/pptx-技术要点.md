# python-pptx 技术要点

这些全是踩过的坑。遇到诡异现象先翻这一页，八成有答案。
工具箱 `scripts/kit.py` 已经把这些都封装好了——**改工具箱之前先读完这份**。

---

## 一 · 中文字体：必须单独写 `<a:ea>`

**`run.font.name` 只写 `<a:latin>`，中文会掉字变方块。**

```python
for tag, name in (('latin', en), ('ea', cn)):
    e = rPr.find(qn('a:' + tag))
    if e is None: e = _el(tag, typeface=name); rPr.append(e)
    else: e.set('typeface', name)
```

`kit.F()` 已经做了。**自检渲染器也要单独读 `<a:ea typeface>`**，
否则它渲染出来的中文全是豆腐块，你会误判成自己的代码有问题。

## 二 · 换行：`\n` 会被吞掉

`run.text` 里的 `\n` python-pptx 不认，多行文案会挤成一团。
必须显式插入 `<a:br/>`：

```python
for i, seg in enumerate(text.split('\n')):
    if i: p._p.append(_el('br'))
    ...
```

`kit.R()` 已经做了。**渲染器也要按 XML 顺序遍历 `par._p`** 才能捞到 `<a:br/>`：

```python
for ch in par._p:
    if ch.tag == qn('a:br'): runs.append(None)      # 换行哨兵
    elif ch.tag == qn('a:r'): runs.append(ch)
```

## 三 · `a:tcPr` 的子元素有固定顺序

`lnL, lnR, lnT, lnB` **必须排在 fill 之前**，顺序错了 PowerPoint 会报文件损坏。

```python
for i, tag in enumerate(('lnL', 'lnR', 'lnT')):
    e = _el(tag, w='0'); e.append(_el('noFill')); tcPr.insert(i, e)
ln = _el('lnB', w='6350'); ...
tcPr.insert(3, ln)          # 注意是 insert 不是 append
```

顺带：`_el()` 会自动补 `a:` 命名空间，**不要再手写 `'a:lnL'`**，
否则会生成 `{...}a:lnL` 直接抛 `Invalid tag name`。

## 四 · 表格要中和掉默认样式

python-pptx 建的表格默认是蓝底斑马纹，和深色版面完全冲突：

```python
t.first_row = False; t.horz_banding = False
t._tbl.find(qn('a:tblPr')).find(qn('a:tableStyleId')).text = \
    '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'    # No Style, No Grid
```

然后自己画 `lnB` 横线。**深底上的表格只留横线，不要竖线不要底色。**

## 五 · 表格是 GraphicFrame，没有 `text_frame`

自检渲染器必须单独处理，否则**表格在预览里根本不出现**——
你会以为页面空了，其实是渲染器没画。

```python
if cls == 'GraphicFrame' or (st is not None and st == 19):
    if sh.has_table: render_table(...)
    continue
```

另外表格单元格没有 `left/top`，渲染时要自己按行列算几何，
所以 `render_text()` 需要一个以英寸为单位的 `box=(x,y,w,h)` 覆盖参数。

## 六 · 体积：靠复用，不靠压缩

python-pptx 按**内容 SHA1 去重**——同一张背景用在 20 页上，文件里只存一份。

所以正确做法是**一套背景家族复用全案**（约 10 张 × 170 KB = 1.7 MB），
而不是每页做一张新图。实测 52 页的幻灯组 1.8 MB，换成每页独立背景会到 35 MB。

单张背景的甜点参数：

```python
im.save(p, 'JPEG', quality=82, optimize=True, progressive=True, subsampling=2)
```

1920×1080 铺满 13.33in 仍有 144dpi。**2560px q92 会让文件大三倍**，
投影仪上看不出差别。

## 七 · 行高：估高不估低

自检渲染器 `render.py` 用固定的 `LH = 1.42` 估行高，**它不读 `line_spacing`**。
所以 `kit.th()` 里有个 `LH_MIN = 1.42` 下限，任何比它更紧的估计都会被抬回来。

估高了只是多留一点白；估低了渲染器就报溢出，而你会花时间去查一个
其实不存在的排版问题。

## 八 · 排版度量：1em 估算

中文排版没法精确测量（字体、字距、断行规则都会影响），用近似值够用了：

| 字符 | 宽度 |
|---|---|
| CJK / 全角 | `1.0 × pt`（粗体 `1.02`） |
| 拉丁字母数字 | `0.53 × pt`（粗体 `0.56`） |
| 空格 | `0.30 × pt` |
| 破折号、省略号、引号、书名号 | `1.0 × pt` |

行高 1.42–1.45。这套估算让**每个原语能自己算高度**，
调用方就不用逐页数行数了。

## 九 · 自检渲染器是必须的，不是可选的

肉眼翻缩略图**看不出溢出**——尤其是框比字大一点点的那些。
`render.py` 会逐页报「文字溢出」和「越界」，并给出框高 vs 需要的高度。

排完一定要跑：

```bash
python3 "$SKILL/scripts/render.py" <文件.pptx> /tmp/r
```

**但要诚实交代：预览是我自己写的渲染器画的，不是 PowerPoint 渲染的。**
它能验文件结构（越界 / 溢出 / 缺字），最终效果得用户打开看。
交付时主动说清这一点，不要让用户以为预览就是成品。

## 十 · 人物图：别抓，让用户给

不抓明星 / 博主头像。**抓到了也没法确认那张脸对着那个名字**，
方案里出现「脸和名字对不上」比留个空位置严重得多。

做法：`kit.portrait(slot=...)` 扫描 `产出/素材/<slot>.jpg`，
**有真图用真图（居中裁切填满 + 底部渐变压暗 + 姓名叠字），没有就出设计过的姓名卡**。

姓名卡不是灰块——深底 + 宋体姓氏 + 姓名/身份/入选理由，
**看着像是有意为之，不像缺图**。这一点很重要：测试版给客户看时不能露怯。

配套出一份 **换图清单**，说清每个位置换成什么、从哪来。

## 十一 · 产出目录是命令行参数时，必须 `cfg.use()`

`cfg.py` 默认靠 cwd 往上找 `_project.json`。但 `bg.py` 的产出目录是**命令行参数**，
从别处运行时 cwd 不对，就会读到一个空配置——`sections` 为空，
章节背景那个循环**跑 0 次，一张不出还不报错**。

要到建 deck 时才炸「找不到 `章节01-xxx.jpg`」，而且很难倒查回 bg.py。

所以：**凡是产出目录走 argv 的脚本，开头就要 `cfg.use(argv[1])` + `cfg.require()`。**
反过来，`kit.py` 这类靠 cwd 的（deck 脚本总在 `产出/_src/` 里跑）不用管。

排错口诀：**结果不对但不报错，先查配置读到哪一份了。**
