# -*- coding: utf-8 -*-
"""读这份方案的配置 _project.json。

每份方案在 产出/ 下放一个 _project.json，把「会变的东西」全关在里面：
输出路径、强调色、章节名与章节色、字体、品牌名。kit.py / bg.py 都从这儿取，
所以换一份方案只要换这个文件，工具箱本身一个字都不用改。

查找顺序（找到就用，不再往下找）：
  1. 环境变量 PROPOSAL_PROJECT 指向的目录
  2. 从当前工作目录逐级往上找 _project.json
  3. 本文件所在目录

第 3 条基本命不中（脚本装在 skill 里，配置在产出目录里），所以
**「产出目录是命令行参数」的脚本必须自己调 use(目录)**，
别指望这三条能猜对。

字体那两条是历史包袱，别改：python-pptx 的 run.font.name 只写 <a:latin>，
中文必须另外写 <a:ea typeface>，否则 PowerPoint 里中文会掉字变方块。
"""
import os, json

_DEFAULTS = {
    'base': None,
    'accent': 'C9A227',
    'brand': '', 'sub_brand': '', 'client': '',
    'fonts': {'cn_serif': 'Songti SC', 'en_serif': 'Georgia',
              'cn_sans': 'PingFang SC', 'en_sans': 'Helvetica Neue'},
    'sections': {},
    'cover_bg': 'B-破晓.jpg', 'content_bg': 'A-雾.jpg', 'quiet_bg': 'D-静夜.jpg',
}


def _find():
    env = os.environ.get('PROPOSAL_PROJECT')
    if env:
        p = os.path.join(env, '_project.json')
        if os.path.exists(p):
            return p
    d = os.getcwd()
    for _ in range(6):
        p = os.path.join(d, '_project.json')
        if os.path.exists(p):
            return p
        nd = os.path.dirname(d)
        if nd == d:
            break
        d = nd
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_project.json')
    return p if os.path.exists(p) else None


CFG = dict(_DEFAULTS)
CFG['_path'] = None


def _load(path):
    """把 _project.json 读进 CFG。

    **就地改，不重新绑定**——kit.py 那边写的是 `from cfg import CFG`，
    换个新 dict 它看不见，还是拿着旧的空配置往下跑。
    """
    CFG.clear()
    CFG.update(_DEFAULTS)
    CFG.update(json.load(open(path, encoding='utf-8')))
    CFG['_path'] = path
    if not CFG.get('base'):
        CFG['base'] = os.path.dirname(path) + '/'
    return CFG


def use(d):
    """显式指定产出目录。

    给「产出目录是命令行参数」的脚本用（bg.py）。没有这个函数的话，
    它们只能靠 cwd 往上找——从别处运行时读不到 sections，章节背景
    一张都不出，而且**不报错**，等建 deck 时才发现少了图，很难倒查。
    """
    d = os.path.abspath(d)
    p = os.path.join(d, '_project.json')
    if not os.path.exists(p):
        raise SystemExit(
            '%s 下没有 _project.json。\n'
            '  产出目录要指向建过配置的那一层，不是它的上级。' % d)
    return _load(p)


_path = _find()
if _path:
    _load(_path)


def require():
    """要真配置的场景（建 deck）调一下，缺了就早报错，别拖到画到一半才炸。"""
    if not CFG.get('_path'):
        raise SystemExit(
            '找不到 _project.json。\n'
            '在产出目录下建一个，或设 PROPOSAL_PROJECT=<产出目录> 再跑。\n'
            '字段说明见 SKILL.md「STEP 4 风格与素材」。')
    if not CFG.get('sections'):
        raise SystemExit('_project.json 里没有 sections，章节页和目录页画不出来。')
    return CFG


if __name__ == '__main__':
    print(json.dumps(CFG, ensure_ascii=False, indent=2))
