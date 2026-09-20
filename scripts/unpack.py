# -*- coding: utf-8 -*-
"""把 风格与素材.html 导出的「素材包.json」解成一张张图片，放进 产出/素材/。

为什么需要这个：素材上传走的是浏览器的 File System Access API，能直接写进
产出/素材/。但 Safari / Firefox 不支持，用户拒绝权限时也不行——那就降级成
下载一个 json，由本脚本解开。两条路殊途同归，图片始终只落在本机，
不上传任何外部服务。

用法：  python3 unpack.py [产出目录]
"""
import json, base64, os, glob, sys

BASE = (sys.argv[1] if len(sys.argv) > 1 else
        os.path.dirname(os.path.abspath(__file__)) + '/../../../../产出/')
DIR = os.path.abspath(BASE) + '/素材/'
os.makedirs(DIR, exist_ok=True)

packs = (glob.glob(DIR + '素材包*.json')
         + glob.glob(os.path.expanduser('~/Downloads/素材包*.json')))
if not packs:
    print('没找到素材包.json。找过：\n  %s\n  ~/Downloads/' % DIR)
    raise SystemExit(1)

for p in packs:
    print('读 %s' % p)
    d = json.load(open(p, encoding='utf-8'))
    n = 0
    for _page, arr in d.get('files', {}).items():
        for item in arr:
            name, data = item['name'], item['data']
            b = base64.b64decode(data.split(',', 1)[1] if ',' in data else data)
            open(DIR + name, 'wb').write(b)
            n += 1
    man = {'style': d.get('style', {}), 'notes': d.get('notes', {}),
           'savedAt': d.get('savedAt')}
    json.dump(man, open(DIR + '_素材清单.json', 'w', encoding='utf-8'),
              ensure_ascii=False, indent=2)
    os.rename(p, p + '.done')      # 改名而不是删除：解错了还能回头查
    print('  ✅ 解开 %d 张图，清单已写入' % n)

print('\n产出/素材/ 现在有：')
for f in sorted(os.listdir(DIR)):
    if not f.endswith('.json'):
        print('  %-22s %6.0f KB' % (f, os.path.getsize(DIR + f) / 1024))
