# -*- coding: utf-8 -*-
"""
译音表转写引擎
根据各语种译音表，将外国人名按音节转写为中文。
表数据参考新华通讯社发布的各语种译音表。
"""

def _build(cons, rows):
    """cons: 辅音列表；rows: {韵: 汉字串}，首字为零声母，其余与 cons 对齐。"""
    table = {}
    for rime, s in rows.items():
        chars = s.split()
        row = {'': chars[0] if chars else ''}
        for i, c in enumerate(cons):
            if i + 1 < len(chars):
                row[c] = chars[i + 1]
        table[rime] = row
    return table


# ===========================================================================
# 英语译音表：数据来自 英语译音表(全).xlsx（由 en_table.py 静态提供）
# 25 个辅音列（含音标键：z dz/s th/ʒ/ʃ/j dʒ/ch tʃ 等）；16 组韵母
# ===========================================================================
import en_table

EN_CONS = en_table.CON_KEYS
EN_TABLE = {}
for _rime, _row in en_table.RIMES.items():
    EN_TABLE[_rime] = dict(_row)

# 英语多字母辅音簇（整体识别为一个声母）
EN_CLUSTERS = {'ch', 'th', 'sh', 'zh', 'wh', 'ph', 'qu', 'tch', 'dge', 'ck', 'ng', 'wr'}

# 英语双元音（作为整体韵母识别）
EN_DIPHTHONGS = ['eau', 'ou', 'ow', 'oi', 'oy', 'ai', 'ay', 'au', 'aw',
                 'ee', 'ea', 'ei', 'ey', 'ie', 'oa', 'oo', 'ue', 'ui', 'eu', 'ew', 'ye']

# 英语拼写 → 表中辅音键的映射
EN_ONSET_MAP = {
    'b': 'b', 'p': 'p', 'd': 'd', 't': 't', 'g': 'g', 'k': 'k',
    'v': 'v', 'w': 'w', 'f': 'f', 'h': 'h', 'm': 'm', 'n': 'n',
    'l': 'l', 'r': 'r', 'c': 'k', 'q': 'k', 'x': 's th',
    's': 's th', 'z': 'z dz', 'ts': 'ts', 'ds': 'z dz',
    'th': 's th', 'sh': 'ʃ', 'zh': 'ʒ', 'ch': 'ch tʃ', 'tch': 'ch tʃ',
    'j': 'j dʒ', 'dge': 'j dʒ', 'ph': 'f', 'wh': 'hw',
    'qu': 'kw', 'ck': 'k', 'ng': 'n', 'wr': 'r', 'y': 'j',
}

# 英语韵母规范化
_EN_RIME_RULES = [
    ('a',   ('a', 'ā', 'â')),
    ('e',   ('e', 'ē', 'ey', 'ei', 'ea')),
    ('i',   ('i', 'ī', 'y', 'ee', 'ie')),
    ('o',   ('o', 'ō', 'ou', 'ow', 'oa', 'oo')),
    ('u',   ('u', 'oo', 'ue')),
    ('ju',  ('ū', 'eu', 'ew')),
    ('ə',   ('er', 'ur', 'ir', 'ər', 'ŭr', 'ů', 'ř')),
    ('ai',  ('ai', 'ay', 'ui')),
    ('au',  ('au', 'aw')),
    ('an',  ('an',)),
    ('ang', ('ang',)),
    ('en',  ('en', 'eng')),
    ('in',  ('in', 'yn')),
    ('ing', ('ing',)),
    ('un',  ('un', 'on', 'oun')),
    ('ong', ('ong',)),
]


def en_normalize_rime(rime, is_final=False):
    r = (rime or '').lower()
    # 先查整韵（含鼻音韵 on/un/ing 等），再剥离 n/ng 查元音
    for norm, aliases in _EN_RIME_RULES:
        if r in aliases:
            return norm
    base = r
    for suffix in ('ng', 'n'):
        if base.endswith(suffix):
            base = base[:-len(suffix)]
            break
    for norm, aliases in _EN_RIME_RULES:
        if base in aliases:
            return norm
    return 'a'





def _pad(cons, rows):
    """给每行补空占位直到 len(cons)+1 个 token（空表示该辅音列不适用）。"""
    padded = {}
    target = len(cons) + 1
    for k, v in rows.items():
        tokens = v.split()
        while len(tokens) < target:
            tokens.append('')
        padded[k] = ' '.join(tokens)
    return padded


# ===========================================================================
# 法语译音表：数据来自 法语译音表(全)_v5.xlsx（由 fr_table.py 静态提供）
# 36 个细分辅音列；PURE 为辅音单独存在(无元音)行；20 组韵母
# ===========================================================================
import fr_table

# 特殊韵母键：表示“辅音单独存在、不带元音”的音节，查纯辅音行
PURE = '\x00'

FR_CONS = fr_table.CON_KEYS
FR_TABLE = {}
for _rime, _row in fr_table.RIMES.items():
    FR_TABLE[_rime] = dict(_row)
# 挂上纯辅音行
FR_TABLE[PURE] = dict(fr_table.PURE)

# 法语二合/三合辅音字母：扫描时整体识别为一个声母（qu 的 u 不发音）
# ill/ge 不自动吞；cq 会误伤 Jacques 之类（c+qu）故不吞；gu 按软元音条件处理
FR_CLUSTERS = {k for k in FR_CONS if len(k) >= 2} - {'ill', 'ge', 'cq', 'gu'}
FR_SOFT_CLUSTERS = {'gu'}  # gu 仅在 e/i/y 前整体发硬音（u 不发音），如 guerre/Guy

# 法语韵母规范化（有序：长复合韵必须排在鼻化单韵之前）
# 每项 = (规范键, 命中拼写集合)；输入只可能是 ASCII 拼写 + n/m 鼻音韵尾
_FR_RIME_RULES = [
    ('iere', ('iere',)),
    ('ieu',  ('ieu',)),
    ('ien',  ('ien', 'yen')),
    ('ion',  ('ion', 'yon', 'iom', 'yom')),
    ('oin',  ('oin', 'oim')),
    ('au',   ('eau', 'eaux', 'aux', 'aou', 'ao', 'au')),
    ('ou',   ('oux', 'ou')),
    ('eu',   ('oeu', 'oeux', 'eu', 'oe')),
    ('ai',   ('ay', 'ai')),
    ('ie',   ('ille', 'ie')),
    ('oi',   ('oie', 'oua', 'oy', 'oi')),
    ('in',   ('ain', 'aim', 'ein', 'eim', 'yn', 'un', 'um', 'im', 'in')),
    ('an',   ('aen', 'ean', 'aon', 'en', 'em', 'am', 'an')),
    ('on',   ('om', 'on')),
    ('a',    ('a',)),
    ('i',    ('y', 'i')),
    ('o',    ('o',)),
    ('u',    ('uy', 'ui', 'u')),
    ('e',    ('ei', 'ey', 'e')),
]


def fr_normalize_rime(rime, is_final=False):
    """把拼写韵母映射到 FR_TABLE 的规范键。词尾裸 e 用开音节行。"""
    r = (rime or '').lower()
    if r == 'e' and is_final:
        return 'e_open'
    for norm, aliases in _FR_RIME_RULES:
        if r in aliases:
            return norm
    # 兜底：按首字母归入同头韵
    if r:
        for norm, aliases in _FR_RIME_RULES:
            if r[0] in aliases[0][0] or any(r.startswith(a[0]) for a in aliases):
                return norm
    return 'e'


# ===========================================================================
# 德意志译音表（原 DE_CONS 20 个，每行补末尾一个词使其 21 词对齐）
# ===========================================================================
DE_CONS = ['b','p','d','t','g','k','v','w','f','z','s','sch','tsch','ch','h','m','n','l','r','j']

DE_ROWS_RAW = {
    'a':   '阿 布 帕 达 塔 加 卡 夫 夫 弗 茨 斯 施 奇 赫 马 纳 拉 尔 亚',
    'e':   '埃 贝 佩 德 特 格 克 韦 费 策 塞 席 切 希 黑 梅 内 莱 雷 耶',
    'i':   '伊 比 皮 迪 蒂 吉 基 菲 维 齐 西 席 希 希 希 米 尼 利 里 伊',
    'o':   '奥 博 波 多 托 戈 科 沃 福 措 索 朔 乔 肖 霍 莫 诺 洛 罗 约',
    'u':   '乌 布 普 杜 图 古 库 武 富 楚 苏 舒 丘 休 胡 穆 努 卢 鲁 尤',
    'au':  '奥 鲍 保 道 陶 高 考 沃 福 曹 骚 绍 乔 肖 豪 毛 瑙 劳 劳 尧',
    'ei':  '艾 拜 派 代 泰 盖 凯 韦 法伊 蔡 赛 沙伊 切伊 沙伊 海 迈 奈 莱 赖 伊',
    'ie':  '伊 比 皮 迪 蒂 吉 基 菲 维 齐 西 席 希 希 希 米 尼 利 里 伊',
    'eu':  '奥伊 博伊 波伊 多伊 托伊 戈伊 科伊 沃伊 福伊 措伊 索伊 朔伊 乔伊 肖伊 霍伊 莫伊 诺伊 洛伊 罗伊 约伊',
    'ä':   '埃 贝 佩 德 特 格 克 韦 费 策 塞 席 切 希 黑 梅 内 莱 雷 耶',
    'ö':   '厄 伯 珀 德 特 格 克 弗 沃 弗 泽 瑟 舍 奇 赫 默 讷 勒 勒 耶',
    'ü':   '于 比 皮 迪 蒂 居 屈 菲 维 齐 西 席 希 希 希 米 尼 吕 吕 于',
    'an':  '安 班 潘 丹 坦 甘 坎 万 凡 赞 灿 桑 尚 昌 汉 曼 南 兰 朗 扬',
    'en':  '恩 本 彭 登 滕 根 肯 文 芬 曾 岑 森 申 琴 亨 门 嫩 伦 伦 因',
    'in':  '因 宾 平 丁 廷 金 金 温 芬 钦 辛 欣 钦 欣 欣 明 宁 林 林 因',
    'ing': '英 宾 平 丁 廷 京 金 温 芬 青 辛 兴 青 兴 兴 明 宁 林 林 英',
    'on':  '翁 邦 蓬 东 通 贡 孔 翁 丰 宗 聪 松 雄 琼 洪 蒙 农 隆 龙 永',
    'ung': '翁 本 蓬 敦 通 贡 孔 翁 丰 宗 聪 松 雄 琼 洪 蒙 农 隆 龙 永',
}
DE_TABLE = _build(DE_CONS, _pad(DE_CONS, DE_ROWS_RAW))


# ===========================================================================
# 俄语译音表（原 RU_CONS 20 个，每行补末尾直到 21 词）
# ===========================================================================
RU_CONS = ['b','p','d','t','g','k','v','f','z','s','zh','sh','j','ch','h','m','n','l','r','y']
RU_ROWS_RAW = {
    'a':   '阿 巴 帕 达 塔 加 卡 瓦 法 扎 萨 扎 沙 亚 恰 哈 马 纳 拉 尔 亚',
    'e':   '叶 别 佩 杰 捷 格 克 韦 费 泽 谢 热 舍 叶 切 赫 梅 涅 列 列 叶',
    'i':   '伊 比 皮 迪 季 吉 基 维 菲 齐 西 日 希 伊 奇 希 米 尼 利 里 伊',
    'o':   '奥 博 波 多 托 戈 科 沃 福 佐 索 若 绍 约 乔 霍 莫 诺 洛 罗 约',
    'u':   '乌 布 普 杜 图 古 库 武 富 祖 苏 茹 舒 尤 丘 胡 穆 努 卢 鲁 尤',
    'y':   '乌 布 普 杜 图 古 库 武 富 祖 苏 茹 舒 尤 丘 胡 穆 努 卢 鲁 尤',
    'ai':  '艾 拜 派 代 泰 盖 凯 韦 法 蔡 赛 扎伊 沙伊 亚伊 恰伊 海 迈 奈 莱 赖 亚伊',
    'ei':  '艾 拜 派 代 泰 盖 凯 韦 法 蔡 赛 扎伊 沙伊 亚伊 恰伊 海 迈 奈 莱 赖 亚伊',
    'ia':  '亚 比亚 皮亚 迪亚 蒂亚 贾 基亚 维亚 菲亚 齐亚 夏 里亚 夏 亚 恰亚 希亚 米亚 尼亚 利亚 里亚 亚',
    'ie':  '耶 别 佩 杰 捷 格 克 韦 费 泽 谢 热 舍 耶 切 赫 梅 涅 列 列 耶',
    'io':  '约 比奥 皮奥 迪奥 蒂奥 吉奥 基奥 维奥 菲奥 齐奥 肖 里奥 绍 约 乔 希奥 米奥 尼奥 利奥 里奥 约',
    'iu':  '尤 比乌 皮乌 久 久 久 丘 丘 维乌 菲乌 久 休 米乌 纽 柳 留 尤',
    'an':  '安 班 潘 丹 坦 甘 坎 万 凡 赞 桑 然 尚 扬 昌 汉 曼 南 兰 扬 扬',
    'en':  '恩 本 彭 登 滕 根 肯 文 芬 曾 森 任 申 因 钦 亨 门 嫩 伦 伦 因',
    'in':  '因 宾 平 丁 京 金 金 温 芬 钦 辛 欣 因 钦 欣 明 宁 林 林 因',
    'ing': '英 宾 平 丁 京 金 温 芬 青 辛 兴 英 青 兴 欣 明 宁 林 林 英',
    'on':  '翁 邦 蓬 东 通 贡 孔 翁 丰 宗 松 容 雄 永 琼 洪 蒙 农 隆 龙 永',
    'un':  '温 本 蓬 敦 通 贡 孔 温 丰 尊 孙 容 春 云 春 洪 蒙 农 伦 伦 永',
}
RU_TABLE = _build(RU_CONS, _pad(RU_CONS, RU_ROWS_RAW))


# ===========================================================================
# 西班牙语译音表：数据来自 西班牙语译音表(全).xlsx（由 es_table.py 静态提供）
# 35 个细分辅音列；PURE 为辅音单独存在(无元音)行；22 组韵母
# ===========================================================================
import es_table

ES_CONS = es_table.CON_KEYS
ES_TABLE = {}
for _rime, _row in es_table.RIMES.items():
    ES_TABLE[_rime] = dict(_row)
ES_TABLE[PURE] = dict(es_table.PURE)

# 西班牙语二合辅音字母：整体识别为一个声母（qu 的 u 不发音）
ES_CLUSTERS = {k for k in ES_CONS if len(k) >= 2} - {'gü', 'gu', 'cc', 'ck', 'cq', 'tch'}
ES_SOFT_CLUSTERS = {'gu'}  # gu 仅在 e/i/y 前整体发硬音（u 不发音），如 guerra

# 西班牙语双元音：y 等同 i（ya/ye/yu）；eu/io 在西语中是元音分立非双元音，故不列入
ES_DIPHTHONGS = ['eau', 'ou', 'au', 'ai', 'ei', 'oi', 'ui', 'oe', 'ia', 'ie',
                 'iu', 'ua', 'ue', 'uo', 'ae', 'ee', 'oo', 'oa', 'ow', 'oy',
                 'ya', 'ye', 'yu', 'yie', 'yhe']

# 西班牙语韵母规范化（有序：长复合韵必须排在鼻化单韵之前）
_ES_RIME_RULES = [
    ('uan', ('uan',)),
    ('ien', ('ien',)),
    ('ue',  ('uei', 'uey', 'ue')),
    ('ui',  ('uy', 'ui')),
    ('ion', ('ion',)),
    ('ia',  ('ya', 'ia')),
    ('ie',  ('yhe', 'yie', 'ye', 'ie')),
    ('iu',  ('yu', 'iu')),
    ('an',  ('aan', 'an')),
    ('au',  ('ao', 'au')),
    ('ai',  ('ae', 'ay', 'ai')),
    ('en',  ('een', 'ein', 'en')),
    ('in',  ('ing', 'yn', 'in')),
    ('on',  ('ung', 'oun', 'on')),
    ('un',  ('uen', 'un')),
    ('uo',  ('uo',)),
    ('ua',  ('ua',)),
    ('a',   ('ah', 'aa', 'a')),
    ('e',   ('ei', 'ey', 'e')),
    ('i',   ('y', 'i')),
    ('o',   ('ou', 'o')),
    ('u',   ('u',)),
]


def es_normalize_rime(rime, is_final=False):
    """把拼写韵母映射到 ES_TABLE 的规范键。"""
    r = (rime or '').lower()
    # 西语中 m 在辅音前/词尾发鼻化音，等同 n：om→on, am→an, em→en, im→in, um→un
    if r.endswith('m') and len(r) >= 2:
        r = r[:-1] + 'n'
    # 'io' 是弱+强双元音，表中无独立行，按强元音 o 行处理
    if r == 'io':
        return 'o'
    # 'eu' 在西语中是元音分立（非双元音），按 e 行处理
    if r == 'eu':
        return 'e'
    # 单独的 y 等同 i
    if r == 'y':
        return 'i'
    for norm, aliases in _ES_RIME_RULES:
        if r in aliases:
            return norm
    return 'a'


# ===========================================================================
# 葡萄牙语译音表：数据来自 葡萄牙语译音表(全).xlsx（由 pt_table.py 静态提供）
# 26 个辅音列；PURE 为辅音单独存在(无元音)行；21 组韵母
# ===========================================================================
import pt_table

PT_CONS = pt_table.CON_KEYS
PT_TABLE = {}
for _rime, _row in pt_table.RIMES.items():
    PT_TABLE[_rime] = dict(_row)
PT_TABLE[PURE] = dict(pt_table.PURE)

# 葡萄牙语二合辅音：lh(/ʎ/)、nh(/ɲ/)、ch(/ʃ/)、qu(/k/,u不发音)
PT_CLUSTERS = {'lh', 'nh', 'ch', 'qu'}
PT_SOFT_CLUSTERS = {'gu'}  # gu 仅在 e/i/y 前整体发硬音 g（u 不发音）

# 葡萄牙语双元音：ãe/õe 暂按元音分立处理
PT_DIPHTHONGS = ['eau', 'ou', 'au', 'ai', 'ei', 'oi', 'ui', 'oe', 'ia', 'ie', 'io',
                 'iu', 'ua', 'ue', 'uo', 'ae', 'ee', 'oo', 'oa', 'ow', 'oy',
                 'ya', 'ye', 'yu']

# 葡萄牙语韵母规范化
_PT_RIME_RULES = [
    ('a~',  ('ã', 'õ')),           # 鼻化元音
    ('an',  ('an',)),
    ('au',  ('ão', 'au')),         # ão 鼻化双元音
    ('en',  ('en', 'em')),
    ('in',  ('in', 'im')),
    ('on',  ('on', 'om')),
    ('ui',  ('ui', 'um')),         # um 鼻化
    ('un',  ('un',)),
    ('ia',  ('ya', 'ia')),
    ('ie',  ('ye', 'ie')),
    ('io',  ('io',)),
    ('iu',  ('iu',)),
    ('oa',  ('oa',)),
    ('ua',  ('ua',)),
    ('ue',  ('ue',)),
    ('a',   ('á', 'a')),
    ('e',   ('é', 'ey', 'ei', 'e')),
    ('i',   ('y', 'i')),
    ('o',   ('ou', 'o')),
    ('u',   ('u',)),
    ('ai',  ('ai',)),
]


def pt_normalize_rime(rime, is_final=False):
    """把拼写韵母映射到 PT_TABLE 的规范键。"""
    r = (rime or '').lower()
    for norm, aliases in _PT_RIME_RULES:
        if r in aliases:
            return norm
    return 'a'


# ===========================================================================
# 意大利译音表（已验证对齐：26 辅音 x 14 行，全部 OK）
# ===========================================================================
IT_CONS = ['b','p','d','t','g','gu','gh','gi','gl','gn','k','c','ch','f','v','z','s','sc','sch','h','m','n','l','r','y','qu']
IT_ROWS = {
    'a':   '阿 巴 帕 达 塔 加 瓜 格 吉 格利亚 尼亚 卡 卡 恰 法 瓦 扎 萨 斯卡 斯卡 哈 马 纳 拉 尔 亚 夸',
    'e':   '埃 贝 佩 代 泰 杰 圭 盖 杰 莱 涅 凯 凯 切 费 韦 泽 塞 谢 谢 赫 梅 内 莱 雷 耶 凯',
    'i':   '伊 比 皮 迪 蒂 吉 古 吉 吉 利 尼 基 基 奇 菲 维 齐 西 希 希 希 米 尼 利 里 伊 丘',
    'o':   '奥 博 波 多 托 戈 古 戈 乔 洛 尼奥 科 科 乔 福 沃 佐 索 斯科 斯科 霍 莫 诺 洛 罗 约 科',
    'u':   '乌 布 普 杜 图 古 古 古 古 卢 努 库 库 丘 富 武 祖 苏 斯库 斯库 胡 穆 努 卢 鲁 乌 库',
    'ia':  '亚 比亚 皮亚 迪亚 蒂亚 贾 瓜亚 吉亚 吉亚 利亚 尼亚 基亚 基亚 恰亚 菲亚 维亚 齐亚 夏 夏 夏 希亚 米亚 尼亚 利亚 里亚 亚 夸亚',
    'ie':  '耶 别 皮耶 迭 铁 杰 圭 盖 杰 列 涅 凯 凯 切耶 费耶 韦耶 泽耶 塞耶 谢耶 谢耶 赫耶 米耶 涅耶 列耶 列耶 耶 凯耶',
    'io':  '约 比奥 皮奥 迪奥 蒂奥 焦 古奥 戈 焦 利奥 尼奥 基奥 基奥 乔 菲奥 维奥 齐奥 肖 肖 肖 希奥 米奥 尼奥 利奥 里奥 约 基奥',
    'iu':  '乌 比乌 皮乌 迪乌 蒂乌 久 古 久 久 留 纽 久 久 丘乌 菲乌 维乌 齐乌 休 休 休 希乌 米乌 纽 留 留 乌 久丘',
    'an':  '安 班 潘 丹 坦 甘 宽 格兰 詹 兰 南 坎 坎 钱 凡 万 赞 桑 斯坎 斯坦 汉 曼 南 兰 朗 扬 宽',
    'en':  '恩 本 彭 登 滕 根 肯 真 根 伦 嫩 肯 肯 琴 芬 文 曾 森 斯肯 斯滕 亨 门 嫩 伦 伦 延 肯',
    'in':  '因 宾 平 丁 廷 金 金 欣 欣 林 尼 金 金 钦 芬 温 津 辛 斯钦 斯廷 欣 明 宁 林 林 因 钦',
    'on':  '翁 邦 蓬 东 通 贡 贡 贡 贡 隆 农 孔 孔 琼 丰 翁 宗 松 斯孔 斯托 洪 蒙 农 隆 隆 永 孔',
    'un':  '温 本 昆 敦 敦 贡 贡 春 春 伦 嫩 贡 贡 春 丰 文 尊 孙 斯昆 斯图恩 洪 蒙 嫩 伦 伦 云 贡',
}
IT_TABLE = _build(IT_CONS, IT_ROWS)


# ===========================================================================
# 各语种“辅音单独存在（不带元音）”行 —— 译音表顶部辅音横栏
# 游离的辅音（词首如 Mbeumo 的 m、词尾、音节间）查此行，而非硬配 e 韵
# 数据逐字核对自各语种译音表原图
# ===========================================================================
EN_PURE = {
    'b':'布','p':'普','d':'德','t':'特','g':'格','k':'克','v':'夫','w':'夫',
    'f':'夫','z':'兹','s':'斯','th':'斯','zh':'日','sh':'什','ch':'奇','j':'奇',
    'h':'赫','m':'姆','n':'恩','l':'尔','r':'尔','y':'伊','c':'克','qu':'库','x':'克斯',
}
DE_PURE = {
    'b':'布','p':'普','d':'德','t':'特','g':'格','k':'克','v':'夫','w':'夫',
    'f':'夫','z':'茨','s':'斯','sch':'施','tsch':'奇','ch':'希','h':'赫',
    'm':'姆','n':'恩','l':'尔','r':'尔','j':'伊',
}
RU_PURE = {
    'b':'布','p':'普','d':'德','t':'特','g':'格','k':'克','v':'夫','f':'夫',
    'z':'兹','s':'斯','zh':'日','sh':'什','j':'奇','ch':'奇','h':'赫',
    'm':'姆','n':'恩','l':'尔','r':'尔','y':'伊',
}
ES_PURE = None  # 已由 es_table.PURE 提供，见 ES_TABLE 构建处
IT_PURE = {
    'b':'布','p':'普','d':'德','t':'特','g':'格','gu':'古','gh':'格','gi':'吉',
    'gl':'尔','gn':'尼','k':'克','c':'克','ch':'克','f':'夫','v':'夫','z':'兹',
    's':'斯','sc':'斯克','sch':'斯克','h':'赫','m':'姆','n':'恩','l':'尔',
    'r':'尔','y':'伊','qu':'库',
}
EN_TABLE[PURE] = dict(en_table.PURE)
DE_TABLE[PURE] = DE_PURE
RU_TABLE[PURE] = RU_PURE
# ES_TABLE[PURE] 已在上方由 es_table.PURE 赋值
IT_TABLE[PURE] = IT_PURE


# ===========================================================================
# 语种到译音表的映射
# ===========================================================================
LANG_TABLE = {
    '英国': EN_TABLE, '法国': FR_TABLE, '德国': DE_TABLE, '俄罗斯': RU_TABLE,
    '西班牙': ES_TABLE, '意大利': IT_TABLE, '葡萄牙': PT_TABLE,
    # 没有专属译音表的语种 -> EN 表（通用拉丁字母译音）
    '荷兰': EN_TABLE, '瑞典': EN_TABLE, '挪威': EN_TABLE, '丹麦': EN_TABLE,
    '捷克': EN_TABLE, '波兰': EN_TABLE, '匈牙利': EN_TABLE, '罗马尼亚': EN_TABLE,
    '希腊': EN_TABLE, '保加利亚': EN_TABLE, '塞尔维亚': EN_TABLE,
    '阿尔巴尼亚': EN_TABLE, '拉脱维亚': EN_TABLE, '爱沙尼亚': EN_TABLE, '马耳他': EN_TABLE,
}
LANG_NAMES = sorted(LANG_TABLE.keys())


# ===========================================================================
# 籍贯缩写 -> 全称
# ===========================================================================
ABBREV_TO_FULL = {
    '英': '英国', '法': '法国', '德': '德国', '俄': '俄罗斯', '西': '西班牙',
    '意': '意大利', '荷': '荷兰', '瑞典': '瑞典', '挪': '挪威', '丹': '丹麦',
    '捷': '捷克', '波': '波兰', '匈': '匈牙利', '罗': '罗马尼亚', '葡': '葡萄牙',
    '希': '希腊', '保': '保加利亚', '塞': '塞尔维亚', '阿尔巴': '阿尔巴尼亚',
    '拉脱维亚': '拉脱维亚', '爱莎尼亚': '爱沙尼亚', '马耳他': '马耳他',
}


# ===========================================================================
# 性别后处理
# ===========================================================================
def _apply_gender(text, gender):
    """根据性别对译写结果做女性化后处理。"""
    if gender == 'female' and text:
        chars = list(text)
        # 最后一个字：常见女性尾音
        last_map = {'亚': '娅', '里': '丽', '纳': '娜', '尼': '妮',
                    '因': '茵', '林': '琳', '尔': '尔'}
        if chars and chars[-1] in last_map:
            chars[-1] = last_map[chars[-1]]
        # 中间的 亚→娅、里→丽
        for i in range(len(chars) - 1):
            if chars[i] == '亚':
                chars[i] = '娅'
            elif chars[i] == '里':
                chars[i] = '丽'
        return ''.join(chars)
    return text


# ===========================================================================
# 音节拆分与转写引擎
# ===========================================================================
import re

VOWELS = set('aeiouyàáâãäåèéêëìíîïòóôõöùúûüýÿæœ')
DIPHTHONGS = ['eau','ou','au','ai','ei','oi','ui','eu','oe','ia','ie','io','iu',
              'ua','ue','uo','ae','ee','oo','oa','ow','oy']


def _cut_consonant_run(run, cons, clusters=None):
    """把一段游离辅音串切成表里的辅音单元（最长匹配，从左到右）。"""
    units = []
    i = 0
    n = len(run)
    onset_set = set(cons)
    if clusters:
        onset_set = onset_set | set(clusters)
    while i < n:
        matched = None
        for L in range(min(3, n - i), 0, -1):
            cand = run[i:i + L]
            if cand in onset_set:
                matched = cand
                break
        if matched:
            units.append(matched)
            i += len(matched)
        else:
            units.append(run[i])
            i += 1
    return units


def _split_syllables(name, table=None, clusters=None, soft_clusters=None,
                     has_pure=False, fold_doubles=False, fold_regex=None,
                     diphthongs=None):
    """把名字拆成 (声母, 韵母) 音节。

    clusters:      该语言整体识别的辅音二合字母（如法语 ch/qu，其中 qu 的 u 不发音）
    soft_clusters: 仅在后面跟软元音 e/i/y 时才整体识别的簇（如法语 gu）
    has_pure:      该表是否含“纯辅音行”；为真时游离辅音单独成 PURE 音节，
                   否则沿用旧行为（硬配 e 韵），保证其他语言零回归。
    """
    name = name.lower().strip()
    parts = re.split(r'[^a-zà-ÿ]+', name)
    syllables = []
    cons = _table_consonants(table) if table else set()
    front_vowels = set('eiyéèêë')
    dips = diphthongs if diphthongs is not None else DIPHTHONGS
    for part in parts:
        if not part:
            continue
        # 双写辅音折叠（法语折叠全部；西语排除 ll/rr 等独立音位）
        if fold_doubles:
            part = re.sub(r'([bcdfghjklmnpqrstvwxz])\1+', r'\1', part)
        elif fold_regex:
            part = fold_regex.sub(r'\1', part)
        i = 0
        n = len(part)
        consonant_run = ''
        while i < n:
            if part[i] in VOWELS:
                nucleus = ''
                for d in dips:
                    if part[i:i+len(d)] == d:
                        nucleus = d
                        i += len(d)
                        break
                if not nucleus:
                    nucleus = part[i]
                    i += 1
                coda = ''
                if i < n and part[i] in 'nm':
                    if i + 1 >= n or part[i+1] not in VOWELS:
                        coda = part[i]
                        i += 1
                        if coda == 'n' and i < n and part[i] == 'g' and (i+1 >= n or part[i+1] not in VOWELS):
                            coda = 'ng'
                            i += 1
                # 切声母：辅音串末尾最长匹配（表辅音键 + 语言簇）
                onset = ''
                leftover = ''
                onset_set = set(cons)
                if clusters:
                    onset_set = onset_set | set(clusters)
                if onset_set:
                    for length in range(len(consonant_run), 0, -1):
                        cand = consonant_run[-length:]
                        if cand in onset_set:
                            onset = cand
                            leftover = consonant_run[:-length]
                            break
                if not onset and consonant_run:
                    onset = consonant_run[-1]
                    leftover = consonant_run[:-1]
                # 元音前游离的辅音前缀
                if leftover:
                    if has_pure:
                        for u in _cut_consonant_run(leftover, cons, clusters):
                            syllables.append((u, PURE))
                    else:
                        syllables.append((leftover, 'e'))
                syllables.append((onset, nucleus + coda))
                consonant_run = ''
            else:
                # 先尝试语言相关的二合/三合辅音簇
                hit = False
                if clusters or soft_clusters:
                    for L in (3, 2):
                        cand = part[i:i+L]
                        if clusters and cand in clusters:
                            consonant_run += cand
                            i += L
                            hit = True
                            break
                        if (soft_clusters and cand in soft_clusters
                                and i + L < n and part[i+L] in front_vowels):
                            consonant_run += cand
                            i += L
                            hit = True
                            break
                if not hit:
                    consonant_run += part[i]
                    i += 1
        # 词尾游离辅音
        if consonant_run:
            if has_pure:
                for u in _cut_consonant_run(consonant_run, cons, clusters):
                    syllables.append((u, PURE))
            else:
                syllables.append((consonant_run, 'e'))
    return syllables


def _table_consonants(table):
    cons = set()
    for row in table.values():
        cons.update(c for c in row if c)
    return cons


def _normalize_onset(onset, table):
    cons = _table_consonants(table)
    if not cons or not onset:
        return onset if onset else ''
    # 最长精确匹配
    sorted_cons = sorted(cons, key=len, reverse=True)
    for c in sorted_cons:
        if onset.startswith(c):
            return c
    # 回退：单字符 → 对应双辅音形式（法语表只有 mm 没有 m 等）
    # 多字符 onset（语言簇）原样返回，由各语种 onset_map 处理
    if len(onset) == 1:
        ch = onset[0]
        if ch:
            if ch in cons:
                return ch
            if ch + ch in cons:
                return ch + ch
    return onset


def _normalize_rime(rime, table):
    if rime in table:
        return rime
    base = rime
    for suffix in ('ng', 'n', 'm'):
        if base.endswith(suffix):
            base = base[:-len(suffix)]
            break
    if base in table:
        return base
    if base:
        first = base[0]
        candidates = [r for r in table if r.startswith(first)]
        if candidates:
            return candidates[0]
    return next(iter(table)) if table else ''


DEFAULT_CONSONANT_CHAR = {
    'b': '布', 'p': '普', 'd': '德', 't': '特', 'g': '格', 'k': '克',
    'v': '夫', 'w': '夫', 'f': '夫', 'z': '兹', 's': '斯', 'th': '斯',
    'zh': '日', 'sh': '什', 'ch': '奇', 'j': '奇', 'h': '赫', 'm': '姆',
    'n': '恩', 'l': '尔', 'r': '尔', 'y': '伊', 'c': '茨', 'x': '克斯',
    'q': '克', 'gu': '古', 'qu': '库',
}


def transliterate(name, lang='英国', gender='male'):
    table = LANG_TABLE.get(lang, EN_TABLE)
    is_fr = (lang == '法国')
    is_es = (lang == '西班牙')
    is_pt = (lang == '葡萄牙')
    is_en = (table is EN_TABLE)  # 英语及共享 EN_TABLE 的语种
    has_pure = PURE in table
    # 西班牙语：ll/rr 是独立音位，不折叠；其余双写辅音折叠
    # 葡萄牙语：lh/nh 是二合字母（已在 clusters 识别），其余双写折叠
    es_fold = re.compile(r'([bcdfghjkmnpqstvwxz])\1+') if is_es else None
    pt_fold = re.compile(r'([bcdfghjkmnpqrstvwxz])\1+') if is_pt else None
    if is_fr:
        clusters, soft, dips, fold_d, fold_r = FR_CLUSTERS, FR_SOFT_CLUSTERS, None, True, None
    elif is_es:
        clusters, soft, dips, fold_d, fold_r = ES_CLUSTERS, ES_SOFT_CLUSTERS, ES_DIPHTHONGS, False, es_fold
    elif is_pt:
        clusters, soft, dips, fold_d, fold_r = PT_CLUSTERS, PT_SOFT_CLUSTERS, PT_DIPHTHONGS, False, pt_fold
    elif is_en:
        clusters, soft, dips, fold_d, fold_r = EN_CLUSTERS, None, EN_DIPHTHONGS, True, None
    else:
        clusters = soft = dips = fold_r = None
        fold_d = False
    syllables = _split_syllables(
        name, table, clusters=clusters, soft_clusters=soft, has_pure=has_pure,
        fold_doubles=fold_d, fold_regex=fold_r, diphthongs=dips,
    )
    if not syllables:
        return name
    total = len(syllables)
    result = []
    for idx, (onset, rime) in enumerate(syllables):
        o = _normalize_onset(onset, table)
        # 英语：拼写字母 → 表中音标键（cluster 如 wh/th 直接查映射，避免被短辅音截断）
        if is_en:
            # 词中 h 不发音（John, Thomas, vehicle 等）；词首 h 发音
            if onset == 'h' and idx > 0:
                continue
            # c/g 软硬音：e/i/y 前发软音，否则硬音
            if onset == 'c':
                o = 's th' if (rime or '')[:1] in 'eiy' else 'k'
            elif onset == 'g':
                o = 'j dʒ' if (rime or '')[:1] in 'eiy' else 'g'
            else:
                o = EN_ONSET_MAP.get(onset) or EN_ONSET_MAP.get(o, o)
        # 西班牙语 x 发 s 音，对应表中 x* 列
        if is_es and o == 'x':
            o = 'x*'
        # 葡萄牙语 h 不发音：跳过
        if is_pt and o == 'h':
            continue
        # 游离辅音：查“辅音单独存在”的纯辅音行，查不到再用通用辅音字
        if rime == PURE:
            pure_row = table.get(PURE, {})
            ch = pure_row.get(o) or DEFAULT_CONSONANT_CHAR.get(o, o or '')
            result.append(ch)
            continue
        rime = rime or 'e'
        is_final = (idx == total - 1)
        if is_fr:
            # 外来名开音节 eu 实际读 /e/（辅音被后一元音音节借走，如 Mbeu-mo）：按 e 行译
            if (rime in ('eu', 'oeu') and not is_final
                    and syllables[idx + 1][1] != PURE):
                r = 'e'
            else:
                r = fr_normalize_rime(rime, is_final)
        elif is_es:
            r = es_normalize_rime(rime, is_final)
        elif is_pt:
            r = pt_normalize_rime(rime, is_final)
        elif is_en:
            r = en_normalize_rime(rime, is_final)
        else:
            r = _normalize_rime(rime, table)
        row = table.get(r, {})
        ch = row.get(o) or row.get('') or DEFAULT_CONSONANT_CHAR.get(o, onset or '')
        result.append(ch)
    text = ''.join(result)
    return _apply_gender(text, gender)
