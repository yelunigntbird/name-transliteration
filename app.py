# -*- coding: utf-8 -*-
"""
外国人名音译查询 —— 后端服务
- 加载 人名翻译.xlsx 作为主词典
- /            -> 前端页面
- /api/translate?name=...&origin=...  -> 查询结果（命中词典或用译音表补充）
- /api/origins -> 返回所有籍贯选项
"""
import os
import sys
import json
import re
import pickle
import openpyxl
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

import translit

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
XLSX_PATH = os.path.join(BASE_DIR, "人名翻译.xlsx")
INDEX_PATH = os.path.join(BASE_DIR, "index.html")
CACHE_PATH = os.path.join(BASE_DIR, ".names_cache.pkl")
CACHE_VERSION = 2  # 解析逻辑变更时递增，强制缓存失效

# 主词典：{name_lower: [(origin_str, chinese_str), ...]}
NAMES_DB = {}
# 所有有效籍贯全称（仅 22 种有译音表的语言）
ORIGINS = translit.LANG_NAMES
# xlsx 缩写 -> 全称
ABBREV_TO_FULL = translit.ABBREV_TO_FULL
# 全称 -> 缩写
FULL_TO_ABBREV = {v: k for k, v in ABBREV_TO_FULL.items()}


def _is_valid_origin(o):
    """只保留在 ABBREV_TO_FULL 中的籍贯缩写。"""
    return o in ABBREV_TO_FULL


# 说明性词汇：含这些词的片段不是译名本身（如"男子教名,源于法语。"）
_NOT_NAME_WORDS = ('教名', '源于', '姓氏', '小名', '昵称', '地名')


def _extract_name(c):
    """从 C 列文本提取真正的中文译名。

    词典里有三种形态：
      1) "约翰(男子教名…)"            -> 约翰
      2) "亚伦，《圣经》中…" / "亚伯,…" -> 亚伦
      3) "奥克亚<丹>诗人…" / "…〈瑞士〉"-> 奥克亚
    统一在第一个说明性分隔符处截断。
    返回 '' 表示该行没有可用译名（整句是注释，如"男子教名,源于法语。"）。
    """
    if c is None:
        return ''
    s = str(c).strip()
    # 注意：<丹>〈瑞士〉是国籍标记，必须作为切分点，不能当 HTML 标签删掉
    # ;；用于分隔多个译名（如"贝利;贝莉"），取第一个
    s = re.split(r'[（(<〈《\[【,，、;；>》〉\]】）)]', s, maxsplit=1)[0]
    s = s.strip(' ·.。:：;；')
    # 必须是纯汉字（含间隔号），长度 1~15，且不是说明性短语
    # 单汉字译名合法（如 King=金、Long=朗、Shaw=肖）
    if not (1 <= len(s) <= 15):
        # 编号列表格式："(1) 主教 (2) 毕晓普" — 取最后一个编号后的译名
        m = re.findall(r'\d+[.)）]\s*([一-鿿·]{2,})', str(c))
        if m:
            s = m[-1].strip()
        else:
            return ''
    if not re.fullmatch(r'[一-鿿·]+', s):
        return ''
    if any(w in s for w in _NOT_NAME_WORDS):
        return ''
    return s


def _name_keys(a):
    """A 列可能是 "姓，名" 形式（如 Aakjaer，Jeppe），额外收录逗号前的姓。"""
    keys = {a.lower()}
    if ('，' in a) or (',' in a):
        surname = re.split(r'[，,]', a, maxsplit=1)[0].strip()
        if len(surname) >= 2:
            keys.add(surname.lower())
    return keys


def load_xlsx():
    """加载词典：优先读 pickle 缓存；xlsx 有更新时才重新解析。"""
    xlsx_mtime = os.path.getmtime(XLSX_PATH)
    # 尝试读缓存
    if os.path.exists(CACHE_PATH):
        try:
            with open(CACHE_PATH, 'rb') as f:
                cached = pickle.load(f)
            if cached.get('version') == CACHE_VERSION and cached.get('mtime') == xlsx_mtime:
                NAMES_DB.update(cached['db'])
                print(f"从缓存加载词典：{sum(len(v) for v in NAMES_DB.values())} 条译名。", flush=True)
                return
        except Exception:
            pass  # 缓存损坏或版本不兼容，重新加载
    # 缓存失效：解析 xlsx
    print("正在加载 xlsx 词典...", flush=True)
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True, read_only=True)
    ws = wb["4"]
    count = 0
    skipped_empty_origin = 0
    skipped_no_name = 0
    for r in ws.iter_rows(values_only=True):
        a, b, c = r[0], r[1], r[2]
        if a is None or c is None:
            continue
        a = str(a).strip()
        b = str(b).strip() if b is not None else ''
        c = str(c).strip()
        if not a or not c:
            continue
        # 跳过标题行
        if a.startswith("第一部分") or a.startswith("第二部分"):
            continue
        # 跳过字母分段行
        if len(a) == 1 and a.isalpha() and a.isupper():
            continue
        # 提取纯译名；整句注释、没有译名的行直接跳过（之后走译音表）
        # 注意：不能因 B 列是生卒年（如 1866-1930）就跳过——具体人物行同样携带权威译名
        chinese_clean = _extract_name(c)
        if not chinese_clean:
            skipped_no_name += 1
            continue
        # 规范化籍贯
        if b:
            origins = re.split(r'[、;；,，]', b)
            origins = [o.strip() for o in origins if o.strip() and _is_valid_origin(o)]
            if origins:
                origin_tag = b                   # 正常多籍贯缩写串，如"英、法"
            elif re.search(r'\d', b):
                origin_tag = b                   # 生卒年（如 1866-1930）：保留，识别为人物条目
            else:
                origin_tag = '*'                 # 无法识别的籍贯：按通用条目处理
        else:
            origin_tag = '*'                     # 无籍贯：通用词头条目
            skipped_empty_origin += 1
        for key in _name_keys(a):
            NAMES_DB.setdefault(key, []).append((origin_tag, chinese_clean))
            count += 1
    wb.close()
    print(f"加载完成：{count} 条译名记录（含 {skipped_empty_origin} 条通用条目；"
          f"跳过 {skipped_no_name} 条无译名的纯注释行），{len(ORIGINS)} 种籍贯。", flush=True)
    # 写入缓存
    try:
        with open(CACHE_PATH, 'wb') as f:
            pickle.dump({'version': CACHE_VERSION, 'mtime': xlsx_mtime, 'db': NAMES_DB}, f, protocol=pickle.HIGHEST_PROTOCOL)
        print("词典缓存已写入。", flush=True)
    except Exception as e:
        print(f"警告：缓存写入失败 ({e})，不影响运行。", flush=True)


# 仅用于候选展示：白名单外、xlsx 中常见的籍贯缩写 -> 全称
_EXTRA_LABELS = {
    '以': '以色列', '科特': '科特迪瓦', '芬': '芬兰', '苏': '苏联', '奥': '奥地利',
    '比': '比利时', '瑞': '瑞士', '美': '美国', '加': '加拿大', '澳': '澳大利亚',
    '爱': '爱尔兰', '阿根': '阿根廷', '墨': '墨西哥', '印': '印度', '日': '日本',
    '韩': '韩国', '越': '越南', '泰': '泰国', '印尼': '印度尼西亚', '菲': '菲律宾',
    '埃': '埃及', '土': '土耳其', '新': '新西兰', '巴西': '巴西',
}


def _origin_label(orig_str):
    """把词条籍贯缩写串转成展示用全称；特殊标记转中文说明。"""
    if orig_str == '*':
        return '通用'
    if re.search(r'\d', orig_str):
        return '人物'
    labels = {**ABBREV_TO_FULL, **_EXTRA_LABELS}
    parts = [labels.get(p.strip(), p.strip())
             for p in re.split(r'[、;；,，]', orig_str) if p.strip()]
    return '、'.join(dict.fromkeys(parts)) or orig_str


def translate(name, origin, gender='male'):
    """返回查询结果字典。origin 为全称，gender 为 male/female。

    主译名严格按所选籍贯：
      1) 词典中籍贯包含所选语种的条目
      2) 词典中无籍贯的通用词头条目（如 John=约翰）
      3) 译音表转写
    其他籍贯的同名条目只出现在 candidates（相关条目）里，不顶替主译名。
    """
    name = name.strip()
    origin = origin.strip()
    if not name:
        return {"error": "请输入名字"}

    # 全称 -> 缩写，用于 xlsx 匹配
    abbrev = FULL_TO_ABBREV.get(origin, '')

    entries = NAMES_DB.get(name.lower(), [])

    exact, universal, other = [], [], []
    for orig_str, ch in entries:
        if orig_str == '*' or not abbrev:
            universal.append((orig_str, ch))
        elif re.search(r'\d', orig_str):
            other.append((orig_str, ch))          # 生卒年人物条目：仅供参考，不做主译名
        elif abbrev in set(re.split(r'[、;；,，]', orig_str)):
            exact.append((orig_str, ch))
        else:
            other.append((orig_str, ch))

    def _dedupe_ordered(pairs):
        seen, out = set(), []
        for orig_str, ch in pairs:
            if ch in seen:
                continue
            seen.add(ch)
            out.append({"origin": _origin_label(orig_str), "chinese": ch})
        return out

    # 相关条目：所选籍贯 + 通用 + 其他籍贯，全部可参考
    related = _dedupe_ordered(exact + universal + other)[:10]

    # 主译名：严格按籍贯 → 通用词头 → 译音表（绝不用其他籍贯顶替）
    if exact:
        main, source = exact[0][1], '词典'
    elif universal:
        main = universal[0][1]
        source = '词典' if not abbrev else '词典（通用条目）'
    else:
        main = translit.transliterate(name, origin if origin else '英国', gender)
        source = '译音表'

    # 始终按所选籍贯计算译音表结果，供对照展示
    phonetic = translit.transliterate(name, origin if origin else '英国', gender)

    return {
        "name": name, "origin": origin, "gender": gender,
        "chinese": main,
        "source": source,
        "phonetic": phonetic,
        "candidates": related,
    }


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

    def _send_json(self, obj, status=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/" or parsed.path == "/index.html":
            try:
                with open(INDEX_PATH, "rb") as f:
                    body = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except FileNotFoundError:
                self._send_json({"error": "index.html 未找到"}, 404)
            return

        if parsed.path == "/api/origins":
            origins = sorted(ORIGINS)
            self._send_json({"origins": origins})
            return

        if parsed.path == "/api/translate":
            qs = parse_qs(parsed.query)
            name = qs.get("name", [""])[0]
            origin = qs.get("origin", [""])[0]
            gender = qs.get("gender", ["male"])[0]
            result = translate(name, origin, gender)
            self._send_json(result)
            return

        self._send_json({"error": "Not Found"}, 404)


def main():
    load_xlsx()
    port = int(os.environ.get("PORT", 8000))
    server = HTTPServer(("0.0.0.0", port), Handler)
    print(f"服务已启动：http://0.0.0.0:{port}", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        server.shutdown()


if __name__ == "__main__":
    main()
