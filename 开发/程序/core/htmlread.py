# -*- coding: utf-8 -*-
"""读取「其实是网页表格」的文件 —— 后缀写成 .xls 的 HTML，以及 .html/.htm 直传。

为什么要有这个模块：
- 很多 ERP / 网银 / 政务系统的「导出 Excel」导出的其实是 HTML 表格，只是文件名写成 .xls；
- 用 Excel 打开会弹「格式与扩展名不符」，有时候还整片乱码（改后缀成 .html 就正常）；
- 平台原来只看后缀，`.xls` 非 OLE2 就一律当 xlsx，最后死在 zipfile 报
  "File is not a zip file"，被前端翻译成「文件可能已损坏」——方向全错。

实现原则（与 xlsread.py / xlsxinfo.py 一致）：**纯标准库 + openpyxl**，不引第三方。
运行时只有 openpyxl 与 duckdb，装不了 lxml/bs4。

对外：looks_html / decode_text / read_tables / to_xlsx，失败一律抛 HtmlError（中文人话）。
"""
import os
import re
from html.parser import HTMLParser

try:
    from openpyxl import Workbook
except Exception:                                   # pragma: no cover
    Workbook = None

#: 整个网页文件上限。再大就该让人用 Excel 另存了，硬撑会吃光内存
MAX_BYTES = 40 << 20
#: 单份文件最多转多少张表（ERP 导出偶尔带几百个小表）
MAX_TABLES = 200
#: 单张表单元格上限（普通 Workbook 也是内存型的）
MAX_CELLS = 400_000
#: 单表列数上限，防 colspan="1000" 这种占位格把表撑爆
MAX_COLS = 512
#: Excel 单元格字符上限
MAX_TEXT = 32767

_BAD_NAME = re.compile(r"[\[\]:*?/\\]")
_CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_WS = re.compile(r"[ \t\r\n\u00a0\u3000]+")
_CHARSET = re.compile(r"charset\s*=\s*[\"']?([\w\-]+)", re.I)
_HEAD_JUNK = re.compile(r"^\s*(?:<!--.*?-->|<\?xml[^>]*\?>|<!doctype[^>]*>)", re.S | re.I)


class HtmlError(Exception):
    """网页表格解析失败。message 一定是能直接给用户看的中文。"""


# ------------------------------------------------------------------ 判定
def looks_html(head, sniff=8192):
    """只看文件头判断「这是不是网页表格」。只有 True 才可能进 HTML 转换通道。

    判据（宁窄勿宽，别把正文里带 < 的文本账表误伤）：
      1) 前 8KB 里出现 <table
      2) 去掉 BOM/注释/XML 声明/DOCTYPE 后以 <html 或 <table 开头
         —— Excel「另存为网页」的 <head> 里样式动辄几 KB，<table 常落在 8KB 之外，
            所以必须有第 2 条，否则刚好漏掉最常见的形态。
    """
    b = bytes(head or b"")
    if not b:
        return False
    if b[:2] in (b"\xff\xfe", b"\xfe\xff"):
        s = b[:sniff].decode("utf-16", "ignore")
    else:
        if b[:3] == b"\xef\xbb\xbf":
            b = b[3:]
        s = b[:sniff].decode("latin-1", "ignore")
    low = s.lower()
    if "<table" in low:
        return True
    t = _HEAD_JUNK.sub("", low).lstrip()
    return t.startswith("<html") or t.startswith("<table") or t.startswith("<body")


def decode_text(raw):
    """返回 (全文文本, 实际编码名)。

    BOM 优先；其次看 meta charset；声明不可信时按 utf-8 → gb18030 → gbk → big5 回退。
    必须**全文 strict 解码**（不能只试开头）：HTML 尾部常出现生僻字，
    采样通过而全文失败会得到一堆 U+FFFD，却当成转换成功。
    中文客户导出普遍标 gb2312 实为 gbk 扩展字集，所以统一升级到 gb18030。
    """
    if not raw:
        return "", "utf-8"
    if raw[:3] == b"\xef\xbb\xbf":
        return raw[3:].decode("utf-8", "replace"), "utf-8-sig"
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return raw.decode("utf-16", "replace"), "utf-16"
    declared = None
    m = _CHARSET.search(raw[:8192].decode("latin-1", "ignore"))
    if m:
        declared = m.group(1).lower().replace("_", "-")
        if declared in ("gb2312", "gbk", "gb-2312", "cp936", "x-gbk"):
            declared = "gb18030"
        if declared == "utf8":
            declared = "utf-8"
    order, seen = [], set()
    for enc in ([declared] if declared else []) + ["utf-8", "gb18030", "gbk", "big5"]:
        if not enc or enc in seen:
            continue
        seen.add(enc)
        try:
            return raw.decode(enc), enc
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode("utf-8", "replace"), "utf-8(uncertain)"


# ------------------------------------------------------------------ 解析
def _int(v, default):
    try:
        return max(1, int(str(v).strip()))
    except Exception:
        return default


def _clean(s):
    """单元格/表名清洗：空白归一、去 Excel 非法控制符、截断。

    注意**不转数字/日期**：平台读 xlsx 用的是 all_varchar=true，
    保持原字符串才不会把 1,234.00 变成 1234、把 2026-01-05 变成序列号 46027。
    """
    s = _CTRL.sub("", s or "")
    s = _WS.sub(" ", s).strip()
    return s[:MAX_TEXT]


class _TableParser(HTMLParser):
    """把 HTML 拆成若干 table 节点。

    与常见的「单份全局状态」写法不同，每个 table 自带 rows/cell 状态（用栈），
    这样嵌套表不会把外层单元格的内容吃掉；数据只落到**最内层**表格的当前单元格，
    避免同一笔金额同时出现在父表和子表里被重复统计。
    """

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables = []          # 先序（打开即登记），保证编号稳定
        self._stack = []
        self._skip = 0            # script/style/title 里不采文本
        self._page_text = []      # 表外文字，作表名候选

    # -- 标签
    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag in ("script", "style", "title", "noscript"):
            self._skip += 1
            return
        if tag == "table":
            node = {"rows": [], "caption": "", "children": [], "prev_text": ""}
            if self._stack:
                self._stack[-1]["children"].append(node)
            else:
                node["prev_text"] = "".join(self._page_text)[-80:]
            self._stack.append(node)
            self.tables.append(node)
            return
        if not self._stack:
            return
        node = self._stack[-1]
        if tag == "caption":
            node["_cap"] = True
        elif tag == "tr":
            node["rows"].append([])
        elif tag in ("td", "th"):
            if not node["rows"]:
                node["rows"].append([])
            d = dict(attrs)
            node["rows"][-1].append({"t": [], "cs": _int(d.get("colspan"), 1),
                                     "rs": _int(d.get("rowspan"), 1)})
        elif tag == "br" and node["rows"] and node["rows"][-1]:
            node["rows"][-1][-1]["t"].append(" ")

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in ("script", "style", "title", "noscript"):
            if self._skip:
                self._skip -= 1
            return
        if tag == "caption" and self._stack:
            self._stack[-1].pop("_cap", None)
            return
        if tag == "table" and self._stack:
            self._stack.pop()

    # -- 文本
    def handle_data(self, data):
        if self._skip:
            return
        if not self._stack:
            self._page_text.append(data)
            if len(self._page_text) > 40:
                self._page_text = self._page_text[-40:]
            return
        node = self._stack[-1]
        if node.get("_cap"):
            node["caption"] += data
            return
        if node["rows"] and node["rows"][-1]:
            node["rows"][-1][-1]["t"].append(data)


def _grid(node):
    """把 table 节点摊成矩形网格。返回 (grid, 有效列数)。

    合并单元格按 **Excel 语义**：值只放左上角，被覆盖的格子留空。
    绝对不能把值复制到每个格子 —— rowspan=2 的「金额 100」复制成两行就是 200，
    合计直接翻倍；而表头合并用留空是安全的，平台的 _header_names() 本来就会左填充。
    """
    occ, maxc, r = {}, 0, 0
    for row in node["rows"]:
        c = 0
        for cell in row:
            while (r, c) in occ:
                c += 1
            cs = max(1, min(int(cell.get("cs") or 1), MAX_COLS))
            rs = max(1, int(cell.get("rs") or 1))
            occ[(r, c)] = _clean("".join(cell.get("t") or []))
            for dr in range(rs):
                for dc in range(cs):
                    if dr == 0 and dc == 0:
                        continue
                    occ.setdefault((r + dr, c + dc), "")
            maxc = max(maxc, c + cs)
            c += cs
        r += 1
    if not occ:
        return [], 0
    maxr = max(k[0] for k in occ) + 1
    maxc = max(maxc, max(k[1] for k in occ) + 1)
    grid = [[occ.get((i, j), "") for j in range(maxc)] for i in range(maxr)]
    while maxc and all(not (row[maxc - 1] or "").strip() for row in grid):
        for row in grid:
            row.pop()
        maxc -= 1
    while grid and all(not (v or "").strip() for v in grid[-1]):
        grid.pop()
    if not grid or not maxc:
        return [], 0
    return grid, maxc


def _table_name(node, idx):
    nm = _clean(node.get("caption") or "") or _clean(node.get("prev_text") or "")
    nm = nm.strip(" ：:—-·\t")
    nm = _BAD_NAME.sub("_", nm)[:31].strip()
    return nm or ("表%d" % idx)


def _read_bytes(src):
    if isinstance(src, (bytes, bytearray)):
        return bytes(src)
    if hasattr(src, "read"):
        return src.read()
    with open(src, "rb") as fh:
        return fh.read()


def read_tables(src):
    """解析网页，返回 [{"index","name","grid","rows","cols","note"}]（已过滤空表/壳表）。

    过滤顺序固定（保证同一文件重复解析得到同样的表名与顺序，缓存才能复用）：
    全空 → 行<2 或 列<2 → 只有 1 个非空格且很短（导航/间距壳）。
    """
    raw = _read_bytes(src)
    if not raw:
        raise HtmlError("文件是空的（0 字节），没有可导入的内容")
    if len(raw) > MAX_BYTES:
        raise HtmlError("网页文件过大（%.1f MB），请用 Excel 打开后另存为 .xlsx 再导入"
                        % (len(raw) / 1048576.0))
    text, enc = decode_text(raw)
    p = _TableParser()
    try:
        p.feed(text)
        p.close()
    except Exception as e:                      # 解析器本身的异常也不外抛栈信息
        raise HtmlError("网页内容解析失败（%s），请用 Excel 打开后另存为 .xlsx 再导入"
                        % str(e)[:60])
    if not p.tables:
        raise HtmlError("这个网页里没有找到表格（<table>）。若是网页上动态加载的数据，"
                        "请在浏览器里「另存为 → 网页，全部」，或用 Excel 打开后另存为 .xlsx")

    out, notes = [], []
    for node in p.tables:
        grid, cols = _grid(node)
        if not grid:
            continue
        rows = len(grid)
        if rows < 2 or cols < 2:
            continue
        filled = [(v or "").strip() for row in grid for v in row]
        nonempty = [v for v in filled if v]
        if len(nonempty) <= 1 and (not nonempty or len(nonempty[0]) <= 4):
            continue                            # 布局外壳/单格残壳
        if rows * cols > MAX_CELLS:
            raise HtmlError("表格太大（%s 个单元格），请用 Excel 打开后另存为 .xlsx 再导入"
                            % format(rows * cols, ","))
        out.append({"index": len(out), "name": _table_name(node, len(out) + 1),
                    "grid": grid, "rows": rows, "cols": cols, "note": ""})
        if len(out) >= MAX_TABLES:
            notes.append("只转换了前 %d 张表" % MAX_TABLES)
            break

    if not out:
        raise HtmlError("网页里有 %d 个表格，但都是空表或布局框，没有可导入的数据"
                        % len(p.tables))
    if "uncertain" in enc:
        notes.append("编码识别不可靠（%s），若有乱码请另存为 UTF-8 后重试" % enc)

    # 表名去重（Excel 工作表名必须唯一）
    used = {}
    for t in out:
        base = t["name"]
        used[base] = used.get(base, 0) + 1
        if used[base] > 1:
            t["name"] = (base[:28] + "_%d" % used[base])
    if notes:
        out[0]["note"] = "；".join(notes)
    return out


def to_xlsx(src, out_path):
    """解析并写成真正的 xlsx，返回 [{"sheet","rows","cols","note"}]。

    必须用普通 Workbook：write_only 模式不写 <dimension>，
    xlsxinfo.list_sheets 会退化成扫描首 1MB，宽表就会少列。
    先写 .part 再 os.replace：半成品 xlsx 会被 list_sheets 的
    (path, mtime, size) 缓存成「0 张表」，比报错更难查。
    """
    if Workbook is None:
        raise HtmlError("Excel 写入组件不可用，无法把网页表格转成 xlsx")
    tables = read_tables(src)
    wb = Workbook()
    wb.remove(wb.active)
    meta = []
    for t in tables:
        ws = wb.create_sheet(t["name"])
        for i, row in enumerate(t["grid"], 1):
            for j, v in enumerate(row, 1):
                if not v:
                    continue
                cell = ws.cell(row=i, column=j, value=v)
                if v[:1] == "=":
                    # 以 = 开头的文本（如 "=1+1"）会被 openpyxl 当公式，导入后变成结果或 0
                    cell.data_type = "s"
        meta.append({"sheet": t["name"], "rows": t["rows"], "cols": t["cols"],
                     "note": t.get("note") or ""})
    tmp = out_path + ".part"
    try:
        wb.save(tmp)
        os.replace(tmp, out_path)
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
    return meta
