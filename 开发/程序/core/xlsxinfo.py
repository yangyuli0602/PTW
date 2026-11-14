# -*- coding: utf-8 -*-
"""不依赖第三方库，直接从 xlsx 的 zip 结构里读出工作表清单和维度。

比 openpyxl 打开快几个数量级（只读 XML 头部几 KB），用于：
1. 导入前展示"有哪些 sheet、大概多少行"
2. 给 duckdb read_xlsx 一个明确的 range（否则它按首行猜列数，遇到合并大标题会只识别 1 列）
"""
import html
import io
import re
import zipfile

_DIM = re.compile(rb'<dimension[^>]*ref="([A-Z]+[0-9]+)(?::([A-Z]+[0-9]+))?"')
_NS = re.compile(rb'xmlns="[^"]*"')

# ---- 轻量预览用的正则（只解析 sheet XML 开头一小段，不解压整表）----
_ROW = re.compile(rb"<row\b([^>]*?)(?:/>|>(.*?)</row>)", re.S)
_ROWNUM = re.compile(rb'\br="(\d+)"')
_CELL = re.compile(rb"<c\b([^>]*?)(?:/>|>(.*?)</c>)", re.S)
_CELLREF = re.compile(rb'\br="([A-Z]+)\d+"')
_CELLT = re.compile(rb'\bt="([^"]+)"')
_VAL = re.compile(rb"<v[^>]*>(.*?)</v>", re.S)
_IST = re.compile(rb"<is>.*?<t[^>]*>(.*?)</t>", re.S)
_SI_T = re.compile(rb"<t[^>]*>(.*?)</t>", re.S)
_RPH = re.compile(rb"<rPh\b.*?</rPh>", re.S)
_CELL_COL = re.compile(rb'<c\b[^>]*?\br="([A-Z]{1,3})')


def col_to_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def num_to_col(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _cells(ref):
    m = re.match(r"([A-Z]+)([0-9]+)", ref or "")
    if not m:
        return 0, 0
    return col_to_num(m.group(1)), int(m.group(2))


def _rels(z):
    """rId -> 目标路径"""
    try:
        raw = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
    except KeyError:
        return {}
    out = {}
    for m in re.finditer(r"<Relationship\b[^>]*>", raw):
        tag = m.group(0)
        rid = re.search(r'Id="([^"]+)"', tag)
        tgt = re.search(r'Target="([^"]+)"', tag)
        if rid and tgt:
            t = tgt.group(1).lstrip("/")
            if not t.startswith("xl/"):
                t = "xl/" + t
            out[rid.group(1)] = t
    return out


def list_sheets(path):
    """返回 [{name, index, path, rows, cols, dim_ok, bytes, broken, reason}]

    打不开时抛 ValueError（人话）—— 别把 BadZipFile 的英文原文漏到界面，
    那会被前端当成"文件损坏"，而真实原因可能是加密、是 HTML、是压缩包。
    """
    out = []
    try:
        zf = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        raise ValueError("不是有效的 Excel 工作簿：文件不是 zip 结构"
                         "（可能已损坏，或后缀与实际格式不符）")
    except OSError as e:
        raise ValueError("读不到文件：%s" % e)
    with zf as z:
        try:
            wb = z.read("xl/workbook.xml").decode("utf-8", "ignore")
        except KeyError:
            return []
        rels = _rels(z)
        names = z.namelist()
        for i, m in enumerate(re.finditer(r"<sheet\b[^>]*/?>", wb)):
            tag = m.group(0)
            nm = re.search(r'name="([^"]*)"', tag)
            rid = re.search(r'r:id="([^"]+)"', tag)
            name = _unxml(nm.group(1)) if nm else "Sheet%d" % (i + 1)
            sp = rels.get(rid.group(1)) if rid else None
            if not sp or sp not in names:
                sp = None
                cand = "xl/worksheets/sheet%d.xml" % (i + 1)
                if cand in names:
                    sp = cand
            rows, cols, dim_ok, nbytes = 0, 0, False, 0
            if sp:
                try:
                    nbytes = z.getinfo(sp).file_size
                except Exception:
                    nbytes = 0
                try:
                    # 只读开头几 KB 找 <dimension>：原来 z.read(sp)[:4000] 会先把
                    # 整张表解压出来再截断（真实序时账 47MB，白花 0.1s+）
                    with z.open(sp) as fh:
                        head = fh.read(8192)
                    dm = _DIM.search(head)
                    if dm:
                        c1, r1 = _cells(dm.group(1).decode())
                        c2, r2 = _cells((dm.group(2) or dm.group(1)).decode())
                        rows, cols = max(r1, r2), max(c1, c2)
                        dim_ok = rows > 1 and cols > 1
                except Exception:
                    pass
            if not dim_ok and sp:
                rows, cols = _scan(z, sp)
            # 有内容但读不出任何单元格 = 多半 sheet XML 坏了（导出中断、被工具改坏）。
            # 不能当成"空表"：真空表的 sheet XML 一般只有几百字节，用户在
            # 「N 张空表已自动跳过」里完全看不出有一张表把数据丢了。
            broken = bool(sp) and (rows <= 0 or cols <= 0) and nbytes >= 2048
            out.append({"name": name, "index": i, "path": sp,
                        "rows": rows, "cols": cols, "dim_ok": dim_ok,
                        "bytes": nbytes, "broken": broken,
                        "reason": ("工作表约 %d KB，但解析不出任何单元格，疑似损坏"
                                   % max(1, nbytes // 1024)) if broken else ""})
    return out


def _scan(z, sp, col_probe_bytes=1 << 20):
    """没有 dimension 时，粗扫 sheet xml：数 <row 得到行数，扫单元格引用得到最大列。

    行数用 bytes.count（C 实现，很快）全量统计；列数只探开头 col_probe_bytes
    （默认 1MB）。理由：没有 dimension 的文件基本都是程序导出，列结构规整，
    表头区就定得出最大列；而逐块对每个 <c> 做正则在 10 万行表上是 250 万次
    匹配，实测白花 2 秒。
    """
    rows = 0
    maxc = 0
    probed = 0
    tail = b""
    with z.open(sp) as f:
        while True:
            chunk = f.read(1 << 20)
            if not chunk:
                break
            rows += chunk.count(b"<row")
            if probed < col_probe_bytes:
                # 拼上上一块尾巴，避免 <c r="AB 这类标签被块边界切断而漏列
                for m in _CELL_COL.finditer(tail + chunk):
                    c = col_to_num(m.group(1).decode())
                    if c > maxc:
                        maxc = c
                probed += len(chunk)
                tail = chunk[-32:]
    return rows, maxc


def _si_text(seg):
    """从一个 <si>…</si> 片段取出可见文本（富文本多段拼接，剔除拼音注音）。"""
    seg = _RPH.sub(b"", seg)
    raw = b"".join(_SI_T.findall(seg))
    return html.unescape(raw.decode("utf-8", "ignore"))


def _sst_prefix(z, max_idx):
    """只解析 sharedStrings 的前 max_idx+1 条，够用即停（大表 sst 有好几 MB）。"""
    if max_idx < 0:
        return []
    try:
        f = z.open("xl/sharedStrings.xml")
    except KeyError:
        return []
    out = []
    buf = b""
    with f:
        while len(out) <= max_idx:
            chunk = f.read(1 << 19)
            if not chunk:
                break
            buf += chunk
            segs = buf.split(b"</si>")
            buf = segs.pop()          # 末段可能被块边界截断，留到下一轮
            for seg in segs:
                out.append(_si_text(seg))
                if len(out) > max_idx:
                    break
    return out


def preview_cells(path, sheet_xml, max_rows=60, max_bytes=24 << 20):
    """直读 sheet XML 开头，返回 {行号: {列字母(bytes): 文本}}（最多 max_rows 行）。

    只解压到够 max_rows 行就停，不碰整张表 —— 真实序时账的 sheet1.xml 解压后
    有 47MB，而前 60 行只占几十 KB。

    返回的文本与 duckdb `read_xlsx(..., all_varchar=true)` 保持一致：日期是 Excel
    序列号原值（如 46190），空单元格不出现，富文本取合并后的可见文字。
    行号取自 <row r="N">，因此中间有空行也不会错位。
    """
    if not sheet_xml:
        return {}
    with zipfile.ZipFile(path) as z:
        try:
            f = z.open(sheet_xml)
        except KeyError:
            return {}
        buf = b""
        with f:
            while buf.count(b"</row>") < max_rows + 1:
                chunk = f.read(256 << 10)
                if not chunk:
                    break
                buf += chunk
                if len(buf) > max_bytes:
                    break
        rows = {}
        max_sst = -1
        seq = 0
        for m in _ROW.finditer(buf):
            if len(rows) > max_rows:
                break
            attrs, body = m.group(1), m.group(2) or b""
            rm = _ROWNUM.search(attrs)
            rn = int(rm.group(1)) if rm else seq + 1
            seq = rn
            cells = {}
            for cm in _CELL.finditer(body):
                cattrs, cbody = cm.group(1), cm.group(2) or b""
                ref = _CELLREF.search(cattrs)
                if not ref:
                    continue
                col = ref.group(1)
                tm = _CELLT.search(cattrs)
                typ = tm.group(1).decode() if tm else "n"
                if typ == "s":
                    vm = _VAL.search(cbody)
                    if not vm:
                        cells[col] = ""
                        continue
                    idx = int(vm.group(1))
                    if idx > max_sst:
                        max_sst = idx
                    cells[col] = idx          # 先用索引占位，下面回填
                elif typ == "inlineStr":
                    im = _IST.search(cbody)
                    cells[col] = html.unescape(
                        (im.group(1) if im else b"").decode("utf-8", "ignore"))
                else:
                    vm = _VAL.search(cbody)
                    cells[col] = html.unescape(
                        vm.group(1).decode("utf-8", "ignore")) if vm else ""
            rows[rn] = cells
        sst = _sst_prefix(z, max_sst) if max_sst >= 0 else []
    for cells in rows.values():
        for col, v in cells.items():
            if isinstance(v, int):            # sharedStrings 索引
                cells[col] = sst[v] if 0 <= v < len(sst) else ""
    return rows


def _unxml(s):
    return (s.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
             .replace("&quot;", '"').replace("&apos;", "'"))


def sniff_csv(path, sample=1 << 20):
    """探测分隔符与编码。返回 (delim, encoding)"""
    raw = open(path, "rb").read(sample)
    enc = "utf-8"
    if raw.startswith(b"\xef\xbb\xbf"):
        enc = "utf-8-sig"
        raw = raw[3:]
    else:
        try:
            raw.decode("utf-8")
        except UnicodeDecodeError:
            for cand in ("gb18030", "gbk", "big5"):
                try:
                    raw.decode(cand)
                    enc = cand
                    break
                except UnicodeDecodeError:
                    continue
    head = raw.split(b"\n", 5)[:5]
    best, bestn = ",", 0
    for d in (",", "\t", ";", "|"):
        n = sum(line.count(d.encode()) for line in head)
        if n > bestn:
            best, bestn = d, n
    return best, enc
