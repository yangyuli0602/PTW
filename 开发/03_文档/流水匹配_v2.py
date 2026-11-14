# -*- coding: utf-8 -*-
"""网银流水 x 账面账户清单 匹配工具  v2.2

用法（Windows，任意 Python 3.8+，只需 openpyxl）：

    python 流水匹配_v2.py --acct 账面账户清单.xlsx --flow 流水.zip
    python 流水匹配_v2.py --acct 账面账户清单.xlsx --flow "D:\\资料\\13、网银流水"
    python 流水匹配_v2.py --acct 清单.xlsx --flow 流水.zip --out D:\\结果.xlsx

输出：同目录（或 --out 指定）一个 xlsx，含 7 张表：
    说明 / 待人工确认 / 匹配明细 / 账户全覆盖 / 有账无流水 / 有流水无账 / 重复文件
同时在控制台打印「待人工确认」清单和原因。

支持格式（按文件头判断，不信扩展名）：
    xlsx(zip) / xls(BIFF8，内置读取器) / xls实为xlsx / xls实为HTML / pdf

设计约束：
    1. 每一行匹配都必须给出证据位置（文件名 / Sheet!单元格 + 表头原文 + 命中原始值）
    2. 拿不到「本方账号」就进待人工确认，绝不用「最早出现的账号」「出现次数最多」去猜
    3. 单文件失败隔离，不中断整批

本脚本不依赖平台代码，可单独拷走运行。
"""

from __future__ import annotations

import argparse
import datetime as _dt
import hashlib
import io
import os
import re
import sys
import traceback
import zipfile
import zlib
from html.parser import HTMLParser

try:
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except Exception:  # noqa: BLE001
    sys.stderr.write("缺少 openpyxl，请先 pip install openpyxl\n")
    raise SystemExit(2)

VERSION = "2.2"

# ==========================================================================
#  内嵌 .xls（BIFF8）读取器 —— 摘自 PTW 平台 core/xlsread.py，纯标准库
# ==========================================================================
"""读取老式 .xls（Excel 97-2003 / BIFF8）—— 纯 Python，零第三方依赖。

为什么自己写：
- PTW平台运行时只有 openpyxl（只认 xlsx）和 duckdb 的 excel 扩展（只认 xlsx 的
  zip 结构，遇到 .xls 直接报 "No xl/workbook.xml found"）；
- 现场底稿/序时账大量是 .xls，装不了 xlrd、也不该依赖本机 Office 转换。

支持：BIFF8（Excel 97/2000/XP/2003）。不支持：加密文件、BIFF5 及更早（Excel 95-）。
输出：read_xls(bytes) -> [{"name": 表名, "rows": [[值, ...], ...]}]
      值保留类型：str / float / datetime / time / bool / None
"""
import re
import struct
from datetime import datetime, date, time, timedelta

__all__ = ["XlsError", "is_xls", "read_xls", "xls_to_xlsx"]

MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
MAXREGSECT = 0xFFFFFFFA


class XlsError(Exception):
    pass


def is_xls(data):
    """靠文件头判断（不能只看扩展名，现场很多 .xls 其实是改了后缀的 xlsx）。"""
    return isinstance(data, (bytes, bytearray)) and bytes(data[:8]) == MAGIC


def _u8(b, o):
    return b[o]


def _u16(b, o):
    return b[o] | (b[o + 1] << 8)


def _u32(b, o):
    return b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)


# ============================================================ OLE2 复合文档
class _Ole:
    """最小实现的 OLE2 / Compound File Binary 读取器（只读流）。"""

    def __init__(self, data):
        if len(data) < 512 or bytes(data[:8]) != MAGIC:
            raise XlsError("不是有效的 .xls 文件（文件头不是 OLE2）")
        self.d = data
        self.ssz = 1 << _u16(data, 30)
        self.mssz = 1 << _u16(data, 32)
        if self.ssz < 128 or self.ssz > 65536:
            raise XlsError("OLE2 扇区大小异常")

        n_fat = _u32(data, 44)
        self.dir_start = _u32(data, 48)
        self.mini_cutoff = _u32(data, 56)
        mini_fat_start = _u32(data, 60)
        difat_start = _u32(data, 68)
        n_difat = _u32(data, 72)

        # ---- DIFAT：头 109 项 + 后续 DIFAT 扇区链
        difat = [_u32(data, 76 + 4 * i) for i in range(109)]
        sec, guard = difat_start, 0
        while sec not in (FREESECT, ENDOFCHAIN) and sec <= MAXREGSECT and guard < 10000:
            blk = self._sector(sec)
            per = self.ssz // 4 - 1
            difat.extend(_u32(blk, 4 * i) for i in range(per))
            sec = _u32(blk, per * 4)
            guard += 1
        fat_secs = [s for s in difat if s <= MAXREGSECT][:n_fat] if n_fat else \
            [s for s in difat if s <= MAXREGSECT]

        # ---- FAT
        self.fat = []
        for s in fat_secs:
            blk = self._sector(s)
            self.fat.extend(_u32(blk, 4 * i) for i in range(self.ssz // 4))

        # ---- MiniFAT
        self.minifat = []
        for s in self._chain(mini_fat_start):
            blk = self._sector(s)
            self.minifat.extend(_u32(blk, 4 * i) for i in range(self.ssz // 4))

        # ---- 目录项
        self.entries = []
        raw = b"".join(self._sector(s) for s in self._chain(self.dir_start))
        for off in range(0, len(raw) - 127, 128):
            e = raw[off:off + 128]
            nlen = _u16(e, 64)
            name = ""
            if 0 < nlen <= 64:
                name = e[:nlen - 2].decode("utf-16-le", "replace")
            self.entries.append({
                "name": name, "type": _u8(e, 66),
                "start": _u32(e, 116), "size": _u32(e, 120),
            })

    def _sector(self, i):
        o = 512 + i * self.ssz
        if o + self.ssz > len(self.d):
            return self.d[o:] + b"\x00" * (o + self.ssz - len(self.d))
        return self.d[o:o + self.ssz]

    def _chain(self, start):
        out, cur, seen = [], start, set()
        while cur <= MAXREGSECT and cur < len(self.fat) and cur not in seen:
            seen.add(cur)
            out.append(cur)
            cur = self.fat[cur]
        return out

    def stream(self, *names):
        """按名字取流（大小写不敏感）。"""
        want = {n.lower() for n in names}
        ent = None
        for e in self.entries:
            if e["type"] == 2 and e["name"].lower() in want:
                ent = e
                break
        if not ent:
            return None
        size, start = ent["size"], ent["start"]
        if size < self.mini_cutoff:
            # 小流放在 mini stream 里（root entry 的流），按 MiniFAT 走
            root = next((e for e in self.entries if e["type"] == 5), None)
            if not root:
                raise XlsError("OLE2 缺少根目录项")
            mini = b"".join(self._sector(s) for s in self._chain(root["start"]))
            out, cur, seen = [], start, set()
            while cur <= MAXREGSECT and cur < len(self.minifat) and cur not in seen:
                seen.add(cur)
                o = cur * self.mssz
                out.append(mini[o:o + self.mssz])
                cur = self.minifat[cur]
            return b"".join(out)[:size]
        data = b"".join(self._sector(s) for s in self._chain(start))
        return data[:size]


# ============================================================ BIFF 基础
def _records(buf, pos=0):
    """把二进制流切成 [(type, body, offset)]。"""
    out = []
    n = len(buf)
    while pos + 4 <= n:
        t = _u16(buf, pos)
        ln = _u16(buf, pos + 2)
        out.append((t, buf[pos + 4:pos + 4 + ln], pos))
        pos += 4 + ln
    return out


class _SegReader:
    """SST / CONTINUE 跨段读取器。

    BIFF8 规定：字符串被 CONTINUE 截断时，新段首字节是该段后续字符的宽度标志。
    只有"正在读字符串"时才会消费这个标志字节，读记录头时不会。
    """

    def __init__(self, segs):
        self.segs = segs
        self.i = 0
        self.pos = 0
        self.wide = False

    def _avail(self, mid_string=False):
        """切到下一段。mid_string=True 时（正在读字符串中间）要跳过 1 字节宽度标志。"""
        while self.i < len(self.segs) and self.pos >= len(self.segs[self.i]):
            self.i += 1
            self.pos = 0
            if self.i < len(self.segs):
                seg = self.segs[self.i]
                if mid_string:
                    self.wide = bool(seg and (seg[0] & 1))
                    self.pos = 1
        return self.i < len(self.segs)

    def u8(self):
        if not self._avail():
            return 0
        v = self.segs[self.i][self.pos]
        self.pos += 1
        return v

    def u16(self):
        return self.u8() | (self.u8() << 8)

    def u32(self):
        return self.u16() | (self.u16() << 16)

    def skip(self, n):
        while n > 0 and self._avail():
            take = min(n, len(self.segs[self.i]) - self.pos)
            self.pos += take
            n -= take

    def chars(self, count):
        """按累计宽度读 count 个字符。"""
        out = []
        while count > 0:
            if not self._avail(mid_string=True):
                break
            seg = self.segs[self.i]
            remain = len(seg) - self.pos
            if self.wide:
                take = min(count, remain // 2)
                if take <= 0:            # 只剩半个字符：跳到下一段
                    self.pos = len(seg)
                    continue
                out.append(seg[self.pos:self.pos + take * 2].decode("utf-16-le", "replace"))
                self.pos += take * 2
            else:
                take = min(count, remain)
                out.append(seg[self.pos:self.pos + take].decode("latin-1"))
                self.pos += take
            count -= take
        return "".join(out)


def _xlstr(body, off, len_bytes=2):
    """记录内的 XLUnicodeString（不会跨记录）。"""
    if len_bytes == 2:
        cch = _u16(body, off)
        flags = body[off + 2]
        p = off + 3
    else:
        cch = body[off]
        flags = body[off + 1]
        p = off + 2
    if p + cch * (2 if flags & 1 else 1) > len(body):
        return ""
    if flags & 1:
        return body[p:p + cch * 2].decode("utf-16-le", "replace")
    return body[p:p + cch].decode("latin-1")


def _sst(segs):
    rd = _SegReader(segs)
    rd.u32()                      # 总引用数
    unique = rd.u32()             # 唯一字符串数
    out = []
    for _ in range(unique):
        cch = rd.u16()
        flags = rd.u8()
        rd.wide = bool(flags & 0x01)      # 本字符串的字符宽度
        n_run = rd.u16() if flags & 0x08 else 0
        ext = rd.u32() if flags & 0x04 else 0
        out.append(rd.chars(cch))
        if n_run:
            rd.skip(4 * n_run)
        if ext:
            rd.skip(ext)
    return out


def _rk(v):
    if v & 2:
        i = v >> 2
        if v & 0x80000000:        # 符号扩展（30 位有符号整数）
            i -= 0x40000000
        val = float(i)
    else:
        val = struct.unpack("<d", b"\x00\x00\x00\x00" + struct.pack("<I", v & 0xFFFFFFFC))[0]
    if v & 1:
        val /= 100.0
    return val


# 内置日期/时间格式号（BIFF8）
_DATE_FMT_IDS = set(range(14, 23)) | {27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
                                     45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58}


def _is_date_fmt(code, fmt_str):
    if code in _DATE_FMT_IDS:
        return True
    if not fmt_str:
        return False
    s = fmt_str.lower()
    s = re.sub(r"\[[^\]]*\]", "", s)       # 去掉 [Red] [$-409] 之类
    s = re.sub(r'"[^"]*"', "", s)          # 去掉引号里的字面量
    s = re.sub(r"\\.", "", s)
    if s.strip() in ("general", "@", ""):
        return False
    return bool(re.search(r"[ymdhs]", s))


# ============================================================ BIFF8 解析
def read_xls(data):
    """把 .xls 读成 [{"name": 表名, "rows": [[值, ...], ...]}]。"""
    ole = _Ole(data)
    buf = ole.stream("Workbook", "Book")
    if buf is None:
        raise XlsError("文件里没有 Workbook 流，不是标准 Excel 文件（可能已损坏）")
    if len(buf) < 4:
        raise XlsError("Workbook 流为空")
    recs = _records(buf)
    if not recs:
        raise XlsError("Workbook 流里没有 BIFF 记录")

    # ---- 第一遍：工作簿全局区
    bsheets, sst, daterel = [], [], 0
    xf_fmt, fmts = [], {}
    for idx, (t, body, off) in enumerate(recs):
        if t == 0x000A:                     # EOF：全局区结束
            break
        if t == 0x0085:                     # BOUNDSHEET
            pos = _u32(body, 0)
            vis = body[4]
            typ = body[5]
            if typ == 0:                    # 0 = 工作表（跳过图表页/宏表）
                bsheets.append({"pos": pos, "name": _xlstr(body, 6, 1), "hidden": vis})
        elif t == 0x00FC:                   # SST
            segs = [body]
            j = idx + 1
            while j < len(recs) and recs[j][0] == 0x003C:   # CONTINUE
                segs.append(recs[j][1])
                j += 1
            try:
                sst = _sst(segs)
            except Exception:
                sst = []
        elif t == 0x00E0:                   # XF
            xf_fmt.append(_u16(body, 2))
        elif t == 0x041E:                   # FORMAT
            fmts[_u16(body, 0)] = _xlstr(body, 2)
        elif t == 0x0022:                   # DATEMODE
            daterel = _u16(body, 0)
        elif t == 0x002F:                   # FILEPASS：加密
            raise XlsError("该 .xls 已加密，请先在 Excel 里另存为不加密的 xlsx")

    if not bsheets:
        raise XlsError("没有解析到任何工作表：文件可能已损坏，或不是 Excel 97-2003 格式")

    date_by_xf = []
    for code in xf_fmt:
        date_by_xf.append(_is_date_fmt(code, fmts.get(code)))

    # ---- 第二遍：逐表解析
    out = []
    for bs in bsheets:
        recs_sheet = [(t, b, o) for (t, b, o) in recs if o >= bs["pos"]]
        out.append({"name": bs["name"] or "Sheet",
                    "rows": _sheet_rows(recs_sheet, sst, date_by_xf, daterel)})
    return out


def _sheet_rows(recs, sst, date_by_xf, daterel):
    base = datetime(1904, 1, 1) if daterel else datetime(1899, 12, 30)
    cells = {}
    max_r = max_c = -1
    pending = None            # FORMULA 的字符串结果在后面的 STRING 记录里

    def put(r, c, v):
        nonlocal max_r, max_c
        cells[(r, c)] = v
        if r > max_r:
            max_r = r
        if c > max_c:
            max_c = c

    def as_date(v, xf):
        if not (0 <= xf < len(date_by_xf)) or not date_by_xf[xf]:
            return v
        if not isinstance(v, float):
            return v
        if v <= 0 or v > 2958465:
            return v
        dt = base + timedelta(days=v)
        if v < 1:                          # 纯时间（如 13:45）
            return dt.time()
        return dt

    for t, body, _off in recs:
        if t == 0x000A and cells:          # EOF
            break
        if t == 0x00FD:                    # LABELSST
            r, c, xf, i = _u16(body, 0), _u16(body, 2), _u16(body, 4), _u32(body, 6)
            put(r, c, sst[i] if 0 <= i < len(sst) else "")
        elif t == 0x0204:                  # LABEL
            r, c, xf = _u16(body, 0), _u16(body, 2), _u16(body, 4)
            put(r, c, _xlstr(body, 6))
        elif t == 0x0203:                  # NUMBER
            r, c, xf = _u16(body, 0), _u16(body, 2), _u16(body, 4)
            v = struct.unpack("<d", body[6:14])[0]
            put(r, c, as_date(v, xf))
        elif t == 0x027E:                  # RK
            r, c, xf = _u16(body, 0), _u16(body, 2), _u16(body, 4)
            put(r, c, as_date(_rk(_u32(body, 6)), xf))
        elif t == 0x00BD:                  # MULRK
            r = _u16(body, 0)
            c1 = _u16(body, 2)
            c2 = _u16(body, len(body) - 2)
            p = 4
            for c in range(c1, c2 + 1):
                xf = _u16(body, p)
                put(r, c, as_date(_rk(_u32(body, p + 2)), xf))
                p += 6
        elif t == 0x0006:                  # FORMULA
            r, c, xf = _u16(body, 0), _u16(body, 2), _u16(body, 4)
            res = body[6:14]
            if len(res) == 8 and res[6:8] == b"\xff\xff":
                sub = res[0:2]
                if sub == b"\x00\x00":
                    pending = (r, c)        # 字符串稍后到
                    continue
                put(r, c, bool(res[2]) if sub == b"\x01\x00" else None)
            else:
                put(r, c, as_date(struct.unpack("<d", res)[0], xf))
        elif t == 0x0207:                  # STRING（公式的字符串结果）
            if pending:
                put(pending[0], pending[1], _xlstr(body, 0))
                pending = None
        elif t in (0x0201, 0x00BE):        # BLANK / MULBLANK
            continue

    if max_r < 0:
        return []
    ncols = max_c + 1
    rows = []
    for r in range(max_r + 1):
        rows.append([cells.get((r, c)) for c in range(ncols)])
    # 去掉尾部全空行
    while rows and not any(v not in (None, "") for v in rows[-1]):
        rows.pop()
    return rows


# ============================================================ 转成 xlsx
_BAD = re.compile(r"[\[\]:*?/\\]")


def xls_to_xlsx(data, out_path):
    """读 .xls 后写成 .xlsx，让后续流程完全复用现有 xlsx 通道。"""
    from openpyxl import Workbook

    sheets = read_xls(data)
    wb = Workbook()
    wb.remove(wb.active)
    used = set()
    for sh in sheets:
        nm = _BAD.sub("_", str(sh["name"] or "Sheet"))[:31] or "Sheet"
        base, k = nm, 2
        while nm in used:
            suffix = "_%d" % k
            nm = (base[:31 - len(suffix)]) + suffix
            k += 1
        used.add(nm)
        ws = wb.create_sheet(nm)
        for r, row in enumerate(sh["rows"], 1):
            for c, v in enumerate(row, 1):
                if v is None or v == "":
                    continue
                cell = ws.cell(r, c, v)
                if isinstance(v, datetime):
                    cell.number_format = "yyyy-mm-dd"
                elif isinstance(v, time):
                    cell.number_format = "hh:mm:ss"
                elif isinstance(v, date):
                    cell.number_format = "yyyy-mm-dd"
    wb.save(out_path)
    return out_path
# ==========================================================================
#  常量与正则
# ==========================================================================

MAGIC_OLE2 = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

LABEL_ACCT = re.compile(
    r"本方账[号號]|本方帐[号號]|本方账户|本方帐号|选择账号|交易账号|查询账号"
    r"|账户号码|银行账[号號]|账[号號]|帐[号號]|卡号|[Aa]ccount\s*(?:No|Number)",
    re.I,
)
LABEL_EXCLUDE = re.compile(r"对方|收方|付方|收款|付款|对手|对方账户")
LABEL_COMPANY = re.compile(r"公司名称|单位名称|客户名称|账户名称|户名|主体|公司简称")
LABEL_BANK = re.compile(r"开户行|开户银行|银行名称|存款机构|银行|机构")
LABEL_CUR = re.compile(r"币种|货币")
LABEL_BAL = re.compile(r"余额|期末")

SKIP_NAME = re.compile(r"^~\$|^\.|^__MACOSX", re.I)
SKIP_DIR = re.compile(r"\.files$|^__MACOSX$|^\.", re.I)

HEAD_FILL = PatternFill("solid", fgColor="1F4E79")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
BAD_FILL = PatternFill("solid", fgColor="FCE4E4")
OK_FILL = PatternFill("solid", fgColor="E2EFDA")


# ==========================================================================
#  通用小工具
# ==========================================================================

def to_half(s):
    """全角转半角（数字、字母、常见标点）。"""
    out = []
    for ch in str(s):
        o = ord(ch)
        if o == 0x3000:
            out.append(" ")
        elif 0xFF01 <= o <= 0xFF5E:
            out.append(chr(o - 0xFEE0))
        else:
            out.append(ch)
    return "".join(out)


def digits_of(v):
    """把任意值归一成纯数字串；数值型单元格尤其要小心精度。"""
    if v is None:
        return ""
    if isinstance(v, bool):
        return ""
    if isinstance(v, float):
        if v != v:  # NaN
            return ""
        if abs(v) >= 1e15:
            # 超过 15 位有效数字，Excel 数值存储本身已经丢精度
            return re.sub(r"\D", "", format(v, ".0f"))
        return re.sub(r"\D", "", format(v, ".0f"))
    if isinstance(v, int):
        return str(v)
    t = to_half(v)
    return re.sub(r"\D", "", t)


def num_like(v):
    """判断一个单元格值是否像账号：8~32 位数字（允许其中夹 - 空格）。

    注意：必须排除日期（20260105 / 2026-01-05），否则会把交易日期当成账号。
    """
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (_dt.datetime, _dt.date, _dt.time)):
        return None
    if isinstance(v, (int, float)):
        d = digits_of(v)
        return d if 8 <= len(d) <= 32 else None
    t = to_half(v).strip()
    if not t or len(t) > 60:
        return None
    has_sep = bool(re.search(r"[-/.．年月日:/]", t))
    core = re.sub(r"[\s\-]", "", t)
    if not re.fullmatch(r"\d{8,32}", core):
        return None
    if has_sep and len(core) < 12:
        # 带分隔符且不足 12 位，基本是日期
        return None
    if len(core) == 8 and _looks_like_yyyymmdd(core):
        return None
    return core


def _looks_like_yyyymmdd(s):
    try:
        y, m, d = int(s[:4]), int(s[4:6]), int(s[6:])
    except ValueError:
        return False
    return 1990 <= y <= 2100 and 1 <= m <= 12 and 1 <= d <= 31


def fmt_bytes(n):
    for u in ("B", "KB", "MB", "GB"):
        if n < 1024 or u == "GB":
            return "%.1f%s" % (n, u)
        n /= 1024.0
    return "%.1fGB" % n


def sha1_of(b):
    return hashlib.sha1(b).hexdigest()


# ==========================================================================
#  输入 1：账面账户清单
# ==========================================================================

def load_grids_from_path(path):
    """读任意 excel（xlsx / xls），返回 [(sheet_name, rows), ...]，rows 为二维 list。"""
    with open(path, "rb") as fh:
        data = fh.read()
    fmt = sniff(data, path)
    if fmt in ("xls", "xls_zip"):
        return parse_grid(data, fmt)
    if fmt == "xlsx":
        return parse_grid(data, "xlsx")
    raise RuntimeError("账户清单格式不支持：%s（只吃 xlsx / xls）" % fmt)


def pick_header_row(rows, need_acct=True):
    """在前 30 行里找表头行，返回 (行号, {类别: 列号})。"""
    best = (-1, None, None)
    for ri, row in enumerate(rows[:30]):
        cmap = {}
        for ci, cell in enumerate(row[:60]):
            if cell is None:
                continue
            t = to_half(cell).strip()
            if not t or len(t) > 30:
                continue
            if "acct" not in cmap and LABEL_ACCT.search(t) and not LABEL_EXCLUDE.search(t):
                cmap["acct"] = ci
            if "company" not in cmap and LABEL_COMPANY.search(t):
                cmap["company"] = ci
            if "bank" not in cmap and LABEL_BANK.search(t):
                cmap["bank"] = ci
            if "cur" not in cmap and LABEL_CUR.search(t):
                cmap["cur"] = ci
            if "bal" not in cmap and LABEL_BAL.search(t):
                cmap["bal"] = ci
        score = len(cmap) + (3 if "acct" in cmap else 0)
        if need_acct and "acct" not in cmap:
            score -= 5
        if score > best[0]:
            best = (score, ri, cmap)
    return best[1], best[2] or {}


def read_accounts(path):
    """返回 (accounts, meta)。accounts 每项：company/acct/digits/bank/currency/balance/status"""
    accounts = []
    meta = {"file": os.path.basename(path), "sheet": "", "header_row": None, "rows": 0}
    for sheet_name, rows in load_grids_from_path(path):
        hri, cmap = pick_header_row(rows)
        if hri is None or "acct" not in cmap:
            continue
        cur_sheet = []
        for row in rows[hri + 1:]:
            if cmap["acct"] >= len(row):
                continue
            raw = row[cmap["acct"]]
            dg = num_like(raw) or digits_of(raw)
            if not dg or len(dg) < 6:
                continue
            acc = {
                "company": _cell_str(row, cmap.get("company")),
                "acct_raw": _cell_str(row, cmap.get("acct")),
                "digits": dg,
                "bank": _cell_str(row, cmap.get("bank")),
                "currency": _cell_str(row, cmap.get("cur")),
                "balance": _num(row, cmap.get("bal")),
            }
            acc["status"] = "零余额" if acc["balance"] == 0 else "非零余额"
            cur_sheet.append(acc)
        if cur_sheet:
            accounts.extend(cur_sheet)
            if not meta["sheet"]:
                meta["sheet"] = sheet_name
                meta["header_row"] = hri + 1
                meta["rows"] = len(cur_sheet)
            break
    if not accounts:
        raise RuntimeError(
            "在账户清单里没找到账号列。请确认：①表头含「账号/账户号码/本方账号」等字样；"
            "②表头在前 30 行内。"
        )
    return accounts, meta


def _cell_str(row, ci):
    if ci is None or ci >= len(row) or row[ci] is None:
        return ""
    v = row[ci]
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return to_half(v).strip()


def _num(row, ci):
    if ci is None or ci >= len(row):
        return None
    v = row[ci]
    if isinstance(v, (int, float)):
        return float(v)
    t = to_half(v).replace(",", "").replace("，", "")
    m = re.search(r"-?\d+(?:\.\d+)?", t)
    return float(m.group(0)) if m else None


# ==========================================================================
#  输入 2：流水文件枚举（目录 或 zip）
# ==========================================================================

def sniff(data, name=""):
    """按文件头判断真实格式，不信扩展名。"""
    head = data[:4096]
    if head[:8] == MAGIC_OLE2:
        return "xls"
    if head[:4] == b"PK\x03\x04":
        low = name.lower()
        if low.endswith((".docx", ".pptx")):
            return "other"
        return "xlsx"
    if head[:5] == b"%PDF-":
        return "pdf"
    if head[:2] == b"\xff\xfe" or head[:2] == b"\xfe\xff":
        return "other"
    low = head.lower()
    if low.lstrip()[:1] == b"<" or b"<table" in low or b"<html" in low or b"<td" in low:
        return "html"
    return "other"


def iter_flow(flow):
    """yield dict(rel, name, ext, size, read())。支持 zip 与目录。"""
    if os.path.isdir(flow):
        root = flow.rstrip("\\/")
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = [d for d in dirnames if not SKIP_DIR.search(d)]
            for fn in filenames:
                if SKIP_NAME.search(fn):
                    continue
                if fn.lower().endswith((".zip", ".rar", ".7z", ".tar", ".gz")):
                    # 文件夹里夹着的压缩包不当作流水（避免把包本身当文件解析）
                    continue
                full = os.path.join(dirpath, fn)
                rel = os.path.relpath(full, root).replace("\\", "/")
                try:
                    size = os.path.getsize(full)
                except OSError:
                    continue
                yield {
                    "rel": rel,
                    "name": fn,
                    "ext": os.path.splitext(fn)[1].lower(),
                    "size": size,
                    "source": "dir",
                    "read": (lambda p=full: open(p, "rb").read()),
                }
        return

    if zipfile.is_zipfile(flow):
        flow_abs = os.path.abspath(flow)
        with zipfile.ZipFile(flow) as zf:
            for info in zf.infolist():
                if info.is_dir():
                    continue
                name = _fix_zip_name(info)
                base = name.split("/")[-1]
                if not base or SKIP_NAME.search(base):
                    continue
                if any(SKIP_DIR.search(p) for p in name.split("/")[:-1]):
                    continue
                if any(SKIP_DIR.search(p) for p in [base]) and "." not in base:
                    continue
                if base.lower().endswith((".db", ".duckdb", ".wal", ".tmp")):
                    continue
                yield {
                    "rel": name,
                    "name": base,
                    "ext": os.path.splitext(base)[1].lower(),
                    "size": info.file_size,
                    "source": "zip",
                    "read": (lambda z=flow_abs, n=info.filename: _read_zip_member(z, n)),
                }
        return

    raise RuntimeError("--flow 既不是文件夹也不是 zip：%s" % flow)


def _read_zip_member(zpath, member):
    """按成员名重新打开 zip 读取（生成器退出后 ZipFile 已关闭，不能复用句柄）。"""
    with zipfile.ZipFile(zpath) as z:
        return z.read(member)


def _fix_zip_name(info):
    """Windows 打的 zip 常用 GBK 且不带 UTF-8 标志，会导致中文名乱码。"""
    name = info.filename
    if not (info.flag_bits & 0x800):
        try:
            name = name.encode("cp437").decode("gbk")
        except Exception:  # noqa: BLE001
            try:
                name = name.encode("cp437").decode("utf-8")
            except Exception:  # noqa: BLE001
                pass
    return name.replace("\\", "/")


# ==========================================================================
#  解析：把任意格式变成 [(sheet_name, rows), ...]
# ==========================================================================

class _TableParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables = []
        self._stack = []
        self._row = None
        self._cell = None
        self._buf = []

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag == "table":
            self._stack.append([])
        elif tag == "tr":
            self._row = []
        elif tag in ("td", "th"):
            d = dict(attrs)
            try:
                cs = max(1, int(d.get("colspan", "1")))
            except ValueError:
                cs = 1
            try:
                rs = max(1, int(d.get("rowspan", "1")))
            except ValueError:
                rs = 1
            self._cell = {"text": "", "cs": cs, "rs": rs}
            self._buf = []
        elif tag == "br" and self._cell is not None:
            self._buf.append(" ")

    def handle_data(self, data):
        if self._cell is not None:
            self._buf.append(data)

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in ("td", "th"):
            if self._cell is not None:
                self._cell["text"] = "".join(self._buf).strip()
                if self._row is not None:
                    self._row.append(self._cell)
                self._cell = None
                self._buf = []
        elif tag == "tr":
            if self._row is not None and self._stack:
                self._stack[-1].append(self._row)
            self._row = None
        elif tag == "table":
            if self._stack:
                self.tables.append(self._stack.pop())


def _html_to_grid(tbl):
    occ = {}
    for r, row in enumerate(tbl):
        c = 0
        for cell in row:
            while (r, c) in occ:
                c += 1
            for dr in range(cell["rs"]):
                for dc in range(cell["cs"]):
                    occ[(r + dr, c + dc)] = cell["text"]
            c += cell["cs"]
    if not occ:
        return []
    maxr = max(k[0] for k in occ)
    maxc = max(k[1] for k in occ)
    return [[occ.get((r, c), "") for c in range(maxc + 1)] for r in range(maxr + 1)]


def parse_grid(data, fmt, scan_rows=80, scan_cols=40, max_sheets=5):
    """返回 [(sheet_name, rows)]；rows 最多 scan_rows 行 × scan_cols 列。"""
    out = []
    if fmt == "xlsx":
        bio = io.BytesIO(data)
        wb = openpyxl.load_workbook(bio, read_only=True, data_only=True)
        try:
            for ws in wb.worksheets[:max_sheets]:
                rows = []
                try:
                    # 显式给 max_row/max_col，绕开错误 dimension 的截断问题
                    it = ws.iter_rows(max_row=scan_rows, max_col=scan_cols, values_only=True)
                except TypeError:
                    it = ws.iter_rows(max_row=scan_rows, max_col=scan_cols, values_only=True)
                for r in it:
                    rows.append(list(r))
                out.append((ws.title, rows))
        finally:
            wb.close()
        return out

    if fmt == "xls":
        for sh in read_xls(data):
            rows = [list(r)[:scan_cols] for r in sh["rows"][:scan_rows]]
            out.append((sh["name"], rows))
            if len(out) >= max_sheets:
                break
        if not out:
            raise RuntimeError("xls 未解析出任何工作表")
        return out

    if fmt == "html":
        p = _TableParser()
        p.feed(data.decode("utf-8", "ignore") if b"charset=\"utf-8\"" in data[:2048].lower()
               else data.decode("gbk", "ignore"))
        grids = [_html_to_grid(t) for t in p.tables]
        grids = [(i + 1, g) for i, g in enumerate(grids)]
        grids = [x for x in grids if x[1]]
        if not grids:
            raise RuntimeError("HTML 里没有可解析的表格")
        for i, g in grids[:max_sheets]:
            out.append(("表格%d" % i, [r[:scan_cols] for r in g[:scan_rows]]))
        return out

    raise RuntimeError("格式 %s 不支持结构化解析" % fmt)


# ==========================================================================
#  内容取证：在一张表里找「本方账号」
# ==========================================================================

def _col_letter(i):
    return get_column_letter(i + 1)


def _column_values(rows, ri, ci, span=400):
    """标签所在列、标签下方出现的所有不同账号值（保持出现顺序）。"""
    vals = []
    for rj in range(ri + 1, min(len(rows), ri + span)):
        row = rows[rj]
        if ci >= len(row):
            continue
        dg = num_like(row[ci])
        if dg and dg not in vals:
            vals.append(dg)
    return vals


def find_own_account(rows, sheet_name, max_scan=80):
    """返回 dict:
        found      : 找到的唯一账号 digits（或 ""）
        value_raw  : 原始值
        location   : 'Sheet「xx」!C5'
        label      : 表头原文
        multi      : 若同列出现多个不同账号，这里是 list
        counter    : 扫描到的「对方账号」类标签数量
        near       : 附近出现的候选账号（用于无本方字段时给人看）
    """
    res = {
        "found": "", "value_raw": "", "location": "", "label": "",
        "multi": [], "counter": 0, "near": [],
    }

    label_hits = []
    for ri, row in enumerate(rows[:max_scan]):
        for ci, cell in enumerate(row):
            if cell is None:
                continue
            t = to_half(cell).strip()
            if not t or len(t) > 40:
                continue
            if LABEL_EXCLUDE.search(t):
                if "账" in t or "帐" in t:
                    res["counter"] += 1
                continue
            if LABEL_ACCT.search(t):
                label_hits.append((ri, ci, t))

    # 第一优先：先判多账户。任一标签列出现 >1 个不同账号，即视为多账户文件。
    for ri, ci, t in label_hits:
        vals = _column_values(rows, ri, ci)
        if len(vals) > 1:
            res.update(multi=vals, label=t,
                       location="Sheet「%s」!%s%d（列，%d 个不同账号）"
                                % (sheet_name, _col_letter(ci), ri + 1, len(vals)))
            return res

    # 第二优先：标签右侧第一个非空单元格
    for ri, ci, t in label_hits:
        row = rows[ri]
        for cj in range(ci + 1, min(ci + 8, len(row))):
            v = row[cj]
            if v is None or (isinstance(v, str) and not v.strip()):
                continue
            dg = num_like(v)
            if dg:
                res.update(found=dg, value_raw=_raw_str(v), label=t,
                           location="Sheet「%s」!%s%d" % (sheet_name, _col_letter(cj), ri + 1))
                return res
            break

    # 第三优先：标签下方第一个非空单元格 / 整列唯一值
    for ri, ci, t in label_hits:
        vals = _column_values(rows, ri, ci)
        if len(vals) == 1:
            res.update(found=vals[0], value_raw=vals[0], label=t,
                       location="Sheet「%s」!%s%d（列，逐行重复）"
                                % (sheet_name, _col_letter(ci), ri + 1))
            return res

    # 都没找到：把附近像账号的值收起来，给人看
    seen = []
    for ri, row in enumerate(rows[:max_scan]):
        for cell in row:
            dg = num_like(cell)
            if dg and dg not in seen:
                seen.append(dg)
                if len(seen) >= 8:
                    break
        if len(seen) >= 8:
            break
    res["near"] = seen
    return res


def _raw_str(v):
    if isinstance(v, float):
        return format(v, ".0f") if abs(v) >= 1e15 else str(v)
    return str(v)


# ==========================================================================
#  文件名取证
# ==========================================================================

def filename_tokens(name):
    return [(m.group(), m.start()) for m in re.finditer(r"\d{8,}", name)]


def relate(token, ad):
    """返回 (关系, 长度差) 或 (None, None)。"""
    if token == ad:
        return "equal", 0
    if len(token) >= 8 and ad.endswith(token):
        d = len(ad) - len(token)
        return ("acct_ends_token", d) if d <= 10 else (None, None)
    if len(ad) >= 8 and token.endswith(ad):
        d = len(token) - len(ad)
        return ("token_ends_acct", d) if d <= 10 else (None, None)
    return None, None


def tier_of(diff):
    if diff == 0:
        return "高"
    if diff <= 3:
        return "中"
    return "低"


def match_filename(name, accounts):
    """返回 dict(tokens, hits[{acct, token, diff, rel}], ambiguous, chosen)"""
    toks = filename_tokens(name)
    hits = []
    for tok, _pos in toks:
        for a in accounts:
            rel, diff = relate(tok, a["digits"])
            if rel:
                hits.append({"acct": a, "token": tok, "diff": diff, "rel": rel})
    out = {"tokens": [t for t, _ in toks], "hits": hits, "ambiguous": False, "chosen": None}
    if not hits:
        return out

    uniq = {}
    for h in hits:
        uniq[("digits", h["acct"]["digits"])] = h
    if len(uniq) == 1:
        out["chosen"] = list(uniq.values())[0]
        return out

    # 多个账户：先看是否都由同一个 token 引出（互为后缀），取最长
    toks_involved = set(h["token"] for h in hits)
    if len(toks_involved) == 1:
        longest = max(uniq.values(), key=lambda h: len(h["acct"]["digits"]))
        same = [h for h in uniq.values() if len(h["acct"]["digits"]) == len(longest["acct"]["digits"])]
        if len(same) == 1:
            longest["_note"] = "文件名另有 %d 个账户符合后缀关系，已取最长者" % (len(uniq) - 1)
            out["chosen"] = longest
            return out
    out["ambiguous"] = True
    return out


# ==========================================================================
#  PDF：只做文件名 / 同源 excel / 粗略文本层，不做 OCR
# ==========================================================================

def pdf_text_tokens(data, limit=200):
    """粗略抽取 PDF 文本层里的数字串。中文 CID 字体大概率抽不出，抽不到就返回空。"""
    chunks = []
    for m in re.finditer(rb"stream\r?\n(.*?)endstream", data, re.S):
        raw = m.group(1)
        try:
            dec = zlib.decompress(raw)
        except Exception:  # noqa: BLE001
            dec = raw
        for lit in re.findall(rb"\((?:\\\\.|[^\\\\()])*\)", dec):
            try:
                chunks.append(lit[1:-1].decode("latin-1"))
            except Exception:  # noqa: BLE001
                continue
        if len(chunks) > 200000:
            break
    text = " ".join(chunks)
    toks = re.findall(r"\d{8,32}", text)
    out = []
    for t in toks:
        if t not in out:
            out.append(t)
        if len(out) >= limit:
            break
    return out


# ==========================================================================
#  裁决
# ==========================================================================

def new_record(f):
    return {
        "file": f, "fmt": "unknown", "sha1": "", "evidence": [],
        "acct": None, "confidence": "", "status": "", "check": "未校验",
        "note": "", "reason": "", "raw": "",
    }


def judge(acc_index, fname_hits, content, pdf_tokens, accounts, fmt=""):
    """把各路证据合成结论。acc_index: digits -> account"""
    ev = []
    picked = None
    status = ""
    conf = ""
    reason = ""
    note = ""

    fn_chosen = fname_hits.get("chosen")
    fn_ambiguous = fname_hits.get("ambiguous")

    # ---- 内容优先（内容证据最强）----
    ct_found = content.get("found") if content else ""
    ct_multi = content.get("multi") if content else []
    if ct_found:
        ev.append({
            "source": "content_header", "location": content["location"],
            "label": content["label"], "value_raw": content["value_raw"],
        })

    # ---- 文件名证据 ----
    if fn_chosen:
        ev.insert(0, {
            "source": "filename", "location": "文件名",
            "label": "", "value_raw": fn_chosen["token"],
        })

    # ---- 冲突判定 ----
    if fn_chosen and ct_found and fn_chosen["acct"]["digits"] != ct_found:
        picked = None
        status = "待确认-证据冲突"
        note = "文件名指向 %s，内容 %s 指向 %s，两者不一致" % (
            fn_chosen["acct"]["digits"], content["location"], ct_found)
        reason = "文件名与文件内容指向不同账号，请打开文件核对本方账号到底是哪一个"
        return ev, picked, conf, status, note, reason

    # ---- 多账户 ----
    if ct_multi:
        picked = None
        if content.get("location"):
            ev.append({"source": "content_header", "location": content["location"],
                       "label": content.get("label", ""), "value_raw": ""})
        if len(ct_multi) > 20:
            status = "待确认-账号列不唯一"
            note = "「%s」列出现 %d 个不同账号（示例 %s），更像对方账号列" % (
                content["label"], len(ct_multi), "、".join(ct_multi[:5]))
            reason = "该列每个值都不同，不能判断哪一个是本方账号，请打开文件确认账户信息区"
        else:
            status = "待确认-多账户"
            note = "「%s」列出现 %d 个不同账号：%s" % (
                content["label"], len(ct_multi), "、".join(ct_multi[:5]))
            reason = "同一文件含多个账户（%s），需确认是否应拆成多行，以及每个账号对应哪个账户" % (
                content["location"])
        return ev, picked, conf, status, note, reason

    # ---- 定案 ----
    if ct_found:
        a = acc_index.get(ct_found)
        if a is None:
            a = _by_suffix(ct_found, accounts)
        if a is not None:
            picked = a
            status = "已匹配-内容"
            if fn_chosen:
                status = "已匹配-文件名+内容一致"
                conf = "高"
                note = "文件名与内容一致"
            else:
                conf = "高"
            if fn_chosen and fn_chosen.get("_note"):
                note = (note + "；" if note else "") + fn_chosen["_note"]
            return ev, picked, conf, status, note, reason
        status = "无匹配-账号不在清单"
        note = "内容 %s 抽到账号 %s，清单中无此账号" % (content["location"], ct_found)
        reason = "该文件的账号不在账面清单里，需确认：是新开户未入账、还是保证金/理财等非银行存款账户"
        return ev, None, "", status, note, reason

    if fn_ambiguous:
        cands = []
        for h in fname_hits["hits"]:
            d = h["acct"]["digits"]
            if d not in cands:
                cands.append(d)
        status = "待确认-文件名多候选"
        note = "文件名含 %d 个账户：%s" % (len(cands), "、".join(cands[:6]))
        reason = "文件名里出现多个账户号，无法唯一确定本方账号（其中可能有对方账号）"
        return ev, None, "", status, note, reason

    if fn_chosen:
        a = fn_chosen["acct"]
        diff = fn_chosen["diff"]
        if diff == 0:
            picked, status, conf = a, "已匹配-文件名", "中"
            note = "仅文件名证据（未做内容校验）"
            return ev, picked, conf, status, note, reason
        if diff <= 3:
            picked, status, conf = a, "已匹配-文件名", "中"
            note = "文件名 %s 与清单 %s 差 %d 位（按后缀/前导0兼容）" % (
                fn_chosen["token"], a["digits"], diff)
            return ev, picked, conf, status, note, reason
        status = "待确认-后缀匹配"
        note = "文件名 %s 只是清单 %s 的尾号（差 %d 位），且内容无法确认" % (
            fn_chosen["token"], a["digits"], diff)
        reason = "仅凭尾号匹配到 %s，证据不足，请人工核对" % a["digits"]
        return ev, None, "", status, note, reason

    # ---- PDF ----
    if pdf_tokens:
        hit = [t for t in pdf_tokens if t in acc_index]
        if len(hit) == 1:
            a = acc_index[hit[0]]
            ev.append({"source": "pdf_text", "location": "PDF 文本层",
                       "label": "", "value_raw": hit[0]})
            return ev, a, "低", "已匹配-PDF文本", "PDF 文本层抽到唯一账号，建议复核首页", ""
        status = "待确认-PDF"
        note = "PDF 文本层抽到 %d 个候选：%s" % (len(pdf_tokens), "、".join(pdf_tokens[:5]))
        reason = "PDF 里出现多个账号，分不清本方/对方，请打开首页看账户名与账号"
        return ev, None, "", status, note, reason

    # ---- PDF 兜底（无文件名、无文本层）----
    if fmt == "pdf":
        status = "待确认-PDF"
        note = "PDF 未命中文件名账号，无同名 Excel 可继承，文本层也未抽到账号"
        reason = "需人工打开 PDF 首页看账户名称与账号；若该账户有 Excel/OFD 导出，补进来可自动匹配"
        return ev, None, "", status, note, reason

    # ---- 什么都没 ----
    near = (content or {}).get("near") or []
    counter = (content or {}).get("counter") or 0
    if counter:
        status = "待确认-无本方字段"
        note = "只找到「对方账号」类表头 %d 个；附近候选：%s" % (
            counter, "、".join(near[:5]) if near else "无")
        reason = "文件里没有「本方账号」字段，只有对方账号。打开文件确认本方账户（可能在页眉/账号栏）"
        return ev, None, "", status, note, reason
    status = "待确认-无本方字段"
    note = "未找到任何账号字段；附近候选：%s" % ("、".join(near[:5]) if near else "无")
    reason = "文件里未识别到账号（可能是图片/特殊版式），请打开确认"
    return ev, None, "", status, note, reason


def _by_suffix(dg, accounts):
    cand = [a for a in accounts if len(dg) >= 8 and a["digits"].endswith(dg)]
    if len(cand) == 1:
        return cand[0]
    cand2 = [a for a in accounts if len(a["digits"]) >= 8 and dg.endswith(a["digits"])]
    if len(cand2) == 1:
        return cand2[0]
    return None


# ==========================================================================
#  主流程
# ==========================================================================

def run(args):
    t0 = _dt.datetime.now()
    print("[1/4] 读账面账户清单：%s" % args.acct)
    accounts, ameta = read_accounts(args.acct)
    acc_index = {}
    for a in accounts:
        acc_index.setdefault(a["digits"], a)
    uniq_acct = len(acc_index)
    print("      %d 行账户，%d 个唯一账号（sheet=%s，表头第 %s 行）"
          % (len(accounts), uniq_acct, ameta["sheet"], ameta["header_row"]))
    if uniq_acct != len(accounts):
        print("      注意：同一账号出现多行（可能是多币种），按账号做键会合并，见「说明」表")

    print("[2/4] 枚举流水文件：%s" % args.flow)
    files = list(iter_flow(args.flow))
    # 排除账户清单本身（清单如果放在流水目录里，不能当成流水）
    acct_base = os.path.basename(args.acct).lower()
    acct_size = os.path.getsize(args.acct)
    files = [f for f in files
             if not (f["name"].lower() == acct_base and f["size"] == acct_size)]
    print("      %d 个文件" % len(files))
    if not files:
        raise RuntimeError("没有枚举到任何流水文件")

    print("[3/4] 逐个匹配 ……")
    records = []
    for i, f in enumerate(files, 1):
        rec = new_record(f)
        try:
            data = f["read"]()
            rec["sha1"] = sha1_of(data)
            fmt = sniff(data, f["name"])
            rec["fmt"] = fmt

            # 文件名
            fn = match_filename(f["name"], accounts)

            content = None
            pdf_tokens = []
            need_content = True
            if fn.get("chosen") and fn["chosen"]["diff"] == 0 and fmt in ("xlsx", "xls", "html"):
                need_content = args.cross_check == "always"
                if not need_content:
                    rec["check"] = "未校验"

            if need_content:
                if fmt in ("xlsx", "xls", "html"):
                    try:
                        grids = parse_grid(data, fmt, args.scan_rows, args.scan_cols, args.max_sheets)
                        best = None
                        for sheet_name, rows in grids:
                            r = find_own_account(rows, sheet_name, args.scan_rows)
                            if r["found"] or r["multi"]:
                                best = r
                                break
                            if best is None and (r["near"] or r["counter"]):
                                best = r
                        content = best or {"found": "", "multi": [], "near": [], "counter": 0,
                                           "location": "", "label": "", "value_raw": ""}
                    except Exception as e:  # noqa: BLE001
                        rec["status"] = "待确认-无法解析"
                        rec["note"] = "%s 解析失败：%s" % (fmt, e)
                        rec["reason"] = "文件打不开或结构异常（%s），请人工打开确认" % e
                        records.append(rec)
                        continue
                elif fmt == "pdf":
                    try:
                        pdf_tokens = pdf_text_tokens(data)
                    except Exception:  # noqa: BLE001
                        pdf_tokens = []
                else:
                    rec["status"] = "待确认-无法解析"
                    rec["note"] = "文件头不属于 xlsx/xls/html/pdf：%r" % data[:12]
                    rec["reason"] = "无法识别文件真实格式，请人工打开确认"
                    records.append(rec)
                    continue

            ev, acct, conf, status, note, reason = judge(
                acc_index, fn, content, pdf_tokens, accounts, fmt)
            rec.update(evidence=ev, acct=acct, confidence=conf,
                       status=status, note=note, reason=reason)
            if content and content.get("found"):
                rec["check"] = "一致" if (fn.get("chosen") and
                                          fn["chosen"]["acct"]["digits"] == content["found"]) else "未校验"
        except Exception as e:  # noqa: BLE001
            rec["status"] = "待确认-处理异常"
            rec["note"] = "%s: %s" % (type(e).__name__, e)
            rec["reason"] = "脚本处理该文件时出错，请人工确认（详见 运行日志）"
            if args.verbose:
                traceback.print_exc()
        records.append(rec)
        if i % 20 == 0 or i == len(files):
            print("      %d/%d" % (i, len(files)))

    print("[4/4] 写结果 ……")
    post_pass_pdf(records)
    out_path = args.out or default_out(args)
    write_result(out_path, records, accounts, ameta, args, t0)
    print_summary(records, out_path)
    return out_path


def post_pass_pdf(records):
    """PDF 兜底：同目录存在同名 Excel 且该 Excel 已唯一匹配时，继承其账号。"""
    by_base = {}
    for rec in records:
        base = os.path.splitext(rec["file"]["name"])[0].lower()
        by_base.setdefault(base, []).append(rec)

    for rec in records:
        if rec["status"] != "待确认-PDF":
            continue
        base = os.path.splitext(rec["file"]["name"])[0].lower()
        sibs = [r for r in by_base.get(base, [])
                if r is not rec and r["acct"] and r["fmt"] in ("xlsx", "xls", "html")]
        if len(sibs) == 1:
            src = sibs[0]
            rec["acct"] = src["acct"]
            rec["status"] = "已匹配-同源PDF"
            rec["confidence"] = "中"
            rec["note"] = "继承同名 Excel「%s」的账号（该 Excel 为 %s）" % (
                src["file"]["name"], src["status"])
            rec["reason"] = ""
            rec["evidence"] = [{
                "source": "sibling", "location": "同目录同名文件：%s" % src["file"]["name"],
                "label": "", "value_raw": src["acct"]["digits"],
            }]


def default_out(args):
    base = os.path.basename(args.flow).rsplit(".", 1)[0]
    d = os.path.dirname(os.path.abspath(args.flow)) if os.path.isdir(args.flow) else os.path.dirname(os.path.abspath(args.flow))
    return os.path.join(d or ".", "网银流水匹配结果_%s_%s.xlsx"
                        % (base, _dt.datetime.now().strftime("%Y%m%d_%H%M")))


def print_summary(records, out_path):
    pending = [r for r in records if r["status"].startswith("待确认")]
    nomatch = [r for r in records if r["status"].startswith("无匹配")]
    ok = [r for r in records if r["status"].startswith("已匹配")]

    print("")
    print("=" * 78)
    print("已匹配 %d 个文件 ｜ 待人工确认 %d ｜ 无匹配(账号不在清单) %d ｜ 合计 %d"
          % (len(ok), len(pending), len(nomatch), len(records)))
    if pending:
        print("-" * 78)
        print("待人工确认（%d）：" % len(pending))
        for i, r in enumerate(pending, 1):
            print("[%2d] %-22s %s" % (i, r["status"], r["file"]["rel"]))
            print("     原因：%s" % (r["reason"] or r["note"]))
            if r["note"] and r["note"] != r["reason"]:
                print("     线索：%s" % r["note"])
    if nomatch:
        print("-" * 78)
        print("无匹配（账号不在清单，%d）：" % len(nomatch))
        for i, r in enumerate(nomatch, 1):
            print("[%2d] %-26s %s" % (i, r["note"][:60], r["file"]["rel"]))
    print("=" * 78)
    print("结果文件：%s" % out_path)


def write_result(path, records, accounts, ameta, args, t0):
    wb = Workbook()

    # ---------- 说明 ----------
    ws = wb.active
    ws.title = "说明"
    ok = [r for r in records if r["status"].startswith("已匹配")]
    pend = [r for r in records if r["status"].startswith("待确认")]
    nomt = [r for r in records if r["status"].startswith("无匹配")]
    matched_acct = set(r["acct"]["digits"] for r in records if r["acct"])
    rows = [
        ("网银流水 × 账面账户清单 匹配结果", ""),
        ("生成时间", t0.strftime("%Y-%m-%d %H:%M:%S")),
        ("脚本版本", VERSION),
        ("流水来源", os.path.abspath(args.flow)),
        ("账户清单", ameta["file"]),
        ("", ""),
        ("流水文件数", len(records)),
        ("已匹配文件数", len(ok)),
        ("待人工确认文件数", len(pend)),
        ("无匹配文件数（账号不在清单）", len(nomt)),
        ("", ""),
        ("账户清单行数", len(accounts)),
        ("账户清单唯一账号数", len(set(a["digits"] for a in accounts))),
        ("有流水的账户数", len(matched_acct)),
        ("有账无流水账户数", len(set(a["digits"] for a in accounts) - matched_acct)),
        ("", ""),
        ("【勾稽检查】", ""),
        ("① 文件总数 = 已匹配 + 待确认 + 无匹配",
         "%d = %d + %d + %d → %s" % (len(records), len(ok), len(pend), len(nomt),
                                     "通过" if len(records) == len(ok) + len(pend) + len(nomt) else "不通过")),
        ("② 账户总数 = 有流水账户 + 有账无流水",
         "%d = %d + %d → %s" % (len(set(a["digits"] for a in accounts)), len(matched_acct),
                                len(set(a["digits"] for a in accounts) - matched_acct),
                                "通过" if len(matched_acct) <= len(set(a["digits"] for a in accounts)) else "不通过")),
        ("", ""),
        ("口径声明", "本表只做「文件 → 账户」的归属匹配，不核对金额、不做截止测试。"),
        ("证据要求", "每一行的『证据位置 / 证据表头 / 命中值』即判定依据；缺失则不进「已匹配」。"),
        ("待确认处理", "见「待人工确认」表，每条都写明原因与下一步看什么。"),
    ]
    for i, (k, v) in enumerate(rows, 1):
        ws.cell(i, 1, k).font = Font(bold=(i == 1 or str(k).startswith("【")))
        ws.cell(i, 2, v)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 90

    # ---------- 待人工确认 ----------
    ws2 = wb.create_sheet("待人工确认")
    head2 = ["序号", "文件", "状态", "原因（人下一步该看什么）", "线索", "格式", "大小"]
    _head(ws2, head2)
    r = 2
    for i, rec in enumerate(pend + nomt, 1):
        ws2.cell(r, 1, i)
        ws2.cell(r, 2, rec["file"]["rel"])
        ws2.cell(r, 3, rec["status"])
        ws2.cell(r, 4, rec["reason"] or "")
        ws2.cell(r, 5, rec["note"])
        ws2.cell(r, 6, rec["fmt"])
        ws2.cell(r, 7, fmt_bytes(rec["file"]["size"]))
        for c in range(1, 8):
            ws2.cell(r, c).fill = WARN_FILL if rec["status"].startswith("待确认") else BAD_FILL
        r += 1
    _widths(ws2, [6, 62, 20, 56, 56, 10, 10])
    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = "A1:G%d" % max(1, r - 1)

    # ---------- 匹配明细 ----------
    ws3 = wb.create_sheet("匹配明细")
    head3 = ["序号", "文件名", "实际格式", "大小", "sha1", "匹配依据", "证据位置",
             "证据表头", "命中值(原始)", "归一化账号", "置信度", "账户名称", "账号(清单)",
             "开户银行", "币种", "期末余额", "匹配状态", "交叉校验", "备注"]
    _head(ws3, head3)
    r = 2
    for i, rec in enumerate(records, 1):
        a = rec["acct"] or {}
        ev = rec["evidence"][0] if rec["evidence"] else {}
        src_map = {"filename": "文件名", "content_header": "内容-表头字段",
                   "pdf_text": "PDF 文本层", "sibling": "同源PDF(继承同名Excel)"}
        vals = [
            i, rec["file"]["rel"], rec["fmt"], fmt_bytes(rec["file"]["size"]), rec["sha1"],
            src_map.get(ev.get("source", ""), "无"), ev.get("location", "-"),
            ev.get("label", ""), ev.get("value_raw", ""),
            (a.get("digits") or ev.get("value_raw", "")),
            rec["confidence"], a.get("company", ""), a.get("digits", ""),
            a.get("bank", ""), a.get("currency", ""), a.get("balance", ""),
            rec["status"], rec["check"], rec["note"],
        ]
        for c, v in enumerate(vals, 1):
            cell = ws3.cell(r, c, v)
            if c == 13:
                cell.number_format = "@"
            if rec["status"].startswith("已匹配"):
                cell.fill = OK_FILL
            elif rec["status"].startswith("待确认"):
                cell.fill = WARN_FILL
            elif rec["status"].startswith("无匹配"):
                cell.fill = BAD_FILL
        r += 1
    _widths(ws3, [6, 60, 11, 9, 14, 16, 40, 14, 22, 22, 7, 30, 22, 26, 7, 14, 20, 10, 46])
    ws3.freeze_panes = "C2"
    ws3.auto_filter.ref = "A1:S%d" % max(1, r - 1)

    # ---------- 账户全覆盖 ----------
    ws4 = wb.create_sheet("账户全覆盖")
    head4 = ["序号", "账户名称", "账号", "开户银行", "币种", "期末余额", "命中文件数", "命中文件", "覆盖状态"]
    _head(ws4, head4)
    by_acct = {}
    for rec in records:
        if rec["acct"]:
            by_acct.setdefault(rec["acct"]["digits"], []).append(rec["file"]["rel"])
    r = 2
    for i, a in enumerate(accounts, 1):
        fl = by_acct.get(a["digits"], [])
        vals = [i, a["company"], a["digits"], a["bank"], a["currency"],
                a["balance"] if a["balance"] is not None else "",
                len(fl), "；".join(fl[:8]) + ("…" if len(fl) > 8 else ""),
                "有流水" if fl else "无流水"]
        for c, v in enumerate(vals, 1):
            cell = ws4.cell(r, c, v)
            if c == 3:
                cell.number_format = "@"
            if not fl:
                cell.fill = BAD_FILL
        r += 1
    _widths(ws4, [6, 34, 22, 30, 7, 16, 9, 60, 10])
    ws4.freeze_panes = "A2"

    # ---------- 有账无流水 ----------
    ws5 = wb.create_sheet("有账无流水")
    head5 = ["序号", "分类", "账户名称", "账号", "开户银行", "币种", "期末余额", "备注"]
    _head(ws5, head5)
    no_flow = [a for a in accounts if a["digits"] not in by_acct]
    no_flow.sort(key=lambda a: (a["status"] != "非零余额", a["company"] or ""))
    r = 2
    for i, a in enumerate(no_flow, 1):
        note = "账面余额非零，重点核查是否漏取流水" if a["status"] == "非零余额" else "账面余额为 0，通常无需流水"
        vals = [i, a["status"], a["company"], a["digits"], a["bank"], a["currency"],
                a["balance"] if a["balance"] is not None else "", note]
        for c, v in enumerate(vals, 1):
            cell = ws5.cell(r, c, v)
            if c == 4:
                cell.number_format = "@"
        r += 1
    _widths(ws5, [6, 10, 34, 22, 30, 7, 16, 40])
    ws5.freeze_panes = "A2"

    # ---------- 有流水无账 ----------
    ws6 = wb.create_sheet("有流水无账")
    head6 = ["序号", "抽到的账号", "来源文件", "证据位置", "备注"]
    _head(ws6, head6)
    r = 2
    for rec in records:
        if rec["status"] == "无匹配-账号不在清单":
            ev = rec["evidence"][0] if rec["evidence"] else {}
            ws6.cell(r, 1, r - 1)
            ws6.cell(r, 2, ev.get("value_raw", "")).number_format = "@"
            ws6.cell(r, 3, rec["file"]["rel"])
            ws6.cell(r, 4, ev.get("location", ""))
            ws6.cell(r, 5, "清单中无此账号")
            r += 1
    _widths(ws6, [6, 24, 60, 40, 30])
    ws6.freeze_panes = "A2"

    # ---------- 重复文件 ----------
    ws7 = wb.create_sheet("重复文件")
    head7 = ["序号", "sha1", "份数", "文件列表"]
    _head(ws7, head7)
    by_sha = {}
    for rec in records:
        if rec["sha1"]:
            by_sha.setdefault(rec["sha1"], []).append(rec["file"]["rel"])
    r = 2
    for i, (h, fs) in enumerate(sorted(by_sha.items(), key=lambda kv: -len(kv[1])), 1):
        if len(fs) < 2:
            continue
        ws7.cell(r, 1, i)
        ws7.cell(r, 2, h)
        ws7.cell(r, 3, len(fs))
        ws7.cell(r, 4, "；".join(fs))
        r += 1
    _widths(ws7, [6, 44, 8, 100])
    ws7.freeze_panes = "A2"

    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    wb.save(path)


def _head(ws, heads):
    for c, h in enumerate(heads, 1):
        cell = ws.cell(1, c, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 22


# ==========================================================================

def main(argv=None):
    ap = argparse.ArgumentParser(
        description="网银流水 × 账面账户清单 匹配工具 v%s" % VERSION,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例：\n"
               "  python 流水匹配_v2.py --acct 清单.xlsx --flow 流水.zip\n"
               "  python 流水匹配_v2.py --acct 清单.xlsx --flow D:\\资料\\13、网银流水\n")
    ap.add_argument("--acct", required=True, help="账面账户清单（xlsx 或 xls）")
    ap.add_argument("--flow", required=True, help="流水文件夹 或 zip 压缩包")
    ap.add_argument("--out", default="", help="输出 xlsx 路径（默认放在流水同级）")
    ap.add_argument("--cross-check", choices=["auto", "always", "never"], default="auto",
                    help="内容交叉校验策略：auto=文件名完全命中时跳过（快）；always=全部校验（慢）")
    ap.add_argument("--scan-rows", type=int, default=80, help="每表扫描行数（默认 80）")
    ap.add_argument("--scan-cols", type=int, default=40, help="每表扫描列数（默认 40）")
    ap.add_argument("--max-sheets", type=int, default=5, help="每文件最多看几个 sheet")
    ap.add_argument("--verbose", action="store_true", help="打印异常堆栈")
    args = ap.parse_args(argv)

    # Windows 控制台可能是 GBK，避免中文路径/符号导致 UnicodeEncodeError
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except Exception:  # noqa: BLE001
            pass

    if not os.path.exists(args.acct):
        sys.stderr.write("找不到账户清单：%s\n" % args.acct)
        return 2
    if not os.path.exists(args.flow):
        sys.stderr.write("找不到流水来源：%s\n" % args.flow)
        return 2
    try:
        run(args)
    except Exception as e:  # noqa: BLE001
        sys.stderr.write("\n[失败] %s\n" % e)
        if args.verbose:
            traceback.print_exc()
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
