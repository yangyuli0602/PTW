# -*- coding: utf-8 -*-
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
