# -*- coding: utf-8 -*-
"""导入数据：网页表格（HTML 伪 .xls / .html 直传）——V1.2.120。

背景：很多 ERP / 网银的「导出 Excel」产出的其实是 HTML 表格，后缀写成 .xls。
用 Excel 打开是乱码（用户描述「全是问号」），改成 .html 就正常。
平台原来只看后缀，`.xls` 非 OLE2 就一律当 xlsx，最后死在 zipfile 报
"File is not a zip file"，被前端翻译成「文件可能已损坏」，方向全错。

判定标准：
  A. 内容识别：GBK/BOM/大 head 的 HTML 都要认出来，并转成真 xlsx
  B. 结构保真：多表各成一个工作表、合并格只留左上、嵌套不重复、空表过滤、行列不少
  C. 端到端：转出来的表能被 DuckDB 正常读入且金额不翻倍
  D. 不误伤/不误导：docx 改后缀、SpreadsheetML、纯文本都要给准确说法，不能报"损坏"
"""
import os
import shutil
import sys
import tempfile
import zipfile

ROOT = os.path.dirname(os.path.abspath(__file__))
while not os.path.isdir(os.path.join(ROOT, "程序")):
    parent = os.path.dirname(ROOT)
    if parent == ROOT:
        break
    ROOT = parent
sys.path.insert(0, os.path.join(ROOT, "程序"))

PROJ = os.path.abspath(os.path.join(ROOT, ".."))            # 平台根
_TMP = os.path.join(PROJ, "tmp")
os.makedirs(_TMP, exist_ok=True)
WORK = tempfile.mkdtemp(prefix="ptw_html_", dir=_TMP)
os.environ["KZ_DATA_DIR"] = WORK

from core import xlsxinfo  # noqa: E402
from core.db import Ledger  # noqa: E402

fails = []
n = [0]


def ck(name, cond, got=None):
    n[0] += 1
    print(("  ok   " if cond else "  FAIL ") + name + ("" if cond else " | %r" % (got,)))
    if not cond:
        fails.append(name)


def html_doc(body, charset="utf-8", head_pad=0):
    head = ('<html><head><meta http-equiv="Content-Type" '
            'content="text/html; charset=%s"><title>导出</title>' % charset)
    head += "<style>td{border:1px}</style>" + (" " * head_pad) + "</head>"
    return (head + "<body>" + body + "</body></html>")


def put(name, raw):
    p = os.path.join(WORK, name)
    with open(p, "wb") as fh:
        fh.write(raw)
    return p


led = Ledger(os.path.join(ROOT, "程序"))


def add(path):
    with open(path, "rb") as fh:
        return led.add_file(os.path.basename(path), fh)["file_id"]


def err_of(fn, *a):
    try:
        fn(*a)
        return ""
    except Exception as e:
        return str(e)


TABLE = ("<table><caption>明细账</caption>"
         "<tr><th>日期</th><th>凭证号</th><th>科目编码</th><th>金额</th></tr>"
         "<tr><td>2026-01-05</td><td>记-1</td><td>1001</td><td>100</td></tr>"
         "<tr><td>2026-01-06</td><td>记-2</td><td>1002</td><td>200</td></tr>"
         "<tr><td>2026-01-07</td><td>记-3</td><td>1001</td><td>300</td></tr>"
         "</table>")

# ---------------------------------------------------------------- A 内容识别
# A1 GBK 伪 .xls（charset 声明写成老式的 gb2312，实际是 gbk 扩展字集）
p = put("明细账.xls", html_doc(TABLE, "gb2312").encode("gbk"))
fid = add(p)
r = err_of(led.inspect, fid)
ins = led.inspect(fid) if not r else {}
f = ins.get("file") or {}
ck("A1 GBK 伪 .xls 能被识别", not r, r)
ck("A2 判成 xlsx 且标记来源", f.get("kind") == "xlsx" and f.get("converted_from") == "html",
   (f.get("kind"), f.get("converted_from")))
prev = (ins.get("sheets") or [{}])[0].get("preview") or []
ck("A3 中文不乱码（gb2312→gb18030 生效）",
   bool(prev) and prev[0][:2] == ["日期", "凭证号"], prev[0][:3] if prev else prev)

# A4 UTF-8 BOM + <table> 藏在 3KB 之后（Excel 另存网页的典型形态）
p2 = put("大头部.html", b"\xef\xbb\xbf" + html_doc(TABLE, "utf-8", head_pad=3000).encode("utf-8"))
ck("A4 BOM + 大 head 仍认出 HTML", not err_of(led.inspect, add(p2)))

# A5 .htm 直传
p3 = put("明细账.htm", html_doc(TABLE, "utf-8").encode("utf-8"))
ck("A5 .htm 后缀直传可用", not err_of(led.inspect, add(p3)))


# A6/A7 有些系统把网页表格存成 .csv / .txt（同样是"改成 .html 才看得见"）
p10 = put("明细账.csv", html_doc(TABLE, "utf-8").encode("utf-8"))
ins10 = led.inspect(add(p10))
ck("A6 .csv 里装 HTML 也能转（不是当 CSV 读成乱码）",
   (ins10.get("file") or {}).get("converted_from") == "html"
   and len(ins10.get("sheets") or []) == 1,
   ((ins10.get("file") or {}).get("converted_from"), len(ins10.get("sheets") or [])))
p11 = put("明细账.txt", html_doc(TABLE, "gbk").encode("gbk"))
ins11 = led.inspect(add(p11))
ck("A7 .txt 里装 HTML 同样转", (ins11.get("file") or {}).get("converted_from") == "html",
   (ins11.get("file") or {}).get("converted_from"))
# 真 CSV 不能被误判
p12 = put("真表格.csv", "日期,凭证号,金额\n2026-01-05,记-1,100\n".encode("utf-8"))
ins12 = led.inspect(add(p12))
ck("A8 真 CSV 仍走 CSV 通道", ins12.get("kind") == "csv"
   and not (ins12.get("file") or {}).get("converted_from"), ins12.get("kind"))

# ---------------------------------------------------------------- B 结构保真
# B1 双表各成一页，名字取 caption
two = ("<table><caption>余额表</caption><tr><td>a</td><td>b</td></tr>"
       "<tr><td>1</td><td>2</td></tr></table>"
       "<table><caption>序时账</caption><tr><td>c</td><td>d</td></tr>"
       "<tr><td>3</td><td>4</td></tr></table>")
p4 = put("两张表.xls", html_doc(two, "utf-8").encode("utf-8"))
ins4 = led.inspect(add(p4))
names = [s["sheet"] for s in ins4.get("sheets") or []]
ck("B1 两张表各成一个工作表且取 caption", names == ["余额表", "序时账"], names)

# B2 rowspan/colspan 只留左上，不复制
merged = ("<table><tr><td rowspan=2 colspan=2>100</td><td>a</td></tr>"
          "<tr><td>b</td></tr><tr><td>c</td><td>d</td><td>e</td></tr></table>")
p5 = put("合并格.xls", html_doc(merged, "utf-8").encode("utf-8"))
fid5 = add(p5)
conv5 = led.inspect(fid5)["file"]["path"]
cells5 = xlsxinfo.preview_cells(conv5, "xl/worksheets/sheet1.xml", 5)
ck("B2 合并格只在左上留值", cells5.get(1, {}).get(b"A") == "100", cells5.get(1))
ck("B3 被覆盖的格子为空（不复制，防金额翻倍）",
   cells5.get(1, {}).get(b"B") in (None, "") and cells5.get(2, {}).get(b"A") in (None, ""),
   (cells5.get(1), cells5.get(2)))

# B4 外层空壳 + 内层真表：只留内表
nest = ("<table><tr><td><table><tr><td>x</td><td>y</td></tr>"
        "<tr><td>1</td><td>2</td></tr></table></td></tr></table>")
p6 = put("嵌套.xls", html_doc(nest, "utf-8").encode("utf-8"))
ins6 = led.inspect(add(p6))
ck("B4 嵌套只留内层真表", len(ins6.get("sheets") or []) == 1,
   [s["sheet"] for s in ins6.get("sheets") or []])

# B5 空表被过滤（不算 broken）
mix = "<table><tr><td> </td></tr></table>" + TABLE
p7 = put("含空表.xls", html_doc(mix, "utf-8").encode("utf-8"))
ins7 = led.inspect(add(p7))
ck("B5 空表被过滤且不算损坏",
   len(ins7.get("sheets") or []) == 1 and not (ins7["sheets"][0].get("broken")),
   [(s["sheet"], s.get("broken")) for s in ins7.get("sheets") or []])

# B6 行列数一致（防 write_only 不写 dimension 导致少列）
sh = (ins4.get("sheets") or [{}])[0]
ck("B6 行列数与预览一致", sh.get("rows") == 2 and sh.get("cols") == 2
   and len((sh.get("preview") or [[]])[0]) == 2,
   (sh.get("rows"), sh.get("cols"), sh.get("preview")))

# ---------------------------------------------------------------- C 端到端
fid8 = add(p)                       # 复用 A1 那份 3 行 4 列
imp = led.import_sheet(fid8, "明细账", header_row=1, header_rows=1, target="journal")
sid = (imp or {}).get("sid") or ""
cnt = led.conn.execute('select count(*) from "raw_%s"' % sid).fetchone()[0]
total = led.conn.execute(
    'select sum(try_cast("金额" as double)) from "raw_%s"' % sid).fetchone()[0]
ck("C1 DuckDB 读入行数正确", cnt == 3, cnt)
ck("C2 金额合计正确（100+200+300=600，没被合并格撑大）", abs(float(total or 0) - 600.0) < 0.01, total)

# ---------------------------------------------------------------- D 缓存/不误导
conv = led.inspect(fid8)["file"]["path"]
m1 = os.stat(conv).st_mtime_ns
led.inspect(fid8)
m2 = os.stat(conv).st_mtime_ns
updir = os.path.dirname(conv)
parts = [x for x in os.listdir(updir) if x.endswith(".part")]
ck("D1 二次识别命中缓存（mtime 不变）", m1 == m2, (m1, m2))
ck("D2 没有 .part 残留", not parts, parts)

# D3 docx 改后缀（zip 但无 xl/workbook.xml）
fake_docx = os.path.join(WORK, "伪装.xls")
with zipfile.ZipFile(fake_docx, "w") as z:
    z.writestr("word/document.xml", "<w/>")
msg = err_of(led.inspect, add(fake_docx))
ck("D3 zip 无工作簿 → 报「不是 Excel 工作簿」", "不是 Excel 工作簿" in msg, msg)
ck("D4 报错里不出现 BadZipFile", "BadZipFile" not in msg and "zipfile" not in msg, msg)

# D5 SpreadsheetML
sml = ('<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>'
       '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"></Workbook>')
msg = err_of(led.inspect, add(put("2003.xml", sml.encode("utf-8"))))
ck("D5 SpreadsheetML 有专门提示", "SpreadsheetML" in msg, msg)

# D6 纯文本伪装 .xls → 明确"无法识别"，不是"损坏"
msg = err_of(led.inspect, add(put("其实是csv.xls", b"a,b,c\n1,2,3\n")))
ck("D6 纯文本伪装报「无法识别的格式」", "无法识别的格式" in msg, msg)

# D7 = 开头的单元格保持文本
eq = "<table><tr><td>公式</td><td>值</td></tr><tr><td>=1+1</td><td>2</td></tr></table>"
p9 = put("等号开头.xls", html_doc(eq, "utf-8").encode("utf-8"))
conv9 = led.inspect(add(p9))["file"]["path"]
import openpyxl  # noqa: E402
wb = openpyxl.load_workbook(conv9)
v = wb.active["A2"].value
wb.close()
ck("D7 = 开头的内容存成文本", v == "=1+1", v)

shutil.rmtree(WORK, ignore_errors=True)

print()
print("%d 项通过，失败 %d" % (n[0] - len(fails), len(fails)))
if fails:
    print("失败：" + "；".join(fails))
    sys.exit(1)
