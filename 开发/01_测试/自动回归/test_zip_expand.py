# -*- coding: utf-8 -*-
"""导入数据：压缩包（.zip）自动解压（V1.2.119）。

背景：客户 ERP 按账导出成 zip（如 DMOL_xxx.zip 里是明细账 xlsx），拖进「导入数据」
被告知"请用导入数据包入口"——因为前端把所有 .zip 当成平台数据包；可放进数据包入口
又因缺 manifest.json 被拒，两条路都堵死。现在改成：上传后由服务端判断。

判定标准（四件事）：
  A. 包内表格要登记成待导入文件（rel 保留包内路径），不能整体当数据包丢掉
  B. 真数据包（含 manifest.json）要被识别出来，不能误当普通压缩包解了
  C. 干扰项（~$ 临时文件、readme.txt、docx、内层 zip）各自说明原因，不能静默吞
  D. 坏包要给人话，而不是 BadZipFile 原始异常
"""
import io
import os
import shutil
import sys
import tempfile
import zipfile

import openpyxl

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
WORK = tempfile.mkdtemp(prefix="ptw_zip_", dir=_TMP)
os.environ["KZ_DATA_DIR"] = WORK

from core.db import Ledger  # noqa: E402

fails = []
n = [0]


def ck(name, cond, got=None):
    n[0] += 1
    print(("  ok   " if cond else "  FAIL ") + name + ("" if cond else " | %r" % (got,)))
    if not cond:
        fails.append(name)


def make_xlsx(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["日期", "凭证号", "科目编码", "科目全名", "借方金额", "贷方金额"])
    ws.append(["2026-01-05", "记-1", "1001", "库存现金", "100", "0"])
    wb.save(path)
    wb.close()


xlsx = os.path.join(WORK, "a.xlsx")
make_xlsx(xlsx)
blob = open(xlsx, "rb").read()

# ---- A/C：像客户 ERP 导出的包，夹带各种干扰项 ----
inner = io.BytesIO()
with zipfile.ZipFile(inner, "w") as z:
    z.writestr("更深.zip", b"PK\x03\x04")
zp = os.path.join(WORK, "DMOL_明细账.zip")
with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("明细账/DMOL_0668_明细账.xlsx", blob)
    z.writestr("明细账/~$DMOL_0668_明细账.xlsx", b"lock")
    z.writestr("明细账/readme.txt", "导出说明".encode("gbk"))
    z.writestr("明细账/说明.docx", b"PK\x03\x04fake")
    z.writestr("更深.zip", inner.getvalue())

led = Ledger(os.path.join(ROOT, "程序"))
with open(zp, "rb") as fh:
    fid = led.add_file("DMOL_明细账.zip", fh)["file_id"]
r = led.expand_zip(fid)
files, skipped = r.get("files") or [], r.get("skipped") or []
ck("A1 包内 xlsx 被登记成待导入", len(files) == 1, [x.get("rel") for x in files])
ck("A2 保留包内相对路径", files and files[0].get("rel") == "明细账/DMOL_0668_明细账.xlsx",
   files and files[0].get("rel"))
ck("A3 类型判成 xlsx", files and files[0].get("kind") == "xlsx", files and files[0].get("kind"))
ck("A4 登记结果带 file_id", files and len(files[0].get("file_id") or "") > 0, files and files[0])
whys = {x.get("name"): x.get("why") for x in skipped}
ck("C1 ~$ 临时文件被跳过并说明", "明细账/~$DMOL_0668_明细账.xlsx" in whys, list(whys))
ck("C2 readme.txt 不收（防变成一张文字表）", any("readme.txt" == os.path.basename(k) for k in whys), list(whys))
ck("C3 docx 判非表格", any(os.path.basename(k) == "说明.docx" for k in whys), list(whys))
ck("C4 内层 zip 提示先解开", any("还有压缩包" in (v or "") for v in whys.values()), list(whys.values()))
ck("C5 干扰项一个都不进 files", len(files) == 1, len(files))

# ---- B：真数据包不能被当普通压缩包吞掉 ----
zp2 = os.path.join(WORK, "数据包.zip")
with zipfile.ZipFile(zp2, "w") as z:
    z.writestr("manifest.json", '{"format":"ptw-data-pack/v1","tables":{}}')
with open(zp2, "rb") as fh:
    fid2 = led.add_file("数据包.zip", fh)["file_id"]
r2 = led.expand_zip(fid2)
ck("B1 含 manifest 判为数据包", r2.get("data_pack") is True, r2.get("data_pack"))
ck("B2 数据包不返回文件", not (r2.get("files") or []), r2.get("files"))

# ---- D：坏包给人话 ----
zp3 = os.path.join(WORK, "坏包.zip")
open(zp3, "wb").write(b"PK\x03\x04" + b"\x00" * 64)
with open(zp3, "rb") as fh:
    fid3 = led.add_file("坏包.zip", fh)["file_id"]
msg = ""
try:
    led.expand_zip(fid3)
except Exception as e:
    msg = str(e)
ck("D1 坏包报「压缩包已损坏」", "压缩包已损坏" in msg, msg)
ck("D2 不抛 BadZipFile 原文", "BadZipFile" not in msg and "zipfile" not in msg, msg)

shutil.rmtree(WORK, ignore_errors=True)

print()
print("%d 项检查，%d 失败" % (n[0], len(fails)))
if fails:
    print("失败：" + "；".join(fails))
    sys.exit(1)
