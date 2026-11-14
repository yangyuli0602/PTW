# -*- coding: utf-8 -*-
"""清洗链路的三个「静默失效」防护（2026-11-11）。

背景：这三处原来都不报错、只是结果不对，属于最难发现的一类问题——
  A. 按关键词删行时，列名填错 → 退化成全列拼接搜索 → 大规模误删（实测 206,244 行）
  B. 按关键列判空删行时，英文/繁体表头一个都不命中 → 退化成「所有列都空才删」→ 一行没删
  C. 双层表头第二层写「借方金额」而非「借方」→ 认不出两层 → 列名整体错位

判定标准：不报错才是问题。填错列名要 noop 说出来，不能猜。
"""
import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
while not os.path.isdir(os.path.join(ROOT, "程序")):
    parent = os.path.dirname(ROOT)
    if parent == ROOT:
        break
    ROOT = parent
sys.path.insert(0, os.path.join(ROOT, "程序"))

from core.clean import _b_drop_keyword, _b_drop_empty_key
from core.fields import detect_header

fails = []
n = [0]


def ck(name, cond, got=None):
    n[0] += 1
    print(("  ok   " if cond else "  FAIL ") + name + ("" if cond else " | %r" % (got,)))
    if not cond:
        fails.append(name)


COLS = ["账簿", "记账日期", "凭证号", "科目编码", "借方金额"]

print("=== A. 按关键词删行：列名填错不能猜 ===")
r = _b_drop_keyword({"col": "日期", "kw": "2026-10"}, COLS, "t")
ck("唯一包含候选 → 用「记账日期」", r["kind"] == "filter" and "记账日期" in r["cond"], r)
ck("并在说明里注明匹配过程", "包含匹配" in r.get("reason", ""), r.get("reason"))

r = _b_drop_keyword({"col": "科目", "kw": "x"}, ["科目编码", "科目名称", "凭证号"], "t")
ck("多个候选 → noop 让人选", r["kind"] == "noop" and "候选" in r["reason"], r)

r = _b_drop_keyword({"col": "不存在的列", "kw": "x"}, COLS, "t")
ck("完全找不到 → noop 且列出可用列",
   r["kind"] == "noop" and "可用列" in r["reason"] and "凭证号" in r["reason"], r)

r = _b_drop_keyword({"col": "", "kw": "制表人"}, COLS, "t")
ck("不填列名 → 仍走全列搜索（预设删页脚靠它）",
   r["kind"] == "filter" and "concat_ws" in r["cond"], r)

r = _b_drop_keyword({"col": "凭证号", "kw": "x"}, COLS, "t")
ck("精确命中 → 用该列", r["kind"] == "filter" and '"凭证号"' in r["cond"], r)

r = _b_drop_keyword({"col": "凭证号", "kw": ""}, COLS, "t")
ck("没填关键词 → noop", r["kind"] == "noop", r)

print("=== B. 按关键列判空：认不出要报，不能静默 ===")
r = _b_drop_empty_key({}, COLS, "t")
ck("中文标准表头 → 命中", r["kind"] == "filter" and '"科目编码"' in r["cond"], r)

r = _b_drop_empty_key({}, ["总帐科目", "总账科目长文本", "Cl. Bal."], "t")
ck("繁体「总帐科目」→ 命中", r["kind"] == "filter" and "总帐科目" in r["cond"], r)

r = _b_drop_empty_key({}, ["account code", "Voucher No", "amount"], "t")
ck("英文表头 → 命中", r["kind"] == "filter", r)

r = _b_drop_empty_key({}, ["A", "B", "C"], "t")
ck("无关键列 → noop 而不是「所有列都空才删」",
   r["kind"] == "noop" and "没认出关键列" in r["reason"], r)

print("=== C. 双层表头：第二层带后缀也要认 ===")
rows = [["科目余额表"], ["公司", "2026年1期至9期"],
        ["", "", "期初余额", "本期发生额"],
        ["科目编码", "科目名称", "借方金额", "贷方金额"],
        ["1001", "现金", "0", "0"]]
ck("「借方金额 / 贷方金额」当第二层 → 认出两层", detect_header(rows, "tb") != (1, 1),
   detect_header(rows, "tb"))
hr, hrows = detect_header(rows, "tb")
ck("表头从分类行起、跨 2 行", hrows == 2 and hr <= 3, (hr, hrows))

rows2 = [["银行存款明细账"], ["科目编码", "科目名称", "借方金额", "贷方金额"],
         ["1001", "现金", "0", "0"]]
ck("大标题 + 单层表头 → 不误判成两层", detect_header(rows2, "tb") == (2, 1), detect_header(rows2, "tb"))

rows3 = [["科目编码", "科目名称", "期初余额", "本期发生额", "期末余额"],
         ["1001", "现金", "0", "0", "0"]]
ck("标准单层表头 → 第 1 行 1 层", detect_header(rows3, "tb") == (1, 1), detect_header(rows3, "tb"))

rows4 = [["凭证日期", "凭证号", "摘要", "科目编码", "借方金额", "贷方金额"],
         ["2026-01-05", "记-1", "收款", "1002", "100", ""]]
ck("序时账单层表头 → 不受影响", detect_header(rows4, "journal") == (1, 1), detect_header(rows4, "journal"))

# 真实场景：导出工具把合并单元格铺平，标题行 9 格都是「科目余额表」，
# 分类行和明细行都重复写着「序号|科目编码|科目名称|币种|账簿|核算维度」。
# 旧版判成 (1,1)，列名整体错位；新版应认出 (3,2)。
rows5 = [["科目余额表"] * 9,
         ["账簿 : DA ZHONG TANG"] * 9,
         ["序号", "科目编码", "科目名称", "币种", "账簿", "核算维度", "年初余额", "年初余额", "期初余额"],
         ["序号", "科目编码", "科目名称", "币种", "账簿", "核算维度", "借方金额", "贷方金额", "借方金额"],
         ["1", "1001", "库存现金", "新加坡元", "GUSHENGTANG", "", "", "", ""]]
hr, hrows = detect_header(rows5, "tb")
ck("铺平的标题行不当表头", hr >= 3, (hr, hrows))
ck("双层都带主体列（前缀相同）→ 认出两层", (hr, hrows) == (3, 2), (hr, hrows))

# 单层表头把「科目编码」和「期初本位币(借)」写在同一行，不能误判成两层
rows6 = [["科目余额表"] * 8,
         ["序号", "科目编码", "科目名称", "币种", "账簿", "期初本位币(借)", "期初本位币(贷)", "期末余额"],
         ["1", "1001", "库存现金", "人民币", "总部", "0", "0", "0"]]
ck("单层但含方向词的列名 → 不误判两层", detect_header(rows6, "tb") == (2, 1), detect_header(rows6, "tb"))

print("\n%s  共 %d 项，失败 %d" % ("ALL PASS" if not fails else "FAILED", n[0], len(fails)))
if fails:
    for f in fails:
        print("   -", f)
sys.exit(1 if fails else 0)
