# -*- coding: utf-8 -*-
"""固定资产 SWP：客户台账（fa / fa_beg）路径的判定器。

判三件事：
1. 两张表都空时 ``load_cards`` 返回 None —— 瑞云那种「没台账、按序时账造明细」
   的老路必须原样保留。
2. 台账一行 = 期初/期末两个数，变动按**净额**拆；纯新增、纯处置、两边都有
   三种形态都要拆对，且 期末 = 期初 + 新增 − 减少。
3. 台账出现 TB 没有的类别要当场报错，不能静默跳过。

**没有 sys.exit 就是假通过**（项目红线）。
"""

import os
import sys
import datetime as dt
from decimal import Decimal as D

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(ROOT, "程序"))

fails = []


def ck(name, cond, extra=""):
    print(("  ✓ " if cond else "  ✗ ") + name + (("  " + str(extra)) if extra and not cond else ""))
    if not cond:
        fails.append(name)


def card_row(src_row, cat, code, orig, dep):
    """仿 fa / fa_beg 的 select 顺序：0 src_row 1 company 2 category 3 code 4 name
    5 dept 6 loc 7 start 8 life 9 resid 10 orig 11 dep 12 impair 13 status"""
    return (src_row, "某公司", cat, code, "设备" + str(src_row), "", "深圳",
            dt.date(2024, 1, 1), 60, 0.0, D(str(orig)), D(str(dep)), D("0"), "")


def main():
    from modules.fa_swp import synth

    cats = {"机器设备": {"fa_beg": D("1000"), "dep_beg": D("200")},
            "电子设备": {"fa_beg": D("0"), "dep_beg": D("0")}}
    prep_res = {"cats": cats, "cat_order": ["机器设备", "电子设备"]}
    bd = dt.date(2026, 6, 30)

    # ---- 1) 三种行形态
    cards = {
        "beg": {1: card_row(1, "机器设备", "A1", 600, 100),
                2: card_row(2, "机器设备", "A2", 400, 100)},
        "end": {1: card_row(1, "机器设备", "A1", 600, 150),      # 两边都有：本期计提 50
                2: card_row(2, "机器设备", "A2", 0, 0),          # 纯处置：期初 400/100 全转
                3: card_row(3, "电子设备", "B1", 300, 30)},      # 纯新增
    }
    rows = synth._cards_rows(prep_res, cards, bd)
    ck("行数 = 三张卡片的并集", len(rows) == 3, len(rows))
    m = {r["code"]: r for r in rows}
    ck("两边都有：净额新增 0", m["A1"]["add"] == 0 and m["A1"]["reduce"] == 0)
    ck("纯处置：减少 400", m["A2"]["reduce"] == D("400"), m["A2"]["reduce"])
    ck("纯新增：新增 300", m["B1"]["add"] == D("300"), m["B1"]["add"])
    tot_beg = sum((r["fa_beg"] for r in rows), D("0"))
    tot_end = sum((r["fa_end"] for r in rows), D("0"))
    net = sum((r["add"] - r["reduce"] for r in rows), D("0"))
    ck("期末 = 期初 + 新增 − 减少", tot_end == tot_beg + net, (tot_beg, net, tot_end))
    ck("净值 = 原值 − 折旧 − 减值",
       all(r["nbv_end"] == r["fa_end"] - r["dep_end"] - r["imp_end"] for r in rows))

    # ---- 2) 类别不在 TB 里 → 报错
    bad = {"beg": {9: card_row(9, "运输设备", "C1", 100, 10)}, "end": {}}
    try:
        synth._cards_rows(prep_res, bad, bd)
        ck("未知类别要报错", False, "没有报错")
    except ValueError as e:
        ck("未知类别要报错", "运输设备" in str(e), e)

    # ---- 3) load_cards：空表返回 None，有数据返回两组
    import duckdb
    conn = duckdb.connect(":memory:")
    for t in ("fa", "fa_beg"):
        conn.execute("create table %s (src_row bigint, company_name varchar, category_name varchar,"
                     " asset_code varchar, asset_name varchar, dept_name varchar, location varchar,"
                     " start_date date, life_months decimal(20,4), residual_rate decimal(20,4),"
                     " orig_cost decimal(20,2), acc_dep decimal(20,2), impair decimal(20,2),"
                     " status varchar)" % t)
    ck("两张表都空 → None", synth.load_cards(conn) is None)
    conn.execute("insert into fa_beg values (1,'某公司','机器设备','A1','x','','深圳',"
                 "'2024-01-01',60,0,600,100,0,'')")
    conn.execute("insert into fa values (1,'某公司','机器设备','A1','x','','深圳',"
                 "'2024-01-01',60,0,600,150,0,'')")
    got = synth.load_cards(conn)
    ck("有数据 → 按 src_row 配对", got is not None and set(got["beg"]) == {1} and set(got["end"]) == {1})
    conn.close()

    print()
    if fails:
        print("失败 %d 项：%s" % (len(fails), "、".join(fails)))
        return 1
    print("全部通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
