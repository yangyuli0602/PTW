# PTW 平台 · AI 作业指南（导入 → 清洗 → 映射 → 合并 → 检查）

> 面向对象：另一个 AI 代理（或人）。
> 目标：拿到一套**未脱敏**的「科目余额表 + 序时账」，自主完成导入、清洗、字段映射、合并，
> 跑完四组数据检查，输出结论与可交付的 Excel。
> 平台版本：**V1.1.02**（清洗阶段开发中）
> 写这份文档的依据：同一套流程已在 208 万行真实脱敏数据上跑通，四组检查 0 差异。

---

## 0. 五条铁律（先读，违反会毁数据）

1. **不要改 `程序/` 下的任何代码**。所有操作都通过 HTTP 接口完成。代码是你操作的工具，不是待改的对象。
2. **不要删 `程序/../data/` 目录**。那里是账套库（`ledger.duckdb`）、上传件、导出件。
3. **原始表只读**。清洗永远写在 `stg_<sid>` 新表，`raw_<sid>` 不动；「恢复原始」永远可用。
4. **不要猜列名**。每一步的列名必须与 `GET /api/state` 里该表的 `columns` 完全一致（含全角括号、空格）。
5. **每一步之后要自检**。本文每节末尾的「✅ 自检」是硬门槛，不通过不许进下一步。

---

## 1. 启动与连接

Windows，解压即用（包里自带 Python 3.13 运行时，无需装 Python）。

```bat
:: 方式一：双击启动（会自动开浏览器，不适合无人值守）
启动PTW平台.bat

:: 方式二：命令行启动（推荐给 AI，可控端口、不开浏览器）
set KZ_PORT=8790
set KZ_NO_BROWSER=1
set KZ_DATA_DIR=D:\ptw_work\data
程序\runtime\python\python.exe 程序\server.py
```

- `KZ_PORT`：固定端口。不设则从 8787 起找空闲端口，实际端口看启动横幅的「服务地址」。
- `KZ_NO_BROWSER=1`：不弹浏览器。
- `KZ_DATA_DIR`：数据目录。**不设就写到 `程序\../data`**。换账套/隔离测试一定要设，避免污染。
  - ⚠️ 启动横幅里打印的「数据目录」是默认值，设了 `KZ_DATA_DIR` 时以环境变量为准，横幅不反映。

服务只监听 `127.0.0.1`，不对外。关闭进程即退出，数据都在 `data/`。

✅ 自检：
```python
GET http://127.0.0.1:8790/api/modules
# 期望返回 {"version": "1.1.02", "modules": [...3 个板块...]}
```

---

## 2. 数据流全貌

```
上传文件 ──POST /api/upload──> file_id
   │
   ├─ POST /api/inspect  看有几个 sheet、表头在第几行、多少行
   ├─ POST /api/import   每张 sheet 落成 raw_<sid>（原始表，只读）
   │
   ├─ POST /api/clean/save + /api/clean/apply   清洗 → stg_<sid>
   │        （或 POST /api/clean/apply-all 一次洗完全部）
   │
   ├─ POST /api/mapping  中文列名 → 标准字段（tb / journal）
   ├─ POST /api/merge    所有同目标表 UNION 成 tb 或 journal
   │
   ├─ POST /api/recon    四组平衡检查
   └─ POST /api/recon/export  导出多 sheet Excel
```

中途想查数据：`POST /api/query {"sql": "..."}`，可以跑任意 DuckDB SQL。表名见下表。

| 表名 | 含义 | 何时存在 |
|---|---|---|
| `raw_<sid>` | 原始表（带内部列 `_src_row`） | 导入后 |
| `stg_<sid>` | 清洗后 | 应用清洗后 |
| `_dropped_<sid>` | 被剔除的行（`_src_row,step,step_name,reason`） | 应用清洗后 |
| `tb` | 合并后的科目余额表 | 合并 tb 后 |
| `journal` | 合并后的序时账 | 合并 journal 后 |
| `_recon_*` | 检查结果表（供导出） | 跑检查后 |

---

## 3. 可直接用的 Python 客户端

存成 `ptw.py`，后面所有步骤都用它。**只用标准库**，不依赖 requests。

```python
# ptw.py
import json, urllib.parse, urllib.request

BASE = "http://127.0.0.1:8790"

def _req(method, path, body=None, raw=None, name=None, timeout=3600):
    url = BASE + path
    data = raw if raw is not None else (
        json.dumps(body, ensure_ascii=False).encode("utf-8") if body is not None else None)
    r = urllib.request.Request(url, data=data, method=method)
    if raw is None and body is not None:
        r.add_header("Content-Type", "application/json")
    if name:                                  # 文件名要 URL 编码
        r.add_header("X-Filename", urllib.parse.quote(name))
    with urllib.request.urlopen(r, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))

def get(path):
    return _req("GET", path)

def post(path, body=None, **kw):
    return _req("POST", path, body if body is not None else {}, **kw)

def upload(path):
    with open(path, "rb") as f:
        blob = f.read()
    import os
    return _req("POST", "/api/upload", raw=blob, name=os.path.basename(path))

def q(sql, limit=2000):
    return post("/api/query", {"sql": sql, "limit": limit})
```

---

## 4. 阶段一：导入

### 4.1 上传 + 探明表头

```python
up  = upload(r"D:\客户资料\科目余额表_2026.xlsx")
insp = post("/api/inspect", {"fid": up["file_id"], "target": "tb"})
print(json.dumps(insp, ensure_ascii=False, indent=2))
```

`/api/inspect` 返回大致结构：

```json
{"file": {"file_id": "...", "file_name": "...", "kind": "xlsx", "size": 434176},
 "sheets": [{"sheet": "Table1", "rows": 9082, "cols": 20, "header_row": 2,
             "preview": [["合并方案:多账簿"], ["方案层级","科目编码","科目名称","币别", "..."]]}]}
```

**要点**：
- `header_row` 是系统自动识别的表头行号，**可以直接用**。余额表常见第 2 行（上面一行是「合并方案:xxx」），凭证常见第 1 行。
- 若 `preview` 里明显是标题/空行而非列名，手工指定 `header_row`，从 1 开始数。
- `target` 只有两个合法值：`tb`（科目余额表）、`journal`（序时账）。

### 4.2 导入

```python
res = post("/api/import", {"fid": up["file_id"], "sheet": "Table1",
                          "header_row": 2, "header_rows": 1, "target": "tb"})
sid_tb = res["sid"]        # 后续所有操作都用这个 sid
```

大文件（百万行）会有进度接口 `GET /api/import/progress`，也可以直接等（实测 100 万行约 11 秒）。

✅ 自检：
```python
st = get("/api/state")
for s in st["sheets"]:
    print(s["sid"], s["file_name"], s["sheet_name"], s["n_rows"], s["target"])
    print("   列:", s["columns"])
```
- `n_rows` 与你在 Excel 里看到的行数对得上吗？差很多说明 `header_row` 选错或多 sheet 只导了一张。
- **把 `columns` 原样记下来**，后面写清洗参数和映射必须一字不差。

---

## 5. 阶段二：诊断（决定怎么清洗）

**不要一上来就套默认配方。** 先用 SQL 看清数据长什么样。这是整个流程里最需要判断力的一步，也是"清洗对不对"的分水岭。

```python
# 换成你的 raw 表名
T = "raw_%s" % sid_tb

q(f'SELECT count(*) FROM "{T}"')                                    # 行数
q(f'SELECT * FROM "{T}" LIMIT 5')                                   # 长得什么样
```

### 5.1 七项体检

| # | 要查什么 | 参考 SQL / 判断方法 | 命中后该做什么 |
|---|---|---|---|
| 1 | 空行、重复表头、页脚 | 看 `SELECT * LIMIT 20` 尾部几行 | `drop_empty`、`drop_keyword`（kw 填「合计,小计,制表人,打印日期」） |
| 2 | **汇总层**（余额表里的「多账簿/合并方案/合计」主体） | `SELECT 列, count(*), sum(借方) FROM T GROUP BY 1` — 如果某主体金额 ≈ 其它所有主体之和，它就是汇总层 | 必须删，否则金额翻倍：`drop_keyword` |
| 3 | **账簿层重复**（序时账里 `GB-*`、`副账簿`、`合计` 账簿） | `SELECT 账簿, count(*), sum(借方) FROM T GROUP BY 1` — 逐账簿金额与余额表逐家对比，与汇总层同额的就是重复层 | 必须删：`drop_keyword`（mode=正则，kw=`GB-\|副账簿\|合计`） |
| 4 | **外币原币行重复计本币**（余额表按币别长表：综合本位币行 + 各外币行，**外币行的本币列也有折算值**） | `SELECT 币别, count(*), sum(本期发生本位币借) FROM T GROUP BY 1` | **不要删外币行**（会丢原币）。保持不变，检查阶段按币别分组即可 |
| 5 | **期间范围是否与另一张表一致** | `SELECT 期间, count(*) FROM T GROUP BY 1 ORDER BY 1` | 两边范围不同就记录在案，检查时必须传 `period_from/period_to` |
| 6 | 行数是不是整百万（1000000 / 2000000） | 直接看行数 | 疑似导出被截断。**必须在报告里写明**，会直接影响"是否核对完整"的结论 |
| 7 | 科目层级与分隔符 | 看 `科目名称`/`科目全名` 样例：`应付账款_结算应付_材料款` 用 `_`；`1001-库存现金` 用 `-`；`1001.01` 用 `.` | 决定 `add_l1` 的 `sep` |

### 5.2 判断「一级科目」怎么来

- 有**独立的一级科目列** → 直接映射，不用生成。
- 科目名称是**层级拼接**（`A_B_C`）→ `add_l1`：`{"col":"科目名称","sep":"_","name":"一级科目"}`。
- 科目**编码**带层级（`1001.01`）且要按「主体」删非末级 → `add_acct_path`（会删掉非末级父行，慎用；只在余额表父子同存导致重复时用）。
- ⚠️ **一级科目要的是名称层级**（「应付账款」），不是编码（「2202」）。审计语境下 `acct_l1` 存名称，表里显示才对。

### 5.3 序时账必做的一步：唯一凭证号

检查第二组的「凭证级平衡」需要按整张凭证汇总借贷。原表通常只有「凭证号」，同一凭证号在不同账簿/月份会撞车。

```python
{"type": "add_voucher_key", "enabled": True,
 "params": {"entity": "账簿", "period": "期间", "vno": "凭证号",
            "sep": "-", "name": "唯一凭证号"}}
```

产物形如 `1A公司-06-记-0012`（某段为空自动跳过，不会拼出 `A--001`）。

---

## 6. 阶段三：清洗

### 6.1 两种执行方式

```python
# 单表：先存配置，再应用
post("/api/clean/save", {"sid": sid_tb, "steps": steps_tb})
r = post("/api/clean/apply", {"sid": sid_tb})

# 全部表：所有表配置好之后一次洗完
r = post("/api/clean/apply-all", {})     # 也接受 {"sids": [...]} 只洗指定表
```

`/api/clean/apply` 返回：`total_before`、`total_after`、`dropped`、`steps[]`（每步 before/after/dropped/note/error）、`error`。

> `/api/clean/apply-all` 返回 `{ok, total, ok_n, skipped, failed, seconds, items[]}`。
> 没配动作的表标 `skip`；单表失败不中断其余表，逐表报 `status/error`。

**先试运行再落库**：`POST /api/clean/preview {"sid": sid}` 只算不落库，返回每步删了几行。大表改配置时先跑它。

### 6.2 动作清单（13 个，全部单条 SQL，千万行可跑）

| type | 名称 | 参数 |
|---|---|---|
| `drop_empty` | 删除空行 | `cols`(列数组，空=全部列), `mode`(`all`/`any`) |
| `drop_keyword` | 按关键词删行 | `col`, `kw`(多词用 `,` 分隔), `mode`(`包含`/`全等`/`正则`), `reason` |
| `drop_dup` | 整行去重 | `cols`(空=全部列) |
| `trim_text` | 文本归一 | `cols` |
| `pad_code` | 编码补零 | `col`, `length`(0=按数据自动推断，推不出来就不执行) |
| `fill_down` | 向下填充 | `col` |
| `fill_const` | 填充固定值 | `name`(新列名), `value` |
| `derive_level` | 科目级次推断 | `col`, `name` |
| `add_l1` | 提取一级科目 | `col`, `sep`(默认 `-_/ .:、`), `name` |
| `add_acct_path` | 生成一级科目（删非末级） | `code`, `name`, `entity`, `sep`, `join` |
| `split_aux` | 辅助核算拆分 | `col`, `sep`(默认 `/`), `name1`/`name2`/`name3`(拆出来的新列名) |
| `add_voucher_key` | 生成唯一凭证号 | `entity`, `period`, `vno`, `sep`, `name` |
| `check_balance` | 勾稽校验（只报不删） | `open`, `dr`, `cr`, `end`, `tol` |

步骤是**有序流水线**，`enabled: false` 保留但不执行。列名不存在的动作会被自动跳过并给 warning，不会炸整条流水线。

### 6.3 两个已验证的参考配置

**科目余额表（多账簿导出格式）**
```python
steps_tb = [
  {"id":"s1","type":"drop_empty","enabled":True,"params":{"cols":[],"mode":"all"}},
  {"id":"s2","type":"drop_keyword","enabled":True,
   "params":{"col":"方案层级","kw":"多账簿","mode":"包含","reason":"合并方案汇总层，与各公司明细重复"}},
  {"id":"s3","type":"add_l1","enabled":True,
   "params":{"col":"科目名称","sep":"_","name":"一级科目"}},
]
```

**序时账（凭证导出格式）**
```python
steps_jr = [
  {"id":"s1","type":"drop_empty","enabled":True,"params":{"cols":[],"mode":"all"}},
  {"id":"s2","type":"drop_keyword","enabled":True,
   "params":{"col":"账簿","kw":"GB-|副账簿|合计","mode":"正则","reason":"GB合并层/副账簿/合计行，与主账簿重复"}},
  {"id":"s3","type":"add_l1","enabled":True,
   "params":{"col":"科目全名","sep":"_","name":"一级科目"}},
  {"id":"s4","type":"add_voucher_key","enabled":True,
   "params":{"entity":"账簿","period":"期间","vno":"凭证号","sep":"-","name":"唯一凭证号"}},
]
```

> ⚠️ 这两套是**针对特定导出格式**的，列名和关键词都写死了。套到别的客户数据前，必须按第 5 节的体检结果改列名和关键词，不能照抄。

✅ 自检：
```python
r = post("/api/clean/apply", {"sid": sid})
assert not r["error"], r["error"]
print(r["total_before"], "->", r["total_after"], "剔除", r["dropped"])
for s in r["steps"]:
    print(s["i"], s["name"], s["before"], "->", s["after"], s.get("note") or s.get("error"))
```
- 每步的 before/after 符合预期吗？（比如「删合计行」应该只删掉少量行）
- **剔除了不该删的？** 立刻 `POST /api/clean/reset {"sid": sid}` 回到原始，改配置重来。
- 想看具体删了哪些行：`POST /api/clean/dropped {"sid": sid, "limit": 200}`，会告诉你第几步、按什么原因。

---

## 7. 阶段四：字段映射

```python
post("/api/mapping", {"sid": sid_tb, "target": "tb", "mapping": MAP_TB})
post("/api/mapping", {"sid": sid_jr, "target": "journal", "mapping": MAP_JR})
```

`mapping` 是 `{标准字段: 源列名}`。也可以先让系统猜：`POST /api/suggest {"sid": sid, "target": "tb"}`。

**列名必须与 `GET /api/state` 的 `columns` 一字不差。** 映射到清洗后新增的列（`一级科目`、`唯一凭证号`）也是允许的 —— 但必须先清洗、再映射。

### 7.1 科目余额表（`target: "tb"`）

```python
MAP_TB = {
  "book_name":      "方案层级",          # 账簿名（多账簿导出用它做主体兜底）
  "company_name":   "<公司名称列，如果有>",
  "acct_code":      "科目编码",
  "acct_name":      "科目名称",
  "acct_full":      "科目名称",          # 科目全称（没有独立列就指向名称列）
  "acct_l1":        "一级科目",          # 清洗生成的列
  "currency":       "币别",              # 综合本位币 / USD / EUR …（原币核对必需）
  "open_debit":     "期初本位币(借)",     "open_credit":     "期初本位币(贷)",
  "open_debit_fx":  "期初余额(借)",       "open_credit_fx":  "期初余额(贷)",
  "cur_debit":      "本期发生本位币(借)", "cur_credit":      "本期发生本位币(贷)",
  "cur_debit_fx":   "本期发生(借)",       "cur_credit_fx":   "本期发生(贷)",
  "ytd_debit":      "本年累计本位币(借)", "ytd_credit":      "本年累计本位币(贷)",
  "ytd_debit_fx":   "本年累计(借)",       "ytd_credit_fx":   "本年累计(贷)",
  "end_debit":      "期末本位币(借)",     "end_credit":      "期末本位币(贷)",
  "end_debit_fx":   "期末余额(借)",       "end_credit_fx":   "期末余额(贷)",
}
```

**本币 / 原币要成对映射**：`*_debit` / `*_credit` 是本币，`*_fx` 是原币。没有外币核算就只映射本币那几对。

### 7.2 序时账（`target: "journal"`）

```python
MAP_JR = {
  "vdate":        "日期",
  "period":       "期间",
  "vtype":        "凭证字",
  "vno":          "凭证号",
  "voucher_key":  "唯一凭证号",     # 清洗生成
  "summary":      "摘要",
  "acct_code":    "科目编码",
  "acct_name":    "科目全名",
  "acct_full":    "科目全名",
  "acct_l1":      "一级科目",       # 清洗生成
  "debit":        "借方金额",
  "credit":       "贷方金额",
  "amount_fx":    "原币金额",       # ⚠️ 原币是单列，不是借/贷两列
  "currency":     "币别",
  "book_name":    "账簿",
  "person":       "制单",  "checker": "审核",  "poster": "过账",
  "counterparty": "往来单位",
}
```

⚠️ **序时账原币通常是「原币金额」单列**（不是借原币/贷原币两列）。映射到 `amount_fx` 后，合并时会按借/贷方向自动拆成 `debit_fx` / `credit_fx`。

✅ 自检：
```python
st = get("/api/state")
for s in st["sheets"]:
    print(s["sheet_name"], "映射字段数:", len(s["mapping"]), "target:", s["target"])
```
- `acct_l1` 映射上了吗？**没有它，检查第三组直接报错**（提示回去拆一级科目）。
- `currency` 映射上了吗？没有就只能核本币，原币核对会显示"无数据"。
- 主体：`company_name` → `book_name` → 全归「(未标注主体)」。多公司账套一定要提供其中之一。

---

## 8. 阶段五：合并成标准表

```python
r1 = post("/api/merge", {"target": "tb"})
r2 = post("/api/merge", {"target": "journal"})
```

把所有映射到同一目标的工作表 UNION 成 `tb` / `journal`，**重复执行整表重建**（幂等）。

✅ 自检：
```python
q("SELECT count(*) FROM tb")            # 应 ≈ 各余额表 sheet 行数之和
q("SELECT count(*) FROM journal")
q("SELECT count(*), count(DISTINCT voucher_key) FROM journal")
```
`warnings` 非空要读：常见是"某表没映射主体"、"唯一凭证号缺失，已现场拼接"。

---

## 9. 阶段六：四组数据检查

```python
r = post("/api/recon", {"period_from": "2026-01", "period_to": "2026-12",
                        "tol": 0.01, "tb_base": "ytd"})
```

| 参数 | 说明 |
|---|---|
| `period_from` / `period_to` | 序时账期间过滤。**只影响第三、四组**（第一、二组是自平衡，永远看全量）。传 `"1"`~`"12"` 或 `"2026-01"` 都行，系统会归一化（**不要再自己拼字符串比较**，历史上这里有字典序 bug） |
| `tol` | 容差，默认 0.01 |
| `tb_base` | 余额表取哪一对本币发生额：`ytd`(本年累计，默认) / `cur`(本期发生) / `open`(期初) / `end`(期末) |

### 四组是什么

| 组 | 内容 | 期望 |
|---|---|---|
| 一 | **余额表**按「公司 × 币别」：本期借贷平衡、本年累计借贷平衡、期初+本期借−本期贷=期末 | 全部「平衡」，勾稽差 0 |
| 二 | **序时账**按「公司」：总额平衡 / 月平衡 / **唯一凭证号平衡**（每张凭证借贷相等，只列不平的） | 全部 0 问题 |
| 三 | 两表按「公司 + **一级科目**」比对（分币别） | 差额 0，单边项应只是零发生额科目 |
| 四 | 两表按「公司 + **科目全称**」比对（分币别） | 同上，粒度更细 |

### 结果怎么读

```python
print(json.dumps(r["summary"], ensure_ascii=False, indent=2))
for c in r["checks"]:
    print(c["no"], c["title"], "| ok=", c["ok"], "bad=", c["bad"], "| err=", c.get("error"))
    for p in c.get("parts", []):
        print("   ", p["key"], "total", p["total"], "bad", p["bad"])
```

**判读规则（重要，别误报）**：

| 现象 | 通常意味着 | 要不要当问题 |
|---|---|---|
| 第一组个别公司「本期发生额不平」 | 余额表里父子科目同存（`1002` 和 `1002.01` 都在），一级级汇总时重复 | **真问题**，该用「生成一级科目」按主体删非末级 |
| 第一组外币行"发生额不平" | 借美元、贷人民币是正常业务 | 不算（系统已只判外币行的余额勾稽） |
| 第三/四组大量「仅余额表有」且金额为 0 | 余额表每个科目×币别都有行，序时账没发生 | 不算（系统已把零金额单边判「一致」） |
| **不符率 > 50%** | **口径错了**：汇总层没删 / 重复导入 / 期间错位 / 币别没分 | 停下来回第 5 节重做诊断，不要当"数据问题"报给用户 |
| 总额对得上、个别科目差固定金额 | 真实串户或漏记 | 真问题，按凭证号定位 |
| 第二组 2.3 那张表是空的 | 所有凭证都平衡（表只列不平的） | 正常，导出时会写说明 |

### 导出

```python
e = post("/api/recon/export", {"name": "TB序时账核对"})
print(e["url"])      # /download?f=xxx.xlsx → 用 GET 下载
```

一个工作簿多张表：`结论汇总` + 各组明细，中文表头，数字保持数值型。

✅ 自检（**这是验收标准**）：
1. 第一组：所有公司「平衡」，勾稽差 0。
2. 第二组：总额差 0、每月差 0、不平凭证 0 张。
3. 第三组：两表本币总额差 < 0.01。
4. 第四组：非零金额的「不符」项 = 0。
5. 若 1~2 不满足 → 是**数据本身**的问题（借贷真不平），据实报告，不要"修"数据。
6. 若 3~4 大面积不符 → 是**你的口径**问题，回第 5 节。

---

## 10. 建议的执行顺序（照抄即可）

```
1. 启动服务 → GET /api/modules 校验版本
2. 上传余额表 → inspect → import(target=tb) → 记下 sid_tb 和 columns
3. 上传序时账（可能多个文件）→ inspect → import(target=journal) → 记下各 sid
4. 对每张表跑第 5 节诊断 SQL，写下结论（尤其：汇总层？重复账簿层？期间范围？行数是否整百万？）
5. 按诊断结果写 steps，逐表 clean/preview 看影响行数 → clean/apply
6. 全部表映射 → /api/mapping
7. /api/merge tb、/api/merge journal
8. /api/recon（期间按第 4 步结论传）→ 判读
9. 不符就按第 9 节表格分清"数据问题"还是"口径问题"，口径问题回第 4/5 步
10. /api/recon/export 导出，下载
11. 输出报告（见第 12 节模板）
```

---

## 11. 已知陷阱（都是实战踩过的）

1. **外币原币行的本币列有值**。余额表 9,082 行里 5,804 行是外币原币行，全取会让本币发生额变成 943 亿（正确的是 629 亿）。
   → 正确做法不是删，而是**按币别分组**：`综合本位币` 行比本币，外币行比原币。
2. **汇总层主体**。「多账簿(01)」这类主体金额正好等于 15 家公司之和（314.5 亿 = 314.5 亿），必须删。
3. **账簿层的重复**。序时账里 `GB-*`（集团合并）、`*副账簿`、`合计` 行，与主账簿重复。全导进去金额放大 4.4 倍。
   → 逐个账簿与余额表逐家比对金额，能精确证明谁是重复层。
4. **期间错位**。余额表 1-6 月、凭证到 12 月。检查必须传 `period_to`，否则比的不是同一段时间。
5. **行数整百万**。三卷各恰好 1,000,000 / 1,000,003 行，疑似导出工具上限。**这会影响"核对是否完整"的结论，必须写进报告**。
6. **一级科目要名称不要编码**。`acct_l1` 存「应付账款」，不是「2202」。
7. **期间字符串比较**。`'10' < '6'`，`period BETWEEN '1' AND '6'` 会把 10-12 月算进来。平台已修（内部 `lpad` 归一到 2 位），但**你自己写 SQL 时也要注意**。

---

## 12. 输出报告模板

```markdown
# TB × 序时账 数据检查报告

## 数据
| 文件 | 工作表 | 原始行数 | 清洗后 | 剔除 | 剔除原因 |
|---|---|---|---|---|---|
| ... | ... | ... | ... | ... | 汇总层 / 重复账簿 / 空行合计行 |

## 口径（做了哪些判断，为什么）
1. 汇总层：xxx 是 N 家合计（金额 A = B，精确相等）→ 删除
2. 账簿层：GB-*/副账簿/合计 与主账簿重复（逐家金额比对见附表）→ 删除
3. 期间：余额表 1-6 月，凭证 1-12 月 → 检查限定 1-6 月
4. 币别：余额表保留全部行，综合本位币行比本币、外币行比原币

## 结果
| 组 | 检查数量 | 问题数 | 结论 |
|---|---|---|---|
| 一 余额表自平衡 | 173 | 0 | 通过 |
| 二 序时账平衡 | 322,513 | 0 | 通过 |
| 三 公司+一级科目 | 1,885 | 0 | 通过 |
| 四 公司+科目全称 | 8,496 | 0 | 通过 |

## 待客户确认
- （如）行数整百万，是否导出被截断
- （如）M副账簿 635 亿 与余额表 2.96 亿 差两个数量级
```

---

## 13. 出问题时

| 症状 | 处理 |
|---|---|
| 清洗删多了 | `POST /api/clean/reset {"sid": sid}` 恢复原始，改配置重来。原始表一直都在 |
| 想删掉合并表里的脏行 | `POST /api/delete-rows {"table": "journal", "col": "...", "mode": "等于", "value": "..."}` |
| 合并后行数不对 | `GET /api/state` 看每张表 `stg` 是否 true（没清洗的表会按原始表参与合并） |
| 检查报"需要一级科目列" | 两张表都要映射 `acct_l1`，且清洗里确实生成了这一列 |
| 检查报"没有借方/贷方金额列" | 序时账映射缺 `debit`/`credit` |
| 接口 500 | 看服务控制台输出；`GET /api/ping` 确认进程活着 |
| 想重来一遍 | 换一个空的 `KZ_DATA_DIR` 重启，最干净 |

---

## 14. 接口速查

| 方法 | 路径 | 用途 |
|---|---|---|
| GET | `/api/modules` | 版本 + 板块清单（探活） |
| GET | `/api/state` | 所有工作表、列名、映射、清洗配置、行数 |
| POST | `/api/upload` | 上传（raw body，头 `X-Filename`） |
| POST | `/api/inspect` | `{fid, target}` 探 sheet / 表头行 |
| POST | `/api/import` | `{fid, sheet, header_row, header_rows, target}` |
| GET | `/api/clean/types` | 动作清单 + 参数定义 |
| POST | `/api/clean/save` | `{sid, steps}` 存配置 |
| POST | `/api/clean/preview` | `{sid}` 试运行不落库 |
| POST | `/api/clean/apply` | `{sid}` 落库 |
| POST | `/api/clean/apply-all` | `{}` 或 `{sids}` 批量洗 |
| POST | `/api/clean/dropped` | `{sid, limit}` 看被剔除的行 |
| POST | `/api/clean/reset` | `{sid}` 恢复原始 |
| POST | `/api/mapping` | `{sid, target, mapping}` |
| POST | `/api/suggest` | `{sid, target}` 自动猜映射 |
| POST | `/api/books/preview`、`/api/books/import` | 账簿 → 公司 对照表（多账簿账套用） |
| POST | `/api/merge` | `{target}` |
| POST | `/api/recon` | 四组检查 |
| POST | `/api/recon/export` | `{name}` 导出 Excel |
| POST | `/api/query` | `{sql, limit}` 任意 SQL |
| POST | `/api/export` | `{sql, fmt: "csv"/"xlsx", name}` 导查询结果 |
| GET | `/api/schema` | 标准表字段结构 |

---

**一句话给执行者**：这套东西的判断力全在「诊断」和「判读」两步。清洗动作本身只是 SQL；真正的风险是**把汇总层当明细算**、**把重复账簿算两遍**、**拿不同期间的两张表对比**。这三件事做错，检查出来几千项不符，全是假的。
