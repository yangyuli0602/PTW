# PTW 平台 · AI 执行手册

> 给 **Agent**（Qoder 等本地 AI 编程助手）看的操作手册，不是给人看的产品说明。
> 读完这份就能干活，不需要读源码。
> 业务规则另外四份文档在包根：`版本变更史` / `AI速查_导入清洗` / `AI速查_WP初稿` / `AI提示词_清洗达标`。

---

## 0. 红线（先看这条，违反会毁掉交付件）

**`程序\` 下的代码是交付件，不是工作区。**

- 遇到报错先查第 6 节的故障表，**不要修改 `程序\` 下的任何 .py 文件**
- 确实需要改的，把现象和建议写进 `issues.md`，不要动代码
- 不要 `pip install`：这台机器**没有外网**，装不上，且会破坏内置 runtime

原因：改了不由版本控制跟踪，下次替换包会被覆盖；更糟的是改错了底稿数字会变，而事后无法追溯是谁改的。

---

## 1. 环境

| 项 | 值 |
|---|---|
| Python | `程序\runtime\python\python.exe`（内置，**不要用系统 python**） |
| 依赖 | 全部内置（duckdb / openpyxl），**不需要也不允许 pip** |
| 数据目录 | `data\`（默认）。要用别处加 `--data-dir` |
| 输出目录 | `data\bkd_out\`（生成的 Excel） |
| 网络 | **不通**。所有能力本地完成 |

启动界面（只在需要人看结果时用）：双击 `启动PTW平台.bat`，浏览器打开 http://127.0.0.1:8787/

**但 Agent 通常不需要开界面**，用下面的命令行。

---

## 2. 命令行入口（推荐）

所有命令都从包根执行，输出单行 JSON。

```bat
程序\runtime\python\python.exe 程序\ptw.py <命令> [参数]
```

成功 `{"ok": true, ...}` 退出码 0；失败 `{"ok": false, "code": "...", "error": "...", "next": "..."}` 退出码 1。
**判断成败看 `ok` 字段，不要只解析中文 error。**

通用参数：

| 参数 | 说明 |
|---|---|
| `--data-dir DIR` | 数据目录，默认 `程序\..\data`。**做验证时务必指到临时目录** |
| `--pretty` | 格式化 JSON（人看时用，Agent 不用） |

---

## 3. 任务一：把 Excel 变成可用的账套

### 3.1 先看现状

```bat
程序\runtime\python\python.exe 程序\ptw.py state
```

返回关键字段：

```json
{"ok":true, "rows":{"tb":455,"journal":0,"flow":0},
 "merged":{"tb":true,"journal":false,"flow":false},
 "ready":{"cash":true,"exp":false},
 "hint":"可以直接 ptw.py cash"}
```

**`ready.cash` 为 true 才能生成货币资金底稿；`ready.exp` 为 true 才能生成费用底稿。**

### 3.2 导入

```bat
程序\runtime\python\python.exe 程序\ptw.py import "D:\资料\科目余额表.xlsx"

:: 已经清洗成标准表时（kanzhang/v1 数据包：zip 内含 tb.parquet 与 journal.parquet），
:: 不必再绕回 xlsx，直接灌库：
程序\runtime\python\python.exe 程序\ptw.py importpack --file 脱敏数据包.zip
```

它会自动做四件事：识别表头行 → 判定数据类别（tb / journal / flow）→ 套默认清洗配方 → 合并成标准表。

`importpack` 按列名映射：包里有的列按标准表类型转换，没有的补 NULL，并在返回的 `missing` 里列出来。

返回：

```json
{"ok":true,
 "imported":[{"sheet":"源表","target":"tb","rows":456,"score":9,"mapped":9}],
 "skipped":[],
 "merged":{"tb":{"rows":455,"sheets":1}}}
```

判读要点：

- `target` 是自动判定的类别。判错了用 `--target tb|journal|flow` 强制指定
- `mapped` 是字段映射命中数。tb/journal 一般 8~10 个；**低于 5 个说明表头不规范**，看 4.2
- `merged.tb.rows` 比 `imported.rows` 少几行是正常的（清洗删了合计行/空行）
- `skipped` 非空要看原因：空表、格式不支持、清洗失败

**同一个文件重复导入不会产生两份数据**（按 md5 去重）。

---

## 4. 任务二：生成底稿

### 4.1 货币资金

```bat
程序\runtime\python\python.exe 程序\ptw.py cash --date 2026-10-31 --customer "某某公司"
```

参数：

| 参数 | 说明 |
|---|---|
| `--date` | 资产负债表日。**必填**，会写进 Lead，也是截止性测试的基准日 |
| `--customer` | 客户名，可空 |
| `--prior 上年BKD.xlsx` | 上年底稿。**给了才有上年数**，不给则上年数为空（BKD 明细区不涂色，缺什么看 `gapList`） |
| `--rate 主体=7.1` | 折算率，可重复。只对**本位币是外币**的主体需要 |
| `--no-cutoff` | 不做截止性测试（默认做） |
| `--no-oar` | 不生成 OAR（默认一起生成） |
| `--restricted 账号=金额` | 受限资金，可重复。**这是人工判断结果，工具不会自己猜**；不填则 BKD 受限列留空，由人在 Excel 里补。同一账号跨科目时要写 `一级科目|账号=金额`，否则报 `RESTRICTED_AMBIGUOUS` |
| `--strict` | 模板结构与上次不一致时不落盘（只报差异） |

一次产出两份（OAR 是附加产物，失败不影响主底稿）：

```json
{"ok":true,
 "main":{"file":"货币资金BKD_20261031_101500.xlsx",
         "rows":162,"gaps":11,"amount":12215000.0,
         "stats":{"matched":140,"newCount":5,"missCount":2,"diff":0},
         "gapList":[{"company":"...","field":"是否函证","row":37,"col":9,"reason":"..."}]},
 "oar":{"file":"OAR货币资金_....xlsx","totalQian":12215.0},
 "warnings":[],
 "next":"gaps 为 0 则底稿可直接用；gaps > 0 表示有格子标黄待人工确认"}
```

**验收三条**（都满足才算做对）：

1. `main.rows` > 0 —— 有账户进去
2. `main.stats.diff` 接近 0 —— 本年合计与上年合计的差额。**不用 ROUND 调平，差额必须有解释**
3. `main.gaps` 是"待人工确认的格子数"，**不是错误**。BKD 明细区不涂色，所以它不在表格上显示，逐格清单在 `gapList`（页签/行/列/字段/原因）和「检查」表末尾

### 4.2 销管费用

```bat
程序\runtime\python\python.exe 程序\ptw.py exp --from 1 --to 9 --pre-end 6
```

```json
{"ok":true,"main":{"file":"费用BKD_....xlsx","vc":175,"vd":120,"amount":70311758.9,
 "bySegment":{"preEnd":6,"pre":34125575.57,"final":36186183.33,"total":70311758.9},
 "leadLinks":{"销售费用":{"sheet":"VC.00 销售费用BKD","sumRow":205},
              "管理费用":{"sheet":"VD.00 管理费用BKD","sumRow":150}},
 "oar":{"status":"ok","sheet":"VC&VD OAR","items":78,"analyzed":20,
        "pending":["管理费用_中介费"],"evidenceFile":"bkd_out/OAR_evidence_费用.json"}}}
```

**`vc` = 销售费用行数，`vd` = 管理费用行数。**

| 参数 | 说明 |
|---|---|
| `--prior 上年费用底稿.xlsx` | 有它才回捞「上期审定数」（K 列），否则整列标黄 |
| `--pre-end 6` | 预审段结束月份。**只影响 `bySegment` 的分段金额，不改 BKD 输出** |

`bySegment` 的两段金额是给 TOD 用的（预提冲销必须按段配对，跨段配会把一段的净额挪到另一段）。费用 BKD 取的是序时账本币，**不涉及外币折算**——汇率只在 TOD 池按段应用。

### 输出是一个文件、四个关键页

| 页 | 内容 |
|---|---|
| `Uexp.00 Lead` | 表1 四列（账面/账表调整/审计调整/上期审定）全部链接到 BKD **合计行**；客户名称在 C2，资产负债表日在 C3 |
| `VC.00 销售费用BKD` | 明细 + 合计 + Rx/A3/Diff；数据区每行格式统一，空白格由模板的条件格式上色 |
| `VD.00 管理费用BKD` | 同上（模板这张的合计行原本没有「合计」两字，已补） |
| `VC&VD OAR` | **现场生成，不抄上年模板**：上半是审定数明细（SUMIF 链接 BKD 的 J/K 列），下半是变动分析 |

`leadLinks` 是 Lead 实际写到的合计行。**插行后一定要看这个**：模板原本写死 `E34`，明细一多合计行就下移，写死会静默少算最后一行的金额。

`oar.analyzed` = 生成的分析段条数，`oar.noReason` = 序时账里找不到对应摘要、只能留「未识别」的科目。

### OAR 页长什么样

**一、明细表**：一行一个费用科目，只取**审定数**（本期 / 上期 / 变动 / 变动率 / 占本期比重 / 是否需进一步调查）。到期审定数走
`SUMIF('VC.00 销售费用BKD'!$D:$D, $D9, 'VC.00 销售费用BKD'!$J:$J) + SUMIF(VD…)`，
BKD 一改这里立刻跟着变，不会有两处数字打架。

**二、变动分析**：按**变动金额**排序（不是本期金额），一段一个科目。

选中规则：`|变动金额| >= 门槛` **且** `|变动率| >= 10%`（两者都不满足才排除）。
门槛默认取 SAD（Lead!C5），没填就 TE×10%、再没有就 1 万。可用 `--oar-chg-min` / `--oar-chg-pct` 调。

> 为什么不按本期金额：本期花 1000 万、上期也 1000 万，说明不了任何问题；
> 本期只花 100 万、上期 20 万，才是要问清楚的。

每条结论是一整条公式，前半段自动、后半段是 AI 写的文字：

```
="「"&$C16&"」本期较上期"&IF(D16>0,"增加","减少")&TEXT(ABS(D16),"#,##0.00")
  &" 元（"&TEXT(ABS(E16),"0.00%")&"），"&"<AI 写的变动原因>"
```

### OAR 结论由 AI 写，工具不套模板

```
① 生成并导出证据
   ptw.py exp --from 1 --to 9 --prior 上年费用底稿.xlsx
            --oar-evidence tmp/OAR_evidence_费用.json

② AI 读证据，写 {"科目名称": "结论文字"} 存成 JSON
   证据长这样（每个科目一块）：
     {"name": "管理费用_中介费", "now": 13048099, "prior": 2870582,
      "change": 10177518, "changePct": 3.546,
      "groups": [{"key": "确认上市费用", "amt": 11482419, "rows": 1, "share": 0.88},
                 {"key": "北京弘讯和科技有限公司", "amt": 496040, "rows": 1, "share": 0.038}]}

③ 回填
   ptw.py oar-fill --file 费用BKD_xxx.xlsx --conclusions conclusions.json
```

`--oar-evidence` 不传也会导出到 `bkd_out/OAR_evidence_费用.json`（界面生成同样会导）。

**工具只做两件事**：算金额、按「对方单位 / 业务类型」分组摆证据。
**「为什么变」不由工具判断** —— 它只写占位符 `【待AI撰写】` 并标黄；
AI 回填后底色转白。没回填的科目在 `oar.pending` 里列着，交底稿前必须清掉。

### 分组口径（踩过的坑）

- **摘要含「结转」不等于结转损益**。「结转----确认上市费用」是 1,148 万的**借方真实费用**。
  正确做法只有一条：`借 −（贷 − 结转损益贷方）`，与 `read_expense` 完全一致。
  按摘要关键字整行剔除会让分组金额少一半（实测 8.69M vs 17.31M）。
- 对方单位从**摘要**抽：真实账套 `counterparty` / `counterparty_name` 填充率是 0。
  支持「与「XX公司」发生的…」「支付给XX银行…分行」「…工作室（个体工商户）」等写法。
- 分组键优先用对方单位；抽不到单位时用业务类型，并跳过「1月/2月/8/428#」这类片段，
  否则股权激励会被拆成 10 组「1月」「2月」而不是一组「股权激励」。

### 4.3 校验与配置（生成会自动带，这三条命令用来看详情）

```bat
程序\runtime\python\python.exe 程序\ptw.py fingerprint --subject cash
程序\runtime\python\python.exe 程序\ptw.py baseline --subject cash
程序\runtime\python\python.exe 程序\ptw.py project show
程序\runtime\python\python.exe 程序\ptw.py project export --customer 某某 --period 2026-09-30
程序\runtime\python\python.exe 程序\ptw.py project import --file 客户配置_某某_2026-09-30.ptwpack
```

每次 `cash` / `exp` 生成后，返回里会多一个 `checks` 字段：

```json
"checks": {
  "fingerprint": {"blocking": 0, "saved": true, "diffs": []},
  "invariants": {"summary": {"pass": 8, "fail": 0, "skip": 2, "total": 10},
                 "items": [{"id":"cash_sum","status":"pass","detail":"..."}]},
  "baseline": {"summary": "与基准完全一致"}
}
```

三条判读规则：

| 字段 | 怎么看 |
|---|---|
| `invariants.fail` | **大于 0 就别用这份底稿**。看 `items` 里 `status=fail` 的 `id` 和 `detail` |
| `invariants.skip` | 正常。是"本版算不了"的项（需要 Excel 计算、或本版未实现），不是失败 |
| `fingerprint.blocking` | 大于 0 表示模板结构和上次不一样。**先人工复核再往下走** |
| `baseline.summary` | 「首次运行，已建立基准」= 没有基准可比；「与基准完全一致」= 和上次跑的数一样 |
| `main.teSource` | `default` 表示模板没填 TE，阈值按 100,000 算，`warnings` 里会有 `TE_NOT_READ` |
| `main.restricted.mode` | `disclosure` = 受限金额已进 Lead 披露行；`reclass` = 已做重分类调整行；`none` = 有受限但 Lead 无处可落（不变量会判 fail） |
| `main.derived.*.status` | `ok` 已填 / `skip` 模板没有或无事可做 / `partial` 有内容没写全 / `manual` 需人工 / `error` 出错 |
| `main.leadRows` | Lead 实际写到了哪几行（表1/表2 三科目 + 合计 + 境外 + 日期）。**换模板后第一眼看这里**：V7 是 37/38/39 与 62/63/64，别的模板会不一样，这是正常的 |

**`--strict`**：`cash --strict` 时，指纹有阻塞级漂移就**不落盘**，只报差异。改代码后想确认"输出没变"时用它：

```bat
程序\runtime\python\python.exe 程序\ptw.py --data-dir tmp\_v cash --date 2026-09-30 --strict
```

### 4.4 单独跑 OAR / 数据检查

```bat
程序\runtime\python\python.exe 程序\ptw.py oar --date 2026-10-31 --prior 上年.xlsx
程序\runtime\python\python.exe 程序\ptw.py check --from 1 --to 6
```

`check` 跑五组检查，返回 `summary` 和 `failed`。**`failed` 里 `error` 非空的项通常是数据没准备好，不是真的不平**——先解决数据，再看差额。

---

## 5. 任务三：AI 判断写回（可选，但这是"判断留痕"的核心）

平台自己不做职业判断（性质、是否函证），这些格子留空。Agent 可以起草，**但只能写 draft，不能替人确认**。

```
GET  /api/judge/plan?scope=bkd.c00&balance_date=2026-10-31
      → items[]：每项含 jid、field、enum（允许取值）、reason
  ↓ Agent 判断
POST /api/judge/apply   {"items":[{"jid":"...","value":"是","basis":"...","confidence":0.9}]}
      → 写 draft，底稿里仍然留空（未确认前不进门）
  ↓ 人复核
POST /api/judge/review  {"jid":"...","action":"confirm"}
```

`apply` 的硬性校验（不合格整条打回，不会静默丢弃）：

| 校验 | 不通过的后果 |
|---|---|
| `jid` 必须来自本批 plan | `不在本批 plan 内` |
| `value` 必须在 `enum` 内 | `取值不在允许范围` |
| `basis` ≥ 8 字 | `依据太短` |
| `confidence` < 0.6 | `证据不足不写值` |

**注意**：这些是 HTTP 接口，需要先起服务。命令行版本的判定接口还没有，如果只是生成底稿不需要它。

---

## 6. 故障表（按 code 查）

| code | 含义 | 怎么办 |
|---|---|---|
| `TB_EMPTY` | 科目余额表还没合并 | 先 `ptw.py import <余额表>` |
| `NO_CASH_ACCT` | tb 里没有货币资金三个一级科目 | 清洗时加「提取一级科目」步骤；或检查 `acct_l1` 是否识别成 库存现金/银行存款/其他货币资金 |
| `NO_EXP_DATA` | 序时账里没有销售/管理费用发生额 | 检查 `--from/--to` 期间；确认序时账已导入 |
| `TPL_MISSING` | 底稿模板没了 | 从交付包把 `data\bkd_tpl\` 整个拷回来 |
| `FILE_NOT_FOUND` | 文件路径不对 | 路径含空格/中文要加引号 |
| `NO_SHEET_IMPORTED` | 一个工作表都没导进去 | 看返回的 `skipped[].why`；多半是表头行识别错，用 GUI 手动定表头行 |
| `BAD_RATE` | 折算率写成别的格式 | `--rate 主体名=7.1` |
| `FP_MISMATCH` | `--strict` 下模板结构与上次不一致 | 先 `ptw.py fingerprint` 看 diff；确认无误去掉 `--strict` |
| `BAD_RESTRICTED` | 受限金额写成别的格式 | `--restricted 账号=金额`，账号照 BKD 里的一字不差 |
| `LEAD_ROWS_DEFAULT`（warning，不是错误码） | Lead 里没找全「库存现金/银行存款/其他货币资金」，已按 V7 的 37/38/39 兜底 | 看 `main.leadRows`；若这份底稿不是 V7 结构，Lead 数字会对错行，须核对后再交付 |
| `PACK_BAD` / `PACK_EMPTY` | 数据包格式不对或里面没有 tb/journal | 确认是 kanzhang/v1（zip 内有 tb.parquet） |
| `DB_LOCKED` | 平台正开着同一个数据目录，DuckDB 文件被独占 | 先关平台再跑 CLI，或 `--data-dir` 指向别的目录 |
| `RESTRICTED_AMBIGUOUS` | 同一账号同时挂在银行存款和其他货币资金下 | 加上一级科目：`--restricted "其他货币资金|账号=金额"` |
| `PRIOR_UNUSABLE` | `exp --prior` 的文件解析不出上期审定数 | 确认是上年**费用**底稿（含 VC.00/VD.00 BKD 与「上期审定数」列） |
| `FP_MISSING` | 还没生成过，没有指纹 | 先跑一次 `cash` 或 `exp` |
| `BASE_MISSING` | 还没生成过，没有基准 | 先跑一次生成，会自动建立 |
| `UNKNOWN` | 未归类的异常 | 看 `error` 原文；stderr 里有完整 traceback |

### 几个"看起来像 bug 其实不是"

| 现象 | 真相 |
|---|---|
| 费用底稿 K 列整列标黄 | 上期审定数本版无数据源，留空是**设计如此**（费用底稿仍标黄，只有 BKD 明细区不涂色） |
| `gaps` 几十个 | 正常。性质/函证/受限资金属职业判断，一律留空。BKD 明细区按设计**不涂任何底色**，看「检查」表末尾的清单 |
| `diff` 不是 0 但很小 | 浮点噪声，±5e-7 已归零；超出就得解释 |
| 折算人民币列 = 本币列 | 折算率全填 1 时的正常表现 |
| `merged.tb.rows` 比导入行数少 | 清洗删掉了合计行/空行 |
| 银行流水 `flow` 表为空 | 只有导入网银流水后才有，不是必须 |
| 不变量里有 1 条 `skip` | 是 `cash_lead_tie`：需要 Excel 算完公式的值，openpyxl 写出的是公式串没有缓存值 |
| `derived.c02` 是 `manual` | **正常**。是否存在调节项必须有客户对账单才能判，工具下结论属 D 级禁止 |
| `derived.c00b` 是 `skip` | BKD 受限资金列全空，无明细可填；先填受限列再重跑 |
| `derived.c02a` 只有 1 行 CNY | **正常**。外币测算只列「本期有余额」的币种，模板预置的 USD/JPY 行但账上没金额就不出现（零余额的账户仍在 BKD 里）。C.02a 的币种写 ISO 代码，与 BKD「币种」列同一套，SUMIF 才对得上 |
| C.02a 的汇率空着、差异列也是空 | 该币种 PBC 当日没有公布中间价（如巴西雷亚尔）。**不能编汇率**，留空等人工填；差异列用 `=IF(D="")` 保护，不会算出"负的账面金额" |
| C.02a 表内没有格线，只有外一圈 | 设计如此。表内画满格子会显得像两张表拼起来 |
| C.02a 的「账面本位币金额」等于「折合本位币」 | F 列取的是 BKD 的「期末账面金额（人民币）」，不是「本币」。海外主体（香港/新加坡）本位币是美元，**必须先把 BKD 折算率改成美元汇率**，否则 F 列还是美元，跟 E 列（人民币）比就会差出上千万 |
| BKD 多出「折算率 / 期末账面金额（人民币）」两列（在原币、本币之后） | 系统插入的两列，位置固定在本币列右侧（原币 → 本币 → 折算率 → 人民币）。插列时会一并搬走列宽、修合并区、复制样式，所以表头和横幅是连续的；看到断开说明模板被手工改过，重新生成即可 |
| BKD 最右边多出「受限资金」一列 | 系统新增列，接在 Notes 之后，人工填写。空 = 不受限 |
| C.01b 的「对账单金额」整列标黄 | 那是客户资料，不是账面能派生的；拿不到就留空，填 0 会造出假差异 |
| 第二次跑报「与基准完全一致」 | 同参数同数据，本来就该一致；数据换了才该有差异 |

---

## 7. 做验证时的一条硬规矩

**任何测试都不要写进 `data\`**，那是用户的真实账套。

```bat
:: 正确：临时目录 + 拷模板
mkdir tmp\_verify\bkd_tpl
copy data\bkd_tpl\*.xlsx tmp\_verify\bkd_tpl\
程序\runtime\python\python.exe 程序\ptw.py --data-dir tmp\_verify cash --date 2026-10-31
```

曾经有一次测试把检查结果表和底稿写进了真实 `data\`，用户打开平台看到一堆自己没跑过的数据，无法分辨哪些是自己的——这是最不能接受的事。

---

## 8. 最小可行流程（照抄即可）

```bat
:: 1. 看数据
程序\runtime\python\python.exe 程序\ptw.py state

:: 2. 没有数据就导入（一类文件一条命令）
程序\runtime\python\python.exe 程序\ptw.py import "D:\资料\科目余额表.xlsx"
程序\runtime\python\python.exe 程序\ptw.py import "D:\资料\序时账.xlsx"

:: 3. 再确认一次
程序\runtime\python\python.exe 程序\ptw.py state

:: 4. 生成货币资金底稿（有上年底稿就加 --prior）
程序\runtime\python\python.exe 程序\ptw.py cash --date 2026-10-31 --customer "某某公司"

:: 5. 生成费用底稿
程序\runtime\python\python.exe 程序\ptw.py exp --from 1 --to 6

:: 6. 看校验结果（生成时已带，这里只是复核）
::    checks.invariants.summary.fail 必须为 0
::    checks.fingerprint.blocking 大于 0 时要人工看
:: 7. 交付：data\bkd_out\ 下的两个 xlsx
```

每步都看 `ok` 字段。失败就按 `code` 查第 6 节。
