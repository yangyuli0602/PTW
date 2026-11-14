# AI 速查 · 导入清洗检查（压缩版）

> 原文：`AI作业指南_导入清洗检查.md`。本版只删背景解释与重复示例，
> **规则、参数、判读、陷阱一条未删**；有歧义回原文。
> 对象：执行清洗的 AI。替代原文喂入，约省 60% token。

## 0. 铁律（违反毁数据）
1. 不改 `程序/` 任何代码，只走 HTTP 接口。
2. 不删 `程序/../data/`（账套库、上传件、导出件都在里面）。
3. raw 只读：清洗写 `stg_<sid>`，`raw_<sid>` 不动，`clean/reset` 永远可回退。
4. 不猜列名：必须与 `GET /api/state` 的 `columns` 一字不差（含全角括号、空格）。
5. 每步后自检，不通过不进下一步。

## 1. 启动
```bat
set KZ_PORT=8790
set KZ_NO_BROWSER=1
set KZ_DATA_DIR=D:\ptw_work\data
程序\runtime\python\python.exe 程序\server.py
```
- 不设 `KZ_PORT` 则从 8787 起找空闲，实际端口看启动横幅。
- 不设 `KZ_DATA_DIR` 写到 `程序\../data`；换账套/隔离测试必设。横幅里的数据目录不反映该变量。
- 只监听 127.0.0.1。探活 `GET /api/modules` 应返回 version + 3 个板块。

## 2. 客户端（只用标准库，存为 ptw.py）
```python
import json, os, urllib.parse, urllib.request
BASE = "http://127.0.0.1:8790"
def _req(m, p, body=None, raw=None, name=None, timeout=3600):
    d = raw if raw is not None else (json.dumps(body, ensure_ascii=False).encode() if body is not None else None)
    r = urllib.request.Request(BASE + p, data=d, method=m)
    if raw is None and body is not None: r.add_header("Content-Type", "application/json")
    if name: r.add_header("X-Filename", urllib.parse.quote(name))
    with urllib.request.urlopen(r, timeout=timeout) as resp:
        return json.loads(resp.read().decode())
def get(p): return _req("GET", p)
def post(p, b=None, **kw): return _req("POST", p, b if b is not None else {}, **kw)
def upload(path):
    return _req("POST", "/api/upload", raw=open(path, "rb").read(), name=os.path.basename(path))
def q(sql, limit=2000): return post("/api/query", {"sql": sql, "limit": limit})
```

## 3. 数据流与表名
`upload → inspect → import(原始表) → clean/save+apply(或 apply-all) → mapping → merge → recon → recon/export`

| 表名 | 含义 | 何时有 |
|---|---|---|
| `raw_<sid>` | 原始表，带内部列 `_src_row` | 导入后 |
| `stg_<sid>` | 清洗后 | 应用清洗后 |
| `_dropped_<sid>` | 被剔除行（`_src_row,step,step_name,reason`） | 应用清洗后 |
| `tb` / `journal` | 合并后的标准表 | merge 后 |
| `_recon_*` | 检查结果表 | 检查后 |

## 4. 导入
```python
up = upload(r"D:\客户资料\科目余额表_2026.xlsx")
insp = post("/api/inspect", {"fid": up["file_id"], "target": "tb"})
res = post("/api/import", {"fid": up["file_id"], "sheet": "Table1",
                          "header_row": 2, "header_rows": 1, "target": "tb"})
sid_tb = res["sid"]
```
- `target` 只有 `tb` / `journal` 两个合法值。
- `header_row` 系统已识别，通常可直接用；preview 里若是标题/空行则手工指定（从 1 数）。两层表头 `header_rows=2`。
- 自检：`GET /api/state` 核对 `n_rows` 与 Excel 行数；**把 `columns` 原样记下**。

## 5. 诊断（最需判断力的一步，不要直接套默认配方）
先 `q(f'SELECT * FROM "{T}" LIMIT 20')` 看长相，再按下表逐项体检。

| # | 查什么 | 怎么判 | 命中后 |
|---|---|---|---|
| 1 | 空行/重复表头/页脚 | 看首尾几行 | `drop_empty`、`drop_keyword` |
| 2 | 汇总层（多账簿/合并方案/合计主体） | 按主体 group by 求和，某主体 ≈ 其余之和 | **必须删**，否则金额翻倍 |
| 3 | 重复账簿层（`GB-*`、`副账簿`、`合计`） | 逐账簿金额与余额表逐家对比 | **必须删**（`mode=正则`） |
| 4 | 外币原币行（本币列也有折算值） | 按币别 group by | **不要删**（会丢原币），检查按币别分组 |
| 5 | 两表期间范围是否一致 | 按期间 group by | 不一致则 recon 必须传 `period_from/to` |
| 6 | 行数是否整百万 | 看行数 | 疑似导出截断，**必须写进报告** |
| 7 | 科目层级分隔符（`_` / `-` / `.`） | 看科目名称样例 | 决定 `add_l1` 的 `sep` |

一级科目怎么来：有独立列→直接映射；名称层级拼接→`add_l1`；编码带层级且父子同存→`add_acct_path`（会删非末级父行，慎用）。
**`acct_l1` 要名称（应付账款），不要编码（2202）。**

序时账必做唯一凭证号（凭证级平衡按整凭证汇总，凭证号跨账簿/月份会撞）：
```python
{"type":"add_voucher_key","enabled":True,
 "params":{"entity":"账簿","period":"期间","vno":"凭证号","sep":"-","name":"唯一凭证号"}}
```

## 6. 清洗
```python
post("/api/clean/save", {"sid": sid, "steps": steps})
r = post("/api/clean/apply", {"sid": sid})        # 或 post("/api/clean/apply-all", {})
```
- 先 `clean/preview` 试运行看每步删几行，再 apply。
- apply 返回 `total_before/total_after/dropped/steps[]/error`；apply-all 返回 `{ok,ok_n,skipped,failed,items[]}`，单表失败不中断。
- 步骤有序，`enabled:false` 保留不执行；列名不存在的动作自动跳过并 warning。

13 个动作（全部单条 SQL）：

| type | 参数 |
|---|---|
| `drop_empty` | `cols`(空=全部), `mode`(all/any) |
| `drop_keyword` | `col`, `kw`(逗号分隔), `mode`(包含/全等/正则), `reason` |
| `drop_dup` | `cols` |
| `trim_text` | `cols` |
| `pad_code` | `col`, `length`(0=自动推断) |
| `fill_down` | `col` |
| `fill_const` | `name`, `value` |
| `derive_level` | `col`, `name` |
| `add_l1` | `col`, `sep`, `name` |
| `add_acct_path` | `code`, `name`, `entity`, `sep`, `join`（另有 `amount/parent_tol/aux/detail_first/group_cols`，见原文） |
| `split_aux` | `col`, `sep`, `strip` |
| `add_voucher_key` | `entity`, `period`, `vno`, `sep`, `name` |
| `check_balance` | `open`, `dr`, `cr`, `end`, `tol`（只报不删） |

参考配置（**列名与关键词按第 5 节诊断结果改，不能照抄**）：
- 余额表：`drop_empty` → `drop_keyword`(方案层级含「多账簿」，reason=汇总层) → `add_l1`(科目名称, sep=`_`)
- 序时账：`drop_empty` → `drop_keyword`(账簿 正则 `GB-|副账簿|合计`) → `add_l1`(科目全名) → `add_voucher_key`

自检：每步 before/after 是否符合预期；删错了立刻 `clean/reset`；看删了哪些行用 `clean/dropped`。

## 7. 映射
`mapping` = `{标准字段: 源列名}`，可先 `POST /api/suggest` 让系统猜。先清洗再映射（才能映射到生成列）。

- **tb**：主体 `company_name` 或 `book_name`（至少一个）；`acct_code/acct_name/acct_full/acct_l1/currency`；金额四组 `open/cur/ytd/end` × `debit/credit`，本币一对 + 原币 `_fx` 一对，没有外币就只映射本币。
- **journal**：`vdate/period/vtype/vno/voucher_key/summary/acct_code/acct_name/acct_full/acct_l1/debit/credit/amount_fx/currency/book_name`，可选 `person/checker/poster/counterparty`。
- 注意：**序时账原币是单列 `amount_fx`**，不是借/贷两列，合并时按方向自动拆 `debit_fx/credit_fx`。
- 自检：`acct_l1` 必须映射（缺了第三组直接报错）；`currency` 缺了只能核本币。
- 主体兜底顺序：`company_name` → `book_name` → 全归「(未标注主体)」。多公司账套必须提供其一。
- 完整动作与参数定义以 `GET /api/clean/types` 为准（另有 `drop_subtotal`、`drop_empty_key` 等选行动作）。

## 8. 合并
`post("/api/merge", {"target":"tb"})`、同法 journal。同目标表 UNION，重复执行幂等整表重建。
自检：`tb`/`journal` 行数 ≈ 各 sheet 之和；`count(DISTINCT voucher_key)`；`warnings` 必读。

## 9. 四组检查
```python
r = post("/api/recon", {"period_from":"2026-01","period_to":"2026-12","tol":0.01,"tb_base":"ytd"})
```
- `period_from/to` **只影响第三、四组**（一、二组永远全量自平衡）。传 `1`~`12` 或 `2026-01` 均可，系统归一化，**不要自己拼字符串比较**。
- `tb_base`：`ytd`(默认)/`cur`/`open`/`end`。

| 组 | 内容 | 期望 |
|---|---|---|
| 一 | 余额表按公司×币别：本期/累计借贷平衡、期初+借−贷=期末 | 全平衡，勾稽差 0 |
| 二 | 序时账按公司：总额/月/唯一凭证号平衡（只列不平凭证） | 0 问题 |
| 三 | 两表按公司+一级科目（分币别） | 差额 0 |
| 四 | 两表按公司+科目全称（分币别） | 差额 0 |

判读（别误报）：

| 现象 | 结论 |
|---|---|
| 一组个别公司本期不平 | 父子科目同存重复汇总，**真问题**，用 `add_acct_path` 删非末级 |
| 一组外币行发生额不平 | 借美元贷人民币正常，系统只判外币行余额勾稽 |
| 三/四组大量「仅余额表有」且金额 0 | 正常，零金额单边已判一致 |
| **不符率 > 50%** | **口径错了**：汇总层没删/重复导入/期间错位/币别没分，回第 5 节 |
| 总额对、个别科目差固定金额 | 真实串户或漏记，按凭证号定位 |
| 二组凭证表为空 | 全部平衡，正常 |

验收：一组全平衡；二组总额/每月/不平凭证均 0；三组本币总额差 <0.01；四组非零不符 =0。
1~2 不满足是**数据本身**问题，据实报告不要修数据；3~4 大面积不符是**你的口径**问题。

导出：`post("/api/recon/export", {"name":"TB序时账核对"})` → 按返回 url 下载。

## 10. 执行顺序
启动探活 → 上传+inspect+import（记 sid 与 columns）→ 逐表诊断并写下结论 → 写 steps 先 preview 再 apply → 全部映射 → merge tb/journal → recon（期间按诊断结论传）→ 按第 9 节区分数据问题/口径问题 → 导出 → 出报告。

## 11. 陷阱（实战踩过）
1. 外币原币行本币列有值，全取会让本币虚增数十倍；做法是按币别分组，不是删。
2. 汇总层主体金额恰等于各公司之和，必须删。
3. `GB-*`/`副账簿`/`合计` 与主账簿重复，全导会放大数倍；用逐家金额比对证明。
4. 两表期间不同必须传 `period_to`，否则比的不是同一段时间。
5. 行数整百万疑似截断，影响"核对是否完整"的结论，必须写进报告。
6. `acct_l1` 存名称不存编码。
7. 期间字符串比较 `'10'<'6'`，平台已 `lpad` 归一，自己写 SQL 同样要注意。

## 12. 报告必须含
数据表（文件/工作表/原始行数/清洗后/剔除/剔除原因）→ 口径（每个判断及依据：汇总层为何删、账簿层为何删、期间如何限定、币别如何处理）→ 四组结果（检查组数/总行数/问题数/结论）→ 待客户确认（截断嫌疑、数量级异常等）。

## 13. 出问题
删多了 `clean/reset`；合并表删脏行 `POST /api/delete-rows`；合并行数不对看各表 `stg` 是否 true；报缺一级科目=两表都要映射且清洗确实生成；报缺借贷=序时账映射缺 `debit/credit`；500 看控制台 + `GET /api/ping`；想重来换空 `KZ_DATA_DIR` 重启。

## 14. 接口速查
GET `/api/modules` `/api/state` `/api/schema` `/api/clean/types` `/api/import/progress` `/api/clean/audit`
POST `/api/upload` `/api/inspect` `/api/import` `/api/clean/save` `/api/clean/preview` `/api/clean/apply` `/api/clean/apply-all` `/api/clean/dropped` `/api/clean/reset` `/api/mapping` `/api/suggest` `/api/merge` `/api/recon` `/api/recon/export` `/api/query` `/api/export` `/api/delete-rows` `/api/books/preview` `/api/books/import`

**核心风险只有三个**：把汇总层当明细算、把重复账簿算两遍、拿不同期间的两张表对比。做错会出几千项假不符。
