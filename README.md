# PTW 平台 · 可执行版工作区

审计底稿工具（PBC → WP）的可执行发行版：**双击即用、内置 AI、一键自动更新**。

---

## 一、目录结构与分区职责

```
01_PTW平台/
├─ PTW.exe                 ★ 启动器（永不变更）——检查更新、拉起主程序、执行替换
├─ version.json            本机版本指针（launcher 读写）
├─ 启动说明.txt             给最终用户看的使用说明
│
├─ runtime/                【发布产物区】主程序。更新时整体替换，不要手改
│   ├─ PTWApp.exe
│   ├─ _internal/          依赖与静态资源（web / modules / rules / bkd_tpl_seed）
│   └─ ...
│
├─ data/                   【运行数据区】用户数据，不参与更新、不进 git
│   ├─ ledger.duckdb           账套库
│   ├─ bkd_tpl/                底稿模板（首次启动从程序资产释放，可换客户版）
│   ├─ uploads/ bkd_out/       导入件 / 生成件
│   ├─ logs/                   app-YYYYMMDD.log、launcher.log
│   ├─ updates/                更新的下载缓存
│   └─ settings.json           AI 配置（含 Key，只存本机）
│
├─ 开发/                    【开发区】进 git，改代码只在这里改
│   ├─ 程序/                    源码：server.py / core / modules / web / rules / tpl
│   ├─ 构建/                    PTWApp.spec、launcher.spec、build.py、launcher.py
│   ├─ 01_测试/                 回归脚本与测试样本（样本不进 git）
│   ├─ 03_文档/                 当前有效文档（AI先读我 是路由入口）
│   └─ .githooks/               pre-commit 结构体检
│
├─ 02_备份/                 【回滚点】改代码前的快照，命名写"改什么之前"
├─ 04_归档/                 【交付件】发布包 PTW_v{ver}_win64.zip + version.json
└─ tmp/                     【临时区】构建环境、验收环境、一次性别名；可随时清
```

### 四条边界（相互隔离）

| 区 | 归谁 | 进 git | 能被更新替换 | 说明 |
|---|---|---|---|---|
| `runtime/` | 程序 | ❌ | ✅ 整体替换 | 只由构建脚本写入 |
| `data/` | 用户 | ❌ | ❌ | 升级不动它，谁也删不掉 |
| `开发/` | 开发者 | ✅ | ❌ | 唯一的代码来源 |
| `tmp/ 02_备份/ 04_归档/` | 过程产物 | 部分 | ❌ | 归档只放对外交付件 |

> **一条铁律**：`runtime/` 里的东西永远从 `开发/程序/` + `build.py` 生成，绝不手改。
> 手改了下次打包就丢，而且会"本地能跑、别人不能跑"。

### 结构由机器守，不靠记性

乱不是"不小心"，是**没有任何东西拦着**。现在有三道闸：

| 闸 | 位置 | 拦什么 |
|---|---|---|
| 构建前 | `build.py` 自动跑 `check_workspace.py` | 根目录出现野目录/野文件、`runtime` 不完整、tmp 膨胀、归档超 3 个包 → **拒绝构建** |
| 提交前 | `.githooks/pre-commit` | 上面这些 + 新增文件位置、规则文件行数、引用工作区外路径 → **拒绝提交** |
| 构建后 | `build.py` 自动清理 | `runtime_old_build` / `tmp/_dist_new` 当场删掉，构建完根目录回到只有固定项 |

**平台根允许存在的东西（就这些）**：

```
PTW.exe  version.json  README.md  启动说明.txt  .gitignore
runtime/  data/  开发/  02_备份/  04_归档/  tmp/
```

多一个就是脏的。手工整理入口：`开发/构建/3_整理工作区.bat`（先演练再执行）。

---

## 二、构建与发布

```bat
:: 全量（launcher + 主程序 + 发布包）
cd 开发\构建
..\..\tmp\buildenv\Scripts\python.exe build.py

:: 只重打主程序（改业务代码后最常用）
..\..\tmp\buildenv\Scripts\python.exe build.py --app
```

产物：
- `runtime/`（主程序，约 63MB）
- `PTW.exe`（启动器，约 10MB）
- `04_归档/PTW_v{ver}_win64.zip`（对外分发包，约 35MB，**不含 data**）
- `04_归档/version.json`（更新清单，publish.py 会填好 url 与 sha256）

### 发版（一条命令，token 放本地）

**不想敲命令就双击 `开发/构建/` 里的两个 bat**：
`1_首次建仓库.bat`（只跑一次）、`2_发版本.bat`（以后每次发版）。
下面是对应的命令行版本。

**第一次**（建仓库、配好更新源、推代码）：

```
# 1. 把 token 粘进这个文件（只放一行，别加引号，绝不提交）
#    C:\Users\<你>\.workbuddy\ptw_token.txt
#    只认这一个文件：不会借用 MDMQ 等其它项目的 token（已在代码里禁止回退）

# 2. 建仓库 + 写 UPDATE_REPO + 推 main
..\..\tmp\buildenv\Scripts\python.exe publish.py --setup

# 3. 发第一版
..\..\tmp\buildenv\Scripts\python.exe publish.py --notes "首个版本"
```

**token 建议用 fine-grained、最小权限**（不是 classic 全权 token）：

| 项 | 值 |
|---|---|
| Repository access | Only select repositories → `yangyuli0602/PTW` |
| Permissions | Contents: Read and write；Administration: Read；Metadata: Read |
| 有效期 | 90 天 |

**以后每次发版**：改 `app_version.py` 的 `APP_VERSION` → 跑

```
..\..\tmp\buildenv\Scripts\python.exe publish.py --notes "本次改了什么"
```

它会自动完成：打包 → 建 Release（tag = v{版本}）→ 传 zip → **最后**更新仓库里的
`version.json`（含真实 sha256 与大小）。用户点「检查更新 → 立即更新」全自动升级。

顺序不能反：先写 `version.json` 再传包，用户会"看到新版本但下载 404"——publish.py 已按正确顺序编排。

其它开关：`publish.py --check`（只验证 token 与仓库）、`--dry-run`（只生成本地清单不联网）、
`--skip-build`（复用上次的包）、`--private`（建私有仓库；私有仓库客户端读不到 raw 清单，
需改用 `PTW_UPDATE_MANIFEST` 指向自己的服务器）。

---

## 三、三项能力的实现位置

| 能力 | 实现 | 用户看到的位置 |
|---|---|---|
| **日志落文件 + 界面下载** | `core/logutil.py`（接管 stdout/stderr、按天滚动、保留 14 天）+ `/api/logs`、`/api/logs/download` | 右上角「日志」→ 查看 / 下载日志包（zip 含 env.txt 与脱敏设置） |
| **内置 AI（B2）** | `core/ptw_ai.py`；三条通道与外部 AI 旁路**写同一批 `*_result.json`** | 右上角「AI 助手」→ 选服务商 / 填 Key / 一键跑三条通道 |
| **自动更新** | `launcher.py`（常驻等待 + 替换）+ `core/updater.py`（检查/下载/sha256/解压） | 右上角「检查更新」→ 有新版弹说明 → 点「立即更新」自动完成 |

### 更新为什么这么设计

Windows 不能覆盖正在运行的程序。所以：

```
用户双击 PTW.exe
   └─ launcher 拉起 runtime/PTWApp.exe 并 wait() 等它
用户点「立即更新」
   ├─ 主程序：下载 zip → 校验 sha256 → 解压 runtime_new/ → 写 update_pending.json → 自己退出
   └─ launcher：发现主程序退出 + 有标记
        ├─ runtime → runtime_old（改名）
        ├─ runtime_new → runtime（改名）
        ├─ version.json 写新版本
        └─ 重新拉起 PTWApp.exe，日志留痕
```

**用户只点一次「是」**，无需手动解压、无需管理员权限、失败自动回滚（保留 runtime_old）。

### AI 的三条通道（与外部旁路等价）

| 任务 | 待办清单 | 结果文件 |
|---|---|---|
| 口径归类 | `data/bkd_work/ai_tasks/oar_bucket_pending.json` | `oar_map_ai_result.json` |
| BKD 表2 结论 | `data/bkd_work/ai_tasks/table2_pending.json` | `table2_concl_result.json` |
| OAR 变动分析 | `data/bkd_work/ai_tasks/oar_reason_pending.json` | `oar_concl_result.json` |

内置 AI 与外置 AI **可以互换、可以都不跑**（人工填也一样生效）——因为读的是同一批文件。

---

## 四、配置项（app_version.py）

```python
APP_VERSION = "1.3.0"          # 版本真相，只改这里
UPDATE_REPO = ""               # 填 "owner/repo" 后「检查更新」才生效
```

内网部署可绕过 GitHub：设环境变量 `PTW_UPDATE_MANIFEST=<清单地址>` 指向自己的 version.json。

---

## 五、已知边界

- 打包用 Python 3.12 + duckdb 2.0.0 + openpyxl 3.1.5（`tmp/buildenv`，与运行时版本一致）
- 主程序 `console=False`：所有输出进日志，界面可下载
- 更新只替换 `runtime/`，**不碰 data/**；模板首次启动从 `_internal/bkd_tpl_seed/` 释放
- 源码态（`python server.py`）不支持自动替换，更新接口会明确报错
