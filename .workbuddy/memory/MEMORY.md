# PTW 平台 · 项目长期约定（可执行版）

PTW = PBC to WP（客户资料 → 底稿）。审计底稿工具，现已做成**双击即用的 exe**：
启动器 `PTW.exe` + 主程序 `runtime/` + 数据 `data/`，支持内置 AI 与一键自动更新。
版本真相只在一处：`开发/程序/app_version.py` 的 `APP_VERSION`。

## 沟通（每次遵守）
- 用户**不是程序员**：只谈业务逻辑与做出来的功能，不谈函数/框架/设计模式。
- **每次回复结尾加一句成语**。
- Excel 外部链接一律**相对路径**；分发包文件名简短、保持内部相对结构。

## 目录铁律（2026-11-14 重构后）
| 区 | 归谁 | 进 git | 能更新 |
|---|---|---|---|
| `runtime/` | 发布产物，只由 `build.py` 写 | ❌ | ✅ 整体替换 |
| `data/` | 用户数据（账套/模板/日志/设置） | ❌ | ❌ 永不触碰 |
| `开发/` | 源码 + 构建 + 测试脚本 + 文档 | ✅ | ❌ |
| `02_备份/ 04_归档/ tmp/` | 回滚点 / 交付件 / 过程 | 部分 | ❌ |

- **`runtime/` 里绝不手改**：手改下次打包就丢，且"本地能跑、别人不能跑"。
- **root 零新增**：根目录只允许 `PTW.exe / version.json / 启动说明.txt / README.md / .gitignore / .githooks/`。
- 新克隆仓库要 `git config core.hooksPath .githooks`（否则结构体检不生效）。
- `data/bkd_tpl` 首次启动从 `_internal/bkd_tpl_seed`（源 `开发/程序/tpl`）释放，已有不覆盖。

## 构建与发版
```bat
cd 开发\构建
..\..\tmp\buildenv\Scripts\python.exe build.py          :: 全量
..\..\tmp\buildenv\Scripts\python.exe build.py --app    :: 只重打主程序
```
- 打包环境 `tmp/buildenv`：Python 3.12 + PyInstaller 6.22 + duckdb 2.0.0 + openpyxl 3.1.5。
- **不能直接 `--distpath` 到根覆盖 runtime**：PyInstaller 的 `--clean` 会批量删旧目录，
  被批量删除保护拦下且失败会留半个 runtime。`build.py` 已改成"临时目录 + 改名切换"。
- 发版顺序：改 `APP_VERSION` → 打包 → 建 Release（tag=`v{版本}`）→ **最后**改仓库 `version.json`。
- 分发包：`04_归档/PTW_v{ver}_win64.zip`（约 35MB，不含 data）。

## 三项能力（别改坏）
1. **日志**：`core/logutil.py` 接管 stdout/stderr、按天滚动 14 天；`/api/logs` 查看、
   `/api/logs/download` 下载 zip（含 env.txt + 脱敏 settings.json）。打包无控制台时靠它排障。
2. **内置 AI（B2）**：`core/ptw_ai.py`，智谱免费 + OpenAI 兼容 BYOK；三条通道（口径归类 /
   表2 结论 / OAR 变动原因）**写与外部旁路同名 `*_result.json`**，两边可互换、可都不跑。
   AI 只写字与分类；数字一律走公式；`has_number()` 做最后一道闸。
3. **自动更新**：`launcher.py`（常驻 wait 主程序）+ `core/updater.py`（4 段降级链、
   sha256、解压、标记）。替换靠**目录改名**，不覆盖运行中的文件；失败保留 `runtime_old`。
   `PTW_LAUNCHED=1` 表示有 launcher 在等 → 主程序直接 `os._exit` 即可。

## 数据与工程纪律（沿用）
- 源数据只读，产物另存；跑任何验证带 `KZ_DATA_DIR=tmp/xxx`，绝不直连 `data/`。
- 位置不写死：文件/sheet/行/列按表头文字、业务键现场定位；找不到上报"找过哪里"，禁止猜。
- 交付红线：真实消费端验证（Excel 用本机 Excel 打开）；三方核对且独立重算换路径。
- 数值分类：A 已外链 / B 有值无来源（保留+标注）/ C 应取却定位失败（bug）/ D 同源不一致（最高优先）。
- 回归：`开发/01_测试/自动回归`（脚本进 git，样本不进）；`tmp/_regdata` 与 `tmp/_fixture` 是测试依赖，别清。

## 细节去哪查
- 目录与构建：`README.md`
- 业务口径：`开发/03_文档/AI先读我.md`（路由）→ 各专题正文
- 唯一必须记住的结论：**DuckDB `read_xlsx` 单线程是硬天花板**，调参只会更慢。
