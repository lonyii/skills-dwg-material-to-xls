# dwg-material-to-xls

**作者：lonyii（lonyii@qq.com）**

把燃气工程 **施工图 / 竣工图 DWG** 里的材料表自动导入 Excel，生成材料平差 / 材料汇总 `.xls`：
自动填好「1-初步比对」，同步「2-汇总」，并联动汇总打印版、变更估价与调拨价。

---

## 一、给 AI 一句话即可安装（推荐）

把下面这句话连同仓库地址发给任意能执行命令的 AI（WorkBuddy / Claude / Cursor / Copilot 等）：

```
请从这个仓库安装技能到我的用户级 skills 目录，装好后告诉我怎么用：
https://github.com/lonyii/skills-dwg-material-to-xls
```

AI 只需要执行 **一条命令**，无需构建、无需先装依赖：

```bash
# macOS / Linux / Git Bash
git clone --depth 1 https://github.com/lonyii/skills-dwg-material-to-xls "$HOME/.workbuddy/skills/dwg-material-to-xls"
```

```powershell
# Windows PowerShell
git clone --depth 1 https://github.com/lonyii/skills-dwg-material-to-xls "$env:USERPROFILE\.workbuddy\skills\dwg-material-to-xls"
```

> **为什么这样就够了**：本仓库根目录本身就是一个完整的技能包（含 `SKILL.md`、`scripts/`、`references/`、`assets/`），克隆到技能目录即安装完成。仓库里的 `README.md`、`install.py`、`update.py` 等是辅助文件，留在目录里不影响技能加载。
>
> 若目标目录已存在：先在该目录运行 `python update.py` 升级，或删除旧目录后再克隆。
> 安装后自检：`python ~/.workbuddy/skills/dwg-material-to-xls/install.py --check`

### 没有 git？用一键安装脚本

`install.py` 只用 Python 标准库，会从仓库地址下载 zip 并装到正确位置：

```powershell
# Windows PowerShell
irm https://raw.githubusercontent.com/lonyii/skills-dwg-material-to-xls/main/install.py -OutFile install.py
python install.py https://github.com/lonyii/skills-dwg-material-to-xls
```

```bash
# macOS / Linux
curl -fsSL -O https://raw.githubusercontent.com/lonyii/skills-dwg-material-to-xls/main/install.py
python3 install.py https://github.com/lonyii/skills-dwg-material-to-xls
```

国内网络不稳时可加镜像，或先手动下载 zip 再本地安装：

```bash
python install.py --mirror https://ghproxy.net/          # 走镜像
python install.py --local ./解压后的目录                  # 完全离线
python install.py --dir "D:/my-skills"                  # 自定义技能根目录
python install.py --check                               # 只看远端版本
```

也可以直接下载 [仓库 zip](https://github.com/lonyii/skills-dwg-material-to-xls/archive/refs/heads/main.zip)，
把解压出的文件夹重命名为 `dwg-material-to-xls`，放到 `~/.workbuddy/skills/` 下。

---

## 二、运行环境

| 依赖 | 说明 |
|---|---|
| Windows + Python 3.9+ | 读写 DWG / xls 依赖 Windows COM |
| `pywin32` | `python -m pip install pywin32` |
| 浩辰 CAD / AutoCAD / 中望 CAD 之一 | 通过 COM 读取 DWG 中的文字实体 |
| Microsoft Excel | 复制模板后写入，保真保留公式、定义名与打印设置 |
| `ezdxf`（可选） | 仅直接读 `.dxf` 时需要 |

安装脚本会自动检查 `pywin32` 并给出提示。没有 CAD 时，可用已保存的实体缓存 JSON 或 `rows.json` 离线运行匹配与写入（仍需 Excel）。

---

## 三、怎么用

安装后把**若干 `.dwg`** 发给 AI，说一句：

```
把这些 dwg 的材料表导出，生成材料平差表
```

AI 不需要你指出哪张是竣工图：脚本按目录/文件名、材料表图签阶段自动判别竣工图与施工图，
按工程编号分组，输出 `<工程编号>+<工程名称>.xls`
（如 `IC24H201+南昌县向塘博乐美幼儿园配套燃气工程.xls`）及同名核对报告。
**只给一张竣工图时，施工图量自动等于竣工图量**（J/K/L 与 C/D/E 相同，O 列「相同」）。

一键命令（AI 会执行）：

```bash
python scripts/auto_export.py "图1.dwg" "图2.dwg" ... --out-dir "D:/输出"
```

常用参数：`--out-dir`、`--cache-dir`（复用图纸缓存、整批只启动一次 CAD）、
`--dry-run`、`--row-base auto|asbuilt|design|merged`、`--template`。

> **施工图数据源**：竣工图 DWG 编制时保留的施工图版材料表是旧版、不准确的数据，
> 一律忽略；施工图数量只取用户发送的施工图文件。仅有竣工图时施工图量=竣工图量。

**一张施工图含多个工程**：施工图标注各工程名称时脚本会识别每张材料表所属工程
（`inspect` 显示 `工程=…`），**以竣工图为准**只提取对应工程的表——竣工图是厨房就取厨房、
是配套就取配套；发来多个工程的竣工图时，同一施工图分别提取、按各自工程编号生成多份成品。
即使只给一张竣工图也会按工程名筛选（不会把别的工程材料并进来）。
实测一张施工图含 13 个工程，配合/厨房等 3 份成品竣工与施工描述 100% 一致。

需要手动精细控制（明确指定哪张是竣工图/施工图）时：

```bash
python scripts/import_materials.py \
  --asbuilt "竣工图.dwg" \
  --design  "施工图.dwg" \
  --out     "D:/输出/IC24H201-材料汇总.xls" \
  --report  "D:/输出/核对报告.json"
```

首次用于新版图框时，建议先核对再生成：

```bash
python scripts/dwg_tables.py inspect "竣工图.dwg" --cache _jg.json
python scripts/dwg_tables.py inspect "施工图.dwg" --cache _sg.json
python scripts/import_materials.py --asbuilt "竣工图.dwg" --design "施工图.dwg" --dry-run --report _check.json
```

### 生成的文件里有什么

| 工作表 | 内容 |
|---|---|
| `1-初步比对` | C/D/E 竣工图描述/单位/数量，J/K/L 施工图，O 列异同公式 |
| `2-汇总` | C/D/E/I 由初步比对同步；H/J/K/L/M/N（平差、增加/减少材料）自动重算 |
| `4-汇总打印版 (新版)` | VLOOKUP 引用 `2-汇总`，自动生成送审打印表 |
| `5-变更估价` / `调拨价` | 引用汇总数据，公式保留 |

### 内置两份模板

| 模板 | 参数 | 明细容量 | 用途 |
|---|---|---|---|
| `assets/材料汇总模板-加强比对版.xls` | 默认 | 115 行 | **生产用**：空白 9 张表，无样例数据残留，含完整汇总/打印版/调拨价 |
| `assets/材料平差模板.xls` | `--template sample` | 40 行 | IC24H201 成品文件，用于复现格式与逐格验收 |

常用参数：`--template summary|sample|路径`、`--no-fill-summary`（只写初步比对）、
`--match-mode group|global|hybrid`、`--row-base asbuilt|design|merged`（自动编排时默认 `auto`）、
`--asbuilt-tables 1`、`--dry-run`。
完整说明见 [SKILL.md](SKILL.md) 与 [references/layout.md](references/layout.md)。

### 批量处理：拖入文件夹（可脱离 AI 运行）

**方式 A：交给 AI** —— 把整个文件夹路径发给 AI，说「处理这个文件夹里的图纸」：

```bash
python scripts/auto_export.py "D:/2026结算图纸" --out-dir "D:/输出"
```

脚本递归扫描文件夹内所有 `.dwg` / `.dxf`（`--no-recursive` 只扫顶层），
自动判别竣工/施工图、按工程分组，一次生成多个工程的平差表。
成品默认归档到**每张图纸所在文件夹下的 `材料平差输出/`**（以竣工图为准），
不写脚本目录或当前目录；核对报告 JSON 写到系统临时目录 `dwg2xls-reports/`，
不出现在成品目录。仅显式 `--out-dir` / `--report-dir` 才统一覆盖。
支持一次传入多个文件夹，也支持文件夹与单张图纸混合；
分层目录（如 `工程A/竣工图/`、`工程A/施工图/`）同样能正确判别。

**方式 B：完全脱离 AI** —— 把文件夹（或多个文件夹、若干 `.dwg`）直接拖到
技能根目录的 **`run_batch.bat`** 图标上：

- 自动定位 Python 3（优先 `py -3`，其次 `python`），缺失时提示安装
- 自动检查并安装 `pywin32`
- 自动启用实体缓存 `%USERPROFILE%\.workbuddy\dwg-entity-cache`；
  **DWG 被修改后缓存自动失效并重读**，不会用到过期数据
- 结果写入拖入文件夹下的 `材料平差输出/`，完整日志见技能目录 `run_batch.log`
- 结束暂停显示结果；无人值守可设 `DWG_BATCH_NOPAUSE=1`

| 环境变量 | 作用 |
|---|---|
| `DWG_BATCH_NOPAUSE=1` | 结束不暂停（计划任务 / 脚本调用） |
| `DWG_BATCH_DRYRUN=1` | 只打印计划，不写 `.xls` |
| `DWG_BATCH_NOCACHE=1` | 每次都从 CAD 重新读取 |

只需 Windows + Python 3 + 浩辰/AutoCAD/中望 CAD + Microsoft Excel，
**不需要 WorkBuddy 或任何 AI**。

### 多实例验证与行序基准（v1.2.0）

已在三个真实工程上回归：

| 工程 | 图框 | 推荐参数 | 与成品逐格比对 |
|---|---|---|---|
| IC24H201 向塘博乐美幼儿园 | 1:1 | 默认 `--row-base asbuilt --template sample --no-fill-summary` | A3:O43 共 410 格 0 差异 |
| NC25D007 天之益锅炉 | 1:1，竣工图两页（一页无图签） | 默认 `--row-base asbuilt` | 描述与行序提取正确；成品另含人工插入项与另一版次数量，需人工复核 |
| ID24E107 天天包子店 | 1:2 缩放 | `--row-base merged` | 41 行 0 差异 |

- 脚本按表头「序号→单位」锚点间距自动换算全部列偏移与行容差，1:1 与 1:2 图框
  共用一套参数；`inspect` 会打印每张表的「比例」便于诊断。
- **竣工图 DWG 编制时常保留施工图版材料表，那是旧版且不准确的数据，一律忽略**；
  施工图数量只取用户发送的施工图文件，仅有竣工图时按「施工图量=竣工图量」处理。
  缺图签的多页会继承邻近页阶段。
- `--row-base`：`asbuilt`（默认，竣工表行序）、`design`（施工表行序，竣工独有追加表尾）、
  `merged`（两表并集按原顺序交错，并合并重复的分组标题）。
- 成品可能含人工插入/合并的行或另一版次数量，差异项在核对报告中按分组列出，
  交付前请人工复核，不要直接作为结算依据。

---

## 四、更新版本

```bash
python update.py --check     # 比较本地 version.json 与远端版本
python update.py             # 升级到最新版（自动备份旧版到 ~/.workbuddy/skill-backups/）
python update.py --force     # 版本相同也重装，用于修复本地文件
python update.py --mirror https://ghproxy.net/
```

版本号与更新日志统一记录在 `version.json`。`update.py` 只覆盖仓库自带文件，不会删除你在技能目录里自己放的东西。

---

## 五、维护者：发布新版本

```bash
python release.py 1.2.0 "新增 XX 图框支持" --push
# 或按位递增
python release.py --bump patch "修正列偏移" --push
```

脚本会更新 `version.json` → `git add -A` → commit → 打 `vX.Y.Z` 标签 →（`--push`）推送分支与标签。
不加 `--push` 时只在本机提交并打印推送命令。使用者随后运行 `python update.py` 即可升级。

> Windows 上若 `git push` 报 `CRYPT_E_NO_REVOCATION_CHECK`，是本机证书链问题，可临时用
> `git -c http.schannelCheckRevoke=false -c http.sslVerify=false push origin HEAD`，
> 或把凭据交给 Git Credential Manager 处理。

---

## 六、目录结构

```
dwg-material-to-xls/
├── SKILL.md                       # 技能说明（AI 读取的入口）
├── version.json                   # 版本号与更新日志
├── README.md                      # 本文件
├── install.py                     # 一键安装（纯标准库）
├── update.py                      # 版本检查与升级
├── release.py                     # 维护者发布
├── run_batch.bat                  # 拖入文件夹即可运行（脱离 AI）
├── scripts/
│   ├── auto_export.py             # 默认入口：若干 DWG 自动判别、分组、命名、生成
│   ├── dwg_tables.py              # DWG 文字提取、材料表识别（inspect / extract）
│   └── import_materials.py        # 分组匹配、核对报告、Excel 保真写入
├── references/
│   └── layout.md                  # 模板列定义、汇总联动、CAD 识别规则
└── assets/
    ├── 材料汇总模板-加强比对版.xls
    └── 材料平差模板.xls
```

---

## 七、常见问题

- **检测到 0 张表 / 有效行 0**：图框列布局不同，按 `references/layout.md` 调整 `dwg_tables.py` 的列偏移；先用 `inspect` 确认。
- **2026 新图框的描述**：表头为「序号/标准号或图号/名称/规格/单位/数量」时，物料描述按成品格式输出「名称原样 + 空格 + 规格原样」（如 `热镀锌钢管_燃气专用 DN50（D60.3×3.8）`），不写入材质列（Q235B 等）。
- **`名称及规格` 合并列图框**：表头为「序号/标准号或图号/名称及规格/…/单位/数量」时，取该列原样为描述、忽略材质列，自动识别。
- **双栏材料表**：一张 DWG 内左右双栏共用一个表头的版式（如 ID26H016）暂不能自动合并，成品会比识别多十几行，核对报告会列出，交付前需人工复核。
- **竣工图选错表**：一份竣工图 DWG 里常同时含 1 张竣工材料表和多页随图附带的施工图版材料表，技能按图签阶段自动选表（这些旧施工图表**不会**用于施工图数量）；阶段缺失时用 `--asbuilt-tables` 手动指定。
- **施工图多出很多材料**：多为户内材料（DN25 管件、超声波表等），成品模板没有对应行，会在核对报告里列出但不写入表格，属预期情况。
- **汇总表没同步**：模板缺少 `2-汇总`，或该表 C/D/E/I 首行本身是公式（脚本会跳过并提示）；`--no-fill-summary` 可显式关闭。
- **输出必须 `.xls`**：用 Excel COM 保真写入，改成 `.xlsx` 会触发格式提示。
- **工程名称显示成文件名**：加强版打印版的工程名称由文件名公式生成，输出文件建议按项目名命名。
- **换了设计院图纸后解析出 0 行**：多为出图比例不同，v1.2.0 已做等比自适应；先看 `inspect` 的「比例」是否正常（0.2～4），仍异常再按 `references/layout.md` 调列偏移。
- **输出行数与手上的成品对不上**：成品常人工插入条目、合并分组或取另一版次数量；先试 `--row-base design` / `merged`，并查看核对报告中 `extras` 的分组分布。
- **下载慢 / 连接失败**：用 `--mirror` 或环境变量 `GITHUB_MIRROR`，也可手动下载 zip 后 `--local` 安装。

---

## 八、版本记录

| 版本 | 日期 | 内容 |
|---|---|---|
| 2.4.7 | 2026-11-11 | 修复材料表第2页漏识别：表头上方的「图纸目录」等同序号列表不再污染段切分；NC26E761 施工图 17→41 行，NC26E546/生米镇找回被丢弃的页 |
| 2.4.6 | 2026-11-11 | 修复一张施工图混放多个工程时串表：按每张表图签的「项目编号」精确筛选，工程名支持「（室内接管）」等括号后缀；NC26C002 材料不再并入 NC26I003 |
| 2.4.5 | 2026-11-10 | run_batch.bat 成品改输出到桌面的「材料平差输出」，按日期时间建子文件夹，完成后自动打开该文件夹（dry-run 不打开） |
| 2.4.4 | 2026-11-10 | 修复 LISP 导出中文整份乱码导致材料表识别为 0 张（gb18030 解码 + 乱码缓存自动重读，NC26E546 恢复 5 表 79 行）；run_batch.bat 增加窗口标题与作者信息 |
| 2.4.3 | 2026-11-10 | 修复施工图误取「碰接点详见」等引用的其它工程编号（优先括号标注、排除引用行）；读图与生成两段进度之间加分隔说明 |
| 2.4.2 | 2026-11-10 | 成品归档到图纸目录「材料平差输出」，核对报告移到临时目录不展示给用户 |
| 2.4.1 | 2026-11-10 | 批处理提示精简工整：自检一行、进度对齐、汇总一行，功能无变化 |
| 2.4.0 | 2026-11-10 | LISP 进程内批量导出替代逐实体 COM 遍历，大图读图提速 10 倍以上，结果逐条一致 |
| 2.3.1 | 2026-11-10 | 模板文件名改为 ASCII（兼容任意解压工具），保留旧中文名回退 |
| 2.3.0 | 2026-11-10 | 批处理提示全部中文化、依赖自检与引导安装、缓存日志移入包内，整包可复制到任意电脑运行 |
| 2.2.0 | 2026-11-10 | 新增无表头材料表识别（单位词列聚类），缓存优先复用与0表停止重读+识别结果汇总 |
| 2.1.5 | 2026-11-10 | run_batch.bat 运行进度实时显示：读图/生成分阶段 ASCII 进度条，日志同步写入 |
| 2.1.4 | 2026-11-10 | 修复 COM 偶发读漏实体导致缓存不完整、竣工图误判施工图 0 表（NC26E213 恢复 35 行） |
| 2.1.3 | 2026-11-10 | 修复 run_batch.bat 拖入文件名含括号时闪退（重构为无括号块结构） |
| 2.1.2 | 2026-11-10 | 修复老图纸扁表格把设计师姓名当物料（师大南路从 1 行人名恢复为 29 行材料） |
| 2.1.1 | 2026-11-10 | 补充作者信息（lonyii / lonyii@qq.com）到 version.json、SKILL.md、README 与 Python 脚本；功能无变化 |
| 2.1.0 | 2026-11-10 | 新增无页码图纸的多表拼接顺序：序号连贯按序号、序号从上到下→从左到右、从下到上→从右到左；默认并入一张成品表 |
| 2.0.1 | 2026-11-10 | run_batch.bat 拖放时显示技能版本与脚本路径，避免误用旧副本；本机历史测试副本统一升级到 v2.0.0 |
| 2.0.0 | 2026-11-10 | 修复 run_batch 生成 IE26I004 只有 2 行超长描述的错误：适配极扁表格（比例 0.013、行距 8）——比例下限降到 0.01、行聚类按实测行距自适应、并排表头防串列；竣工图按本工程名筛表，排除图内附带的其它工程竣工表。从 2 行乱码恢复为 40 行正确材料，旧实例与师大 13 工程样本回归保持 |
| 1.9.0 | 2026-11-10 | 支持一张施工图含多个工程材料表：按表头图签识别每张表所属工程，以竣工图工程编号/名称为准只提取对应表；多张竣工图时同一施工图共享、分别生成各自成品；大图自动跳过巨型模型空间只读布局空间。实测一张图含 13 个工程，3 份成品描述 100% 匹配 |
| 1.8.0 | 2026-11-10 | 重启后随机抽 4 组图纸实测：新增 `名称及规格` 合并列图框识别（0 表→20 行）、分组标题独立成行、材料表大标题区分；ID26E025/NC25E121 竣工侧与成品仅 1 处差异；旧三实例回归保持 |
| 1.7.0 | 2026-11-10 | 按工程编号匹配 582 组「图纸+成品表」并随机抽样，用 634 份成品表交叉验证；修正 2026 新图框描述格式：只输出「名称原样 + 空格 + 规格原样」、不写入材质列（对照成品 ID26E025/ID26H016）；旧三实例回归保持 |
| 1.6.0 | 2026-11-10 | 扫描本机图纸验证并适配 2026 新图框：识别「序号/标准号或图号/名称/规格/单位/数量」+「材料/质量(kg)」双行表头，把名称+规格+材料拼成物料描述；施工图 DWG 无材料表时明确提示；旧三实例回归保持 |
| 1.5.0 | 2026-11-10 | 按业务规则修正施工图数据源：竣工图 DWG 内残留的施工图版材料表一律忽略，施工图数量只取用户发送的施工图文件；移除自动补充逻辑，旧 `--no-supplement` 保留为无操作兼容参数；三实例与批处理回归通过 |
| 1.4.0 | 2026-11-10 | `auto_export.py` 支持文件夹拖入（递归扫描、多工程批量、默认输出到「材料平差输出」）；新增 `run_batch.bat` 可脱离 AI 拖放运行（自动找 Python/装 pywin32、缓存按图纸修改时间失效、运行日志）；控制台编码与 CAD 实例复用优化 |
| 1.3.0 | 2026-11-10 | 新增 `auto_export.py` 默认入口：若干 DWG 自动判别竣工/施工图、按工程分组、按「编号+名称」命名、单竣工图时施工图量=竣工图量、行序自动选择；整批复用一个 CAD 实例并支持实体缓存 |
| 1.2.0 | 2026-11-10 | 两个新工程实例验证：图框比例自适应、缺图签多页阶段继承与页序修正、`--row-base` 三种行序基准；三实例均通过回归（当时的「施工图数据自动补充」已于 v1.5.0 按业务规则停用） |
| 1.1.0 | 2026-11-10 | 内置「70623加强比对版」汇总模板为默认底稿，自动同步 `2-汇总` / 打印版 / 变更估价；模板容量自动探测 |
| 1.0.0 | 2026-11-10 | 首次发布：DWG 施工图/竣工图材料表导入 `1-初步比对`，含成品模板与核对报告 |

---

`assets/` 中的 `.xls` 模板来自实际工程，仅作为格式底稿使用（运行时只复制、不修改原件）。
本仓库未声明开源协议，使用前请遵守所在单位的数据与模板管理规定。
