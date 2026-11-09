# dwg-material-to-xls

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

安装后在 AI 对话里直接说：

```
把这两张 dwg 的材料表导入 excel 做材料平差 / 材料汇总
```

并附上施工图、竣工图两份 `.dwg`。技能会自动识别材料表、按分组匹配、生成核对报告并写出 `.xls`。

也可以直接跑命令行（在技能目录下）：

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
`--match-mode group|global|hybrid`、`--asbuilt-tables 1`、`--dry-run`。
完整说明见 [SKILL.md](SKILL.md) 与 [references/layout.md](references/layout.md)。

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
├── scripts/
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
- **竣工图选错表**：一份竣工图 DWG 里常同时含 1 张竣工材料表和多页随图附带的施工图版材料表，技能按图签阶段自动选表；阶段缺失时用 `--asbuilt-tables` 手动指定。
- **施工图多出很多材料**：多为户内材料（DN25 管件、超声波表等），成品模板没有对应行，会在核对报告里列出但不写入表格，属预期情况。
- **汇总表没同步**：模板缺少 `2-汇总`，或该表 C/D/E/I 首行本身是公式（脚本会跳过并提示）；`--no-fill-summary` 可显式关闭。
- **输出必须 `.xls`**：用 Excel COM 保真写入，改成 `.xlsx` 会触发格式提示。
- **工程名称显示成文件名**：加强版打印版的工程名称由文件名公式生成，输出文件建议按项目名命名。
- **下载慢 / 连接失败**：用 `--mirror` 或环境变量 `GITHUB_MIRROR`，也可手动下载 zip 后 `--local` 安装。

---

## 八、版本记录

| 版本 | 日期 | 内容 |
|---|---|---|
| 1.1.0 | 2026-11-10 | 内置「70623加强比对版」汇总模板为默认底稿，自动同步 `2-汇总` / 打印版 / 变更估价；模板容量自动探测 |
| 1.0.0 | 2026-11-10 | 首次发布：DWG 施工图/竣工图材料表导入 `1-初步比对`，含成品模板与核对报告 |

---

`assets/` 中的 `.xls` 模板来自实际工程，仅作为格式底稿使用（运行时只复制、不修改原件）。
本仓库未声明开源协议，使用前请遵守所在单位的数据与模板管理规定。
