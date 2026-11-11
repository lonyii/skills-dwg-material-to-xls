#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
「发若干 DWG 给 AI，说导出材料表」时使用的一键编排入口。

    python auto_export.py "竣工图.dwg" "施工图.dwg" ... [--out-dir DIR]
    python auto_export.py "拖入的文件夹"                 # 递归扫描文件夹内所有 DWG
    python auto_export.py "文件夹A" "文件夹B" *.dwg     # 目录与文件可混合
    python auto_export.py "只有一张竣工图.dwg" --out-dir DIR
    python auto_export.py <多个DWG> --cache-dir _cache --dry-run

自动完成：
  1. 把文件夹参数递归展开成图纸清单（--no-recursive 只扫顶层）；
  2. 读取每张 DWG 的文字实体（--cache-dir 可复用缓存，整批只启动一次 CAD）；
  3. 按目录/文件名、材料表图签阶段、图签文字判别「竣工图 / 施工图」；
  4. 按工程编号分组（无编号的施工图按名称相似度并入对应工程）；
  5. 从图签或文件名提取「工程编号 + 工程名称」，输出 `<编号>+<名称>.xls`；
     拖入文件夹时默认输出到该文件夹下的「材料平差输出」；
  6. 只有竣工图时，施工图量 = 竣工图量（J/K/L 与 C/D/E 相同，O 列显示「相同」）；
  7. 自动选择行序基准（见 choose_row_base），写入 `1-初步比对` 与 `2-汇总`。

重要业务规则：竣工图 DWG 编制时常保留施工图版材料表，那是旧版且不准确的数据，
一律忽略；施工图数据只取用户发送的施工图文件。
"""

from __future__ import annotations

__author__ = "lonyii"
__email__ = "lonyii@qq.com"

import argparse
import difflib
import json
import os
import re
import sys
import tempfile

if hasattr(sys.stdout, "reconfigure"):
    try:
        # 跟随控制台编码（CMD 中文代码页 / UTF-8 终端），避免批处理输出乱码
        sys.stdout.reconfigure(encoding=sys.stdout.encoding or "utf-8",
                               errors="replace")
    except Exception:                          # noqa: BLE001
        pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dwg_tables import (  # noqa: E402
    assign_groups, clean_text, close_cad_app, detect_tables, load_entities,
    norm_desc, open_cad_app, order_tables, parse_table,
)
from import_materials import (  # noqa: E402
    match_rows, match_rows_design_base, match_rows_merged,
    resolve_template, write_xls,
)

PROJ_RE = re.compile(r"[A-Z]{2}\d{2}[A-Z]\d{3}", re.I)
# 图纸说明里常引用「别的工程」的编号，例如
#   「碰接点详见：NC24A020 南崇朗翠小区管道燃气工程」
#   「已建管详见竣工图：NC17B069 黄溪村农民公寓二期燃气工程」
# 这些都不能当作本图工程编号。
PROJ_REF_RE = re.compile(r"碰接点|详见|参见|另详|参考|已建|既有|原有|图集|标准图")
ILLEGAL = re.compile(r'[\\/:*?"<>|\r\n\t]')
MATCHERS = {"asbuilt": match_rows, "design": match_rows_design_base,
            "merged": match_rows_merged}


def _flat(s):
    return re.sub(r"\s+", "", clean_text(s))


def clean_filename(s):
    return ILLEGAL.sub("_", str(s)).strip().strip(". ") or "未命名工程"


def cache_path(dwg, cache_dir):
    stem = os.path.splitext(os.path.basename(dwg))[0]
    return os.path.join(cache_dir, clean_filename(stem) + ".json")


def cache_is_fresh(dwg, cache_dir):
    """缓存必须比 DWG 新且内容完整；图纸改动或缓存读取不全时自动重读。

    COM 读取偶尔只读到部分布局（实测 NC26E213 一份缓存 461 实体、缺了竣工表区域，
    另一份 1019 实体完整）。这类坏缓存 mtime 比 DWG 新，旧判断会一直复用、导致
    「表 0 张 / 误判施工图」。这里加内容校验：缓存里有材料表表头关键词、但
    detect_tables 判 0 张，视为不完整缓存，强制重读。
    """
    if not cache_dir:
        return False
    cp = cache_path(dwg, cache_dir)
    if not os.path.isfile(cp):
        return False
    # 缓存优先策略：只要缓存能解析出材料表就复用，不强制要求比 DWG 新。
    # 读 CAD 每张要 1 分钟，缓存是秒级；图纸真改动时由内容校验/用户删缓存处理。
    try:
        data = json.load(open(cp, encoding="utf-8"))
        ents = data.get("entities", [])
        if not ents:
            return False
        # 旧版 LISP 导出遇到 GBK 非法字节会整份退回 UTF-8+replace，中文全变
        # U+FFFD、表头词丢失（实测 NC26E546 5 页表被判 0 表）。这类乱码缓存
        # 没有表头词、不会被下面的内容校验发现，必须单独拦截重读。
        bad = sum(str(e.get("t", "")).count("\ufffd") for e in ents)
        if bad >= 3:
            print("  [缓存校验] 缓存文字乱码（%d 处替换符），重新读图：%s" %
                  (bad, os.path.basename(dwg)), flush=True)
            return False
        tables = detect_tables(ents)
        if tables:
            return True                       # 缓存有效（含无表头表），直接复用
        # 缓存解析出 0 张表：区分「真无表图纸」与「缓存读漏」。
        heads = {"序号", "单位", "数量", "物料描述", "名称"}
        n_head = sum(1 for e in ents
                     if clean_text(e.get("t", "")).strip() in heads)
        if n_head >= 1:
            print("  [缓存校验] 缓存疑似不完整（%d 实体 / %d 表头词但 0 张表），"
                  "重新读图：%s" % (len(ents), n_head, os.path.basename(dwg)),
                  flush=True)
            return False                      # 有表头词却 0 表 → 读漏，重读
        # 无表头词、也无材料 → 视为真无表图纸（如封面/目录 DWG），直接用缓存不再读
        return True
    except Exception:                          # noqa: BLE001  缓存损坏也重读
        return False


def expand_inputs(items, recursive=True):
    """把文件夹参数递归展开为 DWG/DXF 清单；文件参数原样保留，并去重保序。"""
    out = []
    for item in items:
        p = os.path.abspath(item)
        if os.path.isfile(p):
            if p.lower().endswith((".dwg", ".dxf")):
                out.append(p)
            continue
        if not os.path.isdir(p):
            print("[跳过] 路径不存在：%s" % item)
            continue
        found = []
        if recursive:
            for root, dirs, names in os.walk(p):
                dirs[:] = [d for d in dirs
                           if not d.startswith((".", "~$")) and d != "材料平差输出"]
                for n in names:
                    if n.lower().endswith((".dwg", ".dxf")) and not n.startswith("~$"):
                        found.append(os.path.join(root, n))
        else:
            for n in os.listdir(p):
                fp = os.path.join(p, n)
                if os.path.isfile(fp) and n.lower().endswith((".dwg", ".dxf")) \
                        and not n.startswith("~$"):
                    found.append(fp)
        print("  文件夹 %s -> %d 张图纸" % (p, len(found)))
        out.extend(sorted(found))
    seen, uniq = set(), []
    for f in out:
        key = os.path.normcase(f)
        if key not in seen:
            seen.add(key)
            uniq.append(f)
    return uniq


def _label_value(ents, labels, max_len=60):
    """取图签标签（如「项 目 编 号」）右侧同一视觉行的文字。"""
    labs = [e for e in ents if _flat(e["t"]) in labels]
    for lab in labs:
        cands = [e for e in ents if abs(e["y"] - lab["y"]) <= 400
                 and e["x"] > lab["x"] + 100]
        if cands:
            best = min(cands, key=lambda e: e["x"] - lab["x"])
            t = clean_text(best["t"]).strip()
            if 0 < len(t) <= max_len:
                return t
    return ""


def find_project_no(ents, path):
    """
    工程编号：优先文件名，其次图签「项 目 编 号」右侧文字，最后才是全图正则。
    图签常印有通用图集号/其它工程的编号，不能随便取第一个匹配：
      - 「碰接点详见：NC24A020 ××工程」「已建管详见竣工图：NC17B069 ××」是引用；
      - 「厨房工程(NC26E815)」「配套工程（NC26E813）」这类括号标注才是本工程编号。
    全图扫描时先排除引用行，再按「括号标注 > 短独立行 > 大段说明」打分取最优。
    """
    m = PROJ_RE.search(os.path.basename(path))
    if m:
        return m.group(0).upper()
    m = PROJ_RE.search(_label_value(ents, ("项目编号", "工程编号", "项目号"), 20))
    if m:
        return m.group(0).upper()

    cands = {}                                  # 编号 -> (优先级, -文字长度)
    counts = {}                                 # 编号 -> 出现次数
    for e in ents:
        text = clean_text(e.get("t", "")).strip()
        if not text:
            continue
        for m in PROJ_RE.finditer(text):
            no = m.group(0).upper()
            ctx = text[max(0, m.start() - 26):m.end() + 26]
            if PROJ_REF_RE.search(ctx):          # 引用别的工程，跳过
                continue
            counts[no] = counts.get(no, 0) + 1
            score = 0
            if re.search(r"[\(（]\s*" + re.escape(no) + r"\s*[\)）]", text, re.I):
                score += 6                       # 工程名后括号标注，最可信
            if len(text) <= 26:
                score += 2                       # 独立成行的短标注
            if len(text) > 120:
                score -= 3                       # 大段设计说明里的编号多为引用
            if "工程" in text or "项目" in text:
                score += 1
            key = (score, -len(text))
            if no not in cands or key > cands[no]:
                cands[no] = key
    if not cands:
        return ""
    return max(cands, key=lambda no: (cands[no][0], counts[no], cands[no][1]))


def find_project_name(ents, path):
    """
    工程名称：文件名带工程编号时优先用文件名中的名称（与结算文件命名一致，
    如 IC24H201-南昌县向塘博乐美幼儿园配套燃气工程.dwg），其次图签，最后文件名主干。
    """
    stem = os.path.splitext(os.path.basename(path))[0]
    m = PROJ_RE.search(stem)
    if m:
        rest = re.sub(r"^[\s\-+_]+", "", stem[m.end():]).strip()
        if len(rest) >= 4:
            return rest
    for e in ents:
        m = re.match(r"^(?:工程名称|项目名称)\s*[:：]\s*(.+)$", clean_text(e["t"]).strip())
        if m and 2 <= len(m.group(1)) <= 60:
            return m.group(1).strip()
    t = _label_value(ents, ("项目名称", "工程名称"))
    if 2 <= len(t) <= 60 and not PROJ_RE.fullmatch(t):
        return t
    stem = os.path.splitext(os.path.basename(path))[0]
    m = PROJ_RE.search(stem)
    rest = stem[m.end():] if m else stem
    return re.sub(r"^[\s\-+_]+", "", rest).strip() or ""


def classify_stage(path, ents, tables):
    """判别整张图纸是竣工图还是施工图。竣工图 DWG 常同时附带施工图版材料表。"""
    parts = path.replace("\\", "/").split("/")
    hint = " ".join(parts[-2:])
    has_ab, has_sg = "竣工图" in hint, "施工图" in hint
    if has_ab and not has_sg:
        return "asbuilt"
    if has_sg and not has_ab:
        return "design"
    if any(t["stage"] == "竣工图" for t in tables):
        return "asbuilt"
    words = {clean_text(e["t"]).strip() for e in ents}
    if "竣工图" in words:
        return "asbuilt"
    return "design"


def load_drawing(path, cache_dir, app=None):
    path = os.path.abspath(path)
    cp = cache_path(path, cache_dir) if cache_dir else None
    heads = {"序号", "单位", "数量", "物料描述", "名称"}

    def _read_once():
        if cache_dir:
            os.makedirs(cache_dir, exist_ok=True)
            if cache_is_fresh(path, cache_dir):
                return load_entities(path, from_json=cp) + ("cache",)
            return load_entities(path, cache=cp, app=app) + ("cad",)
        return load_entities(path, app=app) + ("cad",)

    ents, meta, source = _read_once()
    tables = detect_tables(ents)
    n_head = sum(1 for e in ents if clean_text(e.get("t", "")).strip() in heads)
    # 有表头词却 0 表 → 可能读漏。删坏缓存、重读 1 次即止（读取端内部已重试）。
    if not tables and n_head >= 1 and source == "cad":
        if cp and os.path.isfile(cp):
            try:
                os.remove(cp)
            except OSError:                    # noqa: BLE001
                pass
        ents, meta, source = _read_once()
        tables = detect_tables(ents)
        n_head = sum(1 for e in ents if clean_text(e.get("t", "")).strip() in heads)
        print("  [重读1次] %s：重读后 %d 实体 / %d 张表（仍 0 表则视为无表，"
              "不再重试）" % (os.path.basename(path), len(ents), len(tables)),
              flush=True)
    if not tables:
        # 明确告知用户：这张图反复读不到材料表，停止重读、继续处理下一张。
        print("  [警告] %s：未识别到材料表（%d 实体 / %d 表头词），"
              "本图跳过；若为无表头图纸请检查是否真有材料表。" %
              (os.path.basename(path), len(ents), n_head), flush=True)
    return {
        "path": path,
        "ents": ents,
        "tables": tables,
        "stage": classify_stage(path, ents, tables),
        "no": find_project_no(ents, path),
        "name": find_project_name(ents, path),
        "meta": meta,
        "zero_tables": not tables,
        # 成品归档到该图纸所在目录下的「材料平差输出」，绝不写脚本目录/当前目录
        "out_dir": os.path.join(os.path.dirname(path), "材料平差输出"),
    }


def _project_match(table_project, target_name, target_no="", table_no=""):
    """
    判断材料表是否属于目标工程。
    一张施工图常混放多个工程的表。表图签上的「项目编号」唯一且最可靠，
    优先按编号精确比对；没有编号时才退回工程名精确/包含匹配。
    绝不使用模糊相似度（如「三米粥铺」vs「面夫子」会误匹配）。
    """
    # 编号双方都有：唯一且精确，直接定论（避免相似工程名串表）
    if table_no and target_no:
        return table_no.upper() == target_no.upper()
    if not table_project:
        return False                      # 未标注工程名的表不参与多工程图纸的匹配
    tp, tn = _flat(table_project), _flat(target_name)
    if not tn:
        return False
    if tp == tn or (len(tp) >= 8 and len(tn) >= 8 and (tp in tn or tn in tp)):
        return True
    # 编号出现在表工程名里也视为匹配（编号唯一，可靠）
    if target_no and target_no.upper() in clean_text(table_project).upper():
        return True
    return False


def rows_of(info, stage, target_project=None, target_no="",
            include_unmarked=False, keep_unmarked_project=False):
    """
    提取某阶段材料表行。target_project 给定时（一张图含多个工程），
    只取 table["project"] 与该工程匹配的表；阶段与工程名都用于筛选。

    筛选优先级：工程名（最可靠）> 阶段。竣工图 DWG 常附带其它工程的竣工表，
    必须用本工程名把那些表排除；阶段只在没有工程名信息时用来兜底。
    """
    ents = info["ents"]
    tables = info["tables"]
    # 先按工程名筛（若给出）：匹配到就只在这些表里继续
    if target_project:
        matched = [t for t in tables
                   if _project_match(t.get("project", ""), target_project, target_no,
                                     t.get("no", ""))]
        # 竣工图老图纸的竣工表常无工程名标注（project 为空），它们本就属于本图，
        # 需保留，否则会把真正的竣工表筛掉（实测 IC24H201）。
        if keep_unmarked_project:
            matched += [t for t in tables if not t.get("project")]
        if matched:
            tables = matched
        # 匹配不到时保留全部（退化旧行为），交由调用方提示
    # 再按阶段筛
    chosen = [t for t in tables if t["stage"] == stage]
    if include_unmarked:
        chosen += [t for t in tables if t["stage"] == "未标注"]
    if not chosen:
        chosen = tables                  # 阶段缺失时退化为全部（本工程）表
    rows = []
    for t in order_tables(chosen):      # 无页码时按序号走向/连贯性排多表顺序
        rows.extend(parse_table(ents, t))
    assign_groups(rows)
    return rows


def _name_of(group, no):
    for i in group:
        if i["stage"] == "asbuilt" and i["name"]:
            return i["name"]
    for i in group:
        if i["name"]:
            return i["name"]
    return no or "未命名工程"


def _core_name(s):
    """去掉「燃气工程 / 配套 / 低压 / 版次」等通用词，避免不同工程被误判为相似。"""
    s = re.sub(r"[（(].*?[)）]", "", str(s or ""))
    for w in ("燃气工程", "配套", "低压", "地上", "地下", "工程", "材料",
              "外管", "配合", "版次"):
        s = s.replace(w, "")
    return s.strip()


def _sim(a, b):
    return difflib.SequenceMatcher(None, _core_name(a), _core_name(b)).ratio()


def group_drawings(infos):
    """
    先按工程编号分组，再把「无编号」或「只有施工图」的组按工程名称并入有竣工图的组。
    施工图图签常印有通用图集号，不能仅凭编号分工程。
    """
    groups, unknown = {}, []
    for info in infos:
        (groups.setdefault(info["no"], []) if info["no"] else unknown).append(info)

    def ab_keys():
        return [k for k in groups if any(i["stage"] == "asbuilt" for i in groups[k])]

    for info in unknown:
        targets = ab_keys() or list(groups)
        if not targets:
            groups.setdefault("_未编号", []).append(info)
            continue
        best = max(targets, key=lambda k: _sim(info["name"], _name_of(groups[k], "")))
        if _sim(info["name"], _name_of(groups[best], "")) >= 0.4 or len(targets) == 1:
            groups[best].append(info)
        else:
            groups.setdefault("_未编号_%s" % clean_filename(info["name"])[:16], []).append(info)

    for k in list(groups):
        if k.startswith("_") or any(i["stage"] == "asbuilt" for i in groups[k]):
            continue
        targets = ab_keys()
        if not targets:
            continue
        name = _name_of(groups[k], "")
        best = max(targets, key=lambda b: _sim(name, _name_of(groups[b], "")))
        if _sim(name, _name_of(groups[best], "")) >= 0.4 or len(targets) == 1:
            groups[best].extend(groups[k])
            del groups[k]
    return groups


def choose_row_base(asbuilt, design, single):
    """
    自动选行序：施工图多出的材料若大半属于竣工表没有的分组（如户内/调压前地下），
    说明施工图范围更大，成品以竣工表为骨架 → asbuilt；否则两表规模相当 → merged。
    """
    if single:
        return "asbuilt"
    _, extras, _ = match_rows(asbuilt, design)
    if not extras:
        return "asbuilt"
    groups = {norm_desc(r["group"]) for r in asbuilt if r["is_header"]}
    outside = sum(1 for e in extras if norm_desc(e.get("group", "")) not in groups)
    return "asbuilt" if outside / len(extras) > 0.5 else "merged"


class _Tee:
    """同时写到屏幕与日志文件，让批处理窗口能实时看到进度、又保留日志。"""

    def __init__(self, *streams):
        self.streams = streams

    def write(self, s):
        for st in self.streams:
            try:
                st.write(s)
                st.flush()
            except Exception:                  # noqa: BLE001
                pass

    def flush(self):
        for st in self.streams:
            try:
                st.flush()
            except Exception:                  # noqa: BLE001
                pass


def progress_line(done, total, label=""):
    """行式 ASCII 进度条：CMD 里实时可见，日志里也逐行留痕（不用 \\r 覆盖）。"""
    if total <= 0:
        return
    done = max(0, min(done, total))
    width = 20
    filled = int(width * done / total)
    bar = "#" * filled + "." * (width - filled)
    print("[%s] %3d%% (%d/%d) %s" % (bar, 100 * done // total, done, total, label),
          flush=True)


def main(argv=None):
    p = argparse.ArgumentParser(description="若干 DWG -> 自动判别并生成材料平差 xls")
    p.add_argument("dwgs", nargs="+",
                   help="DWG/DXF 文件或含图纸的文件夹（可多个、可混合；支持直接拖入）")
    p.add_argument("--out-dir", default=None,
                   help="输出目录；省略且拖入文件夹时输出到该文件夹下的「材料平差输出」")
    p.add_argument("--no-recursive", action="store_true",
                   help="只扫描文件夹顶层，不递归子目录")
    p.add_argument("--cache-dir", default=None, help="实体缓存目录，可复用避免重复读 CAD")
    p.add_argument("--template", default="summary",
                   help="summary（默认）/ sample / 任意 .xls 路径")
    p.add_argument("--row-base", default="auto",
                   choices=("auto", "asbuilt", "design", "merged"))
    p.add_argument("--match-mode", default="group", choices=("group", "global", "hybrid"))
    # 兼容旧命令：v1.5.0 起默认且强制忽略竣工图内的施工图材料表，此参数不再有作用
    p.add_argument("--no-supplement", action="store_true", help=argparse.SUPPRESS)
    p.add_argument("--no-fill-summary", action="store_true", help="不同步 2-汇总")
    p.add_argument("--report-dir", default=None, help="核对报告目录（默认同输出目录）")
    p.add_argument("--dry-run", action="store_true", help="只打印判别与生成计划")
    p.add_argument("--log-file", default=None,
                   help="同时把进度输出到屏幕和该日志文件（run_batch.bat 用）")
    args = p.parse_args(argv)

    # 日志 tee：批处理窗口实时显示进度，同时完整写入 run_batch.log
    if args.log_file:
        try:
            log_f = open(args.log_file, "w", encoding="utf-8", errors="replace")
            sys.stdout = _Tee(sys.stdout, log_f)
            sys.stderr = _Tee(sys.stderr, log_f)
        except OSError:                        # noqa: BLE001  日志不可写则退化为仅屏幕
            pass

    # 文件夹参数递归展开成图纸清单（批处理拖入的核心入口）
    raw_inputs = args.dwgs
    dwgs = expand_inputs(raw_inputs, recursive=not args.no_recursive)
    if not dwgs:
        print("未找到 .dwg / .dxf 文件，请检查拖入的文件夹。")
        return 1

    # 默认成品归档到「每张图纸所在目录/材料平差输出」，绝不写脚本目录/当前目录。
    # 仅当用户显式给了 --out-dir 时统一覆盖（CLI 场景保留灵活性）。
    force_out = os.path.abspath(args.out_dir) if args.out_dir else None
    force_report = os.path.abspath(args.report_dir) if args.report_dir else None
    # 核对报告是给 AI 内部核对用的中间产物，默认写到系统临时目录，
    # 不出现在成品目录（用户只看到 .xls）。显式 --report-dir 仍可覆盖。
    default_report_dir = os.path.join(tempfile.gettempdir(), "dwg2xls-reports")

    def dirs_for(info):
        d = force_out or info.get("out_dir") or os.path.dirname(info["path"])
        r = force_report or default_report_dir
        return d, r

    print("图纸 %d 张（成品归档到 %s）" % (
        len(dwgs), force_out or "各自图纸目录的「材料平差输出」"))

    # 只在确实需要读 DWG 时启动 CAD，并且整批只启动一次后复用
    need_cad = False
    for path in dwgs:
        if not os.path.isfile(path):
            continue
        if cache_is_fresh(path, args.cache_dir):
            continue
        need_cad = True
        break
    app, owns = open_cad_app() if need_cad else (None, False)
    infos = []
    zero_table_drawings = []               # 未识别到材料表的图纸，末尾统一汇总
    try:
        for idx, path in enumerate(dwgs, 1):
            name_i = os.path.basename(path)
            if not os.path.isfile(path):
                print("[跳过] 文件不存在：%s" % path)
                continue
            try:
                info = load_drawing(path, args.cache_dir, app=app)
            except Exception as exc:                   # noqa: BLE001
                print("[跳过] %s 读取失败：%s" % (name_i, str(exc)[:120]))
                continue
            infos.append(info)
            if info.get("zero_tables"):
                zero_table_drawings.append(name_i)
            tag = "竣工" if info["stage"] == "asbuilt" else "施工"
            progress_line(idx, len(dwgs), "[%s] %s 表%d %s" % (
                tag, name_i[:24], len(info["tables"]), info["no"] or ""))
    finally:
        close_cad_app(app, owns)
    if not infos:
        print("没有可用图纸。")
        return 1

    # 读图阶段与生成阶段之间留出明显间隔，避免两段进度条连在一起难分辨
    print()
    print("======== 读图完成，开始匹配竣工 / 施工图并生成材料平差表 ========")

    # 一张施工图 DWG 可能混放多个工程的材料表（实测师大四食堂 1 张图含 13 个工程）。
    # 这种「多工程图」不能只归到某一个组，而应作为共享数据源供各竣工图按工程名筛选。
    all_design = [i for i in infos if i["stage"] == "design"]
    multi_design = {}
    for i in all_design:
        projects = sorted({t["project"] for t in i["tables"] if t.get("project")})
        # 「多工程」判定放宽：图内含多个工程名即为候选，随后用
        # 「能匹配到几张不同的竣工图」在分组后最终确认（见下方 multi_targets）。
        multi_design[id(i)] = len(projects) > 1
    # 分组时排除多工程施工图，避免它被当成单工程图只服务某一个组。
    # 但最终是否启用多工程筛选，还要看它能否匹配到 ≥2 张不同的竣工图：
    # 只有 1 个匹配目标时说明是单工程图（图内工程名写法不一），走原流程。
    ab_infos = [i for i in infos if i["stage"] == "asbuilt"]
    for i in all_design:
        if not multi_design.get(id(i)):
            continue
        projects = sorted({t["project"] for t in i["tables"] if t.get("project")})
        # 工程名之间若互为子串（同一工程的不同写法，如「XX低压燃气工程」与
        # 「XX燃气工程」）→ 单工程图，走原流程；若彼此独立（不同店铺名）→ 多工程图，
        # 即使只提供 1 张竣工图也要按工程名筛选其中对应部分。
        cores = [_core_name(p) for p in projects]
        nested = any(a in b or b in a
                     for k, a in enumerate(cores) for b in cores[k + 1:])
        hit_ab = set()
        for ab in ab_infos:
            for t in i["tables"]:
                if _project_match(t.get("project", ""), ab["name"], ab["no"],
                                  t.get("no", "")):
                    hit_ab.add(ab["path"])
                    break
        if nested or not hit_ab:
            multi_design[id(i)] = False
        else:
            print("  [多工程施工图] %s：含 %d 个独立工程，%d 个可对应到竣工图" % (
                os.path.basename(i["path"]), len(projects), len(hit_ab)))
    single_infos = [i for i in infos if not multi_design.get(id(i))]
    groups = group_drawings(single_infos)

    results = []
    gen_total = len(groups)
    gen_done = 0
    for no, group in groups.items():
        gen_done += 1
        ab = [i for i in group if i["stage"] == "asbuilt"]
        sg = [i for i in group if i["stage"] == "design"]
        cur_name = _name_of(group, "" if (not no or no.startswith("_")) else no)
        # 竣工图为基准：把多工程施工图也纳入本组的候选数据源，
        # 由 rows_of 按本组工程名筛选其中属于本工程的表。
        if ab:
            target_name = _name_of(group, no if not no.startswith("_") else "")
            for i in all_design:
                if multi_design.get(id(i)) and i not in sg:
                    sg = sg + [i]
        if not ab:
            print("  [跳过] %s：未找到竣工图，无法生成材料平差表" % (no or "(未编号)"))
            continue
        name = _name_of(group, "" if no.startswith("_") else no)
        proj_no = "" if no.startswith("_") else no
        base = clean_filename((proj_no + "+" + name) if proj_no else name)
        # 以本组竣工图所在目录为准归档（竣工图是成品基准）
        group_out, group_report = dirs_for(ab[0] if ab else group[0])
        if not args.dry_run:
            os.makedirs(group_out, exist_ok=True)
            os.makedirs(group_report, exist_ok=True)
        out = os.path.join(group_out, base + ".xls")

        asbuilt = []
        for i in ab:
            # 竣工图 DWG 常附带其它工程的竣工/施工图材料表（实测 IE26I004 附带
            # 「横三路」工程的表）。以本图工程名筛表，只取本工程那一张。
            asbuilt.extend(rows_of(i, "竣工图", target_project=i["name"],
                                   target_no=i["no"], keep_unmarked_project=True))
        single = not sg
        if single:
            design = [dict(r) for r in asbuilt]      # 施工图量 = 竣工图量
        else:
            # 施工图数据只取用户发送的施工图文件；竣工图 DWG 内残留的
            # 施工图版材料表是编制时留下的旧数据，一律忽略。
            design = []
            for i in sg:
                # 多工程施工图按本组竣工图的工程名/编号筛选对应材料表；
                # 普通单工程施工图保持原行为（target_project=None 取全部表）。
                if multi_design.get(id(i)):
                    design.extend(rows_of(i, "施工图", target_project=target_name,
                                          target_no=proj_no))
                else:
                    design.extend(rows_of(i, "施工图"))
            if not design:
                print("  [提示] %s：施工图文件未检测到本工程的材料表，施工图数量将全部留空；"
                      "请确认施工图是否为最终版。" % base)

        base_mode = (choose_row_base(asbuilt, design, single)
                     if args.row_base == "auto" else args.row_base)
        rows, extras, stats = MATCHERS[base_mode](asbuilt, design, mode=args.match_mode)
        stats.update({"project_no": proj_no, "project_name": name,
                      "row_base": base_mode, "single_asbuilt": single,
                      "design_source": "用户施工图文件（竣工图内施工图材料表已忽略）",
                      "drawings": [os.path.basename(i["path"]) for i in group]})
        label = "竣工%d 施工%d -> %d行 %s" % (
            len(asbuilt), len(design), len(rows), base[:20])
        if args.dry_run:
            results.append({"file": out, "stats": stats})
            progress_line(gen_done, gen_total, label)
            continue
        write_xls(resolve_template(args.template), out, rows,
                  fill_summary=not args.no_fill_summary)
        rep = os.path.join(group_report, base + "-核对报告.json")
        with open(rep, "w", encoding="utf-8") as f:
            json.dump({"stats": stats, "rows": rows, "extras": extras}, f,
                      ensure_ascii=False, indent=1)
        results.append({"file": out, "report": rep, "stats": stats})
        progress_line(gen_done, gen_total, label)

    print("识别：%d 张图纸，%s%d 份成品" % (
        len(infos), "计划 " if args.dry_run else "", len(results)))
    if zero_table_drawings:
        print("  未识别：%s" % "、".join(zero_table_drawings))
    if not results:
        print("没有生成任何文件。")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
