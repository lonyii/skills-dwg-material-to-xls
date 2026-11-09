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
"""
from __future__ import annotations

import argparse
import difflib
import json
import os
import re
import sys

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
    norm_desc, open_cad_app, parse_table,
)
from import_materials import (  # noqa: E402
    match_rows, match_rows_design_base, match_rows_merged, merge_design_rows,
    resolve_template, write_xls,
)

PROJ_RE = re.compile(r"[A-Z]{2}\d{2}[A-Z]\d{3}", re.I)
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
    """缓存必须比 DWG 新；图纸被改动时自动重读，避免批处理用到过期数据。"""
    if not cache_dir:
        return False
    cp = cache_path(dwg, cache_dir)
    return os.path.isfile(cp) and os.path.getmtime(cp) >= os.path.getmtime(dwg)


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
    施工图图签常印有通用图集号/其它编号，不能随便取第一个匹配。
    """
    m = PROJ_RE.search(os.path.basename(path))
    if m:
        return m.group(0).upper()
    m = PROJ_RE.search(_label_value(ents, ("项目编号", "工程编号", "项目号"), 20))
    if m:
        return m.group(0).upper()
    for e in ents:
        m = PROJ_RE.search(clean_text(e["t"]))
        if m:
            return m.group(0).upper()
    return ""


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
    if cache_dir:
        os.makedirs(cache_dir, exist_ok=True)
        cp = cache_path(path, cache_dir)
        if cache_is_fresh(path, cache_dir):
            ents, meta = load_entities(path, from_json=cp)
        else:
            ents, meta = load_entities(path, cache=cp, app=app)
    else:
        ents, meta = load_entities(path, app=app)
    tables = detect_tables(ents)
    return {
        "path": path,
        "ents": ents,
        "tables": tables,
        "stage": classify_stage(path, ents, tables),
        "no": find_project_no(ents, path),
        "name": find_project_name(ents, path),
        "meta": meta,
    }


def rows_of(info, stage):
    ents = info["ents"]
    chosen = [t for t in info["tables"] if t["stage"] == stage]
    if not chosen and stage == "施工图":
        chosen = info["tables"]          # 阶段缺失时退化为全部表
    rows = []
    for t in sorted(chosen, key=lambda x: x["id"]):
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
    p.add_argument("--no-supplement", action="store_true",
                   help="不从竣工图 DWG 附带的施工图版材料表补充施工图数据")
    p.add_argument("--no-fill-summary", action="store_true", help="不同步 2-汇总")
    p.add_argument("--report-dir", default=None, help="核对报告目录（默认同输出目录）")
    p.add_argument("--dry-run", action="store_true", help="只打印判别与生成计划")
    args = p.parse_args(argv)

    # 文件夹参数递归展开成图纸清单（批处理拖入的核心入口）
    raw_inputs = args.dwgs
    print("输入路径：%d 个" % len(raw_inputs))
    dwgs = expand_inputs(raw_inputs, recursive=not args.no_recursive)
    if not dwgs:
        print("未找到 .dwg / .dxf 文件，请检查拖入的文件夹。")
        return 1
    print("共展开 %d 张图纸" % len(dwgs))

    if args.out_dir:
        out_dir = os.path.abspath(args.out_dir)
    else:
        first_dir = next((os.path.abspath(x) for x in raw_inputs if os.path.isdir(x)), None)
        out_dir = os.path.join(first_dir, "材料平差输出") if first_dir else os.getcwd()
    report_dir = os.path.abspath(args.report_dir or out_dir)
    print("输出目录：%s" % out_dir)
    if not args.dry_run:
        os.makedirs(out_dir, exist_ok=True)
        os.makedirs(report_dir, exist_ok=True)

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
    try:
        for path in dwgs:
            if not os.path.isfile(path):
                print("[跳过] 文件不存在：%s" % path)
                continue
            try:
                info = load_drawing(path, args.cache_dir, app=app)
            except Exception as exc:                   # noqa: BLE001
                print("[跳过] %s 读取失败：%s" % (os.path.basename(path), str(exc)[:120]))
                continue
            infos.append(info)
    finally:
        close_cad_app(app, owns)
    if not infos:
        print("没有可用图纸。")
        return 1

    print("图纸判别（共 %d 张）：" % len(infos))
    for i in infos:
        n_ab = sum(1 for t in i["tables"] if t["stage"] == "竣工图")
        n_sg = sum(1 for t in i["tables"] if t["stage"] == "施工图")
        print("  [%s] %s | 表 %d 张（竣工 %d / 施工 %d）| 编号 %s | 名称 %s" % (
            "竣工图" if i["stage"] == "asbuilt" else "施工图",
            os.path.basename(i["path"]), len(i["tables"]), n_ab, n_sg,
            i["no"] or "-", i["name"] or "-"))

    groups = group_drawings(infos)
    print("\n工程分组：%d 组" % len(groups))
    results = []
    for no, group in groups.items():
        ab = [i for i in group if i["stage"] == "asbuilt"]
        sg = [i for i in group if i["stage"] == "design"]
        if not ab:
            print("  [跳过] %s：未找到竣工图，无法生成材料平差表" % (no or "(未编号)"))
            continue
        name = _name_of(group, "" if no.startswith("_") else no)
        proj_no = "" if no.startswith("_") else no
        base = clean_filename((proj_no + "+" + name) if proj_no else name)
        out = os.path.join(out_dir, base + ".xls")

        asbuilt = []
        for i in ab:
            asbuilt.extend(rows_of(i, "竣工图"))
        single = not sg
        if single:
            design = [dict(r) for r in asbuilt]      # 施工图量 = 竣工图量
        else:
            design = []
            for i in sg:
                design.extend(rows_of(i, "施工图"))
            if not args.no_supplement:
                extra = []
                for i in ab:
                    extra.extend(rows_of(i, "施工图"))
                design, added = merge_design_rows(design, extra)
            else:
                added = 0
        if not single and not args.no_supplement:
            print("  %s：施工图补充 %d 项" % (base, added))

        base_mode = (choose_row_base(asbuilt, design, single)
                     if args.row_base == "auto" else args.row_base)
        rows, extras, stats = MATCHERS[base_mode](asbuilt, design, mode=args.match_mode)
        stats.update({"project_no": proj_no, "project_name": name,
                      "row_base": base_mode, "single_asbuilt": single,
                      "drawings": [os.path.basename(i["path"]) for i in group]})
        print("  %s：竣工 %d 行 / 施工 %d 行 -> 写入 %d 行，行序=%s%s" % (
            base, len(asbuilt), len(design), len(rows), base_mode,
            "（单竣工图：施工图量=竣工图量）" if single else ""))

        if args.dry_run:
            results.append({"file": out, "stats": stats})
            continue
        write_xls(resolve_template(args.template), out, rows,
                  fill_summary=not args.no_fill_summary)
        rep = os.path.join(report_dir, base + "-核对报告.json")
        with open(rep, "w", encoding="utf-8") as f:
            json.dump({"stats": stats, "rows": rows, "extras": extras}, f,
                      ensure_ascii=False, indent=1)
        results.append({"file": out, "report": rep, "stats": stats})
        print("    已生成：%s" % out)

    if not results:
        print("没有生成任何文件。")
        return 1
    print("\n%s共 %d 份材料平差表。" % ("[dry-run] 计划生成 " if args.dry_run else "", len(results)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
