# -*- coding: utf-8 -*-

"""
把 DWG 施工图 / 竣工图材料表分别导入「1-初步比对」工作表，生成新的 .xls，
并可选把明细同步到「2-汇总」，供打印版、变更估价等下游表公式联动。

列映射（从第 4 行开始，逐行以竣工图材料表为基准）：
  1-初步比对：A 编号 | B 标记 | C 竣工图-物料描述 | D 单位 | E 数量
              F/G/H 材料平差（不动）
              I 标记 | J 施工图-物料描述 | K 单位 | L 数量
              M/N 领量/退量（不动） | O 异同（保留模板公式，自动重算）
  2-汇总：    C 物料描述 | D 单位 | E 竣工数量 | I 施工图数量
              A 编号沿用模板、B 标记与 F/G 领退量不动，H/J/K/L/M/N 全部保留模板公式

匹配规则：默认以竣工行顺序为基准（--row-base design 可改为施工图行序），
在同分组内按物料描述精确匹配并「消耗」，同名物料跨分组不会串位。
未匹配项仍写描述与单位、数量留空。

施工图数据源规则：竣工图 DWG 在编制时常保留施工图版材料表，那是旧版且不准确的数据，
一律忽略；施工图数量只取用户发送的施工图文件。若用户只给了竣工图，
则由 auto_export.py 按「施工图量=竣工图量」处理。

写入引擎使用本机 Excel COM：先复制模板文件再打开写入，
从而完整保留原模板的所有工作表、公式、定义名、宏与打印设置。

CLI：
  python import_materials.py --asbuilt 竣工图.dwg --design 施工图.dwg \
      [--template summary|sample|模板.xls] [--out 输出.xls] [--report 核对.json] [--dry-run]
  python import_materials.py --asbuilt-rows rows_jg.json --design-rows rows_sg.json --dry-run
"""

from __future__ import annotations

__author__ = "lonyii"
__email__ = "lonyii@qq.com"

import argparse
import difflib
import json
import os
import shutil
import sys

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding=sys.stdout.encoding or "utf-8",
                               errors="replace")
    except Exception:                          # noqa: BLE001
        pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dwg_tables import extract_rows, norm_desc  # noqa: E402

DEFAULT_SHEET = "1-初步比对"
SUMMARY_SHEET = "2-汇总"
FIRST_ROW = 4
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS_DIR = os.path.join(SKILL_DIR, "assets")
def _asset_file(ascii_name, legacy_cn_name):
    """优先 ASCII 文件名（兼容任意解压工具），找不到再回退旧中文名。"""
    p = os.path.join(ASSETS_DIR, ascii_name)
    if os.path.isfile(p):
        return p
    return os.path.join(ASSETS_DIR, legacy_cn_name)


TEMPLATE_SUMMARY = _asset_file("template_summary.xls", "材料汇总模板-加强比对版.xls")
TEMPLATE_SAMPLE = _asset_file("template_sample.xls", "材料平差模板.xls")
DEFAULT_TEMPLATE = TEMPLATE_SUMMARY      # 生产用空白加强版：115 行明细 + 完整汇总/打印版
TEMPLATE_ALIASES = {
    "summary": TEMPLATE_SUMMARY,          # 空白加强比对版（9 张表，明细容量 115 行）
    "sample": TEMPLATE_SAMPLE,            # IC24H201 成品文件（用于复现与逐格验收）
}


def resolve_template(value):
    """模板支持别名 summary / sample，也支持任意 .xls 路径。"""
    return TEMPLATE_ALIASES.get((value or "").strip().lower(), os.path.abspath(value))


# --------------------------------------------------------------------------- #
# 匹配
# --------------------------------------------------------------------------- #
def _find_cand(pool, used, key, group, mode):
    """group：严格同分组；global：忽略分组；hybrid：先同组再全局。"""
    if mode in ("group", "hybrid"):
        for i, d in enumerate(pool):
            if i in used or d["is_header"]:
                continue
            if norm_desc(d["desc"]) == key and norm_desc(d["group"]) == group:
                return i
    if mode in ("global", "hybrid"):
        for i, d in enumerate(pool):
            if i in used or d["is_header"]:
                continue
            if norm_desc(d["desc"]) == key:
                return i
    return None


def _stats(asbuilt, design, out, extras):
    return {
        "asbuilt_rows": len(asbuilt),
        "asbuilt_items": sum(1 for r in asbuilt if not r["is_header"]),
        "design_rows": len(design),
        "design_items": sum(1 for r in design if not r["is_header"]),
        "matched": sum(1 for r in out if r["matched"]),
        "unmatched": sum(1 for r in out if r["matched"] is False),
        "design_extras": len(extras),
    }


def match_rows(asbuilt, design, mode="group"):
    """以竣工图材料表行顺序为基准（默认）。"""
    used = set()
    out = []
    for a in asbuilt:
        key = norm_desc(a["desc"])
        if a["is_header"]:
            out.append({
                "seq": len(out) + 1,
                "a_desc": a["desc"], "a_unit": "", "a_qty": None,
                "j_desc": a["desc"], "k_unit": "", "l_qty": None,
                "is_header": True, "group": a["group"], "matched": None,
            })
            continue
        cand = _find_cand(design, used, key, norm_desc(a["group"]), mode)
        if cand is not None:
            used.add(cand)
            d = design[cand]
            j_desc, k_unit, l_qty = d["desc"], d["unit"], d["qty"]
        else:
            j_desc, k_unit, l_qty = a["desc"], a["unit"], None
        out.append({
            "seq": len(out) + 1,
            "a_desc": a["desc"], "a_unit": a["unit"], "a_qty": a["qty"],
            "j_desc": j_desc, "k_unit": k_unit, "l_qty": l_qty,
            "is_header": False, "group": a["group"], "matched": cand is not None,
        })
    extras = [d for i, d in enumerate(design) if i not in used and not d["is_header"]]
    return out, extras, _stats(asbuilt, design, out, extras)


def match_rows_design_base(asbuilt, design, mode="group"):
    """
    以施工图材料表行顺序为基准：竣工数量回填到 E；施工图独有项 C 列照写、E 留空；
    竣工独有项追加到表尾、L 留空。适用于成品按施工图清单编排的项目。
    """
    used = set()
    out = []
    # 多页续表常重复同一分组标题（如两页都以「低压地上材料：」开头），
    # 只保留最后一次出现，避免比对表里出现多余空标题行。
    last_header = {}
    for i, d in enumerate(design):
        if d["is_header"]:
            last_header[norm_desc(d["desc"])] = i
    for i, d in enumerate(design):
        key = norm_desc(d["desc"])
        if d["is_header"]:
            if last_header.get(key) != i:
                continue
            out.append({
                "seq": len(out) + 1,
                "a_desc": d["desc"], "a_unit": "", "a_qty": None,
                "j_desc": d["desc"], "k_unit": "", "l_qty": None,
                "is_header": True, "group": d["group"], "matched": None,
            })
            continue
        cand = _find_cand(asbuilt, used, key, norm_desc(d["group"]), mode)
        if cand is not None:
            used.add(cand)
            a = asbuilt[cand]
            a_desc, a_unit, a_qty = a["desc"], a["unit"], a["qty"]
        else:
            a_desc, a_unit, a_qty = d["desc"], d["unit"], None
        out.append({
            "seq": len(out) + 1,
            "a_desc": a_desc, "a_unit": a_unit, "a_qty": a_qty,
            "j_desc": d["desc"], "k_unit": d["unit"], "l_qty": d["qty"],
            "is_header": False, "group": d["group"], "matched": cand is not None,
        })
    extras = []
    for i, a in enumerate(asbuilt):
        if i in used or a["is_header"]:
            continue
        extras.append(a)
        out.append({
            "seq": len(out) + 1,
            "a_desc": a["desc"], "a_unit": a["unit"], "a_qty": a["qty"],
            "j_desc": a["desc"], "k_unit": a["unit"], "l_qty": None,
            "is_header": False, "group": a["group"], "matched": None,
            "appended": True,
        })
    return out, extras, _stats(asbuilt, design, out, extras)


def _row_key(r):
    return ("H:" if r["is_header"] else "D:") + norm_desc(r["desc"])


def _dedupe_headers(rows):
    """多页续表重复同一分组标题时，只保留最后一次出现。"""
    last = {}
    for i, r in enumerate(rows):
        if r["is_header"]:
            last[norm_desc(r["desc"])] = i
    return [r for i, r in enumerate(rows)
            if not r["is_header"] or last.get(norm_desc(r["desc"])) == i]


def match_rows_merged(asbuilt, design, mode="group"):
    """
    以施工图序列为骨架，用序列对齐把竣工独有项插回它在竣工表中的相对位置，
    得到「两表并集、各自保持原顺序」的行序。适用于成品按并集人工编排的项目。
    """
    asbuilt = _dedupe_headers(asbuilt)
    design = _dedupe_headers(design)
    ak = [_row_key(r) for r in asbuilt]
    dk = [_row_key(r) for r in design]
    sm = difflib.SequenceMatcher(None, dk, ak, autojunk=False)
    out, extras = [], []

    def emit(d=None, a=None):
        if d is not None and a is not None:
            h = d["is_header"] or a["is_header"]
            out.append({
                "seq": len(out) + 1,
                "a_desc": d["desc"] if h else a["desc"],
                "a_unit": "" if h else a["unit"], "a_qty": None if h else a["qty"],
                "j_desc": d["desc"], "k_unit": "" if h else d["unit"],
                "l_qty": None if h else d["qty"],
                "is_header": h, "group": d["group"] or a["group"],
                "matched": None if h else True,
            })
        elif d is not None:
            h = d["is_header"]
            out.append({
                "seq": len(out) + 1,
                "a_desc": d["desc"], "a_unit": "" if h else d["unit"], "a_qty": None,
                "j_desc": d["desc"], "k_unit": "" if h else d["unit"],
                "l_qty": None if h else d["qty"],
                "is_header": h, "group": d["group"],
                "matched": None if h else False,
            })
        else:
            h = a["is_header"]
            extras.append(a)
            out.append({
                "seq": len(out) + 1,
                "a_desc": a["desc"], "a_unit": "" if h else a["unit"],
                "a_qty": None if h else a["qty"], "j_desc": a["desc"],
                "k_unit": "" if h else a["unit"], "l_qty": None,
                "is_header": h, "group": a["group"],
                "matched": None, "appended": True,
            })

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                emit(design[i1 + k], asbuilt[j1 + k])
        else:
            for d in design[i1:i2]:
                emit(d=d)
            for a in asbuilt[j1:j2]:
                emit(a=a)
    return out, extras, _stats(asbuilt, design, out, extras)


# --------------------------------------------------------------------------- #
# 写 Excel
# --------------------------------------------------------------------------- #
def write_xls(template, out_path, rows, sheet_name=DEFAULT_SHEET, fill_summary=True):
    try:
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError("写入 xls 需要 pywin32：python -m pip install pywin32") from exc
    if not os.path.isfile(template):
        raise FileNotFoundError("找不到模板：%s" % template)

    out_path = os.path.abspath(out_path)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    shutil.copy2(template, out_path)

    XL_UP = -4162
    app = win32.DispatchEx("Excel.Application")
    wb = None
    summary_done = False
    try:
        app.Visible = False
        app.DisplayAlerts = False
        try:
            app.AutomationSecurity = 3        # 禁用宏自动运行
        except Exception:                      # noqa: BLE001
            pass
        wb = app.Workbooks.Open(out_path)
        ws = wb.Worksheets(sheet_name)
        last = FIRST_ROW + len(rows) - 1
        # 明细容量由模板 O 列比对公式的最后一行决定：成品模板 43、加强版 118
        template_last = max(FIRST_ROW, ws.Cells(ws.Rows.Count, 15).End(XL_UP).Row)
        clear_last = max(last, template_last)

        # 只清本技能负责的列，F/G/H（材料平差）与 M/N（领量退量）保持不动
        ws.Range("A%d:E%d" % (FIRST_ROW, clear_last)).ClearContents()
        ws.Range("I%d:L%d" % (FIRST_ROW, clear_last)).ClearContents()

        # 超出模板原有行数时，把 O 列比对公式向下填充
        if last > template_last:
            ws.Range("O%d:O%d" % (template_last, last)).FillDown()

        for i, r in enumerate(rows):
            row = FIRST_ROW + i
            ws.Cells(row, 1).Value = i + 1
            marker = 2 if r["is_header"] and "不用入物资系统" not in r["a_desc"] else None
            ws.Cells(row, 2).Value = marker
            ws.Cells(row, 3).Value = r["a_desc"]
            ws.Cells(row, 4).Value = r["a_unit"] or None
            ws.Cells(row, 5).Value = r["a_qty"]
            ws.Cells(row, 9).Value = marker
            ws.Cells(row, 10).Value = r["j_desc"]
            ws.Cells(row, 11).Value = r["k_unit"] or None
            ws.Cells(row, 12).Value = r["l_qty"]

        # 同步「2-汇总」：C/D/E/I 是常量区，H/J/K/L/M/N 与打印版全部是公式，
        # 只要填入这四列，增加/减少材料、平差量、打印版 VLOOKUP 会自动联动。
        if fill_summary:
            names = [wb.Worksheets(i + 1).Name for i in range(wb.Worksheets.Count)]
            if SUMMARY_SHEET in names:
                sw = wb.Worksheets(SUMMARY_SHEET)
                # 若模板把明细列改成公式（引用别的表），不要覆盖，避免破坏模板设计
                if any(sw.Range("%s%d" % (col, FIRST_ROW)).HasFormula
                       for col in ("C", "D", "E", "I")):
                    print("提示：%s 的明细列含公式，已跳过汇总同步（可检查模板）" % SUMMARY_SHEET)
                else:
                    sum_last = max(FIRST_ROW, sw.Cells(sw.Rows.Count, 8).End(XL_UP).Row)
                    if last > sum_last:
                        for col in ("H", "J", "K", "L", "M", "N"):
                            sw.Range("%s%d:%s%d" % (col, sum_last, col, last)).FillDown()
                    sum_clear = max(last, template_last)
                    sw.Range("C%d:E%d" % (FIRST_ROW, sum_clear)).ClearContents()
                    sw.Range("I%d:I%d" % (FIRST_ROW, sum_clear)).ClearContents()
                    for i, r in enumerate(rows):
                        row = FIRST_ROW + i
                        sw.Cells(row, 3).Value = r["a_desc"]
                        sw.Cells(row, 4).Value = r["a_unit"] or None
                        sw.Cells(row, 5).Value = r["a_qty"]
                        sw.Cells(row, 9).Value = r["l_qty"]
                    summary_done = True
        wb.Save()
        return out_path, summary_done
    finally:
        try:
            if wb is not None:
                wb.Close(True)
        except Exception:                      # noqa: BLE001
            pass
        try:
            app.Quit()
        except Exception:                      # noqa: BLE001
            pass


# --------------------------------------------------------------------------- #
# 数据输入
# --------------------------------------------------------------------------- #
def _rows_from_source(args, prefix):
    """优先使用现成的 rows.json；否则调用 dwg_tables 从 DWG/DXF 提取。"""
    rows_json = getattr(args, prefix + "_rows")
    if rows_json:
        return json.load(open(rows_json, encoding="utf-8"))["rows"]
    drawing = getattr(args, prefix + "_path") or "drawing"
    result = extract_rows(
        drawing,
        from_json=getattr(args, prefix + "_cache"),
        cache=getattr(args, prefix + "_save_cache"),
        stage=getattr(args, prefix + "_stage"),
        tables=getattr(args, prefix + "_tables"),
    )
    return result["rows"]


def _ids(s):
    return [int(x) for x in s.split(",") if x.strip()] if s else None


def main(argv=None):
    p = argparse.ArgumentParser(description="DWG 施工图/竣工图材料表 -> 材料平差 xls")
    p.add_argument("--asbuilt", dest="asbuilt_path", help="竣工图 DWG")
    p.add_argument("--design", dest="design_path", help="施工图 DWG")
    p.add_argument("--asbuilt-rows", help="已提取好的竣工图 rows.json（离线复用）")
    p.add_argument("--design-rows", help="已提取好的施工图 rows.json（离线复用）")
    p.add_argument("--asbuilt-cache", help="竣工图实体缓存 JSON（--from-json）")
    p.add_argument("--design-cache", help="施工图实体缓存 JSON（--from-json）")
    p.add_argument("--asbuilt-save-cache", help="把竣工图实体另存为缓存")
    p.add_argument("--design-save-cache", help="把施工图实体另存为缓存")
    p.add_argument("--asbuilt-stage", default="竣工图")
    p.add_argument("--design-stage", default="施工图")
    p.add_argument("--asbuilt-tables", type=_ids, default=None)
    p.add_argument("--design-tables", type=_ids, default=None)
    p.add_argument("--match-mode", default="group", choices=("group", "global", "hybrid"))
    p.add_argument("--row-base", default="asbuilt",
                   choices=("asbuilt", "design", "merged"),
                   help="输出行序基准：asbuilt（默认，按竣工图表）/ "
                        "design（按施工图表，竣工独有追加表尾）/ "
                        "merged（两表并集，独有项按各自原顺序交错插入）")
    # 兼容旧命令：v1.5.0 起默认且强制忽略竣工图内的施工图材料表，此参数不再有作用
    p.add_argument("--no-supplement", action="store_true", help=argparse.SUPPRESS)
    p.add_argument("--template", default="summary",
                   help="模板：summary（默认，空白加强比对版汇总模板）/ "
                        "sample（IC24H201 成品模板，用于复现验收）/ 任意 .xls 路径")
    p.add_argument("--no-fill-summary", action="store_true",
                   help="只写「1-初步比对」，不同步「2-汇总」")
    p.add_argument("--out", default=None, help="输出 .xls 路径")
    p.add_argument("--sheet", default=DEFAULT_SHEET)
    p.add_argument("--report", default=None, help="核对报告 JSON 路径")
    p.add_argument("--dry-run", action="store_true", help="只核对匹配，不写 Excel")
    args = p.parse_args(argv)

    if not (args.asbuilt_rows or args.asbuilt_path or args.asbuilt_cache):
        p.error("需要 --asbuilt 或 --asbuilt-rows / --asbuilt-cache")
    if not (args.design_rows or args.design_path or args.design_cache):
        p.error("需要 --design 或 --design-rows / --design-cache")

    # 施工图数据只取用户发送的施工图文件；竣工图 DWG 里残留的施工图版材料表一律忽略
    asbuilt = _rows_from_source(args, "asbuilt")
    design = _rows_from_source(args, "design")
    if not design:
        print("提示：施工图资料未提取到材料表，施工图数量将全部留空（请确认施工图文件）。")

    design_base = args.row_base in ("design", "merged")
    matcher = {"asbuilt": match_rows, "design": match_rows_design_base,
               "merged": match_rows_merged}[args.row_base]
    rows, extras, stats = matcher(asbuilt, design, mode=args.match_mode)
    stats["row_base"] = args.row_base
    stats["design_source"] = "用户施工图文件（竣工图内施工图材料表已忽略）"

    print("行序基准：%s | 竣工图 %d 行（材料 %d 项） | 施工图 %d 行（材料 %d 项）" % (
        args.row_base, stats["asbuilt_rows"], stats["asbuilt_items"],
        stats["design_rows"], stats["design_items"]))
    print("写入 %d 行；匹配 %d 项；%s %d 项；%s %d 项" % (
        len(rows), stats["matched"],
        "施工图独有" if design_base else "竣工图独有", stats["unmatched"],
        "竣工独有" if design_base else "施工图多余", stats["design_extras"]))
    if stats["unmatched"]:
        print("\n[%s]" % ("施工图独有（竣工数量 E 留空，已写入）" if design_base
                          else "竣工图独有（施工图数量留空）"))
        for r in rows:
            if r["matched"] is False:
                print("  -", r["a_desc"].replace("\n", " / ")[:60])
    if extras:
        print("\n[%s]" % ("竣工独有（已写入，施工图数量留空）" if design_base
                          else "施工图多出（未写入，通常是户内材料，需人工确认是否参与平差）"))
        for d in extras[:20]:
            print("  -", d["desc"].replace("\n", " / ")[:60], d["unit"], d["qty"])
        if len(extras) > 20:
            print("  ... 另有 %d 项" % (len(extras) - 20))

    if args.report:
        with open(args.report, "w", encoding="utf-8") as f:
            json.dump({"stats": stats, "rows": rows, "extras": extras},
                      f, ensure_ascii=False, indent=1)
        print("\n核对报告：%s" % os.path.abspath(args.report))

    if args.dry_run:
        print("\n[dry-run] 未写入 Excel。")
        return 0

    out = args.out
    if not out:
        base = os.path.splitext(os.path.basename(args.asbuilt_path or args.asbuilt_rows or "材料平差"))[0]
        out = os.path.join(os.getcwd(), base + "-材料平差.xls")
    template = resolve_template(args.template)
    path, summary_done = write_xls(template, out, rows, sheet_name=args.sheet,
                                   fill_summary=not args.no_fill_summary)
    print("\n已生成：%s" % path)
    print("模板：%s | 明细 %d 行 | 2-汇总：%s" % (
        os.path.basename(template), len(rows), "已同步" if summary_done else "未改动"))
    print("提示：打开后请核对「1-初步比对」数量与「4-汇总打印版 (新版)」的联动结果。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
