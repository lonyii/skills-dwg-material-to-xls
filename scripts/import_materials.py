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

匹配规则：以竣工行顺序为基准，在施工图同分组内按物料描述精确匹配并「消耗」，
因此同名物料出现在不同分组时不会互相串位。竣工图中施工图没有的条目，
仍写入描述与单位，数量留空（与成品模板一致）。

写入引擎使用本机 Excel COM：先复制模板文件再打开写入，
从而完整保留原模板的所有工作表、公式、定义名、宏与打印设置。

CLI：
  python import_materials.py --asbuilt 竣工图.dwg --design 施工图.dwg \
      [--template summary|sample|模板.xls] [--out 输出.xls] [--report 核对.json] [--dry-run]
  python import_materials.py --asbuilt-rows rows_jg.json --design-rows rows_sg.json --dry-run
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dwg_tables import extract_rows, norm_desc  # noqa: E402

DEFAULT_SHEET = "1-初步比对"
SUMMARY_SHEET = "2-汇总"
FIRST_ROW = 4
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS_DIR = os.path.join(SKILL_DIR, "assets")
TEMPLATE_SUMMARY = os.path.join(ASSETS_DIR, "材料汇总模板-加强比对版.xls")
TEMPLATE_SAMPLE = os.path.join(ASSETS_DIR, "材料平差模板.xls")
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
def match_rows(asbuilt, design, mode="group"):
    """返回 (写入行列表, 施工图多余项列表, 统计)。"""
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

        cand = None
        if mode in ("group", "hybrid"):
            for i, d in enumerate(design):
                if i in used or d["is_header"]:
                    continue
                if norm_desc(d["desc"]) == key and norm_desc(d["group"]) == norm_desc(a["group"]):
                    cand = i
                    break
        if cand is None and mode in ("global", "hybrid"):
            for i, d in enumerate(design):
                if i in used or d["is_header"]:
                    continue
                if norm_desc(d["desc"]) == key:
                    cand = i
                    break

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
    stats = {
        "asbuilt_rows": len(asbuilt),
        "asbuilt_items": sum(1 for r in asbuilt if not r["is_header"]),
        "design_rows": len(design),
        "design_items": sum(1 for r in design if not r["is_header"]),
        "matched": sum(1 for r in out if r["matched"]),
        "unmatched": sum(1 for r in out if r["matched"] is False),
        "design_extras": len(extras),
    }
    return out, extras, stats


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

    asbuilt = _rows_from_source(args, "asbuilt")
    design = _rows_from_source(args, "design")
    rows, extras, stats = match_rows(asbuilt, design, mode=args.match_mode)

    print("竣工图：%d 行（材料 %d 项） | 施工图：%d 行（材料 %d 项）" % (
        stats["asbuilt_rows"], stats["asbuilt_items"],
        stats["design_rows"], stats["design_items"]))
    print("写入 %d 行；匹配到施工图数量 %d 项；竣工图独有 %d 项；施工图多余 %d 项" % (
        len(rows), stats["matched"], stats["unmatched"], stats["design_extras"]))
    if stats["unmatched"]:
        print("\n[竣工图独有（施工图数量留空）]")
        for r in rows:
            if r["matched"] is False:
                print("  -", r["a_desc"].replace("\n", " / ")[:60])
    if extras:
        print("\n[施工图多出（未写入，通常是户内材料，需人工确认是否参与平差）]")
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
