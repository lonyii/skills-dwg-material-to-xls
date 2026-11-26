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
import re
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

# 2026-11-26：v2.5.26 起材料归一默认完全采用模板宏 StartHere 的规则——
# 分组名相同、描述+单位字符串完全相同才配对，不再做别名/模糊归一。
# 下列别名与模糊指纹仅作旧项目兼容，设置环境变量 DWG2XLS_LEGACY_MATCH=1 才启用。
LEGACY_DESC_ALIASES = {
    "闸阀_Z47W-16C_DN300(单闸板)": "闸阀_Z47W-16C_DN300",
    "玻璃标示桩_100mm*100mm*800m": "标识桩",
}
LEGACY_MATCH = os.environ.get("DWG2XLS_LEGACY_MATCH", "") == "1"
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
def _match_fp(desc):
    """匹配指纹：抹平同一物料的防腐工艺/标准号写法差异，保留规格主体。

    实测 NC26I013：竣工写「20#钢无缝管_D273*8_GB/T8163-2018」，施工图写
    「20#钢无缝管_D273*8_双层熔结环氧粉末防腐」，规格相同仅防腐/标准号写法不同，
    精确匹配失败被误判为施工图独有、施工数量丢失。这里把这些等价写法去掉，
    两边指纹一致即可匹配；规格（管径/壁厚/mm 等）完全不动，不会串料。
    """
    # 图签人名常换行追加在描述末尾（实测「...防腐\n张三峰\n李 颖」），
    # 指纹只取第一行物料本身，避免同名项因噪声匹配不上。
    # 注意：必须先按原始 desc 截断换行再 norm_desc——norm_desc 会把 \n 直接删掉，
    # 之后再 split 已无法分离人名。
    x = norm_desc(str(desc).split("\n")[0].split("\\P")[0])
    for p in (r"GB/T\d+-\d+", r"GB\d+-\d+", r"SY/T\d+-\d+",
              r"双层熔结环氧粉末防腐", r"3PE防腐", r"熔结环氧粉末防腐",
              r"单层环氧粉末防腐", r"环氧粉末防腐", r"加强级防腐",
              r"普通级防腐", r"三层PE防腐"):
        x = re.sub(p, "", x, flags=re.I)
    x = re.sub(r"_+", "_", x).strip("_")
    # 实测别名：施工图在型号后补注「（单闸板）」、竣工图简称「标识桩」等，
    # 统一到同一描述后再匹配，避免同一物料被拆成“增加/减少”两行。
    return LEGACY_DESC_ALIASES.get(x, x) if LEGACY_MATCH else x


def _macro_key(r):
    """StartHere 宏的物料键：C&D（描述+单位）原始字符串拼接。

    宏里 lN = C & D、rn = J & K，只在同分组 lb=rb 时按该字符串精确查找；
    除 StartHere 开头的 %%d→°（clean_text 已处理）外不做任何归一。
    """
    if r.get("is_header"):
        return r.get("desc", "")
    return str(r.get("desc", "")) + str(r.get("unit", ""))


def _find_cand(pool, used, key, group, mode, desc=None):
    """group：严格同分组；global：忽略分组；hybrid：先同组再全局。

    desc 给定时，精确匹配失败后再用「匹配指纹」找一次（防腐/标准号写法差异）。
    """
    if mode in ("group", "hybrid"):
        for i, d in enumerate(pool):
            if i in used or d["is_header"]:
                continue
            # StartHere 规则：描述+单位（宏 lN/rn = C&D / J&K）完全相同，分组完全相同
            if _macro_key(d) == key and d["group"] == group:
                return i
    if mode in ("global", "hybrid"):
        for i, d in enumerate(pool):
            if i in used or d["is_header"]:
                continue
            if d["desc"] == key:
                return i
    # 指纹回退：仅当调用方传入原始描述、且精确匹配全部失败时启用。
    # 作用有二：①防腐/标准号写法差异（fp!=key）；②描述相同但两侧分组标注
    # 不同（实测 NC26I013 保护沟：施工图归入「不入系统材料」组、竣工无分组，
    # 严格同组匹配失败）。两种情况都按描述指纹再找一次。
    if desc is not None and LEGACY_MATCH:
        fp = _match_fp(desc)
        for i, d in enumerate(pool):
            if i in used or d["is_header"]:
                continue
            if _match_fp(d["desc"]) == fp:
                return i
    return None


def match_rows_starthere(asbuilt, design, mode="group"):
    """逐动作复刻模板宏 StartHere（施工图及竣工图材料同步.bas）。

    复刻范围：CheckBefore 分组校验、O 列 =IF(C=J,"相同","") 判等、
    CompareAndSync 的同分部 Cut/Insert 移动与整列插入、StartHere 末尾的
    A 列 FillDown、按 A 升序排序、重排序号。isLeftorRight=-1。
    输入会先按两侧标题生成相同骨架（宏 CheckBefore 的前提）。
    """
    def cell(r, unit=""):
        if r is None:
            return None
        return dict(r)

    def txt(r):
        if r is None:
            return ""
        return str(r.get("desc", "")) + str(r.get("unit") or "")

    def desc(r):
        return "" if r is None else str(r.get("desc", ""))

    def hdr(r):
        return bool(r and r.get("is_header"))

    def grp(r):
        return "" if r is None else str(r.get("group", ""))

    # ---- 1. 生成两侧一致的标题骨架 ----
    ah = [r for r in asbuilt if hdr(r)]
    dh = [r for r in design if hdr(r)]
    titles = []
    for r in ah + dh:
        if r["desc"] not in titles:
            titles.append(r["desc"])
    def bucket(rows):
        d = {"(主体)": []}
        d.update({t: [] for t in titles})
        cur = "(主体)"
        for r in rows:
            if hdr(r):
                cur = r["desc"] if r["desc"] in d else "(主体)"
            d[cur].append(dict(r))
        return d
    am, dm = bucket(asbuilt), bucket(design)
    groups = (["(主体)"] if am.get("(主体)") or dm.get("(主体)") else []) + titles
    left, right = [], []
    for g in groups:
        for side, source in ((left, am), (right, dm)):
            rows = source[g]
            has_header = any(hdr(r) for r in rows)
            if g != "(主体)" and not has_header:
                side.append({"desc": g, "unit": "", "qty": None,
                             "is_header": True, "group": g, "A": None})
            for r in rows:
                rr = dict(r)
                rr["A"] = None
                side.append(rr)
    # 宏 GetRows 取 UsedRange 最后一行；加强版模板 O 列公式预置到第 118 行。
    # 复刻时保留这些尾部空行，否则 rn="" 分支会少执行。
    n = max(len(left), len(right), 118)
    left += [None] * (n - len(left))
    right += [None] * (n - len(right))

    def get(side, i):
        return side[i] if 0 <= i < len(side) else None

    def ins(side, i, value):
        if i > len(side):
            side.extend([None] * (i - len(side)))
        side.insert(i, value)
        return max(len(left), len(right))

    # ---- 2. StartHere 初始化：A 序号、O 公式（只比较 C=J） ----
    def assign_A():
        for i in range(n):
            if left[i] is not None:
                left[i]["A"] = i - 3 + 1       # Excel 行 4 -> A=1
    assign_A()

    # ---- 3. CompareAndSync 循环（每次调用最多做一个动作） ----
    for _ in range(1000):
        changed = False
        lb = rb = ""
        for i in range(n):
            L, R = get(left, i), get(right, i)
            if hdr(L):
                lb = txt(L)
            if hdr(R):
                rb = txt(R)
            lN, rn = txt(L), txt(R)
            # O(i) = IF(C=J,"相同","")
            if desc(L) != desc(R) and lb == rb:
                changed = True
                if not lN:                      # 左侧空：A:G 插入，B:D=I:K
                    newL = dict(R, A=None) if R else None
                    if newL is not None:
                        newL.update({"desc": desc(R), "unit": (R or {}).get("unit", ""),
                                     "qty": None, "is_header": hdr(R)})
                    n = ins(left, i, newL)
                    break
                if not rn:                      # 右侧空：I:N 插入，I:K=B:D
                    newR = dict(L) if L else None
                    if newR is not None:
                        newR.update({"desc": desc(L), "unit": (L or {}).get("unit", ""),
                                     "qty": None, "is_header": hdr(L)})
                    n = ins(right, i, newR)
                    break
                # isLeftorRight=-1：在本分部内向下找左侧匹配行，Cut/Insert 移动
                j = None
                lb_tmp = lb
                for l in range(i, n):
                    LL = get(left, l)
                    if hdr(LL):
                        lb_tmp = txt(L)         # 宏原文用 C(i)&D(i)，不是 C(l)
                    ln_tmp = txt(LL)
                    if lb_tmp == rb:
                        if ln_tmp == rn:
                            j = l
                            break
                    else:
                        break
                if j is not None:
                    row = left.pop(j) if j < len(left) else None
                    if j > len(left):
                        left.extend([None] * (j - len(left)))
                    n = ins(left, i, row)
                    break
                # 找不到：A:G 插入，B:D=I:K（数量留空）
                newL = dict(R, A=None) if R else None
                if newL is not None:
                    newL.update({"desc": desc(R), "unit": (R or {}).get("unit", ""),
                                 "qty": None, "is_header": hdr(R)})
                n = ins(left, i, newL)
                break
            elif lb != rb:
                if lN == lb:                    # 左侧到达下一分部头，复制右侧项目到左侧
                    newL = dict(R, A=None) if R else None
                    if newL is not None:
                        newL.update({"desc": desc(R), "unit": (R or {}).get("unit", ""),
                                     "qty": None, "is_header": hdr(R)})
                    n = ins(left, i, newL)
                    changed = True
                    break
                if rn == rb:                    # 反之复制到右侧
                    newR = dict(L) if L else None
                    if newR is not None:
                        newR.update({"desc": desc(L), "unit": (L or {}).get("unit", ""),
                                     "qty": None, "is_header": hdr(L)})
                    n = ins(right, i, newR)
                    changed = True
                    break
        if not changed:
            break

    # ---- 4. StartHere 收尾：A 列 FillDown -> 按 A 升序排序 -> 重排序号 ----
    lastA = None
    for i in range(n):
        if left[i] is not None:
            if left[i].get("A") in (None, ""):
                left[i]["A"] = lastA
            else:
                lastA = left[i]["A"]
    order = sorted(range(n), key=lambda i: (left[i] is None or left[i].get("A") is None,
                                            left[i].get("A") if left[i] else 0, i))
    left = [left[i] if i < len(left) else None for i in order]
    right = [right[i] if i < len(right) else None for i in order]
    for i in range(n):
        if left[i] is not None:
            left[i]["A"] = i - 3 + 1

    out, extras = [], []
    for L, R in zip(left, right):
        if L is None and R is None:
            continue
        h = hdr(L) or hdr(R)
        a, d = L or R, R or L
        out.append({
            "seq": len(out) + 1,
            "a_desc": a["desc"], "a_unit": "" if h else a.get("unit", ""),
            "a_qty": None if h or L is None else L.get("qty"),
            "j_desc": d["desc"], "k_unit": "" if h else d.get("unit", ""),
            "l_qty": None if h or R is None else R.get("qty"),
            "is_header": h, "group": grp(L) or grp(R),
            "matched": None if h else (desc(L) == desc(R) and bool(desc(L))),
        })
        if R is not None and L is None and not h:
            extras.append(R)
    stats = _stats(asbuilt, design, out, extras)
    stats["row_base"] = "starthere"
    return out, extras, stats


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
        key = _macro_key(a)
        if a["is_header"]:
            out.append({
                "seq": len(out) + 1,
                "a_desc": a["desc"], "a_unit": "", "a_qty": None,
                "j_desc": a["desc"], "k_unit": "", "l_qty": None,
                "is_header": True, "group": a["group"], "matched": None,
            })
            continue
        cand = _find_cand(design, used, key, a["group"], mode,
                          desc=a["desc"])
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
            last_header[d["desc"]] = i
    for i, d in enumerate(design):
        key = _macro_key(d)
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
        cand = _find_cand(asbuilt, used, key, d["group"], mode)
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
    # 2026-11-26 修复：与 match_rows 的指纹回退保持一致，否则 merged 行序
    # 只按原描述对齐，实测「闸阀_Z47W-16C_DN300（单闸板）」与竣工图
    # 「闸阀_Z47W-16C_DN300」无法配对，同一物料被拆成增加/减少两行。
    # StartHere 规则：描述+单位原始拼接，不做别名/模糊归一
    return ("H:" if r["is_header"] else "D:") + _macro_key(r)


def _dedupe_headers(rows):
    """多页续表重复同一分组标题时，只保留最后一次出现。"""
    last = {}
    for i, r in enumerate(rows):
        if r["is_header"]:
            last[r["desc"]] = i
    return [r for i, r in enumerate(rows)
            if not r["is_header"] or last.get(r["desc"]) == i]


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
    p.add_argument("--row-base", default="starthere",
                   choices=("asbuilt", "design", "merged", "starthere"),
                   help="输出行序基准：starthere（默认，100%% 复刻模板宏 StartHere 的同分部移动对齐）/ "
                        "asbuilt（按竣工图表）/ design（按施工图表）/ merged（两表并集）")
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
               "merged": match_rows_merged, "starthere": match_rows_starthere}[args.row_base]
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
