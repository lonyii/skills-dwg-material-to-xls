# -*- coding: utf-8 -*-
"""
从 DWG/DXF 图纸中识别并重建燃气工程的「材料表」。

适用图纸特征（详见 references/layout.md）：
  * 表头一行含「序号 / 物 料 描 述 / 单位 / 数量」（还可有 标准号、备注 列）
  * 材料表由独立 TEXT/MTEXT 文字加表格线构成（非 AutoCAD TABLE 对象）
  * 一张 DWG 可含多个材料表：多页「工艺设备材料表」，或另有一张「竣工图」材料表
  * 表格可能自上而下排，也可能自下而上排；可能分上下两段续表共用一个表头

核心算法：
  1. 以表头文字的 x 坐标作为列锚点，相邻锚点中点作为列边界
  2. 以「序号」列中连续数字作为行锚点，相邻序号 y 坐标的中点作为行边界
     —— 与排版方向无关，且能自动把跨两行的物料描述合并成 "主描述\\n续行"
  3. 序号从 1 重新开始时自动切分为新段（续表 / 新页），按出现顺序拼接

实体来源（按优先级）：
  * Windows COM：浩辰CAD / GstarCAD / AutoCAD / 中望CAD，读 TEXT/MTEXT
  * 缓存 JSON：inspect --cache 保存一次，之后 --from-json 可离线反复解析
  * DXF：安装 ezdxf 后可直接传 .dxf（可选依赖）

CLI：
  python dwg_tables.py inspect <drawing> [--cache entities.json] [--from-json entities.json]
  python dwg_tables.py extract <drawing> [--stage auto|竣工图|施工图] [--tables 1,2]
                       [--cache entities.json] [--from-json entities.json] [--out rows.json]
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding=sys.stdout.encoding or "utf-8",
                               errors="replace")
    except Exception:                          # noqa: BLE001
        pass

HEADER_KEYS = ("序号", "物料描述", "名称", "规格", "单位", "数量",
               "标准号现行", "备注")
REQUIRED_HEADERS = ("序号", "物料描述", "单位", "数量")
# 2026 年新图框：表头为「序号 / 标准号或图号 / 名称 / 规格 / 单位 / 数量」，
# 另有一行「材料 / 质量(kg)」；物料描述拆成「名称 + 规格 + 材料」三列。
NEW_LAYOUT_HEADERS = ("序号", "名称", "单位", "数量")
# 参考图框「序号→单位」两个表头锚点的间距（图形单位）。
# 不同设计院/不同比例出图会整体缩放，所有固定偏移与容差都按此比例换算，
# 这样同一套算法可同时处理 1:1 图框与 1:2 缩放图框（实测 NC25D007 / ID24E107）。
REF_SEQ_UNIT = 10126.0
CAD_PROGIDS = (
    "GStarCAD.Application",
    "Gcad.Application",
    "AutoCAD.Application",
    "ZWCAD.Application",
)


# --------------------------------------------------------------------------- #
# 文本清洗
# --------------------------------------------------------------------------- #
def clean_text(raw: str) -> str:
    """清理 CAD 文字里的格式码与转义，保留物料描述本身。"""
    if raw is None:
        return ""
    s = str(raw)
    for a, b in (("%%D", "°"), ("%%d", "°"), ("%%C", "φ"), ("%%c", "φ"),
                 ("%%P", "±"), ("%%p", "±")):
        s = s.replace(a, b)
    s = re.sub(r"\\[Pp]", "\n", s)                 # MText 换行
    s = re.sub(r"\{\\[^;{}]*;", "", s)             # {\C2; ... 格式组起始
    s = re.sub(r"\\[A-Za-z][^;{}\\]*;", "", s)     # \fArial|b0|...; 等
    s = s.replace("{", "").replace("}", "")
    s = s.replace("\r", "").strip()
    # 全角空格 / 连续空格规整（不动物料描述内部的下划线与括号）
    return s


def norm_header(s: str) -> str:
    return re.sub(r"[\s\u3000（）()]", "", clean_text(s))


def norm_desc(s: str) -> str:
    """物料描述归一化，用于施工/竣工两侧匹配。"""
    s = clean_text(s)
    s = s.replace("（", "(").replace("）", ")").replace("＊", "*")
    s = s.replace("～", "~").replace("—", "-").replace("–", "-")
    s = re.sub(r"\s+", "", s)
    return s


def is_header_row(desc: str, unit: str, qty) -> bool:
    if not desc:
        return False
    if unit or qty is not None:
        return False
    return desc.endswith(("：", ":")) or ("材料" in desc or "部分" in desc) and len(desc) < 24


def parse_qty(s: str):
    s = clean_text(s).replace(",", "").strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


# --------------------------------------------------------------------------- #
# 实体读取
# --------------------------------------------------------------------------- #
def open_cad_app(retries=3):
    """
    打开一个 CAD COM 实例，返回 (app, owns)。
    批量读图时应只调用一次并复用 app：反复 DispatchEx/Quit 会让 DCOM
    偶发「无效的类字符串」，因此这里同时做多 ProgID 与多轮重试。
    """
    try:
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError("需要 pywin32：python -m pip install pywin32") from exc
    last = None
    for attempt in range(retries):
        for pid in CAD_PROGIDS:
            for factory in ("DispatchEx", "Dispatch"):
                try:
                    app = getattr(win32, factory)(pid)
                    return app, factory == "DispatchEx"
                except Exception as exc:       # noqa: BLE001
                    last = exc
        if attempt < retries - 1:
            time.sleep(1.5)
    raise RuntimeError("未找到可用的 CAD COM 组件（浩辰/AutoCAD/中望）。最后一次错误：%s" % last)


def close_cad_app(app, owns=True):
    """只关闭自己创建的实例，避免影响用户已打开的 CAD。"""
    if app is not None and owns:
        try:
            app.Quit()
        except Exception:                      # noqa: BLE001
            pass


def read_entities_com(path: str, app=None):
    """通过 CAD COM 读取所有 TEXT/MTEXT 的插入点与文字；app 可复用以减少启动次数。"""
    owns_app = False
    if app is None:
        app, owns_app = open_cad_app()

    try:
        try:
            app.Visible = False
        except Exception:                     # noqa: BLE001
            pass
        try:
            app.DisplayAlerts = False
        except Exception:                     # noqa: BLE001  浩辰不支持该属性
            pass
        doc = app.Documents.Open(os.path.abspath(path))
        ents = []
        seen = set()

        def walk(space, tag):
            try:
                count = space.Count
            except Exception:                 # noqa: BLE001
                return
            for i in range(count):
                try:
                    e = space.Item(i)
                    name = str(getattr(e, "EntityName", "") or getattr(e, "ObjectName", ""))
                    if "Text" not in name and "Attribute" not in name:
                        continue
                    txt = clean_text(getattr(e, "TextString", ""))
                    if not txt:
                        continue
                    pt = e.InsertionPoint
                    x, y = float(pt[0]), float(pt[1])
                    key = (round(x, 2), round(y, 2), txt)
                    if key in seen:
                        continue
                    seen.add(key)
                    ents.append({"t": txt, "x": x, "y": y,
                                 "kind": "MTEXT" if "MText" in name else "TEXT",
                                 "space": tag, "layer": str(getattr(e, "Layer", ""))})
                except Exception:             # noqa: BLE001  跳过坏实体
                    continue

        walk(doc.ModelSpace, "MODEL")
        walk(doc.PaperSpace, "PAPER")
        # 部分 CAD 的 PaperSpace 只对应当前布局，补遍历其他布局
        try:
            for lay in doc.Layouts:
                if str(getattr(lay, "Name", "")).lower() == "model":
                    continue
                walk(lay.Block, "LAYOUT")
        except Exception:                     # noqa: BLE001
            pass
        try:
            doc.Close(False)
        except Exception:                     # noqa: BLE001
            pass
        return ents
    finally:
        close_cad_app(app, owns_app)


def read_entities_dxf(path: str):
    """可选后端：用 ezdxf 读取 DXF（pip install ezdxf）。"""
    try:
        import ezdxf
    except ImportError as exc:
        raise RuntimeError("读取 DXF 需要：python -m pip install ezdxf") from exc
    doc = ezdxf.readfile(path)
    ents = []
    for layout in doc.layouts:
        for e in layout:
            if e.dxftype() not in ("TEXT", "MTEXT", "ATTRIB"):
                continue
            txt = clean_text(e.text if e.dxftype() == "MTEXT" else e.dxf.text)
            if not txt:
                continue
            p = e.dxf.insert
            ents.append({"t": txt, "x": float(p.x), "y": float(p.y),
                         "kind": e.dxftype(), "space": layout.name,
                         "layer": str(e.dxf.layer)})
    return ents


def load_entities(drawing: str, from_json: str | None = None,
                  cache: str | None = None, app=None):
    if from_json:
        data = json.load(open(from_json, encoding="utf-8"))
        return data["entities"], data
    ext = os.path.splitext(drawing)[1].lower()
    ents = read_entities_dxf(drawing) if ext == ".dxf" else read_entities_com(drawing, app=app)
    meta = {"drawing": os.path.abspath(drawing), "version": 1, "entities": ents}
    if cache:
        with open(cache, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False)
    return ents, meta


# --------------------------------------------------------------------------- #
# 表格识别与重建
# --------------------------------------------------------------------------- #
def _table_scale(a):
    """表头「序号→单位」间距相对参考图框的比例，用于换算所有绝对容差。"""
    w = a["单位"] - a["序号"]
    if w <= 0:
        return 1.0
    return max(0.2, min(4.0, w / REF_SEQ_UNIT))


def detect_tables(ents):
    """返回按 (y, x) 排序的材料表定义列表，并自带比例与阶段标注。"""
    headers = []
    for c in ents:
        if norm_header(c["t"]) != "序号":
            continue
        # 同一水平线可并排放多张表：每个表头列取距离本「序号」最近的那个，
        # 避免把隔壁表的列锚点错配过来。
        near = [e for e in ents
                if abs(e["y"] - c["y"]) <= 320 and -1000 <= e["x"] - c["x"] <= 22000]
        anchors = {}
        for key in HEADER_KEYS:
            matches = [e for e in near if norm_header(e["t"]) == key]
            if matches:
                anchors[key] = min(matches, key=lambda e: abs(e["x"] - c["x"]))["x"]
        old_ok = all(k in anchors for k in REQUIRED_HEADERS)
        new_ok = all(k in anchors for k in NEW_LAYOUT_HEADERS)
        if old_ok or new_ok:
            headers.append({"y": c["y"], "anchors": anchors,
                            "layout": "old" if old_ok else "name-spec"})
    # 同一图纸中的多页：先按表头 y 降序（上下并排时上页在前），
    # 同一水平线再按 x 升序（左右并排时左页在前）。
    headers.sort(key=lambda h: (-round(h["y"] / 1000), h["anchors"]["序号"]))

    # 去重（同一表头可能被多个布局重复遍历到）
    uniq = []
    for h in headers:
        lim = 5000 * _table_scale(h["anchors"])
        if not any(abs(h["y"] - u["y"]) < 200 and
                   abs(h["anchors"]["序号"] - u["anchors"]["序号"]) < lim for u in uniq):
            uniq.append(h)

    tables = []
    for i, h in enumerate(uniq):
        a = h["anchors"]
        seq_x = a["序号"]
        sc = _table_scale(a)
        next_seq = [u["anchors"]["序号"] for u in uniq
                    if u["anchors"]["序号"] > seq_x + 5000 * sc
                    and abs(u["y"] - h["y"]) < 8000 * sc]
        right_by_col = a.get("备注", a.get("标准号现行",
                                           a["数量"] + 2000 * sc)) + 3000 * sc
        x_right = min([right_by_col] + ([min(next_seq) - 1000 * sc] if next_seq else []))
        x_left = seq_x - 1500 * sc
        table = {
            "id": i + 1,
            "header_y": h["y"],
            "anchors": a,
            "scale": sc,
            "layout": h.get("layout", "old"),
            "x_left": x_left,
            "x_right": x_right,
            "bounds": _column_bounds(a, x_left, x_right, sc, h.get("layout", "old")),
            "stage": detect_stage(ents, h, x_left, x_right, sc),
            "title": detect_title(ents, h, x_left, x_right, sc),
        }
        tables.append(table)

    # 同一图框左右并排的多页表，常只有一页带图签阶段文字：
    # 让 y 接近、x 相邻的「未标注」表继承最近已标注表，避免漏掉整页材料。
    for t in tables:
        if t["stage"] != "未标注":
            continue
        w = t["anchors"]["单位"] - t["anchors"]["序号"]
        cands = [u for u in tables
                 if u["stage"] != "未标注" and u["id"] != t["id"]
                 and abs(u["header_y"] - t["header_y"]) < 2.5 * w
                 and 0 < abs(u["anchors"]["序号"] - t["anchors"]["序号"]) < 4 * w]
        if cands:
            best = min(cands, key=lambda u: abs(u["header_y"] - t["header_y"])
                       + 0.1 * abs(u["anchors"]["序号"] - t["anchors"]["序号"]))
            t["stage"] = best["stage"]
            t["stage_inherited"] = True
    return tables


def _column_bounds(a, x_left, x_right, scale=1.0, layout="old"):
    """
    列边界。

    注意：CAD 表头文字常居中排版（如 " 物   料   描   述 " 带大量前导空格），
    其插入点 x 明显右于同列数据文字，不能直接用表头锚点的中点切列。
    这里以稳定表头锚点加固定偏移定界；所有偏移按 scale（本图框相对参考
    图框的缩放比例）换算，兼容不同比例出图。

    layout="name-spec" 对应 2026 年新图框：物料描述拆成「名称 / 规格 /
    材料」三列，表头还多一列「标准号或图号」，列序与旧图框完全不同。
    """
    seq, unit, qty = a["序号"], a["单位"], a["数量"]
    if layout == "name-spec":
        std = a.get("标准号现行", seq + 1000 * scale)
        name = a.get("名称", std + 2500 * scale)
        spec = a.get("规格", name + 2200 * scale)
        mat = a.get("材料", spec + 2500 * scale)
        note = a.get("备注", qty + 2500 * scale)
        b = {
            "序号": (x_left, std - 400 * scale),
            "标准号现行": (std - 400 * scale, name - 1200 * scale),
            "名称": (name - 1200 * scale, spec - 1200 * scale),
            "规格": (spec - 1200 * scale, mat - 1200 * scale),
            "材料": (mat - 1200 * scale, unit - 400 * scale),
            "单位": (unit - 400 * scale, qty - 300 * scale),
            "数量": (qty - 300 * scale, note - 1200 * scale),
            "备注": (note - 1200 * scale, x_right),
        }
    else:
        std = a.get("标准号现行", qty + 2500 * scale)
        note = a.get("备注", std + 2800 * scale)
        b = {
            "序号": (x_left, seq + 600 * scale),
            "物料描述": (seq + 600 * scale, unit - 1500 * scale),
            "单位": (unit - 1500 * scale, qty - 600 * scale),
            "数量": (qty - 600 * scale, qty + 700 * scale),
            "标准号现行": (qty + 700 * scale, note - 1500 * scale),
            "备注": (note - 1500 * scale, x_right),
        }
    keep = set(REQUIRED_HEADERS) | {"名称", "规格", "材料"}
    return {k: v for k, v in b.items() if k in a or k in keep}


def detect_stage(ents, header, x_left, x_right, scale=1.0):
    """从图签里判断本表属于「施工图」还是「竣工图」。"""
    hits = []
    for e in ents:
        if not (x_left <= e["x"] <= x_right):
            continue
        if abs(e["y"] - header["y"]) > 12000 * scale:
            continue
        t = clean_text(e["t"]).strip()
        if t in ("竣工图", "施工图"):
            hits.append((abs(e["y"] - header["y"]), t))
    if not hits:
        return "未标注"
    hits.sort()
    return hits[0][1]


def detect_title(ents, header, x_left, x_right, scale=1.0):
    for e in ents:
        if x_left <= e["x"] <= x_right and abs(e["y"] - header["y"]) < 8000 * scale:
            t = clean_text(e["t"])
            if "材料表" in t:
                return t[:40]
    return ""


def _seq_entities(ents, table):
    lo, hi = table["bounds"]["序号"]
    sc = table.get("scale", 1.0)
    out = []
    for e in ents:
        if not (lo <= e["x"] < hi):
            continue
        if abs(e["y"] - table["header_y"]) > 42000 * sc:
            continue
        t = clean_text(e["t"]).strip()
        if re.fullmatch(r"\d{1,3}", t):
            out.append({"y": e["y"], "seq": int(t)})
    # 同一 y 去重
    out.sort(key=lambda s: s["y"])
    dedup = []
    for s in out:
        if dedup and abs(s["y"] - dedup[-1]["y"]) < 60 * sc:
            continue
        dedup.append(s)
    return dedup


def _split_segments(seqs):
    """把序号序列按排版方向与「序号从 1 重启」切成段。"""
    if not seqs:
        return []
    asc = [s["seq"] for s in seqs]
    # 判断 y 升序时序号是递增还是递减
    diffs = [asc[i + 1] - asc[i] for i in range(min(6, len(asc) - 1))]
    increasing = sum(d > 0 for d in diffs) >= sum(d < 0 for d in diffs)
    ordered = seqs if increasing else list(reversed(seqs))
    segments, cur = [], []
    for s in ordered:
        if cur and s["seq"] == 1 and cur[-1]["seq"] > 1:
            segments.append(cur)
            cur = []
        cur.append(s)
    if cur:
        segments.append(cur)
    # 只保留从 1 开始、且大体连续的段，过滤图签里的零散数字
    return [seg for seg in segments
            if seg and seg[0]["seq"] == 1 and len(seg) >= 2]


SIGN_WORDS = (
    "编 制", "编制", "校 核", "校核", "审 核", "审核", "审 定", "日 期",
    "版 次", "第    页", "共    页", "设 计", "制 图", "图 号", "建设单位",
    "项 目 名 称", "专业负责人", "项目负责人", "设计阶段", "比    例",
    "专    业", "设 计 阶 段", "工 程 名 称",
)


def _is_sign_text(s):
    return any(w in s for w in SIGN_WORDS)


def _valid_unit(s):
    return (0 < len(s) <= 3 and " " not in s and "\u3000" not in s
            and not re.search(r"\d", s) and not _is_sign_text(s))


def parse_table(ents, table):
    """
    先把文字按 y 聚成「视觉行」，再把每个视觉行按 x 归入各列；
    以带序号的视觉行为记录锚点，把相邻的、无序号的文字行归入最近的锚点。
    这样既能合并拆成两行的物料描述，又不会把相邻行或图签文字串进来。
    """
    segs = _split_segments(_seq_entities(ents, table))
    if not segs:
        return []
    bounds = table["bounds"]
    sc = table.get("scale", 1.0)
    all_y = [s["y"] for seg in segs for s in seg]
    y_lo, y_hi = min(all_y) - 900 * sc, max(all_y) + 900 * sc

    items = []
    for e in ents:
        if not (table["x_left"] <= e["x"] <= table["x_right"]):
            continue
        if not (y_lo <= e["y"] <= y_hi):
            continue
        t = clean_text(e["t"]).strip()
        if t:
            items.append((e["y"], e["x"], t))
    items.sort(key=lambda a: (-a[0], a[1]))

    lines = []
    for it in items:
        if lines and abs(lines[-1]["y"] - it[0]) <= 300 * sc:
            lines[-1]["items"].append(it)
        else:
            lines.append({"y": it[0], "items": [it]})

    for ln in lines:
        ln["is_header_line"] = any(norm_header(t) in HEADER_KEYS for _, _, t in ln["items"])
        cols = {k: [] for k in bounds}
        for _, x, t in ln["items"]:
            for key, (blo, bhi) in bounds.items():
                if blo <= x < bhi:
                    cols[key].append(t)
                    break
        seqs = [t for t in cols.get("序号", []) if re.fullmatch(r"\d{1,3}", t)]
        ln["seq"] = int(seqs[0]) if seqs and not ln["is_header_line"] else None
        if ln["is_header_line"]:
            ln["descs"] = []
        elif table.get("layout") == "name-spec":
            # 新图框：成品表只取「名称 + 规格」两列，物料材质列（如 Q235B、
            # KTH350-10）不写入物料描述；名称列内部的下划线/空格原样保留。
            ln["descs"] = [t for key in ("名称", "规格")
                           for t in cols.get(key, [])
                           if len(t) >= 2 and not re.fullmatch(r"[\d.\s]+", t)
                           and not _is_sign_text(t)]
        else:
            ln["descs"] = [t for t in cols.get("物料描述", [])
                           if len(t) >= 2 and not re.fullmatch(r"[\d.\s]+", t)
                           and not _is_sign_text(t)]
        units = [] if ln["is_header_line"] else [t for t in cols.get("单位", []) if _valid_unit(t)]
        ln["unit"] = units[0] if units else ""
        qtys = [] if ln["is_header_line"] else [
            q for q in (parse_qty(t) for t in cols.get("数量", [])) if q is not None]
        ln["qty"] = qtys[0] if qtys else None

    # 上下两段续表共用表头时，段区间用相邻段中点切开，避免同一视觉行被解析两次
    seg_ranges = []
    for k, seg in enumerate(segs):
        ys = [s["y"] for s in seg]
        lo, hi = min(ys) - 700 * sc, max(ys) + 700 * sc
        if k > 0:
            lo = max(lo, (max(s["y"] for s in segs[k - 1]) + min(ys)) / 2)
        if k + 1 < len(segs):
            hi = min(hi, (max(ys) + min(s["y"] for s in segs[k + 1])) / 2)
        seg_ranges.append((lo, hi))

    rows = []
    for seg, (lo, hi) in zip(segs, seg_ranges):
        seg_lines = [ln for ln in lines
                     if lo <= ln["y"] <= hi and not ln["is_header_line"]]
        increasing_y = seg[0]["y"] < seg[-1]["y"]
        seg_lines.sort(key=lambda l: l["y"] if increasing_y else -l["y"])
        anchors = [i for i, ln in enumerate(seg_lines) if ln["seq"] is not None]
        if not anchors:
            continue
        owner = {}
        for i, ln in enumerate(seg_lines):
            if ln["seq"] is not None:
                continue
            if not (ln["descs"] or ln["unit"] or ln["qty"] is not None):
                continue
            owner[i] = min(anchors, key=lambda j: abs(seg_lines[j]["y"] - ln["y"]))
        for idx in anchors:
            anchor = seg_lines[idx]
            extra = [seg_lines[i] for i, j in owner.items() if j == idx]
            group = [anchor] + extra
            group.sort(key=lambda l: -l["y"])       # 描述按视觉从上到下拼接
            if table.get("layout") == "name-spec":
                # 成品表格式：名称列原样（含内部空格/下划线）+ 单个空格 + 规格列原样，
                # 不拼接材料列，也不把内部空格换成下划线。
                # 例：热镀锌钢管_燃气专用 DN50（D60.3×3.8）
                desc = " ".join(d.strip() for ln in group
                                for d in ln["descs"]).strip()
            else:
                desc = "\n".join(d for ln in group for d in ln["descs"]).strip()
            if not desc:
                continue
            unit = anchor["unit"] or next((ln["unit"] for ln in extra if ln["unit"]), "")
            qty = anchor["qty"]
            if qty is None:
                qty = next((ln["qty"] for ln in extra if ln["qty"] is not None), None)
            rows.append({
                "seq": anchor["seq"],
                "desc": desc,
                "unit": unit,
                "qty": qty,
                "is_header": is_header_row(desc, unit, qty),
                "table": table["id"],
                "stage": table["stage"],
            })
    return rows


def assign_groups(rows):
    """给每行标注所属分组（跨表/跨页继承最近的分组标题）。"""
    group = ""
    for r in rows:
        if r["is_header"]:
            group = r["desc"]
            r["group"] = group
        else:
            r["group"] = group
    return rows


def extract_rows(drawing, from_json=None, cache=None, stage="auto", tables=None):
    ents, meta = load_entities(drawing, from_json=from_json, cache=cache)
    detected = detect_tables(ents)
    chosen = [t for t in detected if t["id"] in tables] if tables else \
             [t for t in detected if stage == "auto" or t["stage"] == stage]
    if not chosen and tables is None and stage != "auto":
        # 图签缺少阶段标注时退化为全部表，避免静默产出空表
        chosen = detected
    # 手动指定表编号时，按用户给出的顺序拼接（可调整左右页先后）；
    # 自动选表时按检测顺序（先 y 后 x）。
    if tables:
        by_id = {t["id"]: t for t in chosen}
        chosen = [by_id[i] for i in tables if i in by_id]
    else:
        chosen = sorted(chosen, key=lambda x: x["id"])
    rows = []
    for t in chosen:
        rows.extend(parse_table(ents, t))
    assign_groups(rows)
    return {"meta": {k: v for k, v in meta.items() if k != "entities"},
            "tables": [{k: v for k, v in t.items() if k != "bounds"} for t in detected],
            "selected": [t["id"] for t in chosen],
            "rows": rows}


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def _ids(s):
    return [int(x) for x in s.split(",") if x.strip()] if s else None


def main(argv=None):
    p = argparse.ArgumentParser(description="识别并提取 DWG 图纸中的材料表")
    p.add_argument("command", choices=("inspect", "extract"))
    p.add_argument("drawing", help="DWG/DXF 路径；配 --from-json 时仅为标识，可任意填")
    p.add_argument("--from-json", help="使用 inspect --cache 保存的实体 JSON，离线解析")
    p.add_argument("--cache", help="把实体缓存写入该 JSON 文件")
    p.add_argument("--stage", default="auto", choices=("auto", "竣工图", "施工图"))
    p.add_argument("--tables", help="手动指定表编号，逗号分隔，如 1,2（配合 inspect 使用）")
    p.add_argument("--out", help="extract 结果 JSON 输出路径（不指定则打印到 stdout）")
    args = p.parse_args(argv)

    ids = _ids(args.tables)
    if args.command == "inspect":
        ents, meta = load_entities(args.drawing, from_json=args.from_json, cache=args.cache)
        tables = detect_tables(ents)
        print("图纸：%s" % meta.get("drawing", args.drawing))
        print("文字实体：%d 个；检测到材料表：%d 张" % (len(ents), len(tables)))
        for t in tables:
            rows = parse_table(ents, t)
            filled = [r for r in rows if r["desc"]]
            print("  #%d  [%s%s]  %s  表头y=%.0f  序号列x=%.0f  比例=%.2f  有效行=%d  首行=%s" % (
                t["id"], t["stage"], "(继承)" if t.get("stage_inherited") else "",
                t["title"] or "(未取到表名)", t["header_y"],
                t["anchors"]["序号"], t.get("scale", 1.0), len(filled),
                repr(filled[0]["desc"])[:46] if filled else "-"))
        print("\n提示：用 --tables 指定要提取的表；extract 输出 JSON 供导入脚本使用。")
        return 0

    result = extract_rows(args.drawing, from_json=args.from_json, cache=args.cache,
                          stage=args.stage, tables=ids)
    payload = json.dumps(result, ensure_ascii=False, indent=1)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(payload)
        print("已输出 %d 行 -> %s" % (len(result["rows"]), args.out))
        print("选用表：%s" % result["selected"])
    else:
        print(payload)
    return 0


if __name__ == "__main__":
    sys.exit(main())
