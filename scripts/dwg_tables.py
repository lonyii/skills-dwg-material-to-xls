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

__author__ = "lonyii"
__email__ = "lonyii@qq.com"

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

HEADER_KEYS = ("序号", "物料描述", "名称及规格", "名称", "规格", "单位", "数量",
               "标准号现行", "备注")
REQUIRED_HEADERS = ("序号", "物料描述", "单位", "数量")
# 2026 年新图框：表头为「序号 / 标准号或图号 / 名称 / 规格 / 单位 / 数量」，
# 另有一行「材料 / 质量(kg)」；物料描述拆成「名称 + 规格 + 材料」三列。
NEW_LAYOUT_HEADERS = ("序号", "名称", "单位", "数量")
# 2025 年起另一版图框：表头为「序号 / 标准号或图号 / 名称及规格 /
# （或性能等级）/ 单位 / 数量 / 单总 / 备注」，名称与规格合并成一列，
# 另有一行「材料 / 质量(kg)」。物料描述只取「名称及规格」这一列。
MERGED_LAYOUT_HEADERS = ("序号", "名称及规格", "单位", "数量")
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
    if desc.endswith(("：", ":")):
        return True
    if ("材料" in desc or "部分" in desc) and len(desc) < 24:
        return True
    # 分组标题常见写法：「配合工程（长度不超过10m无外管编号）」「配套工程部分」等，
    # 通常以「工程」结尾或含「工程（」「工程(」，且整行只有标题文字。
    if len(desc) < 40 and ("工程" in desc and
                           (desc.endswith("工程") or "工程（" in desc or "工程(" in desc
                            or desc.endswith("工程部分"))):
        return True
    return False


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
    """通过 CAD COM 读取所有 TEXT/MTEXT 的插入点与文字；app 可复用以减少启动次数。

    浩辰 COM 遍历偶发漏实体（实测同一张图读到 461 / 1019 两种结果，少的那份缺了
    竣工表区域）。读取后做完整性自检：若存在材料表表头词、却识别不到任何表，
    关闭文档重开再读，最多 3 次。大图用文件大小辅助判断是否跳过模型空间，
    避免 COM Count 误报导致去遍历十几万实体而漏读。
    """
    owns_app = False
    if app is None:
        app, owns_app = open_cad_app()

    path = os.path.abspath(path)
    try:
        app.Visible = False
    except Exception:                         # noqa: BLE001
        pass
    try:
        app.DisplayAlerts = False
    except Exception:                         # noqa: BLE001  浩辰不支持该属性
        pass

    big = os.path.isfile(path) and os.path.getsize(path) > 1_500_000
    heads = {"序号", "单位", "数量", "物料描述", "名称"}
    best = []
    try:
        for attempt in range(3):
            doc = app.Documents.Open(path)
            ents, seen = [], set()

            def walk(space, tag):
                try:
                    count = space.Count
                except Exception:             # noqa: BLE001
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
                    except Exception:         # noqa: BLE001  跳过坏实体
                        continue

            # 大图优化：建筑底图在 ModelSpace 有十几万实体，逐个 COM 访问会超时/漏读。
            # 文件大小 >1.5MB 直接跳过；否则再按 COM Count 判断（>50000 跳过）。
            try:
                model_count = doc.ModelSpace.Count
            except Exception:                 # noqa: BLE001
                model_count = 0
            if big or model_count > 50000:
                print("  [大图] %s %d 实体 / %.1fMB，跳过模型空间只读布局空间" %
                      ("ModelSpace" if model_count > 50000 else "文件",
                       model_count, os.path.getsize(path) / 1e6), flush=True)
            else:
                walk(doc.ModelSpace, "MODEL")
            walk(doc.PaperSpace, "PAPER")
            try:
                for lay in doc.Layouts:
                    if str(getattr(lay, "Name", "")).lower() == "model":
                        continue
                    walk(lay.Block, "LAYOUT")
            except Exception:                 # noqa: BLE001
                pass
            try:
                doc.Close(False)
            except Exception:                 # noqa: BLE001
                pass

            if len(ents) > len(best):
                best = ents
            n_head = sum(1 for e in ents
                         if clean_text(e.get("t", "")).strip() in heads)
            # 自检：有材料表表头词却识别不到任何表 → 可能是漏读，重试。
            incomplete = n_head >= 1 and not detect_tables(ents)
            if not incomplete:
                break
            print("  [读取自检] 第%d次读到 %d 实体、%d 表头词但 0 张表，"
                  "重开文档再读…" % (attempt + 1, len(ents), n_head), flush=True)
        return best
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
    # 下限放到 0.01：实测 IE26I004 在模型空间画了极扁的材料表（序号→单位仅 128，
    # 正常约 10126），比例约 0.013；旧下限 0.2 会把列偏移放大 16 倍导致整表串列。
    return max(0.01, min(4.0, w / REF_SEQ_UNIT))


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
        # 扁表格（实测 IE26I004，列间距仅约 128）里，并排表的锚点会互相串：
        # 右栏表头会把左栏的单位/数量当成自己的。用第一步估出的 scale 收紧 x 下界，
        # 重算一次：列锚点应在本表序号右侧，不允许越过左边界去取隔壁表的列。
        if "单位" in anchors and "序号" in anchors:
            w0 = anchors["单位"] - anchors["序号"]
            sc0 = max(0.01, min(4.0, w0 / REF_SEQ_UNIT))
            # 仅扁表格（比例 < 0.1）需要收紧 x 下界防止并排表串列；
            # 正常图保持 -1000，避免切掉正常居中排版的列锚点。
            if sc0 < 0.1:
                left_bound = -60 * sc0
                near2 = [e for e in near if e["x"] - c["x"] >= left_bound]
                anchors = {}
                for key in HEADER_KEYS:
                    matches = [e for e in near2 if norm_header(e["t"]) == key]
                    if matches:
                        anchors[key] = min(matches,
                                           key=lambda e: abs(e["x"] - c["x"]))["x"]
        old_ok = all(k in anchors for k in REQUIRED_HEADERS)
        new_ok = all(k in anchors for k in NEW_LAYOUT_HEADERS)
        merged_ok = all(k in anchors for k in MERGED_LAYOUT_HEADERS)
        if old_ok or new_ok or merged_ok:
            layout = "old" if old_ok else ("name-spec" if new_ok else "merged-spec")
            headers.append({"y": c["y"], "anchors": anchors, "layout": layout})
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
            "project": detect_table_project(ents, h, x_left, x_right, sc),
        }
        table["row_pitch"] = _estimate_row_pitch(ents, table)
        tables.append(table)

    # 记录每张表的序号走向与范围，供多表拼接排序使用：
    # top-down = 序号1在最上（沿 y 递减方向递增）；bottom-up = 序号1在最下。
    for t in tables:
        seqs = _seq_entities(ents, t)
        nums = [s["seq"] for s in seqs if s["seq"] <= 200]
        if len(seqs) >= 2:
            diffs = [seqs[i + 1]["seq"] - seqs[i]["seq"]
                     for i in range(min(8, len(seqs) - 1))]
            inc = sum(d > 0 for d in diffs) >= sum(d < 0 for d in diffs)
            t["seq_dir"] = "bottom-up" if inc else "top-down"
        else:
            t["seq_dir"] = "unknown"
        t["seq_min"] = min(nums) if nums else 9999
        t["seq_max"] = max(nums) if nums else 0

    # 同一图框并排的多栏表，工程名常只写在其中一栏的表头附近：
    # 让 y 接近、x 相邻的「未标注工程」表继承相邻表的工程名。
    for t in tables:
        if t.get("project"):
            continue
        near = [u for u in tables if u.get("project") and u["id"] != t["id"]
                and abs(u["header_y"] - t["header_y"]) <= 8000 * t["scale"]
                and abs(u["anchors"]["序号"] - t["anchors"]["序号"]) <= 20000 * t["scale"]]
        if near:
            nearest = min(near, key=lambda u: abs(
                u["anchors"]["序号"] - t["anchors"]["序号"]))
            t["project"] = nearest["project"]

    # 多工程图：给每张表打标，parse_table 据此只保留「首行序号紧贴本表表头」的段，
    # 避免序号跨越到相邻工程的表格。单工程图不打标，保持原有跨页长表行为。
    multi_project = len({t["project"] for t in tables if t.get("project")}) > 1
    if multi_project:
        for t in tables:
            t["multi_project"] = True

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


def order_tables(tables):
    """
    决定同一张图里多张材料表的拼接顺序（图纸无「第X页」标注时的习惯规则）：

      1. 序号连贯（各表主序号段首尾相接）→ 按序号从小到大；
      2. 序号从上到下（表头在上、序号1在最上）→ 整表从左到右（x 升序）；
      3. 序号从下到上（序号1在最下）→ 整表从右到左（x 降序）。

    只影响「多张表」的先后；单表原样返回。调用方应先把不同工程/不同阶段的表分开，
    本函数只排已经属于同一成品的那些表。
    """
    tables = list(tables)
    if len(tables) <= 1:
        return tables
    mins = [t.get("seq_min", 9999) for t in tables]
    # 序号连贯：各表起始序号不同，且按起始排序后前一表的最大序号紧接后一表的最小序号
    # （区间首尾相接、不重叠），如 1-26 / 27-52。
    by_start = sorted(tables, key=lambda t: t.get("seq_min", 9999))
    contiguous = (len(set(mins)) == len(mins) and all(
        by_start[i].get("seq_max", 0) + 1 == by_start[i + 1].get("seq_min", 9999)
        for i in range(len(by_start) - 1)))
    if contiguous:
        return by_start
    # 按多数表的序号走向决定左右顺序
    ups = sum(1 for t in tables if t.get("seq_dir") == "bottom-up")
    downs = sum(1 for t in tables if t.get("seq_dir") == "top-down")
    if downs > ups:                       # 序号从上到下 → 从左到右
        return sorted(tables, key=lambda t: t["anchors"]["序号"])
    return sorted(tables, key=lambda t: -t["anchors"]["序号"])   # 从下到上 → 从右到左


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
    elif layout == "merged-spec":
        # 「名称及规格」合并列：序号/标准号/名称及规格/(性能等级)/单位/数量/单总/备注。
        # 单位与数量锚点很近（本图框约 1000），用较小偏移避免吞并相邻列。
        std = a.get("标准号现行", seq + 1000 * scale)
        mname = a["名称及规格"]
        note = a.get("备注", qty + 3000 * scale)
        # 材料/性能等级列（如 Q235B、KTH350-10）位于名称列与单位列之间，
        # 成品不写入它，故名称列右界取在材料列左侧：单位锚点 - 4000*scale。
        name_right = unit - 4000 * scale
        b = {
            "序号": (x_left, std - 300 * scale),
            "标准号现行": (std - 300 * scale, mname - 1500 * scale),
            "名称及规格": (mname - 1500 * scale, name_right),
            "单位": (name_right, qty - 300 * scale),
            "数量": (qty - 300 * scale, note - 1500 * scale),
            "备注": (note - 1500 * scale, x_right),
        }
    else:
        std = a.get("标准号现行", qty + 2500 * scale)
        note = a.get("备注", std + 2800 * scale)
        # 扁表格（比例<0.1）里表头文字居中严重偏右，而数据文字紧跟序号列：
        # 如师大南路老图「物料描述」表头 x=786、实际物料 x=765、序号 x=761。
        # 用 600*scale 作序号列右界会把物料误判为序号，只有图签人名落入描述列。
        # 扁图改用 seq+2 作分界；正常图保持 600*scale 不变。
        seq_right = seq + (2.0 if scale < 0.1 else 600 * scale)
        b = {
            "序号": (x_left, seq_right),
            "物料描述": (seq_right, unit - 1500 * scale),
            "单位": (unit - 1500 * scale, qty - 600 * scale),
            "数量": (qty - 600 * scale, qty + 700 * scale),
            "标准号现行": (qty + 700 * scale, note - 1500 * scale),
            "备注": (note - 1500 * scale, x_right),
        }
    keep = set(REQUIRED_HEADERS) | {"名称", "规格", "名称及规格", "材料"}
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


# 一张施工图 DWG 内可能并排放多个工程的材料表。每张表的表头上方/附近会写
# 所属工程名（如「三米粥铺师大四食堂店厨房燃气工程」）。下面常量用于识别它。
_PROJECT_SUFFIXES = ("燃气工程", "天然气工程", "煤气工程", "工程")


def detect_table_project(ents, header, x_left, x_right, scale=1.0):
    """
    返回本材料表所属的工程名称（用于一张施工图含多个工程时按竣工图筛选表）。

    工程名通常写在表头 y 上方 1500~4500、x 在表序号列右侧 0~16000 的图签区；
    排除设计院名、通用图集名与「设计阶段/版次」等图签标签。找不到返回 ""。
    """
    seq_x = header["anchors"]["序号"]
    cands = []
    for e in ents:
        dy = e["y"] - header["y"]
        # 表头上方为主（dy>0），也容忍表头正下方一点（个别图纸）
        if not (1000 * scale <= dy <= 5000 * scale):
            continue
        if not (seq_x - 2000 * scale <= e["x"] <= x_right + 16000 * scale):
            continue
        t = clean_text(e["t"]).strip()
        if not (8 <= len(t) <= 48):
            continue
        if not any(t.endswith(sx) for sx in _PROJECT_SUFFIXES):
            continue
        if any(bad in t for bad in ("设计有限公司", "设计院", "通用图集", "技术标准",
                                     "燃气集团", "建设单位", "项目名称", "项 目 名 称")):
            continue
        # 距离表头越近越可能是本表的工程名；同高时 x 越靠近表头列越优先
        cands.append((dy, abs(e["x"] - (seq_x + 5000 * scale)), t))
    if not cands:
        return ""
    cands.sort(key=lambda x: (x[0], x[1]))
    return cands[0][2]


def _estimate_row_pitch(ents, table):
    """
    从序号列的相邻 y 间距估计真实行高（图形单位）。

    正常图框行距数百、表头「序号→单位」间距约 10126；但某些图（实测 IE26I004）
    把材料表压得很扁：列间距仅 128、行距仅 8，此时按表头间距推算的 scale=0.2
    会把 8 的行距当成噪声合并，导致几十行并成 1 行。这里直接用实测行距，
    让后续去重/聚类容差自适应，而不是依赖表头间距比例。
    """
    lo, hi = table["bounds"]["序号"]
    hy = table["header_y"]
    ys = sorted({round(e["y"], 2) for e in ents
                 if lo <= e["x"] < hi and abs(e["y"] - hy) <= 42000
                 and re.fullmatch(r"\d{1,3}", clean_text(e["t"]).strip())})
    gaps = [ys[i + 1] - ys[i] for i in range(len(ys) - 1)
            if 0 < ys[i + 1] - ys[i] < 5000]
    if not gaps:
        return 0.0
    gaps.sort()
    return float(gaps[len(gaps) // 2])      # 中位数，抗零散图签数字干扰


def _seq_entities(ents, table):
    lo, hi = table["bounds"]["序号"]
    sc = table.get("scale", 1.0)
    hy = table["header_y"]
    out = []
    for e in ents:
        if not (lo <= e["x"] < hi):
            continue
        if abs(e["y"] - hy) > 42000 * sc:
            continue
        t = clean_text(e["t"]).strip()
        if re.fullmatch(r"\d{1,3}", t):
            out.append({"y": e["y"], "seq": int(t)})
    # 同一 y 去重：容差取动态行距的一部分，避免扁表格被合并
    pitch = table.get("row_pitch", 0.0)
    dedup_tol = min(60 * sc, pitch * 0.5) if pitch else 60 * sc
    out.sort(key=lambda s: s["y"])
    dedup = []
    for s in out:
        if dedup and abs(s["y"] - dedup[-1]["y"]) < dedup_tol:
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
    # 多工程图：序号搜索会跨越相邻工程的表头混入其它表的数据段，
    # 只保留「首行序号紧贴本表表头」的段（6000*sc 内）。
    # 单工程图/跨页长表/1:2 缩放图框不打这个标，行为完全不变。
    if table.get("multi_project"):
        hy = table["header_y"]
        def _top(seg):
            return max(s["y"] for s in seg)
        near = [seg for seg in segs if 0 <= (hy - _top(seg)) <= 6000 * sc]
        if near:
            segs = near
    all_y = [s["y"] for seg in segs for s in seg]
    y_lo, y_hi = min(all_y) - 900 * sc, max(all_y) + 900 * sc

    # 图签行整行排除：扁表格里图签与最后一条材料 y 只差 1~2，
    # 若不过滤，物料描述末尾会混入设计师姓名（实测师大南路老图「毕倩如」）。
    # 先找签名标签（编制/校核/审核/日期…）所在的视觉 y，再把同 y 的人名一并排除。
    sign_ys = [e["y"] for e in ents
               if y_lo <= e["y"] <= y_hi and _is_sign_text(clean_text(e["t"]).strip())]

    items = []
    for e in ents:
        if not (table["x_left"] <= e["x"] <= table["x_right"]):
            continue
        if not (y_lo <= e["y"] <= y_hi):
            continue
        t = clean_text(e["t"]).strip()
        if not t:
            continue
        # 容差取 0.6：签名标签与人名插入点 y 完全相同，而相邻材料行至少差 4，
        # 用 2 会误删紧邻的材料行（实测师大南路 DN40 被删）。
        if any(abs(e["y"] - sy) <= 0.6 for sy in sign_ys):
            continue
        items.append((e["y"], e["x"], t))
    items.sort(key=lambda a: (-a[0], a[1]))

    # 行聚类容差：正常图用 300*sc；扁表格（行距很小）改用实测行距的一部分，
    # 避免把相邻材料行并成一条超长描述（实测 IE26I004 行距仅 8）。
    pitch = table.get("row_pitch", 0.0)
    line_tol = min(300 * sc, pitch * 0.5) if pitch else 300 * sc
    lines = []
    for it in items:
        if lines and abs(lines[-1]["y"] - it[0]) <= line_tol:
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
        elif table.get("layout") == "merged-spec":
            # 「名称及规格」合并列：成品描述就是这一列原样，忽略材料列。
            ln["descs"] = [t for t in cols.get("名称及规格", [])
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
        # 无序号行分两类：分组标题（如「厨房部分」「调压前材料：」）应独立成行，
        # 其余（续行描述）才归给最近的材料锚点。
        # 材料表大标题（如「厨房工程（ID26E025）」）位于第一条材料之前，
        # 括号内是纯工程编号；数据区中部同形式的文字则是真正的分组标题，应保留。
        first_anchor_idx = anchors[0]

        def _desc_of(ln):
            return " ".join(x.strip() for x in ln["descs"]).strip()

        def _numbered_title_at_top(ln, idx):
            """位于第一条材料之前、括号内为纯工程编号的材料表大标题。"""
            if idx >= first_anchor_idx:
                return False
            d = _desc_of(ln)
            return bool(re.search(r"[（(][A-Z]{2}\d{2}[A-Z]\d{3}[）)]", d))

        def _standalone_header(ln, idx):
            if ln["seq"] is not None or ln["is_header_line"]:
                return False
            if ln["unit"] or ln["qty"] is not None:
                return False
            d = _desc_of(ln)
            if not d or not is_header_row(d, "", None):
                return False
            return not _numbered_title_at_top(ln, idx)

        standalone = {i for i, ln in enumerate(seg_lines) if _standalone_header(ln, i)}
        table_titles = {i for i, ln in enumerate(seg_lines)
                        if ln["seq"] is None and not ln["is_header_line"]
                        and not ln["unit"] and ln["qty"] is None
                        and _numbered_title_at_top(ln, i)}
        owner = {}
        for i, ln in enumerate(seg_lines):
            if i in standalone or i in table_titles or ln["seq"] is not None:
                continue
            if not (ln["descs"] or ln["unit"] or ln["qty"] is not None):
                continue
            owner[i] = min(anchors, key=lambda j: abs(seg_lines[j]["y"] - ln["y"]))

        def _line_desc(group):
            if table.get("layout") in ("name-spec", "merged-spec"):
                # 成品表格式：名称列原样（含内部空格/下划线）+ 单个空格 + 规格列原样，
                # 不拼接材料列，也不把内部空格换成下划线。
                # 例：热镀锌钢管_燃气专用 DN50（D60.3×3.8）
                return " ".join(d.strip() for ln in group
                                for d in ln["descs"]).strip()
            return "\n".join(d for ln in group for d in ln["descs"]).strip()

        made = {}   # idx -> row
        for idx in anchors:
            anchor = seg_lines[idx]
            extra = [seg_lines[i] for i, j in owner.items() if j == idx]
            group = [anchor] + extra
            group.sort(key=lambda l: -l["y"])       # 描述按视觉从上到下拼接
            desc = _line_desc(group)
            if not desc:
                continue
            unit = anchor["unit"] or next((ln["unit"] for ln in extra if ln["unit"]), "")
            qty = anchor["qty"]
            if qty is None:
                qty = next((ln["qty"] for ln in extra if ln["qty"] is not None), None)
            made[idx] = {
                "seq": anchor["seq"],
                "desc": desc,
                "unit": unit,
                "qty": qty,
                "is_header": is_header_row(desc, unit, qty),
                "table": table["id"],
                "stage": table["stage"],
                "_y": anchor["y"],
            }
        for i in standalone:
            ln = seg_lines[i]
            desc = _line_desc([ln])
            if not desc:
                continue
            made[i] = {
                "seq": None, "desc": desc, "unit": "", "qty": None,
                "is_header": True, "table": table["id"], "stage": table["stage"],
                "_y": ln["y"],
            }
        # 按视觉顺序（与 seg_lines 一致）输出
        order = {id(ln): n for n, ln in enumerate(seg_lines)}
        ordered = sorted(made.items(),
                         key=lambda kv: order.get(id(seg_lines[kv[0]]), 0))
        for _, r in ordered:
            r.pop("_y", None)
            rows.append(r)
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
        # 无「第X页」标注时按序号走向/连贯性决定多表拼接顺序
        chosen = order_tables(chosen)
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
            print("  #%d  [%s%s]  %s | 工程=%s  表头y=%.0f  序号列x=%.0f  比例=%.2f  有效行=%d  首行=%s" % (
                t["id"], t["stage"], "(继承)" if t.get("stage_inherited") else "",
                t["title"] or "(未取到表名)", t.get("project") or "(未标注)", t["header_y"],
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
