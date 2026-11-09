#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
维护者发布脚本：更新版本记录、生成 commit 与 tag，可选推送。

    python release.py 1.2.0 "新增 XX 设计院图框支持"
    python release.py 1.2.0 "修复汇总同步" --push
    python release.py --bump patch "修正列偏移"

发布流程：修改 version.json → git add -A → commit → tag vX.Y.Z →（--push）推送。
使用者不需要此脚本，日常更新请用 update.py。
"""
from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent
VERSION_FILE = ROOT / "version.json"

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass


def run(cmd, check=True):
    print("  $", " ".join(cmd), flush=True)
    r = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True, encoding="utf-8",
                       errors="replace")
    if r.stdout.strip():
        print("   ", r.stdout.strip()[:400])
    if check and r.returncode != 0:
        raise RuntimeError("命令失败：%s\n%s" % (" ".join(cmd), r.stderr.strip()[:400]))
    return r


def main(argv=None):
    p = argparse.ArgumentParser(description="发布 dwg-material-to-xls 新版本")
    p.add_argument("version", nargs="?", help="新版本号，如 1.2.0")
    p.add_argument("message", nargs="?", default="", help="本次更新说明")
    p.add_argument("--bump", choices=("major", "minor", "patch"),
                   help="按当前版本自动递增，代替手写版本号")
    p.add_argument("--push", action="store_true", help="提交后推送分支与 tag")
    args = p.parse_args(argv)

    info = json.loads(VERSION_FILE.read_text(encoding="utf-8"))
    old = str(info.get("version", "0.0.0"))
    nums = [int(x) for x in re.findall(r"\d+", old)][:3]
    while len(nums) < 3:
        nums.append(0)

    if args.bump:
        idx = {"major": 0, "minor": 1, "patch": 2}[args.bump]
        nums[idx] += 1
        for i in range(idx + 1, 3):
            nums[i] = 0
        new = ".".join(map(str, nums))
    else:
        new = (args.version or "").strip().lstrip("v")
    if not re.fullmatch(r"\d+\.\d+\.\d+", new):
        raise SystemExit("版本号格式应为 x.y.z（或改用 --bump patch/minor/major）")
    if tuple(map(int, new.split("."))) <= tuple(nums):
        raise SystemExit("新版本 %s 必须大于当前版本 %s" % (new, old))

    tag = "v" + new
    exists = run(["git", "tag", "-l", tag], check=False)
    if exists.stdout.strip():
        raise SystemExit("tag %s 已存在，请先处理或换新版本号" % tag)

    today = date.today().isoformat()
    info["version"] = new
    info["released"] = today
    entry = "%s (%s) %s" % (new, today, args.message or "例行更新")
    info.setdefault("changelog", []).insert(0, entry)
    VERSION_FILE.write_text(json.dumps(info, ensure_ascii=False, indent=2) + "\n",
                            encoding="utf-8")
    print("版本：%s -> %s" % (old, new))

    run(["git", "add", "-A"])
    run(["git", "commit", "-m", "release: %s %s" % (tag, args.message or "")])
    run(["git", "tag", "-a", tag, "-m", entry])
    if args.push:
        run(["git", "push", "origin", "HEAD"])
        run(["git", "push", "origin", tag])
        print("\n已发布 %s。使用者运行 python update.py 即可升级。" % tag)
    else:
        print("\n本地已提交 %s。推送命令：" % tag)
        print("  git push origin HEAD && git push origin %s" % tag)
    return 0


if __name__ == "__main__":
    sys.exit(main())
