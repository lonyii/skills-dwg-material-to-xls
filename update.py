#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
dwg-material-to-xls 版本检查与升级（纯 Python 标准库）。

    python update.py --check     # 只检查是否有新版本
    python update.py             # 升级到最新版（自动备份旧版本）
    python update.py --force     # 版本相同也强制重装（修复本地文件损坏）
    python update.py --ref v1.2.0 --mirror https://ghproxy.net/

版本来源：远端 version.json；副本会备份到 ~/.workbuddy/skill-backups/。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from install import (  # noqa: E402
    DEFAULT_REPO, SKILL_NAME, backup, copy_skill, default_skills_dir, download_zip,
    extract_zip, http_get, installed_version, locate_root, log, parse_repo,
    resolve_branch, target_dir, validate, with_mirror,
)

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass


def vtuple(text: str):
    nums = [int(x) for x in re.findall(r"\d+", str(text))]
    return tuple((nums + [0, 0, 0])[:3])


def remote_version(owner, repo, ref, mirror, token):
    urls = [
        "https://raw.githubusercontent.com/%s/%s/%s/version.json" % (owner, repo, ref),
        "https://cdn.jsdelivr.net/gh/%s/%s@%s/version.json" % (owner, repo, ref),
        "https://github.com/%s/%s/raw/%s/version.json" % (owner, repo, ref),
    ]
    errors = []
    for url in urls:
        # jsDelivr 一般不走 GitHub 镜像，避免拼接出无效地址
        full = url if "jsdelivr" in url else with_mirror(url, mirror)
        try:
            info = json.loads(http_get(full, timeout=20, token=token).decode("utf-8"))
            if info.get("version"):
                return info, full
        except Exception as exc:  # noqa: BLE001
            errors.append("%s -> %s" % (full, str(exc)[:100]))
    raise RuntimeError("读取远端版本失败：\n  " + "\n  ".join(errors))


def main(argv=None):
    p = argparse.ArgumentParser(description="检查/升级 dwg-material-to-xls 技能")
    p.add_argument("--check", action="store_true", help="只检查版本，不修改文件")
    p.add_argument("--force", action="store_true", help="强制重装当前远端版本")
    p.add_argument("--dir", dest="skills_dir", default=None, help="技能根目录")
    p.add_argument("--repo", default=DEFAULT_REPO, help="仓库地址或 owner/repo")
    p.add_argument("--ref", default=None, help="分支或 tag，默认仓库默认分支")
    p.add_argument("--mirror", default=os.environ.get("GITHUB_MIRROR"))
    p.add_argument("--token", default=os.environ.get("GITHUB_TOKEN"))
    args = p.parse_args(argv)

    base = Path(args.skills_dir).expanduser() if args.skills_dir else default_skills_dir()
    dst = target_dir(base)
    owner, repo = parse_repo(args.repo)
    ref = args.ref or resolve_branch(owner, repo, None)

    local = installed_version(dst)
    try:
        info, url = remote_version(owner, repo, ref, args.mirror, args.token)
    except Exception as exc:  # noqa: BLE001
        log("[版本检查失败] %s" % exc)
        return 1
    remote = str(info.get("version"))
    log("本地版本：%s" % local)
    log("远端版本：%s（%s）" % (remote, info.get("released", "")))
    log("版本来源：%s" % url)

    if args.check:
        if vtuple(remote) > vtuple(local):
            log("→ 有新版本，运行 python update.py 升级")
        else:
            log("→ 已是最新版本")
        return 0

    if vtuple(remote) == vtuple(local) and not args.force:
        log("已是最新版本，无需升级（可用 --force 强制重装）")
        return 0
    if vtuple(remote) < vtuple(local) and not args.force:
        log("远端版本低于本地，已中止（可用 --force 强制降级）")
        return 0

    try:
        data, zip_url = download_zip(owner, repo, ref, args.mirror, args.token)
        log("已下载：%s (%d KB)" % (zip_url, len(data) // 1024))
        with tempfile.TemporaryDirectory(prefix="dwg-skill-upd-") as td:
            root = Path(td) / "src"
            root.mkdir()
            extract_zip(data, root)
            src = locate_root(root)
            validate(src)
            bak = backup(dst)
            if bak:
                log("已备份旧版本：%s" % bak)
            copy_skill(src, dst)
    except Exception as exc:  # noqa: BLE001
        log("[升级失败] %s" % exc)
        return 1

    log("已升级到 %s：%s" % (installed_version(dst), dst))
    changes = info.get("changelog") or []
    if changes:
        log("\n本次更新内容：")
        for line in changes[:3]:
            log("  - %s" % line)
    return 0


if __name__ == "__main__":
    sys.exit(main())
