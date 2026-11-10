#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
dwg-material-to-xls 一键安装器（纯 Python 标准库，无需 git、无需第三方包）。

最简用法（把仓库地址发给 AI 时，AI 可以直接执行这一条）：
    python install.py https://github.com/lonyii/skills-dwg-material-to-xls

其它用法：
    python install.py                                  # 用默认仓库地址安装
    python install.py --check                          # 只查看远端版本，不安装
    python install.py --local .                        # 从当前已下载/克隆的目录安装
    python install.py --dir "D:/skills"                # 自定义技能根目录
    python install.py --ref main --mirror https://ghproxy.net/

安装目标默认：~/.workbuddy/skills/dwg-material-to-xls
（Windows 为 %USERPROFILE%\\.workbuddy\\skills\\dwg-material-to-xls）
可用环境变量 WORKBUDDY_SKILLS_DIR 改变技能根目录，GITHUB_MIRROR 设置下载镜像，
GITHUB_TOKEN 用于私有仓库。
"""

from __future__ import annotations

__author__ = "lonyii"
__email__ = "lonyii@qq.com"

import argparse
import io
import json
import os
import re
import shutil
import sys
import tempfile
import urllib.request
import zipfile
from datetime import datetime
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass

SKILL_NAME = "dwg-material-to-xls"
DEFAULT_REPO = "https://github.com/lonyii/skills-dwg-material-to-xls"
FALLBACK_BRANCH = "main"
REQUIRED_FILES = ("SKILL.md", "scripts/import_materials.py", "scripts/dwg_tables.py")
EXCLUDE_NAMES = {".git", ".github", "__pycache__", ".backup", ".backups",
                 ".idea", ".vscode", "skill-backups"}
IGNORE_PATTERNS = shutil.ignore_patterns(
    "__pycache__", "*.pyc", "*.pyo", ".DS_Store", "Thumbs.db", "~$*", "*.tmp")


def log(msg=""):
    print(msg, flush=True)


def default_skills_dir() -> Path:
    env = os.environ.get("WORKBUDDY_SKILLS_DIR")
    return Path(env).expanduser() if env else Path.home() / ".workbuddy" / "skills"


def target_dir(base: Path) -> Path:
    base = Path(base).expanduser()
    return base if base.name == SKILL_NAME else base / SKILL_NAME


def parse_repo(text: str):
    """支持 https://github.com/owner/repo(.git)、git@github.com:owner/repo.git、owner/repo。"""
    s = (text or DEFAULT_REPO).strip().rstrip("/")
    s = re.sub(r"\.git$", "", s)
    m = re.search(r"github\.com[:/]([^/]+)/([^/#?]+)", s)
    if m:
        return m.group(1), m.group(2)
    m = re.fullmatch(r"([^/\s]+)/([^/\s]+)", s)
    if m:
        return m.group(1), m.group(2)
    raise ValueError("无法识别仓库地址：%s" % text)


def http_get(url: str, timeout=30, token=None) -> bytes:
    headers = {"User-Agent": "dwg-material-to-xls-installer/2.4.2"}
    if token:
        headers["Authorization"] = "Bearer " + token
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def with_mirror(url: str, mirror: str | None) -> str:
    if not mirror:
        return url
    m = mirror.rstrip("/")
    return url if url.startswith(m) else m + "/" + url


def resolve_branch(owner: str, repo: str, ref: str | None) -> str:
    if ref:
        return ref
    try:
        data = json.loads(http_get(
            "https://api.github.com/repos/%s/%s" % (owner, repo), timeout=15
        ).decode("utf-8"))
        return data.get("default_branch") or FALLBACK_BRANCH
    except Exception:  # noqa: BLE001 - 离线时退回常见分支名
        return FALLBACK_BRANCH


def zip_urls(owner: str, repo: str, ref: str):
    return [
        "https://codeload.github.com/%s/%s/zip/refs/heads/%s" % (owner, repo, ref),
        "https://github.com/%s/%s/archive/refs/heads/%s.zip" % (owner, repo, ref),
        "https://codeload.github.com/%s/%s/zip/refs/tags/%s" % (owner, repo, ref),
    ]


def download_zip(owner: str, repo: str, ref: str, mirror=None, token=None):
    errors = []
    for url in zip_urls(owner, repo, ref):
        full = with_mirror(url, mirror)
        try:
            log("  尝试下载 %s" % full)
            data = http_get(full, timeout=90, token=token)
            if data[:2] != b"PK":
                errors.append("%s -> 返回的不是 zip" % full)
                continue
            return data, full
        except Exception as exc:  # noqa: BLE001
            errors.append("%s -> %s" % (full, str(exc)[:120]))
    raise RuntimeError(
        "下载仓库失败：\n  " + "\n  ".join(errors) +
        "\n可设置 --mirror 或环境变量 GITHUB_MIRROR 使用镜像；"
        "也可先手动下载 zip 后用 --local <解压目录> 安装。")


def extract_zip(data: bytes, dest: Path):
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        for name in zf.namelist():
            parts = Path(name.replace("\\", "/")).parts
            if any(p == ".." for p in parts) or name.startswith("/"):
                raise RuntimeError("压缩包含非法路径：%s" % name)
        zf.extractall(dest)


def locate_root(base: Path) -> Path:
    """GitHub zip 解压后通常有一层 repo-main/ 目录。"""
    for cand in (base, *sorted(p for p in base.iterdir() if p.is_dir())):
        if (cand / "SKILL.md").is_file():
            return cand
    for found in base.rglob("SKILL.md"):
        return found.parent
    raise RuntimeError("压缩包中找不到 SKILL.md，确认地址指向本技能仓库。")


def validate(src: Path):
    missing = [f for f in REQUIRED_FILES if not (src / f).is_file()]
    if missing:
        raise RuntimeError("技能包不完整，缺少：%s" % ", ".join(missing))
    head = (src / "SKILL.md").read_text(encoding="utf-8", errors="replace")[:600]
    if "dwg-material-to-xls" not in head:
        raise RuntimeError("SKILL.md 的技能名不匹配，拒绝安装到 %s" % SKILL_NAME)
    if not list((src / "assets").glob("*.xls")):
        raise RuntimeError("assets/ 下没有 .xls 模板，技能包不完整。")


def backup(dst: Path) -> Path | None:
    if not dst.exists() or not any(dst.iterdir()):
        return None
    ver = "unknown"
    vf = dst / "version.json"
    if vf.is_file():
        try:
            ver = json.loads(vf.read_text(encoding="utf-8")).get("version", ver)
        except Exception:  # noqa: BLE001
            pass
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    bak_root = dst.parent.parent / "skill-backups"
    bak = bak_root / ("%s-%s-%s" % (SKILL_NAME, ver, stamp))
    bak.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(dst, bak, ignore=shutil.ignore_patterns(".git", "__pycache__", "*.pyc"))
    return bak


def copy_skill(src: Path, dst: Path):
    """只覆盖仓库自带文件，不删除目标目录里的其它文件（保留用户本地内容）。"""
    dst.mkdir(parents=True, exist_ok=True)
    for item in sorted(src.iterdir()):
        if item.name in EXCLUDE_NAMES or item.name.startswith(".backup"):
            continue
        to = dst / item.name
        if item.is_dir():
            shutil.copytree(item, to, dirs_exist_ok=True, ignore=IGNORE_PATTERNS)
        else:
            if item.suffix.lower() in (".pyc", ".pyo", ".tmp", ".log"):
                continue
            shutil.copy2(item, to)


def installed_version(dst: Path) -> str:
    vf = dst / "version.json"
    if vf.is_file():
        try:
            return json.loads(vf.read_text(encoding="utf-8")).get("version", "?")
        except Exception:  # noqa: BLE001
            pass
    return "未记录"


def post_install(dst: Path):
    log()
    log("环境自检：")
    ok = True
    try:
        import win32com.client  # noqa: F401
        log("  [OK] pywin32 已安装")
    except ImportError:
        ok = False
        log("  [待安装] pywin32 —— 运行：python -m pip install pywin32")
    if sys.platform == "win32":
        log("  [需确认] 本机需有 Microsoft Excel 与浩辰/AutoCAD/中望 CAD 之一（COM 读取 DWG）")
    else:
        log("  [提示] 本技能的 CAD/Excel 自动化需要 Windows；其它系统仅能离线使用 rows.json")
    log("  %s" % ("依赖就绪。" if ok else "装好缺失依赖后即可使用。"))
    log()
    log("安装完成！在 AI 对话里直接说：")
    log('  「把这些 dwg 的材料表导出，生成材料平差表」并附上若干图纸即可')
    log("若 WorkBuddy 已打开，刷新/重启一次以加载新技能。")


def main(argv=None):
    p = argparse.ArgumentParser(description="安装 dwg-material-to-xls 技能")
    p.add_argument("repo", nargs="?", default=DEFAULT_REPO,
                   help="GitHub 仓库地址或 owner/repo，默认本仓库")
    p.add_argument("--dir", dest="skills_dir", default=None,
                   help="技能根目录，默认 ~/.workbuddy/skills")
    p.add_argument("--ref", default=None, help="分支或 tag，默认读取仓库默认分支")
    p.add_argument("--mirror", default=os.environ.get("GITHUB_MIRROR"),
                   help="下载镜像前缀，例如 https://ghproxy.net/")
    p.add_argument("--token", default=os.environ.get("GITHUB_TOKEN"),
                   help="私有仓库访问令牌（建议用环境变量 GITHUB_TOKEN）")
    p.add_argument("--local", default=None,
                   help="从本地已解压/克隆的目录安装，跳过网络下载")
    p.add_argument("--check", action="store_true", help="只查看远端版本，不安装")
    p.add_argument("--dry-run", action="store_true", help="下载并校验，但不写入磁盘")
    args = p.parse_args(argv)

    base = Path(args.skills_dir).expanduser() if args.skills_dir else default_skills_dir()
    dst = target_dir(base)
    owner, repo = parse_repo(args.repo)
    ref = args.ref or FALLBACK_BRANCH
    log("技能：%s" % SKILL_NAME)
    log("仓库：https://github.com/%s/%s" % (owner, repo))
    log("目标：%s" % dst)

    try:
        if args.check:
            if not args.ref:
                ref = resolve_branch(owner, repo, None)
            url = with_mirror(
                "https://raw.githubusercontent.com/%s/%s/%s/version.json" % (owner, repo, ref),
                args.mirror)
            info = json.loads(http_get(url, timeout=20, token=args.token).decode("utf-8"))
            log("远端版本：%s（%s）" % (info.get("version"), info.get("released", "")))
            log("本地版本：%s" % installed_version(dst))
            return 0

        if args.local:
            src = Path(args.local).expanduser().resolve()
            if src.name != SKILL_NAME and (src / SKILL_NAME).is_dir():
                src = src / SKILL_NAME
            src = locate_root(src)
            validate(src)
            log("本地来源：%s" % src)
        else:
            if not args.ref:
                ref = resolve_branch(owner, repo, None)
            log("分支：%s" % ref)
            data, url = download_zip(owner, repo, ref, args.mirror, args.token)
            log("已下载：%s (%d KB)" % (url, len(data) // 1024))
            with tempfile.TemporaryDirectory(prefix="dwg-skill-") as td:
                root = Path(td) / "src"
                root.mkdir()
                extract_zip(data, root)
                src = locate_root(root)
                validate(src)
                log("技能包校验通过（SKILL.md / scripts / references / assets）")
                return finish(src, dst, args)

        return finish(src, dst, args)
    except Exception as exc:  # noqa: BLE001
        log("\n[安装失败] %s" % exc)
        return 1


def finish(src: Path, dst: Path, args) -> int:
    if args.dry_run:
        log("\n[dry-run] 校验通过，未写入磁盘。")
        return 0
    bak = backup(dst)
    if bak:
        log("已备份旧版本：%s" % bak)
    copy_skill(src, dst)
    log("已安装版本：%s" % installed_version(dst))
    post_install(dst)
    return 0


if __name__ == "__main__":
    sys.exit(main())
