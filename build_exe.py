#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""一键把本技能打包成单文件 EXE（材料平差批量生成器）。

用法：
    python build_exe.py            # 生成 dist/材料平差批量生成器.exe
    python build_exe.py --open     # 打包完成后打开 dist 目录

中文编码处理：
- 本文件与 gui_app.py 均为 UTF-8；
- 运行时把 stdout/stderr 切到 UTF-8，并把 Windows 控制台切到 65001；
- EXE 名称、输出路径含中文时不依赖 cmd 引号转义，全部用 subprocess 参数列表传递；
- 构建产物写入 dist/、build/，已加入 .gitignore，不会污染仓库。
"""

from __future__ import annotations

import glob
import json
import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(ROOT, "version.json"), encoding="utf-8") as _f:
    VERSION = json.load(_f)["version"]

# 打包前把 GUI 里的 APP_VERSION 同步为 version.json 的版本，
# 否则「检查更新」会用旧版本号判断，导致同一版本被反复下载。
_GUI = os.path.join(ROOT, "gui_app.py")
with open(_GUI, encoding="utf-8") as _f:
    _gui_src = _f.read()
_gui_src = re.sub(r'APP_VERSION = "\d+\.\d+\.\d+"', 'APP_VERSION = "%s"' % VERSION, _gui_src, count=1)
with open(_GUI, "w", encoding="utf-8") as _f:
    _f.write(_gui_src)
# 文件名带版本号，便于「检查更新」在不运行 EXE 的情况下比较版本
EXE_NAME = "材料平差批量生成器_v%s" % VERSION
DIST = os.path.join(ROOT, "dist")
BUILD = os.path.join(ROOT, "build")
ASSETS = os.path.join(ROOT, "assets")
TEMPLATE = os.path.join(ASSETS, "template_summary.xls")
GUI_SOURCE = os.path.join(ROOT, "gui_app.py")


def setup_console_utf8():
    """尽量让中文日志在 cmd / PowerShell / AI 终端里不乱码。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        except Exception:
            pass


def has_module(python_exe, module):
    try:
        r = subprocess.run(
            [python_exe, "-c", "import %s" % module],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=20,
        )
        return r.returncode == 0
    except Exception:
        return False


def find_python_with_tkinter():
    """优先使用当前解释器；缺 tkinter 时自动找 py 启动器或常见安装路径。"""
    candidates = [sys.executable]
    if os.name == "nt":
        for ver in ("3.12", "3.11", "3.10", "3.9", "3"):
            candidates.append("py:-" + ver)
        patterns = [
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\Python\Python3*\python.exe"),
            r"C:\Python3*\python.exe",
        ]
        for pattern in patterns:
            candidates.extend(sorted(glob.glob(pattern), reverse=True))
    for cand in candidates:
        if cand.startswith("py:"):
            exe, arg = "py", cand[3:]
            cmd = [exe, arg, "-c", "import sys; print(sys.executable)"]
        else:
            cmd = [cand, "-c", "import sys; print(sys.executable)"]
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=20)
        except Exception:
            continue
        if r.returncode != 0:
            continue
        real = r.stdout.strip().splitlines()[-1] if r.stdout.strip() else cand
        if real and has_module(real, "tkinter"):
            return real
    return None


def ensure_current_interpreter():
    """若当前 Python 没有 tkinter，自动切换到带 tkinter 的解释器重新执行。"""
    if has_module(sys.executable, "tkinter"):
        return
    target = find_python_with_tkinter()
    if not target:
        raise SystemExit(
            "未找到带 tkinter 的 Python。请安装 python.org 的 Python 3.9+ "
            "（勾选 tcl/tk），或先把整个技能目录交给 AI 处理。"
        )
    target = os.path.abspath(target)
    if os.path.normcase(target) == os.path.normcase(os.path.abspath(sys.executable)):
        return
    print("当前 Python 缺少 tkinter，已切换到：%s" % target)
    os.execv(target, [target, os.path.abspath(__file__)] + sys.argv[1:])


def main():
    setup_console_utf8()
    ensure_current_interpreter()

    for path in (GUI_SOURCE, TEMPLATE):
        if not os.path.isfile(path):
            raise SystemExit("缺少打包所需文件：%s" % path)

    if not has_module(sys.executable, "PyInstaller"):
        print("未检测到 PyInstaller，正在自动安装…")
        r = subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"])
        if r.returncode != 0:
            raise SystemExit("PyInstaller 安装失败，请手动执行 pip install pyinstaller")

    if not has_module(sys.executable, "tkinterdnd2"):
        print("未检测到 tkinterdnd2，正在自动安装…")
        r = subprocess.run([sys.executable, "-m", "pip", "install", "tkinterdnd2"])
        if r.returncode != 0:
            raise SystemExit("tkinterdnd2 安装失败，请手动执行 pip install tkinterdnd2")

    os.makedirs(DIST, exist_ok=True)
    os.makedirs(BUILD, exist_ok=True)

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile", "--noconsole", "--clean", "--noconfirm",
        "--name", EXE_NAME,
        "--distpath", DIST,
        "--workpath", BUILD,
        "--specpath", ROOT,
        "--paths", os.path.join(ROOT, "scripts"),
        "--add-data", TEMPLATE + os.pathsep + "assets",
        "--hidden-import", "tkinterdnd2",
        "--collect-all", "tkinterdnd2",
        "--hidden-import", "auto_export",
        "--hidden-import", "dwg_tables",
        "--hidden-import", "import_materials",
        "--hidden-import", "win32com.client",
        "--hidden-import", "win32com.client.gencache",
        "--hidden-import", "pythoncom",
        "--hidden-import", "win32clipboard",
        "--hidden-import", "win32con",
        "--hidden-import", "winreg",
        GUI_SOURCE,
    ]
    print("开始打包：%s" % EXE_NAME)
    r = subprocess.run(cmd, cwd=ROOT)
    if r.returncode != 0:
        raise SystemExit("打包失败，请查看上方 PyInstaller 日志。")

    exe = os.path.join(DIST, EXE_NAME + ".exe")
    if not os.path.isfile(exe):
        raise SystemExit("打包结束但未找到 EXE：%s" % exe)
    print("打包完成：%s（%.1f MB）" % (exe, os.path.getsize(exe) / 1024 / 1024))

    if "--open" in sys.argv[1:] and os.name == "nt":
        try:
            os.startfile(DIST)
        except OSError:
            pass


if __name__ == "__main__":
    main()
