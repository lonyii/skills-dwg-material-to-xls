# -*- coding: utf-8 -*-
"""依赖自检：Python / pywin32 / CAD / Excel，给 run_batch.bat 启动前调用。

只读注册表判断 COM 组件是否注册，不启动 CAD/Excel（快、无副作用）。
返回码：0=全部就绪；1=缺少关键依赖（缺 CAD 或 Excel 或 pywin32）。
"""
import sys

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding=sys.stdout.encoding or "utf-8", errors="replace")
    except Exception:                      # noqa: BLE001
        pass

CAD_PROGIDS = (
    ("GStarCAD.Application", "浩辰 CAD"),
    ("Gcad.Application", "浩辰 CAD(旧)"),
    ("AutoCAD.Application", "AutoCAD"),
    ("ZWCAD.Application", "中望 CAD"),
)
EXCEL_PROGIDS = (
    ("Excel.Application", "Microsoft Excel"),
    ("KET.Application", "WPS 表格"),
)


def com_registered(progid):
    """查 HKCR\\<ProgID>\\CLSID 是否存在（只读、不启动程序）。"""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + "\\CLSID"):
            return True
    except OSError:
        pass
    # 部分机器 ProgID 只有 CLSID 值在根键
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid) as k:
            winreg.QueryValue(k, "")
        return True
    except OSError:
        return False


def main():
    ok = True
    problems = []
    pyver = sys.version.split()[0]
    if sys.version_info < (3, 8):
        problems.append("Python 需 3.8+（当前 %s）" % pyver)
        ok = False

    try:
        import win32com.client  # noqa: F401
    except ImportError:
        import subprocess
        subprocess.run([sys.executable, "-m", "pip", "install", "pywin32"],
                       capture_output=True, text=True)
        try:
            import win32com.client  # noqa: F401
        except ImportError:
            problems.append("缺少 pywin32：%s -m pip install pywin32" % sys.executable)
            ok = False

    cads = [name for pid, name in CAD_PROGIDS if com_registered(pid)]
    if not cads:
        problems.append("未检测到 CAD，请安装并打开一次浩辰/AutoCAD/中望 CAD")
        ok = False

    excels = [name for pid, name in EXCEL_PROGIDS if com_registered(pid)]
    if not excels:
        problems.append("未检测到 Excel/WPS 表格，请安装并打开一次")
        ok = False

    # 正常时一行带过；有问题才逐条列出
    if ok:
        def short(names):
            out = []
            for n in dict.fromkeys(names):
                if "浩辰" in n:
                    out.append("浩辰")
                elif "AutoCAD" in n:
                    out.append("AutoCAD")
                elif "中望" in n:
                    out.append("中望")
                elif "WPS" in n:
                    out.append("WPS")
                elif "Excel" in n:
                    out.append("Excel")
                else:
                    out.append(n)
            return "/".join(dict.fromkeys(out))
        print("环境：Python %s | CAD %s | 表格 %s | 正常" % (
            pyver, short(cads), short(excels)))
    else:
        for p in problems:
            print("  [缺] %s" % p)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
