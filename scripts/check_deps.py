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
    print("=" * 46)
    print("  运行环境自检")
    print("=" * 46)

    print("Python 版本：%s" % sys.version.split()[0])
    if sys.version_info < (3, 8):
        print("  [错误] 需要 Python 3.8 及以上，请升级。")
        ok = False

    try:
        import win32com.client  # noqa: F401
        print("pywin32：已安装")
    except ImportError:
        print("pywin32：[缺失] 正在尝试自动安装……")
        import subprocess
        r = subprocess.run([sys.executable, "-m", "pip", "install", "pywin32"],
                           capture_output=True, text=True)
        try:
            import win32com.client  # noqa: F401
            print("  pywin32 安装成功")
        except ImportError:
            print("  [错误] pywin32 自动安装失败，请手动运行：")
            print("     %s -m pip install pywin32" % sys.executable)
            print("     %s" % (r.stderr or "").strip()[:200])
            ok = False

    cads = [name for pid, name in CAD_PROGIDS if com_registered(pid)]
    if cads:
        print("CAD 组件：已检测到 %s" % "、".join(dict.fromkeys(cads)))
    else:
        print("CAD 组件：[未检测到] 请先安装并打开一次浩辰/AutoCAD/中望 CAD，")
        print("  （CAD 首次启动会注册 COM 组件，之后本工具才能读取图纸）")
        ok = False

    excels = [name for pid, name in EXCEL_PROGIDS if com_registered(pid)]
    if excels:
        print("表格组件：已检测到 %s" % "、".join(dict.fromkeys(excels)))
    else:
        print("表格组件：[未检测到] 请先安装 Microsoft Excel 或 WPS 表格，")
        print("  （用于生成 .xls 成品；首次打开一次以注册 COM 组件）")
        ok = False

    print("-" * 46)
    if ok:
        print("自检通过，开始处理图纸……")
    else:
        print("[依赖不全] 请按上面提示安装缺失项后，重新拖放运行。")
    print()
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
