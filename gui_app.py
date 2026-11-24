# -*- coding: utf-8 -*-
"""材料平差批量生成器（单文件 EXE 入口）。

把竣工图/施工图 DWG 拖入窗口，或从剪贴板粘贴文件/文件夹，批量生成材料平差表。
成品统一写入桌面的「材料平差输出_时间戳」文件夹，完成后自动打开该文件夹。
"""

from __future__ import annotations

import ctypes
import os
import queue
import re
import sys
import tempfile
import threading
import time
import traceback
import winreg
from ctypes import c_int, c_ssize_t, c_void_p, create_unicode_buffer, wintypes
from ctypes import CFUNCTYPE
from ctypes.wintypes import HWND, LPARAM, UINT, WPARAM

import pythoncom
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32clipboard
import win32con

# 技能包布局：脚本在 scripts/；独立打包布局：脚本与本文件同目录。
_SCRIPT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "scripts")
if os.path.isdir(_SCRIPT_DIR):
    sys.path.insert(0, _SCRIPT_DIR)

import auto_export
import import_materials


# --------------------------------------------------------------------------- #
# 资源路径：兼容源码运行与 PyInstaller onefile
# --------------------------------------------------------------------------- #
BASE_DIR = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
TEMPLATE_XLS = os.path.join(ASSETS_DIR, "template_summary.xls")

# 让内置模板成为技能脚本的默认模板，无需外部模板文件。
import_materials.ASSETS_DIR = ASSETS_DIR
import_materials.TEMPLATE_SUMMARY = TEMPLATE_XLS
import_materials.TEMPLATE_SAMPLE = TEMPLATE_XLS
import_materials.DEFAULT_TEMPLATE = TEMPLATE_XLS
import_materials.TEMPLATE_ALIASES["summary"] = TEMPLATE_XLS
import_materials.TEMPLATE_ALIASES["sample"] = TEMPLATE_XLS


VALID_EXT = (".dwg", ".dxf")
WM_DROPFILES = 0x0233
GWL_WNDPROC = -4


def desktop_dir() -> str:
    """获取真实桌面路径（兼容重定向到 D 盘等情况）。"""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
        )
        with key:
            value, _ = winreg.QueryValueEx(key, "Desktop")
            path = os.path.expandvars(value)
            if os.path.isdir(path):
                return path
    except Exception:
        pass
    fallback = os.path.join(os.path.expanduser("~"), "Desktop")
    return fallback if os.path.isdir(fallback) else os.path.abspath(".")


def stamp() -> str:
    return time.strftime("%Y%m%d_%H%M%S")


def normalize(path: str) -> str:
    return os.path.normcase(os.path.abspath(path))


# --------------------------------------------------------------------------- #
# Windows 文件拖放（WM_DROPFILES），不依赖第三方拖放库
# --------------------------------------------------------------------------- #
LRESULT = c_ssize_t
WNDPROC = CFUNCTYPE(LRESULT, HWND, UINT, WPARAM, LPARAM)
_user32 = ctypes.windll.user32
_shell32 = ctypes.windll.shell32

_set_window_proc = getattr(_user32, "SetWindowLongPtrW", _user32.SetWindowLongW)
_set_window_proc.argtypes = [HWND, c_int, c_void_p]
_set_window_proc.restype = c_void_p

_call_window_proc = _user32.CallWindowProcW
_call_window_proc.argtypes = [c_void_p, HWND, UINT, WPARAM, LPARAM]
_call_window_proc.restype = LRESULT

_shell32.DragAcceptFiles.argtypes = [HWND, wintypes.BOOL]
_shell32.DragQueryFileW.argtypes = [c_void_p, UINT, wintypes.LPWSTR, UINT]
_shell32.DragQueryFileW.restype = UINT
_shell32.DragFinish.argtypes = [c_void_p]

_DROP_STATES = []


def enable_file_drop(widget, callback) -> None:
    widget.update_idletasks()
    try:
        hwnd = int(widget.winfo_id())
    except Exception:
        return
    if hwnd <= 0:
        return
    state = {"old": 0, "proc": None}

    def wndproc(h_wnd, msg, w_param, l_param):
        if msg == WM_DROPFILES:
            hdrop = c_void_p(w_param)
            count = _shell32.DragQueryFileW(hdrop, 0xFFFFFFFF, None, 0)
            files = []
            for i in range(count):
                length = _shell32.DragQueryFileW(hdrop, i, None, 0)
                buf = create_unicode_buffer(length + 1)
                _shell32.DragQueryFileW(hdrop, i, buf, length + 1)
                files.append(buf.value)
            _shell32.DragFinish(hdrop)
            widget.after(0, lambda: callback(files))
            return 0
        return _call_window_proc(state["old"], h_wnd, msg, w_param, l_param)

    state["proc"] = WNDPROC(wndproc)
    state["old"] = _set_window_proc(
        hwnd, GWL_WNDPROC, ctypes.cast(state["proc"], c_void_p).value
    ) or 0
    _shell32.DragAcceptFiles(hwnd, True)
    _DROP_STATES.append(state)


def clipboard_paths():
    """读取剪贴板中的 CF_HDROP 文件/文件夹列表。"""
    try:
        win32clipboard.OpenClipboard()
        try:
            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_HDROP):
                return list(win32clipboard.GetClipboardData(win32con.CF_HDROP))
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        pass
    return []


class QueueWriter:
    """把子进程/工作线程 stdout 转发到界面日志队列。"""

    def __init__(self, q: queue.Queue):
        self.q = q

    def write(self, text):
        if text:
            self.q.put(("log", str(text)))
        return len(text)

    def flush(self):
        pass

    def isatty(self):
        return False


# --------------------------------------------------------------------------- #
# 界面
# --------------------------------------------------------------------------- #
class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("材料平差批量生成器")
        self.root.geometry("1040x700")
        self.root.minsize(900, 600)
        self.root.configure(bg="#F3F6FB")

        self.paths = []            # [{"path": str, "kind": "图纸"|"文件夹"}]
        self.seen = set()
        self.running = False
        self.output_dir = ""
        self.q = queue.Queue()

        self._build_style()
        self._build_ui()
        self._bind_keys()

        self.root.update_idletasks()
        enable_file_drop(self.root, self.add_paths)
        enable_file_drop(self.tree, self.add_paths)
        enable_file_drop(self.log_text, self.add_paths)

    def _build_style(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Root.TFrame", background="#F3F6FB")
        style.configure("Card.TFrame", background="#FFFFFF")
        style.configure(
            "Title.TLabel", background="#F3F6FB", foreground="#111827",
            font=("Microsoft YaHei UI", 18, "bold"),
        )
        style.configure(
            "Sub.TLabel", background="#F3F6FB", foreground="#6B7280",
            font=("Microsoft YaHei UI", 10),
        )
        style.configure(
            "Card.TLabel", background="#FFFFFF", foreground="#4B5563",
            font=("Microsoft YaHei UI", 10),
        )
        style.configure(
            "Hint.TLabel", background="#F3F6FB", foreground="#6B7280",
            font=("Microsoft YaHei UI", 9),
        )
        style.configure("TButton", padding=(10, 6), font=("Microsoft YaHei UI", 10))
        style.configure(
            "Accent.TButton", background="#2563EB", foreground="#FFFFFF",
            borderwidth=0, focuscolor="#2563EB",
            font=("Microsoft YaHei UI", 10, "bold"), padding=(14, 7),
        )
        style.map(
            "Accent.TButton",
            background=[("active", "#1D4ED8"), ("disabled", "#93C5FD")],
            foreground=[("disabled", "#FFFFFF")],
        )
        style.configure(
            "Treeview", rowheight=28, fieldbackground="#FFFFFF",
            background="#FFFFFF", font=("Microsoft YaHei UI", 9),
        )
        style.configure(
            "Treeview.Heading", background="#EEF2F7", foreground="#374151",
            font=("Microsoft YaHei UI", 9, "bold"), padding=(6, 6),
        )
        style.configure(
            "TProgressbar", troughcolor="#E5E7EB", background="#2563EB",
            thickness=8, borderwidth=0,
        )

    def _build_ui(self):
        root = self.root
        root.columnconfigure(0, weight=1)
        root.rowconfigure(2, weight=1)

        # 标题
        head = ttk.Frame(root, style="Root.TFrame", padding=(2, 6, 2, 10))
        head.grid(row=0, column=0, sticky="ew")
        ttk.Label(head, text="材料平差批量生成器", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            head,
            text="拖入 DWG / DXF 文件或文件夹，自动生成材料平差表（内置加强比对版模板）",
            style="Sub.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        # 工具栏
        tools = ttk.Frame(root, style="Card.TFrame", padding=10)
        tools.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        for i in range(8):
            tools.columnconfigure(i, weight=0)
        tools.columnconfigure(6, weight=1)

        self.btn_add_file = ttk.Button(tools, text="添加文件", command=self.add_files_dialog)
        self.btn_add_dir = ttk.Button(tools, text="添加文件夹", command=self.add_folder_dialog)
        self.btn_delete = ttk.Button(tools, text="删除选中", command=self.delete_selected)
        self.btn_clear = ttk.Button(tools, text="清空", command=self.clear_all)
        self.btn_start = ttk.Button(tools, text="开始生成", style="Accent.TButton", command=self.start)

        self.btn_add_file.grid(row=0, column=0, padx=(0, 8))
        self.btn_add_dir.grid(row=0, column=1, padx=(0, 8))
        self.btn_delete.grid(row=0, column=2, padx=(0, 8))
        self.btn_clear.grid(row=0, column=3, padx=(0, 8))
        ttk.Label(tools, text="右键列表可增删 · 可拖入 · 可 Ctrl+V 粘贴", style="Card.TLabel").grid(
            row=0, column=4, sticky="w", padx=(8, 0)
        )
        self.btn_start.grid(row=0, column=7, sticky="e")

        # 文件列表
        list_card = ttk.Frame(root, style="Card.TFrame", padding=1)
        list_card.grid(row=2, column=0, sticky="nsew")
        list_card.columnconfigure(0, weight=1)
        list_card.rowconfigure(0, weight=1)

        columns = ("no", "name", "kind", "path")
        self.tree = ttk.Treeview(
            list_card, columns=columns, show="headings", selectmode="extended", height=10
        )
        self.tree.heading("no", text="#", anchor="center")
        self.tree.heading("name", text="名称", anchor="w")
        self.tree.heading("kind", text="类型", anchor="center")
        self.tree.heading("path", text="路径", anchor="w")
        self.tree.column("no", width=48, anchor="center", stretch=False)
        self.tree.column("name", width=250, anchor="w", stretch=True)
        self.tree.column("kind", width=80, anchor="center", stretch=False)
        self.tree.column("path", width=520, anchor="w", stretch=True)
        self.tree.tag_configure("odd", background="#F8FAFC")
        self.tree.tag_configure("even", background="#FFFFFF")
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=10)

        scroll = ttk.Scrollbar(list_card, orient="vertical", command=self.tree.yview)
        scroll.grid(row=0, column=1, sticky="ns", pady=10, padx=(4, 10))
        self.tree.configure(yscrollcommand=scroll.set)

        # 右键菜单
        self.menu = tk.Menu(root, tearoff=0, font=("Microsoft YaHei UI", 10))
        self.menu.add_command(label="添加文件…", command=self.add_files_dialog)
        self.menu.add_command(label="添加文件夹…", command=self.add_folder_dialog)
        self.menu.add_separator()
        self.menu.add_command(label="删除选中", command=self.delete_selected)
        self.menu.add_command(label="清空列表", command=self.clear_all)
        self.tree.bind("<Button-3>", self.on_right_click)

        # 日志
        log_card = ttk.Frame(root, style="Card.TFrame", padding=1)
        log_card.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        log_card.columnconfigure(0, weight=1)
        ttk.Label(log_card, text="运行日志", style="Card.TLabel").grid(
            row=0, column=0, sticky="w", padx=10, pady=(8, 4)
        )
        self.log_text = tk.Text(
            log_card, height=7, wrap="none", relief="flat", bd=0,
            bg="#FFFFFF", fg="#374151", font=("Consolas", 9),
            insertbackground="#374151",
        )
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.log_text.configure(state="disabled")
        log_scroll = ttk.Scrollbar(log_card, orient="vertical", command=self.log_text.yview)
        log_scroll.grid(row=1, column=1, sticky="ns", pady=(0, 10), padx=(0, 10))
        self.log_text.configure(yscrollcommand=log_scroll.set)

        # 底部状态
        bottom = ttk.Frame(root, style="Root.TFrame", padding=(2, 10, 2, 0))
        bottom.grid(row=4, column=0, sticky="ew")
        bottom.columnconfigure(1, weight=1)
        ttk.Label(bottom, text="成品目录", style="Hint.TLabel").grid(row=0, column=0, sticky="w")
        self.out_var = tk.StringVar(value="桌面 / 材料平差输出_" + stamp() + "（点击开始生成时创建）")
        ttk.Label(bottom, textvariable=self.out_var, style="Hint.TLabel").grid(
            row=0, column=1, sticky="w", padx=(8, 0)
        )
        self.progress = ttk.Progressbar(root, mode="determinate", maximum=100, value=0)
        self.progress.grid(row=5, column=0, sticky="ew", pady=(8, 0))
        self.status_var = tk.StringVar(value="就绪：添加图纸后点击「开始生成」")
        ttk.Label(root, textvariable=self.status_var, style="Hint.TLabel").grid(
            row=6, column=0, sticky="w", pady=(5, 0)
        )

        self.refresh_list()

    def _bind_keys(self):
        self.root.bind_all("<Control-v>", lambda e: self.paste_clipboard())
        self.root.bind_all("<Control-V>", lambda e: self.paste_clipboard())
        self.tree.bind("<Delete>", lambda e: self.delete_selected())
        self.tree.bind("<Control-a>", lambda e: self.tree.selection_set(self.tree.get_children()))
        self.tree.bind("<Control-A>", lambda e: self.tree.selection_set(self.tree.get_children()))

    # ------------------------------------------------------------------ #
    # 列表操作
    # ------------------------------------------------------------------ #
    def refresh_list(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for idx, item in enumerate(self.paths, 1):
            tag = "odd" if idx % 2 else "even"
            self.tree.insert(
                "", "end",
                values=(idx, os.path.basename(item["path"]), item["kind"], item["path"]),
                tags=(tag,),
            )
        self.status_var.set("已添加 %d 项" % len(self.paths))

    def add_paths(self, paths, silent=False):
        added = 0
        skipped = 0
        for raw in paths or []:
            path = os.path.abspath(str(raw).strip().strip('"'))
            if not os.path.exists(path):
                skipped += 1
                continue
            if os.path.isdir(path):
                kind = "文件夹"
            elif path.lower().endswith(VALID_EXT):
                kind = "图纸"
            else:
                skipped += 1
                continue
            key = normalize(path)
            if key in self.seen:
                continue
            self.seen.add(key)
            self.paths.append({"path": path, "kind": kind})
            added += 1
        self.refresh_list()
        if added:
            self.status_var.set("已添加 %d 项（共 %d 项）" % (added, len(self.paths)))
        if skipped and not silent:
            messagebox.showwarning(
                "已跳过", "已忽略 %d 个无效路径（仅支持 .dwg / .dxf 文件或文件夹）。" % skipped
            )
        return added

    def add_files_dialog(self):
        files = filedialog.askopenfilenames(
            title="选择 DWG / DXF 图纸（可多选）",
            filetypes=[("DWG / DXF 图纸", "*.dwg *.dxf"), ("所有文件", "*.*")],
        )
        self.add_paths(files)

    def add_folder_dialog(self):
        folder = filedialog.askdirectory(title="选择包含图纸的文件夹（递归扫描）")
        if folder:
            self.add_paths([folder])

    def delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            return
        indexes = sorted(
            (int(self.tree.item(item, "values")[0]) - 1 for item in selected), reverse=True
        )
        for idx in indexes:
            if 0 <= idx < len(self.paths):
                self.seen.discard(normalize(self.paths[idx]["path"]))
                del self.paths[idx]
        self.refresh_list()

    def clear_all(self):
        self.paths.clear()
        self.seen.clear()
        self.refresh_list()

    def on_right_click(self, event):
        item = self.tree.identify_row(event.y)
        if item and item not in self.tree.selection():
            self.tree.selection_set(item)
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()

    def paste_clipboard(self):
        paths = clipboard_paths()
        if not paths:
            return False
        self.add_paths(paths)
        return True

    # ------------------------------------------------------------------ #
    # 生成流程
    # ------------------------------------------------------------------ #
    def append_log(self, text):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        for line in text.splitlines():
            m = re.search(r"(\d+)%", line)
            if m:
                self.progress.configure(value=min(100, int(m.group(1))))

    def set_busy(self, busy: bool):
        self.running = busy
        state = "disabled" if busy else "normal"
        for btn in (self.btn_add_file, self.btn_add_dir, self.btn_delete, self.btn_clear, self.btn_start):
            btn.configure(state=state)

    def start(self):
        if self.running:
            return
        if not self.paths:
            messagebox.showinfo("提示", "请先添加 DWG / DXF 文件或文件夹。")
            return
        if not os.path.isfile(TEMPLATE_XLS):
            messagebox.showerror("错误", "内置模板缺失：%s" % TEMPLATE_XLS)
            return

        out_dir = os.path.join(desktop_dir(), "材料平差输出_" + stamp())
        try:
            os.makedirs(out_dir, exist_ok=True)
        except OSError as exc:
            messagebox.showerror("错误", "无法创建成品目录：%s" % exc)
            return

        self.output_dir = out_dir
        self.out_var.set(out_dir)
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self.progress.configure(value=0)
        self.status_var.set("正在读取图纸并生成材料平差表…")
        self.set_busy(True)
        self.q = queue.Queue()
        threading.Thread(target=self.work, args=(out_dir,), daemon=True).start()
        self.root.after(80, self.poll_queue)

    def work(self, out_dir):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        sys.stdout = sys.stderr = QueueWriter(self.q)
        com_ready = False
        try:
            pythoncom.CoInitialize()
            com_ready = True
            argv = [item["path"] for item in self.paths]
            argv += [
                "--out-dir", out_dir,
                "--template", TEMPLATE_XLS,
                "--report-dir", os.path.join(tempfile.gettempdir(), "dwg2xls-reports"),
            ]
            cache_dir = os.environ.get("DWG2XLS_CACHE_DIR", "")
            if cache_dir and os.path.isdir(cache_dir):
                argv += ["--cache-dir", cache_dir]
            code = auto_export.main(argv)
            self.q.put(("done", code))
        except BaseException:
            self.q.put(("error", traceback.format_exc()))
        finally:
            if com_ready:
                pythoncom.CoUninitialize()
            sys.stdout, sys.stderr = old_stdout, old_stderr

    def poll_queue(self):
        while True:
            try:
                kind, value = self.q.get_nowait()
            except queue.Empty:
                break
            if kind == "log":
                self.append_log(value)
            elif kind == "error":
                self.finish(None, value)
                return
            elif kind == "done":
                self.finish(value)
                return
        if self.running:
            self.root.after(80, self.poll_queue)

    def finish(self, code, tb=None):
        self.set_busy(False)
        if tb:
            self.append_log("\n生成失败：\n" + tb)
            self.status_var.set("生成失败，请查看日志")
            messagebox.showerror("错误", "生成过程中出现异常，请查看运行日志。")
            return
        if code == 0:
            self.progress.configure(value=100)
            self.status_var.set("完成：%s" % self.output_dir)
            self.append_log("\n完成，成品目录：%s\n" % self.output_dir)
            try:
                os.startfile(self.output_dir)
            except OSError:
                pass
            messagebox.showinfo("完成", "材料平差表已生成：\n%s" % self.output_dir)
        else:
            self.status_var.set("未生成任何成品，请查看日志")
            self.append_log("\n未生成任何成品，请检查图纸。\n")
            messagebox.showwarning("提示", "未生成任何成品，请查看运行日志。")


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
