#!/usr/bin/env python3
"""
批量文件智能整理工具 v2.0 - 时尚版
- 按文件类型分类整理
- 按修改日期分类整理
- 支持拖拽或选择文件夹
- 全新现代化 UI 设计
"""

import os
import shutil
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from datetime import datetime
from pathlib import Path
import threading
import sys

# ── 文件类型分类规则 ──
FILE_CATEGORIES = {
    "文档": [
        ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
        ".pdf", ".txt", ".md", ".csv", ".json", ".xml", ".html", ".htm",
        ".wps", ".et", ".dps"
    ],
    "图片": [
        ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".tif",
        ".webp", ".svg", ".ico", ".raw", ".psd", ".ai"
    ],
    "视频": [
        ".mp4", ".avi", ".mkv", ".mov", ".wmv", ".flv", ".webm",
        ".m4v", ".mpg", ".mpeg", ".rmvb", ".ts"
    ],
    "音频": [
        ".mp3", ".wav", ".flac", ".aac", ".ogg", ".wma", ".m4a",
        ".ape", ".midi"
    ],
    "压缩包": [
        ".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz",
        ".iso", ".dmg"
    ],
    "代码": [
        ".py", ".js", ".ts", ".jsx", ".tsx", ".java", ".cpp", ".c",
        ".h", ".hpp", ".css", ".scss", ".less", ".php", ".rb", ".go",
        ".rs", ".swift", ".kt", ".sql", ".sh", ".bat", ".yaml", ".yml",
        ".toml", ".ini", ".cfg"
    ],
    "安装包": [
        ".exe", ".msi", ".dmg", ".pkg", ".app", ".deb", ".rpm",
        ".apk", ".ipa"
    ],
    "表格数据": [
        ".xls", ".xlsx", ".csv", ".tsv", ".numbers"
    ],
}

# ── 时尚配色方案 ──
C = {
    "bg": "#f0f2f5",
    "card": "#ffffff",
    "primary": "#4F46E5",
    "primary_light": "#EEF2FF",
    "primary_text": "#4338CA",
    "success": "#059669",
    "success_light": "#ECFDF5",
    "success_text": "#047857",
    "text_main": "#111827",
    "text_sec": "#6B7280",
    "border": "#E5E7EB",
    "header_bg": "#ffffff",
}

# 字体
FONT = ("PingFang SC", 11)
FONT_BOLD = ("PingFang SC", 11, "bold")
TITLE_FONT = ("PingFang SC", 18, "bold")
SUBTITLE_FONT = ("PingFang SC", 10)
STATUS_FONT = ("PingFang SC", 10)


def _make_rounded_btn(parent, text, bg, fg, cmd, **kw):
    """创建一个现代风格的按钮（用 Frame 模拟圆角效果）"""
    btn_font = kw.pop("font", FONT_BOLD)
    padx = kw.pop("padx", 18)
    pady = kw.pop("pady", 8)
    # 用 LabelFrame 作为容器模拟圆角
    container = tk.Frame(parent, bg=bg, highlightbackground=bg, highlightthickness=0)
    btn = tk.Button(
        container, text=text, font=btn_font,
        bg=bg, fg=fg,
        activebackground=bg, activeforeground=fg,
        relief="flat", bd=0, cursor="hand2",
        padx=padx, pady=pady,
        highlightthickness=0,
        command=cmd
    )
    btn.pack()
    return container, btn


def _make_outline_btn(parent, text, color, cmd, **kw):
    """创建一个描边风格的按钮"""
    btn_font = kw.pop("font", FONT)
    padx = kw.pop("padx", 14)
    pady = kw.pop("pady", 6)
    container = tk.Frame(parent, bg=C["card"], highlightbackground=color, highlightthickness=2, highlightcolor=color)
    btn = tk.Button(
        container, text=text, font=btn_font,
        bg=C["card"], fg=color,
        activebackground=C["primary_light"], activeforeground=color,
        relief="flat", bd=0, cursor="hand2",
        padx=padx, pady=pady,
        highlightthickness=0,
        command=cmd
    )
    btn.pack()
    return container, btn


class FileOrganizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("批量文件智能整理工具 v2.0")
        self.root.geometry("720x580")
        self.root.minsize(620, 500)
        self.root.configure(bg=C["bg"])

        # 变量
        self.source_dir = tk.StringVar()
        self.org_mode = tk.StringVar(value="type")
        self.progress_value = tk.DoubleVar(value=0)
        self.status_text = tk.StringVar(value='就绪 · 选择文件夹后点击「开始整理」')
        self.file_count_var = tk.StringVar(value="待扫描文件数：-")
        self.is_running = False
        self.files_to_process = []
        self.path_entry = None
        self.start_btn = None

        self._build_ui()

    # ── 构建 UI ──
    def _build_ui(self):
        # ====== 顶栏 ======
        top_bg = C["header_bg"]
        header = tk.Frame(self.root, bg=top_bg, highlightthickness=0)
        header.pack(fill="x", pady=(0, 0))

        # 加一条底部细阴影
        tk.Frame(self.root, bg=C["border"], height=1).pack(fill="x")

        header_inner = tk.Frame(header, bg=top_bg, padx=28, pady=(20, 16))
        header_inner.pack(fill="x")

        tk.Label(
            header_inner, text="🧹 文件智能整理工具",
            font=TITLE_FONT, bg=top_bg, fg=C["text_main"]
        ).pack(anchor="w")

        tk.Label(
            header_inner, text="自动按类型或日期分类 · 跨平台 · 纯本地运行",
            font=SUBTITLE_FONT, bg=top_bg, fg=C["text_sec"]
        ).pack(anchor="w", pady=(3, 0))

        # ====== 主区域 ======
        main = tk.Frame(self.root, bg=C["bg"])
        main.pack(fill="both", expand=True, padx=28, pady=(20, 16))

        # ── 卡片：文件夹选择 ──
        card1 = self._make_card(main)
        card1.pack(fill="x")

        tk.Label(
            card1, text="📁 选择文件夹",
            font=FONT_BOLD, bg=C["card"], fg=C["text_main"]
        ).pack(anchor="w")

        path_row = tk.Frame(card1, bg=C["card"])
        path_row.pack(fill="x", pady=(10, 0))

        self.path_entry = tk.Entry(
            path_row, textvariable=self.source_dir,
            font=("PingFang SC", 10), bg=C["bg"],
            relief="flat", bd=0, highlightthickness=1.5,
            highlightcolor=C["primary"], highlightbackground=C["border"],
            insertbackground=C["text_main"]
        )
        self.path_entry.pack(side="left", fill="x", expand=True, ipady=6, ipadx=8)

        _, self.folder_btn = _make_outline_btn(
            path_row, "选择文件夹", C["primary_text"],
            cmd=self.select_folder, font=FONT_BOLD, padx=16, pady=6
        )
        self.folder_btn.master.pack(side="right", padx=(10, 0))

        # ── 卡片：整理模式 ──
        card2 = self._make_card(main)
        card2.pack(fill="x", pady=(12, 0))

        tk.Label(
            card2, text="⚙️ 整理模式",
            font=FONT_BOLD, bg=C["card"], fg=C["text_main"]
        ).pack(anchor="w")

        modes_row = tk.Frame(card2, bg=C["card"])
        modes_row.pack(fill="x", pady=(10, 0))

        # 自定义单选按钮（用 Frame 模拟好看的选项）
        self._radio_option(modes_row, "按文件类型分类", "type", 0)
        self._radio_option(modes_row, "按修改日期分类（年/月）", "date", 1)

        # ── 卡片：功能介绍 ──
        card3 = self._make_card(main)
        card3.pack(fill="x", pady=(12, 0))

        info_items = [
            ("📂", "按文件类型", "自动创建「文档」「图片」「视频」等分类文件夹"),
            ("📅", "按修改日期", "自动按「年/月」创建目录层级结构"),
            ("🔒", "安全可靠", "仅移动文件，不删除任何内容"),
            ("🎯", "广泛兼容", "支持几乎所有常见文件格式"),
        ]

        tk.Label(
            card3, text="💡 功能亮点",
            font=FONT_BOLD, bg=C["card"], fg=C["text_main"]
        ).pack(anchor="w")

        grid = tk.Frame(card3, bg=C["card"])
        grid.pack(fill="x", pady=(8, 0))

        for i, (icon, title, desc) in enumerate(info_items):
            col = i % 2
            row = i // 2
            item = tk.Frame(grid, bg=C["bg"], padx=12, pady=8, highlightthickness=0)
            item.grid(row=row, column=col, sticky="ew", padx=4, pady=4)
            grid.columnconfigure(col, weight=1)

            tk.Label(item, text=f"{icon} {title}", font=FONT_BOLD,
                     bg=C["bg"], fg=C["text_main"], anchor="w").pack(fill="x")
            tk.Label(item, text=desc, font=("PingFang SC", 9),
                     bg=C["bg"], fg=C["text_sec"], anchor="w").pack(fill="x")

        # ── 进度条区域 ──
        progress_card = tk.Frame(main, bg=C["card"], padx=20, pady=14)
        progress_card.pack(fill="x", pady=(12, 0))

        # 自定义进度条样式
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "custom.Horizontal.TProgressbar",
            background=C["primary"],
            troughcolor=C["border"],
            bordercolor=C["card"],
            lightcolor=C["primary_light"],
            darkcolor=C["primary"],
            thickness=8
        )

        self.progress_bar = ttk.Progressbar(
            progress_card, variable=self.progress_value,
            length=600, mode="determinate", style="custom.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(fill="x")

        status_row = tk.Frame(progress_card, bg=C["card"])
        status_row.pack(fill="x", pady=(6, 0))

        self.file_count_label = tk.Label(
            status_row, textvariable=self.file_count_var,
            font=STATUS_FONT, bg=C["card"], fg=C["text_sec"]
        )
        self.file_count_label.pack(side="left")

        self.status_label = tk.Label(
            status_row, textvariable=self.status_text,
            font=STATUS_FONT, bg=C["card"], fg=C["text_sec"]
        )
        self.status_label.pack(side="right")

        # ── 底部操作栏 ──
        bottom = tk.Frame(main, bg=C["bg"])
        bottom.pack(fill="x", pady=(16, 0))

        # 支持作者
        tk.Button(
            bottom, text="❤️ 支持作者",
            font=("PingFang SC", 9), bg=C["bg"], fg=C["text_sec"],
            relief="flat", cursor="hand2", bd=0,
            activebackground=C["bg"], activeforeground=C["text_main"],
            command=self.show_support
        ).pack(side="left")

        # 开始整理按钮
        start_container = tk.Frame(bottom, bg=C["bg"])
        start_container.pack(side="right")

        self.start_btn = tk.Button(
            start_container, text="🚀 开始整理",
            font=("PingFang SC", 14, "bold"),
            bg=C["primary"], fg="#ffffff",
            relief="flat", bd=0, cursor="hand2",
            padx=36, pady=12,
            activebackground="#6366F1", activeforeground="#ffffff",
            highlightthickness=0,
            command=self.start_organize
        )
        self.start_btn.pack()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _make_card(self, parent):
        """创建一个白色卡片容器"""
        card = tk.Frame(parent, bg=C["card"], padx=20, pady=16)
        return card

    def _radio_option(self, parent, text, value, col):
        """创建一个风格化的单选选项"""
        frame = tk.Frame(parent, bg=C["card"])
        frame.grid(row=0, column=col, sticky="w", padx=(0, 32))

        rb = tk.Radiobutton(
            frame, text=text, variable=self.org_mode, value=value,
            font=("PingFang SC", 11), bg=C["card"], fg=C["text_main"],
            activebackground=C["card"], activeforeground=C["primary"],
            selectcolor=C["card"],
            indicatoron=1
        )
        rb.pack(anchor="w")

    def select_folder(self):
        folder = filedialog.askdirectory(title="选择需要整理的文件夹")
        if folder:
            self.source_dir.set(folder)
            self._scan_folder(folder)

    def _scan_folder(self, folder):
        """扫描文件夹中的文件（仅顶层，不递归）"""
        self.files_to_process = []
        try:
            for f in os.listdir(folder):
                full_path = os.path.join(folder, f)
                if os.path.isfile(full_path) and not f.startswith('.'):
                    self.files_to_process.append(full_path)
            self.file_count_var.set(f"待整理文件数：{len(self.files_to_process)} 个")
        except Exception as e:
            self.file_count_var.set(f"扫描出错：{str(e)}")

    def start_organize(self):
        if self.is_running:
            return

        folder = self.source_dir.get()
        if not folder:
            messagebox.showwarning("提示", "请先选择一个要整理的文件夹")
            return
        if not os.path.isdir(folder):
            messagebox.showerror("错误", "选择的文件夹不存在")
            return

        # 二次扫描确认
        self._scan_folder(folder)
        if not self.files_to_process:
            messagebox.showinfo("提示", "该文件夹中没有需要整理的文件")
            return

        # 确认对话框
        mode_text = "按文件类型" if self.org_mode.get() == "type" else "按修改日期"
        if not messagebox.askyesno(
            "确认操作",
            f"将在「{folder}」内{mode_text}整理 {len(self.files_to_process)} 个文件。\n\n"
            "⚠️ 仅移动文件到分类子文件夹，不会删除任何内容。\n\n确定继续？"
        ):
            return

        self.is_running = True
        self.start_btn.config(state="disabled", text="⏳ 整理中...")
        self.progress_value.set(0)

        thread = threading.Thread(
            target=self._organize_thread, args=(folder,), daemon=True
        )
        thread.start()

    def _organize_thread(self, base_dir):
        """后台执行整理操作"""
        files = self.files_to_process[:]
        total = len(files)
        success = 0
        skipped = 0
        errors = []

        for i, file_path in enumerate(files):
            if not self.is_running:
                break
            try:
                result = self._move_file(file_path, base_dir)
                if result == "moved":
                    success += 1
                else:
                    skipped += 1
            except Exception as e:
                errors.append(f"{os.path.basename(file_path)}: {str(e)}")

            # 更新进度
            progress = ((i + 1) / total) * 100
            self.progress_value.set(progress)
            self.status_text.set(
                f"正在整理：{os.path.basename(file_path)[:30]}... "
                f"({i+1}/{total})"
            )

        # 完成
        self.is_running = False
        self.root.after(0, lambda: self._on_complete(success, skipped, errors))

    def _move_file(self, file_path, base_dir):
        """将文件移动到对应的分类文件夹"""
        ext = Path(file_path).suffix.lower()
        filename = os.path.basename(file_path)

        if self.org_mode.get() == "type":
            # 按类型分类
            target_dir = self._get_type_folder(base_dir, ext)
        else:
            # 按日期分类
            target_dir = self._get_date_folder(base_dir, file_path)

        if not target_dir:
            return "skipped"

        os.makedirs(target_dir, exist_ok=True)
        dest = os.path.join(target_dir, filename)

        # 处理重名
        counter = 1
        while os.path.exists(dest):
            name, ext_ = os.path.splitext(filename)
            dest = os.path.join(target_dir, f"{name}_{counter}{ext_}")
            counter += 1

        shutil.move(file_path, dest)
        return "moved"

    def _get_type_folder(self, base_dir, ext):
        """根据文件扩展名返回对应的类型文件夹名"""
        for category, extensions in FILE_CATEGORIES.items():
            if ext in extensions:
                return os.path.join(base_dir, category)
        # 未识别的归为"其他"
        return os.path.join(base_dir, "其他")

    def _get_date_folder(self, base_dir, file_path):
        """根据文件修改日期返回年/月文件夹路径"""
        try:
            mtime = os.path.getmtime(file_path)
            dt = datetime.fromtimestamp(mtime)
            year = str(dt.year)
            month = f"{dt.month:02d}月"
            return os.path.join(base_dir, year, month)
        except Exception:
            return os.path.join(base_dir, "日期未知")

    def _on_complete(self, success, skipped, errors):
        """整理完成后更新界面"""
        msg = f"✅ 整理完成！成功移动 {success} 个文件"
        if skipped:
            msg += f"，{skipped} 个已跳过"
        if errors:
            msg += f"，{len(errors)} 个出错"

        self.status_text.set(msg)
        self.start_btn.config(state="normal", text="🚀 开始整理")

        detail = f"成功：{success} 个\n跳过：{skipped} 个\n"
        if errors:
            detail += f"\n错误详情：\n" + "\n".join(errors[:5])
        detail += "\n\n❤️ 如果好用，欢迎支持作者继续更新～\n（运行工具后可见支持方式）"
        messagebox.showinfo("整理完成", detail)

        # 重新扫描
        self._scan_folder(self.source_dir.get())

    def show_support(self):
        """显示支持作者的信息"""
        messagebox.showinfo(
            "❤️ 支持作者",
            "如果这个工具帮到了你，欢迎请作者喝杯咖啡 ☕\n\n"
            "支付宝：18010415816\n\n"
            "感谢你的支持！这能帮助我持续更新和维护。"
        )

    def on_close(self):
        if self.is_running:
            self.is_running = False
        self.root.destroy()


def main():
    # macOS 上设置正确的进程名称
    if sys.platform == "darwin":
        os.environ.pop('__PYAPP__', None)

    root = tk.Tk()
    # 设置 macOS 风格
    if sys.platform == "darwin":
        try:
            root.tk.call("::tk::unsupported::MacWindowStyle", "style",
                         root._w, "document", "closeBox")
        except Exception:
            pass

    app = FileOrganizerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
