#!/usr/bin/env python3
"""
批量文件智能整理工具 v1.0
- 按文件类型分类整理
- 按修改日期分类整理
- 支持拖拽或选择文件夹
- 简单易用的图形界面
"""

import os
import shutil
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from datetime import datetime
from pathlib import Path
import threading
import sys

# 文件类型分类规则
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

# 字体
DEFAULT_FONT = ("PingFang SC", 11)
TITLE_FONT = ("PingFang SC", 16, "bold")
STATUS_FONT = ("PingFang SC", 10)


class FileOrganizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("批量文件智能整理工具 v1.0")
        self.root.geometry("720x580")
        self.root.minsize(600, 480)
        self.root.configure(bg="#f5f5f7")

        # 变量
        self.source_dir = tk.StringVar()
        self.org_mode = tk.StringVar(value="type")
        self.progress_value = tk.DoubleVar(value=0)
        self.status_text = tk.StringVar(value='就绪 | 选择文件夹后点击「开始整理」')
        self.file_count_var = tk.StringVar(value="待扫描文件数：-")
        self.is_running = False
        self.files_to_process = []

        self._build_ui()

    def _build_ui(self):
        # 标题
        header = tk.Frame(self.root, bg="#f5f5f7")
        header.pack(fill="x", padx=24, pady=(20, 0))

        tk.Label(
            header, text="📂 批量文件智能整理工具",
            font=TITLE_FONT, bg="#f5f5f7", fg="#1d1d1f"
        ).pack(anchor="w")

        tk.Label(
            header, text="自动按类型或日期分类，让文件不再杂乱",
            font=STATUS_FONT, bg="#f5f5f7", fg="#6e6e73"
        ).pack(anchor="w", pady=(2, 0))

        # 主容器
        main = tk.Frame(self.root, bg="#f5f5f7")
        main.pack(fill="both", expand=True, padx=24, pady=(16, 0))

        # 选择文件夹区域
        folder_frame = tk.LabelFrame(
            main, text=" 选择要整理的文件夹 ",
            font=("PingFang SC", 11), bg="#ffffff", fg="#1d1d1f",
            padx=16, pady=12, relief="groove", bd=1
        )
        folder_frame.pack(fill="x", pady=(0, 12))

        path_row = tk.Frame(folder_frame, bg="#ffffff")
        path_row.pack(fill="x")

        self.path_entry = tk.Entry(
            path_row, textvariable=self.source_dir,
            font=("PingFang SC", 10), bg="#f5f5f7",
            relief="flat", bd=0, highlightthickness=1,
            highlightcolor="#0071e3", highlightbackground="#d2d2d7"
        )
        self.path_entry.pack(side="left", fill="x", expand=True, ipady=4)

        tk.Button(
            path_row, text="选择文件夹",
            font=("PingFang SC", 10), bg="#0071e3", fg="white",
            relief="flat", padx=14, pady=4, cursor="hand2",
            activebackground="#0077ed", activeforeground="white",
            command=self.select_folder
        ).pack(side="right", padx=(8, 0))

        # 整理模式选择
        mode_frame = tk.LabelFrame(
            main, text=" 整理模式 ",
            font=("PingFang SC", 11), bg="#ffffff", fg="#1d1d1f",
            padx=16, pady=12, relief="groove", bd=1
        )
        mode_frame.pack(fill="x", pady=(0, 12))

        modes_row = tk.Frame(mode_frame, bg="#ffffff")
        modes_row.pack(fill="x")

        tk.Radiobutton(
            modes_row, text="按文件类型分类",
            variable=self.org_mode, value="type",
            font=("PingFang SC", 11), bg="#ffffff",
            activebackground="#ffffff", fg="#1d1d1f",
            selectcolor="#ffffff"
        ).pack(side="left", padx=(0, 24))

        tk.Radiobutton(
            modes_row, text="按修改日期分类（年/月）",
            variable=self.org_mode, value="date",
            font=("PingFang SC", 11), bg="#ffffff",
            activebackground="#ffffff", fg="#1d1d1f",
            selectcolor="#ffffff"
        ).pack(side="left")

        # 功能说明
        info_frame = tk.LabelFrame(
            main, text=" 功能介绍 ",
            font=("PingFang SC", 11), bg="#ffffff", fg="#1d1d1f",
            padx=16, pady=12, relief="groove", bd=1
        )
        info_frame.pack(fill="x", pady=(0, 12))

        info_text = (
            "• 按文件类型：自动创建「文档」「图片」「视频」等文件夹分类存放\n"
            "• 按修改日期：自动按「年/月」创建目录层级\n"
            "• 仅整理文件，不删除任何内容，安全可靠\n"
            "• 支持几乎所有常见文件格式"
        )
        tk.Label(
            info_frame, text=info_text, justify="left",
            font=("PingFang SC", 10), bg="#ffffff", fg="#515154",
            anchor="w"
        ).pack(fill="x")

        # 进度条和状态
        progress_frame = tk.Frame(main, bg="#f5f5f7")
        progress_frame.pack(fill="x", pady=(0, 8))

        self.progress_bar = ttk.Progressbar(
            progress_frame, variable=self.progress_value,
            length=600, mode="determinate"
        )
        self.progress_bar.pack(fill="x")

        status_row = tk.Frame(main, bg="#f5f5f7")
        status_row.pack(fill="x")

        self.file_count_label = tk.Label(
            status_row, textvariable=self.file_count_var,
            font=STATUS_FONT, bg="#f5f5f7", fg="#6e6e73"
        )
        self.file_count_label.pack(side="left")

        self.status_label = tk.Label(
            status_row, textvariable=self.status_text,
            font=STATUS_FONT, bg="#f5f5f7", fg="#6e6e73"
        )
        self.status_label.pack(side="right")

        # 底部信息
        footer_frame = tk.Frame(main, bg="#f5f5f7")
        footer_frame.pack(fill="x", pady=(4, 4))

        support_btn = tk.Button(
            footer_frame, text="❤️ 支持作者",
            font=("PingFang SC", 9), bg="#f5f5f7", fg="#86868b",
            relief="flat", cursor="hand2", bd=0,
            activebackground="#f5f5f7", activeforeground="#1d1d1f",
            command=self.show_support
        )
        support_btn.pack()

        # 开始按钮
        btn_frame = tk.Frame(main, bg="#f5f5f7")
        btn_frame.pack(fill="x", pady=(4, 0))

        self.start_btn = tk.Button(
            btn_frame, text="🚀 开始整理",
            font=("PingFang SC", 13, "bold"), bg="#34c759", fg="white",
            relief="flat", padx=32, pady=10, cursor="hand2",
            activebackground="#30d158", activeforeground="white",
            command=self.start_organize
        )
        self.start_btn.pack()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

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
