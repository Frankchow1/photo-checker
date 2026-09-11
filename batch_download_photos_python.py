#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
照片批量下载工具 (并发·断点续传·GUI)
功能：支持从 CSV / TSV / Excel 表格中自动提取图片 URL，支持多线程并发下载、断点续传、失败重试、冲突重命名与实时进度监控。
"""

from __future__ import annotations

import os
import sys
import time
import csv
import re
import hashlib
import threading
import subprocess
import argparse
from pathlib import Path
from typing import List, Tuple, Dict, Optional, Any, Set
from concurrent.futures import ThreadPoolExecutor, wait, FIRST_COMPLETED
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError
from urllib.parse import unquote, urlsplit

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

USER_AGENT = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) PhotoBatchDownloader/2.0"


# ==============================================================================
# 核心下载引擎 (PhotoDownloadEngine)
# ==============================================================================

class PhotoDownloadEngine:
    @staticmethod
    def filename_from_url(url: str) -> str:
        try:
            name = Path(unquote(urlsplit(url).path)).name
            return name if name else ""
        except Exception:
            return ""

    @staticmethod
    def detect_text_file_params(file_path: str) -> Tuple[str, str]:
        encodings = ["utf-8-sig", "utf-8", "gb18030", "gbk"]
        best_enc = "utf-8"
        sample_bytes = b""
        with open(file_path, "rb") as f:
            sample_bytes = f.read(16384)

        for enc in encodings:
            try:
                sample_bytes.decode(enc)
                best_enc = enc
                break
            except UnicodeDecodeError:
                continue

        text_sample = sample_bytes.decode(best_enc, errors="ignore")
        first_line = text_sample.splitlines()[0] if text_sample.splitlines() else ""

        tab_count = first_line.count("	")
        comma_count = first_line.count(",")
        semicolon_count = first_line.count(";")

        if tab_count > 0 and tab_count >= comma_count and tab_count >= semicolon_count:
            delimiter = "	"
        elif comma_count > 0 and comma_count >= semicolon_count:
            delimiter = ","
        elif semicolon_count > 0:
            delimiter = ";"
        else:
            delimiter = ","

        return best_enc, delimiter

    @classmethod
    def inspect_file(cls, file_path: str) -> Dict[str, Any]:
        """快速探测文件列名与元数据"""
        if not os.path.isfile(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")

        file_ext = os.path.splitext(file_path)[1].lower()
        is_excel = file_ext in [".xlsx", ".xls", ".xlsm"]
        sheets = []
        columns = []
        recommended_url_col = ""

        if is_excel:
            excel_file = pd.ExcelFile(file_path)
            sheets = excel_file.sheet_names
            first_sheet = sheets[0] if sheets else 0
            df_preview = pd.read_excel(excel_file, sheet_name=first_sheet, nrows=5, dtype=str)
            columns = [str(c) for c in df_preview.columns]
        else:
            encoding, delimiter = cls.detect_text_file_params(file_path)
            try:
                df_preview = pd.read_csv(file_path, sep=delimiter, encoding=encoding, nrows=5, dtype=str, on_bad_lines="warn")
            except Exception:
                df_preview = pd.read_csv(file_path, sep=delimiter, encoding=encoding, nrows=5, dtype=str, quoting=csv.QUOTE_NONE, on_bad_lines="warn")
            columns = [str(c) for c in df_preview.columns]

        # 智能寻找疑似 URL 列
        for col in columns:
            col_lower = col.lower()
            if any(k in col_lower for k in ["url", "link", "photo", "pic", "img", "doc", "照片", "图片", "链接"]):
                recommended_url_col = col
                break

        return {
            "file_path": file_path,
            "is_excel": is_excel,
            "sheets": sheets,
            "columns": columns,
            "recommended_url_col": recommended_url_col
        }

    @classmethod
    def extract_urls(
        cls,
        file_path: str,
        sheet_name: Optional[str] = None,
        target_col: Optional[str] = None,
        progress_callback=None
    ) -> List[str]:
        """从文件提取所有唯一的图片 URL 链接"""
        file_ext = os.path.splitext(file_path)[1].lower()
        is_excel = file_ext in [".xlsx", ".xls", ".xlsm"]
        urls = []
        seen = set()

        if progress_callback:
            progress_callback(f"正在读取并解析文件: {os.path.basename(file_path)} ...")

        if is_excel:
            df = pd.read_excel(file_path, sheet_name=sheet_name or 0, dtype=str)
            if target_col and target_col != "(全表自动扫描)" and target_col in df.columns:
                series = df[target_col].dropna().astype(str)
                for val in series:
                    val = val.strip()
                    if val.startswith(("http://", "https://")) and val not in seen:
                        if cls.filename_from_url(val):
                            seen.add(val)
                            urls.append(val)
            else:
                for col in df.columns:
                    for val in df[col].dropna().astype(str):
                        val = val.strip()
                        if val.startswith(("http://", "https://")) and val not in seen:
                            if cls.filename_from_url(val):
                                seen.add(val)
                                urls.append(val)
        else:
            encoding, delimiter = cls.detect_text_file_params(file_path)
            with open(file_path, "r", encoding=encoding, errors="replace", newline="") as handle:
                reader = csv.reader(handle, delimiter=delimiter)
                header = None
                target_col_idx = None

                for row_idx, row in enumerate(reader):
                    if row_idx == 0:
                        header = [c.strip() for c in row]
                        if target_col and target_col != "(全表自动扫描)" and target_col in header:
                            target_col_idx = header.index(target_col)
                        continue

                    if target_col_idx is not None and target_col_idx < len(row):
                        val = row[target_col_idx].strip()
                        if val.startswith(("http://", "https://")) and val not in seen:
                            if cls.filename_from_url(val):
                                seen.add(val)
                                urls.append(val)
                    else:
                        for val in row:
                            val = val.strip()
                            if val.startswith(("http://", "https://")) and val not in seen:
                                if cls.filename_from_url(val):
                                    seen.add(val)
                                    urls.append(val)

                    if progress_callback and row_idx % 20000 == 0:
                        progress_callback(f"已扫描 {row_idx:,} 行，累计发现 {len(urls):,} 个唯一照片链接...")

        if progress_callback:
            progress_callback(f"链接提取完成，共收集到 {len(urls):,} 个有效照片链接。")
        return urls

    @classmethod
    def build_tasks(cls, urls: List[str], output_dir: Path) -> Tuple[List[Tuple[str, str, Path]], int]:
        tasks = []
        owner = {}
        collisions = 0
        for url in urls:
            filename = cls.filename_from_url(url)
            if filename in owner and owner[filename] != url:
                collisions += 1
                source = Path(filename)
                token = hashlib.sha1(url.encode("utf-8")).hexdigest()[:10]
                filename = f"{source.stem}__{token}{source.suffix}"
            owner[filename] = url
            tasks.append((url, filename, output_dir / filename))
        return tasks, collisions

    @staticmethod
    def expected_total(response, resume_offset: int) -> Optional[int]:
        content_range = response.headers.get("Content-Range", "")
        match = re.search(r"/(\d+)$", content_range)
        if match:
            return int(match.group(1))
        content_length = response.headers.get("Content-Length")
        if content_length and content_length.isdigit():
            length = int(content_length)
            return resume_offset + length if response.getcode() == 206 else length
        return None

    @classmethod
    def download_one(
        cls,
        task: Tuple[str, str, Path],
        retries: int,
        timeout: int,
        stop_event: Optional[threading.Event] = None
    ) -> Dict[str, Any]:
        url, filename, target = task
        part = target.with_name(target.name + ".part")

        if stop_event and stop_event.is_set():
            return {"status": "stopped", "url": url, "filename": filename, "bytes": 0, "error": "User cancelled"}

        if target.is_file() and target.stat().st_size > 0:
            return {"status": "existing", "url": url, "filename": filename, "bytes": target.stat().st_size, "error": ""}

        last_error = ""
        for attempt in range(1, retries + 1):
            if stop_event and stop_event.is_set():
                return {"status": "stopped", "url": url, "filename": filename, "bytes": 0, "error": "User cancelled"}

            try:
                offset = part.stat().st_size if part.exists() else 0
                headers = {"User-Agent": USER_AGENT, "Accept": "*/*"}
                if offset > 0:
                    headers["Range"] = f"bytes={offset}-"

                request = Request(url, headers=headers)
                with urlopen(request, timeout=timeout) as response:
                    status = response.getcode()
                    if status not in (200, 206):
                        raise RuntimeError(f"HTTP {status}")

                    if offset > 0 and status == 206:
                        mode = "ab"
                        write_offset = offset
                    else:
                        mode = "wb"
                        write_offset = 0

                    total = cls.expected_total(response, write_offset)
                    with part.open(mode) as handle:
                        while True:
                            if stop_event and stop_event.is_set():
                                return {"status": "stopped", "url": url, "filename": filename, "bytes": 0, "error": "User cancelled"}
                            chunk = response.read(256 * 1024)
                            if not chunk:
                                break
                            handle.write(chunk)

                final_size = part.stat().st_size
                if total is not None and final_size != total:
                    raise RuntimeError(f"文件大小不完整：{final_size}/{total}")
                if final_size <= 0:
                    raise RuntimeError("下载结果为空文件")

                os.replace(part, target)
                return {"status": "downloaded", "url": url, "filename": filename, "bytes": final_size, "error": ""}

            except HTTPError as error:
                last_error = f"HTTPError {error.code}: {error.reason}"
                if error.code in (401, 403, 404):
                    break
            except (URLError, TimeoutError, OSError, RuntimeError) as error:
                last_error = repr(error)
            except Exception as error:
                last_error = repr(error)

            if attempt < retries:
                if stop_event and stop_event.is_set():
                    break
                time.sleep(min(2 ** (attempt - 1), 10))

        current_size = part.stat().st_size if part.exists() else 0
        return {"status": "failed", "url": url, "filename": filename, "bytes": current_size, "error": last_error}


# ==============================================================================
# GUI 界面类 (PhotoDownloaderGUI)
# ==============================================================================

class PhotoDownloaderGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("📸 照片批量下载工具 (并发·续传·重试)")
        self.root.geometry("960x780")
        self.root.minsize(860, 640)

        # 状态控制
        self.stop_event = threading.Event()
        self.download_thread: Optional[threading.Thread] = None
        self.is_running = False

        # 文件与路径变量
        self.import_file_path = tk.StringVar()
        self.output_dir_path = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.url_col_var = tk.StringVar(value="(全表自动扫描)")

        # 参数变量
        self.workers_var = tk.IntVar(value=12)
        self.retries_var = tk.IntVar(value=8)
        self.timeout_var = tk.IntVar(value=60)
        self.skip_existing_var = tk.BooleanVar(value=True)

        # 统计看板变量
        self.stat_total_var = tk.StringVar(value="0")
        self.stat_existing_var = tk.StringVar(value="0")
        self.stat_downloaded_var = tk.StringVar(value="0")
        self.stat_failed_var = tk.StringVar(value="0")
        self.stat_progress_text_var = tk.StringVar(value="等待开始...")

        self.file_info: Optional[Dict[str, Any]] = None

        self._init_style()
        self._build_ui()

    def _init_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        base_font = ("PingFang SC", 10) if sys.platform == "darwin" else ("Microsoft YaHei", 9)
        bold_font = (base_font[0], base_font[1], "bold")
        title_font = (base_font[0], base_font[1] + 4, "bold")

        style.configure(".", font=base_font)
        style.configure("Title.TLabel", font=title_font, foreground="#1d2129")
        style.configure("Subtitle.TLabel", font=base_font, foreground="#86909c")
        style.configure("Bold.TLabel", font=bold_font)
        style.configure("Header.TLabelframe.Label", font=bold_font, foreground="#165dff")
        style.configure("Primary.TButton", font=bold_font, foreground="#ffffff", background="#165dff")
        style.configure("Danger.TButton", font=bold_font, foreground="#ffffff", background="#f53f3f")

        # 看板数字样式
        style.configure("StatNum.TLabel", font=(base_font[0], 16, "bold"), foreground="#165dff")
        style.configure("StatDesc.TLabel", font=(base_font[0], 9), foreground="#86909c")

    def _build_ui(self):
        main_frame = ttk.Frame(self.root, padding="14")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. 顶部标题
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))

        title_lbl = ttk.Label(top_frame, text="📸 照片批量并发下载工具", style="Title.TLabel")
        title_lbl.pack(anchor=tk.W)
        sub_lbl = ttk.Label(
            top_frame,
            text="选择包含图片链接的表格文件与目标下载文件夹，支持并发极速下载、HTTP Range 断点续传、自动去重与失败重试。",
            style="Subtitle.TLabel"
        )
        sub_lbl.pack(anchor=tk.W, pady=(2, 0))

        # 2. 文件与目录设置区域
        path_group = ttk.LabelFrame(main_frame, text=" 1. 导入与导出设置", style="Header.TLabelframe", padding="10")
        path_group.pack(fill=tk.X, pady=(0, 8))

        # 导入文件选择行
        row_in = ttk.Frame(path_group)
        row_in.pack(fill=tk.X, pady=2)
        ttk.Label(row_in, text="导入表格:", width=9).pack(side=tk.LEFT)
        self.entry_import = ttk.Entry(row_in, textvariable=self.import_file_path)
        self.entry_import.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        ttk.Button(row_in, text="选择文件...", command=self._browse_import_file, width=10).pack(side=tk.LEFT)

        # 字段与 Sheet 选择行
        row_col = ttk.Frame(path_group)
        row_col.pack(fill=tk.X, pady=4)
        ttk.Label(row_col, text="Sheet:").pack(side=tk.LEFT)
        self.combo_sheet = ttk.Combobox(row_col, textvariable=self.sheet_var, state="readonly", width=14)
        self.combo_sheet.pack(side=tk.LEFT, padx=(4, 16))
        self.combo_sheet.bind("<<ComboboxSelected>>", self._on_sheet_changed)

        ttk.Label(row_col, text="URL 链接列:", style="Bold.TLabel").pack(side=tk.LEFT)
        self.combo_url_col = ttk.Combobox(row_col, textvariable=self.url_col_var, state="readonly", width=22)
        self.combo_url_col.pack(side=tk.LEFT, padx=(4, 16))

        self.lbl_file_hint = ttk.Label(row_col, text="未选择文件", foreground="#86909c")
        self.lbl_file_hint.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 导出目录选择行
        row_out = ttk.Frame(path_group)
        row_out.pack(fill=tk.X, pady=4)
        ttk.Label(row_out, text="导出目录:", width=9).pack(side=tk.LEFT)
        self.entry_output = ttk.Entry(row_out, textvariable=self.output_dir_path)
        self.entry_output.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        ttk.Button(row_out, text="选择文件夹...", command=self._browse_output_dir, width=10).pack(side=tk.LEFT)

        # 3. 参数配置与控制栏
        opt_ctrl_frame = ttk.Frame(main_frame)
        opt_ctrl_frame.pack(fill=tk.X, pady=(0, 8))

        # 参数面板
        opt_group = ttk.LabelFrame(opt_ctrl_frame, text=" 2. 下载参数配置", style="Header.TLabelframe", padding="10")
        opt_group.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))

        row_param = ttk.Frame(opt_group)
        row_param.pack(fill=tk.X, pady=2)

        ttk.Label(row_param, text="并发线程:").pack(side=tk.LEFT)
        spin_workers = ttk.Spinbox(row_param, from_=1, to=32, textvariable=self.workers_var, width=5)
        spin_workers.pack(side=tk.LEFT, padx=(4, 14))

        ttk.Label(row_param, text="重试次数:").pack(side=tk.LEFT)
        spin_retries = ttk.Spinbox(row_param, from_=1, to=20, textvariable=self.retries_var, width=5)
        spin_retries.pack(side=tk.LEFT, padx=(4, 14))

        ttk.Label(row_param, text="超时(秒):").pack(side=tk.LEFT)
        spin_timeout = ttk.Spinbox(row_param, from_=5, to=300, textvariable=self.timeout_var, width=5)
        spin_timeout.pack(side=tk.LEFT, padx=(4, 14))

        ttk.Checkbutton(row_param, text="跳过已存在完整照片", variable=self.skip_existing_var).pack(side=tk.LEFT, padx=6)

        # 操作控制按钮
        ctrl_group = ttk.LabelFrame(opt_ctrl_frame, text=" 3. 执行控制", style="Header.TLabelframe", padding="10")
        ctrl_group.pack(side=tk.RIGHT, fill=tk.BOTH, expand=False, padx=(6, 0))

        btn_box = ttk.Frame(ctrl_group)
        btn_box.pack(fill=tk.BOTH, expand=True)

        self.btn_start = ttk.Button(btn_box, text="🚀 开始批量下载", style="Primary.TButton", command=self._start_download_thread, width=14)
        self.btn_start.pack(side=tk.LEFT, padx=(0, 6), ipady=3)

        self.btn_stop = ttk.Button(btn_box, text="⏹ 停止下载", style="Danger.TButton", command=self._stop_download, width=10, state="disabled")
        self.btn_stop.pack(side=tk.LEFT, ipady=3)

        # 4. 实时看板与进度条
        board_group = ttk.LabelFrame(main_frame, text=" 4. 实时统计看板", style="Header.TLabelframe", padding="8")
        board_group.pack(fill=tk.X, pady=(0, 8))

        # 4 个小指标卡
        kpi_frame = ttk.Frame(board_group)
        kpi_frame.pack(fill=tk.X, pady=(0, 6))

        for idx, (label, var, color) in enumerate([
            ("唯一链接总数", self.stat_total_var, "#1d2129"),
            ("已存在 (跳过)", self.stat_existing_var, "#86909c"),
            ("本次新下载成功", self.stat_downloaded_var, "#00b42a"),
            ("下载失败", self.stat_failed_var, "#f53f3f")
        ]):
            card = ttk.Frame(kpi_frame, borderwidth=1, relief="solid")
            card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4, ipady=2)
            lbl_desc = ttk.Label(card, text=label, style="StatDesc.TLabel")
            lbl_desc.pack(anchor=tk.CENTER)
            lbl_num = ttk.Label(card, textvariable=var, font=("PingFang SC", 15, "bold") if sys.platform == "darwin" else ("Microsoft YaHei", 14, "bold"), foreground=color)
            lbl_num.pack(anchor=tk.CENTER)

        # 进度条
        prog_row = ttk.Frame(board_group)
        prog_row.pack(fill=tk.X, pady=2)
        self.progress_bar = ttk.Progressbar(prog_row, orient="horizontal", mode="determinate")
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        self.lbl_progress_percent = ttk.Label(prog_row, textvariable=self.stat_progress_text_var, width=18, anchor=tk.E)
        self.lbl_progress_percent.pack(side=tk.RIGHT)

        # 5. 实时控制台日志
        log_group = ttk.LabelFrame(main_frame, text=" 5. 运行控制台日志", style="Header.TLabelframe", padding="8")
        log_group.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_group, height=10, wrap=tk.WORD, font=("Menlo" if sys.platform == "darwin" else "Consolas", 9))
        log_scroll = ttk.Scrollbar(log_group, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)

        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.log("准备就绪。请选择导入表格文件及导出文件夹。")

    # --------------------------------------------------------------------------
    # 日志输出与辅助
    # --------------------------------------------------------------------------
    def log(self, message: str):
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def _browse_import_file(self):
        filetypes = [
            ("表格文件", "*.xlsx *.xls *.csv *.tsv *.txt"),
            ("CSV/TSV 文件", "*.csv *.tsv *.txt"),
            ("Excel 文件", "*.xlsx *.xls"),
            ("所有文件", "*.*")
        ]
        path = filedialog.askopenfilename(title="选择包含照片链接的表格", filetypes=filetypes)
        if not path:
            return

        self.import_file_path.set(path)
        self.log(f"已选择导入文件: {path}")

        # 默认自动推导同级 photos 目录作为导出文件夹
        dir_name = os.path.dirname(path)
        base_name = os.path.splitext(os.path.basename(path))[0]
        suggested_dir = os.path.join(dir_name, f"{base_name}_photos")
        if not self.output_dir_path.get().strip():
            self.output_dir_path.set(suggested_dir)

        try:
            info = PhotoDownloadEngine.inspect_file(path)
            self.file_info = info
            cols = ["(全表自动扫描)"] + info["columns"]

            if info["is_excel"]:
                self.combo_sheet["values"] = info["sheets"]
                if info["sheets"]:
                    self.sheet_var.set(info["sheets"][0])
                self.combo_sheet.configure(state="readonly")
            else:
                self.combo_sheet["values"] = ["(无Sheet/纯文本)"]
                self.sheet_var.set("(无Sheet/纯文本)")
                self.combo_sheet.configure(state="disabled")

            self.combo_url_col["values"] = cols
            if info["recommended_url_col"]:
                self.url_col_var.set(info["recommended_url_col"])
                self.log(f"自动推荐照片链接列: '{info['recommended_url_col']}'")
            else:
                self.url_col_var.set("(全表自动扫描)")

            self.lbl_file_hint.config(text=f"解析成功 | 字段数: {len(info['columns'])}", foreground="#00b42a")

            # 后台快速预统计链接数量
            threading.Thread(target=self._preview_url_count, args=(path,), daemon=True).start()

        except Exception as e:
            self.lbl_file_hint.config(text=f"解析错误: {e}", foreground="#f53f3f")
            messagebox.showerror("文件解析失败", str(e))

    def _on_sheet_changed(self, event=None):
        pass

    def _preview_url_count(self, file_path: str):
        try:
            sheet = self.sheet_var.get()
            if self.file_info and not self.file_info["is_excel"]:
                sheet = None
            col = self.url_col_var.get()
            urls = PhotoDownloadEngine.extract_urls(file_path, sheet_name=sheet, target_col=col)
            self.stat_total_var.set(f"{len(urls):,}")
            self.log(f"表格链接扫描完毕，有效唯一链接数: {len(urls):,} 个")
        except Exception as e:
            self.log(f"预扫描提示: {e}")

    def _browse_output_dir(self):
        path = filedialog.askdirectory(title="选择照片保存文件夹")
        if not path:
            return
        self.output_dir_path.set(path)
        self.log(f"已选择导出目录: {path}")

        # 统计已有照片数
        out_p = Path(path)
        if out_p.exists():
            count = sum(1 for f in out_p.iterdir() if f.is_file() and not f.name.endswith(".part") and not f.name.startswith("."))
            self.stat_existing_var.set(f"{count:,}")
            self.log(f"导出目录检测到已有完整文件: {count:,} 个")

    # --------------------------------------------------------------------------
    # 开始下载流程 (多线程)
    # --------------------------------------------------------------------------
    def _start_download_thread(self):
        in_path = self.import_file_path.get().strip()
        out_path = self.output_dir_path.get().strip()

        if not in_path or not os.path.isfile(in_path):
            messagebox.showwarning("提示", "请选择有效的导入表格文件！")
            return
        if not out_path:
            messagebox.showwarning("提示", "请指定导出照片的文件夹路径！")
            return

        workers = self.workers_var.get()
        if workers < 1 or workers > 64:
            messagebox.showwarning("提示", "并发线程数请设置在 1 到 64 之间！")
            return

        self.stop_event.clear()
        self.is_running = True
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.progress_bar["value"] = 0
        self.stat_progress_text_var.set("正在准备任务...")

        self.download_thread = threading.Thread(
            target=self._run_download_worker,
            args=(in_path, out_path, workers, self.retries_var.get(), self.timeout_var.get()),
            daemon=True
        )
        self.download_thread.start()

    def _stop_download(self):
        if self.is_running:
            self.stop_event.set()
            self.log("⚠️ 已发送停止指令，正在等待当前正在传输的切片安全退出...")
            self.btn_stop.config(state="disabled")

    def _run_download_worker(
        self,
        in_path: str,
        out_dir_str: str,
        workers: int,
        retries: int,
        timeout: int
    ):
        t_start = time.time()
        output_dir = Path(out_dir_str)
        output_dir.mkdir(parents=True, exist_ok=True)
        control_dir = output_dir / "_python_download_control"
        control_dir.mkdir(parents=True, exist_ok=True)

        manifest_path = control_dir / "download_manifest.csv"
        results_path = control_dir / "download_results.csv"
        missing_path = control_dir / "download_missing.csv"

        try:
            self.log("========== 开始照片批量下载任务 ==========")
            self.log(f"源文件: {in_path}")
            self.log(f"保存目录: {output_dir}")
            self.log(f"并发数: {workers} | 重试: {retries} | 超时: {timeout}秒")

            sheet = self.sheet_var.get()
            if self.file_info and not self.file_info["is_excel"]:
                sheet = None
            col = self.url_col_var.get()

            # 1. 提取所有链接
            urls = PhotoDownloadEngine.extract_urls(in_path, sheet_name=sheet, target_col=col, progress_callback=self.log)
            if not urls:
                self.log("⚠️ 表格中未找到任何以 http:// 或 https:// 开头的有效图片链接！")
                self.root.after(0, lambda: messagebox.showinfo("提示", "未找到任何有效图片链接，请检查选择的列。"))
                return

            tasks, collisions = PhotoDownloadEngine.build_tasks(urls, output_dir)
            total = len(tasks)
            self.stat_total_var.set(f"{total:,}")

            # 写入清单
            with manifest_path.open("w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.writer(handle)
                writer.writerow(["url", "filename", "target"])
                for u, fn, target in tasks:
                    writer.writerow([u, fn, str(target)])

            # 统计初试本地已存数量
            existing_before = sum(1 for _, _, target in tasks if target.is_file() and target.stat().st_size > 0)
            self.stat_existing_var.set(f"{existing_before:,}")
            self.log(f"构建任务完成: 共 {total:,} 条 | 已存在: {existing_before:,} | 文件名冲突加盐: {collisions:,}")

            # 2. 多线程并发调度执行
            counts = {"existing": 0, "downloaded": 0, "failed": 0, "stopped": 0}
            completed = 0
            iterator = iter(tasks)
            buffer_size = max(workers * 3, workers)

            with results_path.open("a", encoding="utf-8-sig", newline="") as res_handle:
                writer = csv.DictWriter(res_handle, fieldnames=["status", "url", "filename", "bytes", "error"])
                if results_path.stat().st_size == 0:
                    writer.writeheader()

                executor = ThreadPoolExecutor(max_workers=workers)
                pending = set()

                try:
                    for _ in range(min(buffer_size, total)):
                        try:
                            task = next(iterator)
                        except StopIteration:
                            break
                        pending.add(executor.submit(
                            PhotoDownloadEngine.download_one,
                            task, retries, timeout, self.stop_event
                        ))

                    last_ui_update = 0
                    while pending:
                        if self.stop_event.is_set():
                            break

                        done, pending = wait(pending, return_when=FIRST_COMPLETED)
                        for future in done:
                            res = future.result()
                            writer.writerow(res)
                            res_handle.flush()
                            st = res["status"]
                            counts[st] = counts.get(st, 0) + 1
                            completed += 1

                            if st == "failed":
                                self.log(f"❌ 下载失败: {res['filename']} ({res['error']})")

                            now = time.time()
                            if now - last_ui_update > 0.2 or completed == total:
                                last_ui_update = now
                                self._update_kpi_ui(completed, total, counts)

                            if not self.stop_event.is_set():
                                try:
                                    task = next(iterator)
                                    pending.add(executor.submit(
                                        PhotoDownloadEngine.download_one,
                                        task, retries, timeout, self.stop_event
                                    ))
                                except StopIteration:
                                    pass

                finally:
                    executor.shutdown(wait=False, cancel_futures=True)

            # 3. 统计缺失与输出报告
            complete_count = 0
            missing = []
            for url, filename, target in tasks:
                part = target.with_name(target.name + ".part")
                if target.is_file() and target.stat().st_size > 0:
                    complete_count += 1
                else:
                    missing.append({
                        "url": url,
                        "filename": filename,
                        "part_bytes": part.stat().st_size if part.exists() else 0
                    })

            with missing_path.open("w", encoding="utf-8-sig", newline="") as handle:
                m_writer = csv.DictWriter(handle, fieldnames=["url", "filename", "part_bytes"])
                m_writer.writeheader()
                m_writer.writerows(missing)

            cost = time.time() - t_start
            self._update_kpi_ui(completed, total, counts)

            self.log("----------------------------------------")
            if self.stop_event.is_set():
                self.log(f"⏹ 任务已手动中止。已下载与断点均已妥善保留，随时可以继续。")
            else:
                self.log(f"🎉 批量下载已完成! 耗时: {cost:.1f} 秒")

            self.log(f"📊 汇总统计:")
            self.log(f"  • 唯一照片链接总数: {total:,}")
            self.log(f"  • 本轮新下载成功: {counts.get('downloaded', 0):,}")
            self.log(f"  • 本轮识别已存在: {counts.get('existing', 0):,}")
            self.log(f"  • 本轮下载失败: {counts.get('failed', 0):,}")
            self.log(f"  • 目标目录完整文件: {complete_count:,} / {total:,}")
            self.log(f"  • 缺失清单保存至: {missing_path}")
            self.log("========================================")

            self.root.after(0, lambda: self._show_finish_dialog(output_dir, total, complete_count, counts, cost))

        except Exception as e:
            self.log(f"❌ 运行异常: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("错误", f"下载过程发生异常:\n{str(e)}"))
        finally:
            self.is_running = False
            self.root.after(0, self._reset_ui_after_run)

    def _update_kpi_ui(self, completed: int, total: int, counts: Dict[str, int]):
        pct = (completed / total * 100) if total > 0 else 0.0
        self.progress_bar["value"] = pct
        self.stat_progress_text_var.set(f"{completed:,}/{total:,} ({pct:.1f}%)")
        self.stat_downloaded_var.set(f"{counts.get('downloaded', 0):,}")
        self.stat_existing_var.set(f"{counts.get('existing', 0):,}")
        self.stat_failed_var.set(f"{counts.get('failed', 0):,}")

    def _reset_ui_after_run(self):
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")

    def _show_finish_dialog(self, output_dir: Path, total: int, complete_count: int, counts: Dict[str, int], cost: float):
        dialog = tk.Toplevel(self.root)
        dialog.title("下载任务汇报")
        dialog.geometry("520x330")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        content_frame = ttk.Frame(dialog, padding="20")
        content_frame.pack(fill=tk.BOTH, expand=True)

        is_stopped = self.stop_event.is_set()
        title_text = "⏹ 下载已中止" if is_stopped else "✅ 批量下载任务已完成！"
        title_color = "#f53f3f" if is_stopped else "#00b42a"

        title = ttk.Label(content_frame, text=title_text, font=("PingFang SC", 13, "bold") if sys.platform == "darwin" else ("Microsoft YaHei", 12, "bold"), foreground=title_color)
        title.pack(anchor=tk.W, pady=(0, 10))

        rate = (complete_count / total * 100) if total > 0 else 0.0
        info_text = (
            f"• 链接总数: {total:,} 张\n"
            f"• 目录内完整保存: {complete_count:,} 张 (达成率: {rate:.1f}%)\n"
            f"• 本次新下载: {counts.get('downloaded', 0):,} 张\n"
            f"• 已存在跳过: {counts.get('existing', 0):,} 张\n"
            f"• 本次失败数: {counts.get('failed', 0):,} 张\n"
            f"• 本次总耗时: {cost:.1f} 秒\n\n"
            f"存储路径: {output_dir}"
        )
        msg_lbl = ttk.Label(content_frame, text=info_text, justify=tk.LEFT, wraplength=480)
        msg_lbl.pack(anchor=tk.W, pady=(0, 16))

        btn_box = ttk.Frame(content_frame)
        btn_box.pack(fill=tk.X, side=tk.BOTTOM)

        def _open_folder():
            try:
                if sys.platform == "darwin":
                    subprocess.run(["open", str(output_dir)], check=False)
                elif sys.platform == "win32":
                    os.startfile(str(output_dir))
                else:
                    subprocess.run(["xdg-open", str(output_dir)], check=False)
            except Exception as e:
                messagebox.showerror("打开错误", str(e))

        ttk.Button(btn_box, text="打开导出文件夹", command=_open_folder).pack(side=tk.LEFT)
        ttk.Button(btn_box, text="确定", command=dialog.destroy, style="Primary.TButton", width=8).pack(side=tk.RIGHT)


# ==============================================================================
# 兼容原有命令行运行方式 (CLI Mode)
# ==============================================================================

def run_cli_mode(args):
    print("=" * 60)
    print("纯 Python 照片批量下载 (命令行模式)")
    print("=" * 60)

    csv_path = Path(args.csv)
    output_dir = Path(args.output)
    if not csv_path.is_file():
        raise SystemExit(f"找不到导入文件：{csv_path}")
    if not 1 <= args.workers <= 64:
        raise SystemExit("workers 必须在 1 到 64 之间")

    output_dir.mkdir(parents=True, exist_ok=True)
    control_dir = output_dir / "_python_download_control"
    control_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = control_dir / "download_manifest.csv"
    results_path = control_dir / "download_results.csv"
    missing_path = control_dir / "download_missing.csv"

    print(f"读取文件：{csv_path}")
    urls = PhotoDownloadEngine.extract_urls(str(csv_path), progress_callback=print)
    tasks, collisions = PhotoDownloadEngine.build_tasks(urls, output_dir)

    with manifest_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["url", "filename", "target"])
        for u, fn, target in tasks:
            writer.writerow([u, fn, str(target)])

    existing_before = sum(1 for _, _, target in tasks if target.is_file() and target.stat().st_size > 0)

    print("\n=== 下载准备完成 ===")
    print(f"唯一有效链接：{len(tasks):,}")
    print(f"已存在照片：{existing_before:,}")
    print(f"待处理链接：{len(tasks) - existing_before:,}")
    print(f"文件名冲突：{collisions:,}")
    print(f"并发数：{args.workers}")
    print(f"清单：{manifest_path}")

    if args.prepare_only:
        print("prepare-only 模式：未开始下载。")
        return

    counts = {"existing": 0, "downloaded": 0, "failed": 0}
    total = len(tasks)
    completed = 0
    iterator = iter(tasks)
    buffer_size = max(args.workers * 3, args.workers)

    with results_path.open("a", encoding="utf-8-sig", newline="") as result_handle:
        writer = csv.DictWriter(result_handle, fieldnames=["status", "url", "filename", "bytes", "error"])
        if results_path.stat().st_size == 0:
            writer.writeheader()

        executor = ThreadPoolExecutor(max_workers=args.workers)
        pending = set()
        try:
            for _ in range(min(buffer_size, total)):
                try:
                    task = next(iterator)
                except StopIteration:
                    break
                pending.add(executor.submit(PhotoDownloadEngine.download_one, task, args.retries, args.timeout))

            while pending:
                done, pending = wait(pending, return_when=FIRST_COMPLETED)
                for future in done:
                    res = future.result()
                    writer.writerow(res)
                    result_handle.flush()
                    counts[res["status"]] = counts.get(res["status"], 0) + 1
                    completed += 1
                    if completed % 100 == 0 or completed == total:
                        print(
                            f"进度 {completed:,}/{total:,} | "
                            f"新下载 {counts.get('downloaded', 0):,} | "
                            f"已存在 {counts.get('existing', 0):,} | "
                            f"失败 {counts.get('failed', 0):,}"
                        )
                    try:
                        task = next(iterator)
                    except StopIteration:
                        continue
                    pending.add(executor.submit(PhotoDownloadEngine.download_one, task, args.retries, args.timeout))
        except KeyboardInterrupt:
            print("\n收到中断，正在停止新任务；已下载文件和 .part 断点会保留。")
            executor.shutdown(wait=False, cancel_futures=True)
            raise
        else:
            executor.shutdown(wait=True)

    complete = sum(1 for _, _, target in tasks if target.is_file() and target.stat().st_size > 0)
    missing = len(tasks) - complete
    print("\n=== 下载结束 ===")
    print(f"本轮新下载：{counts.get('downloaded', 0):,}")
    print(f"本轮识别已存在：{counts.get('existing', 0):,}")
    print(f"本轮失败：{counts.get('failed', 0):,}")
    print(f"目录内完整文件：{complete:,}/{len(tasks):,}")
    print(f"缺失或未完成：{missing:,}")
    print("如有失败，直接重新运行同一条命令，会跳过完成文件并续传 .part。")


# ==============================================================================
# 入口函数
# ==============================================================================

def main():
    parser = argparse.ArgumentParser(description="纯 Python 照片批量下载：并发、重试、断点续传")
    parser.add_argument("--csv", help="导入表格文件路径")
    parser.add_argument("--output", help="导出文件夹路径")
    parser.add_argument("--workers", type=int, default=12, help="并发数，默认 12")
    parser.add_argument("--retries", type=int, default=8, help="每个链接重试次数，默认 8")
    parser.add_argument("--timeout", type=int, default=60, help="单次网络超时秒数，默认 60")
    parser.add_argument("--prepare-only", action="store_true", help="只扫描链接，不下载")

    # 如果命令行传递了核心参数 --csv 或者 --prepare-only 或 -h，使用 CLI 模式
    if len(sys.argv) > 1 and ("--csv" in sys.argv or "--prepare-only" in sys.argv or "-h" in sys.argv or "--help" in sys.argv):
        args = parser.parse_args()
        if args.csv and args.output:
            run_cli_mode(args)
            return
        elif args.prepare_only and args.csv:
            args.output = "./downloaded_photos"
            run_cli_mode(args)
            return
        else:
            parser.print_help()
            return

    # 默认启动图形化 GUI 界面
    root = tk.Tk()
    app = PhotoDownloaderGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
