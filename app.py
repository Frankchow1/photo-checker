import os
import sys
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List, Dict, Any, Optional
from datetime import datetime
from PIL import Image
import customtkinter as ctk

from ai_client import (
    ArkVisionClient,
    DEFAULT_API_URL,
    DEFAULT_MODEL,
    DEFAULT_API_KEY,
    DEFAULT_COMPARE_PROMPT,
    DEFAULT_STOREFRONT_PROMPT
)
from sampling import (
    calculate_aql_sample_size,
    calculate_confidence_sample_size,
    stratified_sample
)
from concurrent_runner import (
    ConcurrentRunner,
    RunnerConfig,
    DEFAULT_WORKERS,
    suggest_workers
)
from data_processor import (
    PhotoIndex,
    ExcelTaskParser,
    collect_images_from_dir,
    export_results_to_excel,
    export_storefront_results_to_excel,
    export_annotated_original_excel_p1
)
try:
    from data_processor import export_annotated_original_excel_p2
except ImportError:
    export_annotated_original_excel_p2 = None


ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class PhotoCheckerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("绍兴照片智能 AI 质检系统 (相同照片比对 / 真实门头照识别)")
        self.geometry("1400x950")
        self.minsize(1180, 780)

        # 业务状态
        self.current_mode = "相同照片比对核验"

        # 模式一状态
        self.excel_path_p1 = ""
        self.photo_dir_p1 = ""
        self.photo_indexer_p1: Optional[PhotoIndex] = None
        self.all_pairs: List[Dict[str, Any]] = []
        self.ready_pairs: List[Dict[str, Any]] = []
        self.sampled_pairs: List[Dict[str, Any]] = []
        self.results_p1: List[Dict[str, Any]] = []
        self.show_prompt_p1 = False

        # 模式二状态
        self.photo_dir_p2 = ""
        self.excel_path_p2 = ""
        self.all_storefront_items: List[Dict[str, Any]] = []
        self.sampled_storefront_items: List[Dict[str, Any]] = []
        self.results_p2: List[Dict[str, Any]] = []
        self.single_test_img_path = ""

        # 线程与并发控制
        self.is_running = False
        self.stop_requested = False
        self.msg_queue = queue.Queue()
        self.run_started_at: Optional[datetime] = None
        self.runtime_stats = {"workers": DEFAULT_WORKERS, "retries": 0, "failed": 0, "throttled": 0}

        self._build_ui()
        self.after(100, self._process_queue)

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # 1. 顶部栏 (API与模式)
        self.top_frame = ctk.CTkFrame(self, corner_radius=10)
        self.top_frame.grid(row=0, column=0, padx=15, pady=(12, 5), sticky="ew")
        self._build_top_bar()

        # 2. 动态业务面板
        self.panel_container = ctk.CTkFrame(self, corner_radius=10)
        self.panel_container.grid(row=1, column=0, padx=15, pady=4, sticky="ew")
        self._build_panel_p1()
        self._build_panel_p2()

        # 3. 中间工作台
        self.work_split = ctk.CTkFrame(self, corner_radius=10)
        self.work_split.grid(row=2, column=0, padx=15, pady=6, sticky="nsew")
        self.work_split.grid_columnconfigure(0, weight=3) # 列表
        self.work_split.grid_columnconfigure(1, weight=2) # 预览与报告
        self.work_split.grid_rowconfigure(0, weight=1)

        self._build_left_table()
        self._build_right_viewer()

        # 4. 底部状态栏
        self.bottom_frame = ctk.CTkFrame(self, corner_radius=10, height=45)
        self.bottom_frame.grid(row=3, column=0, padx=15, pady=(4, 12), sticky="ew")
        self._build_bottom_bar()

        self._switch_mode("相同照片比对核验")

    def _build_top_bar(self):
        box = ctk.CTkFrame(self.top_frame, fg_color="transparent")
        box.pack(fill="x", padx=10, pady=6)

        ctk.CTkLabel(box, text="质检业务模式:", font=("SF Pro", 13, "bold")).pack(side="left", padx=(5, 6))
        self.seg_mode = ctk.CTkSegmentedButton(
            box,
            values=["相同照片比对核验", "真实门头照质检识别"],
            command=self._on_mode_switched,
            font=("SF Pro", 12, "bold")
        )
        self.seg_mode.set("相同照片比对核验")
        self.seg_mode.pack(side="left", padx=4)

        sep = ttk.Separator(box, orient="vertical")
        sep.pack(side="left", fill="y", padx=12, pady=4)

        ctk.CTkLabel(box, text="API链接:", font=("SF Pro", 11)).pack(side="left", padx=(4, 2))
        self.entry_url = ctk.CTkEntry(box, width=220, font=("SF Pro", 11))
        self.entry_url.insert(0, DEFAULT_API_URL)
        self.entry_url.pack(side="left", padx=2)

        ctk.CTkLabel(box, text="模型:", font=("SF Pro", 11)).pack(side="left", padx=(6, 2))
        self.entry_model = ctk.CTkEntry(box, width=140, font=("SF Pro", 11))
        self.entry_model.insert(0, DEFAULT_MODEL)
        self.entry_model.pack(side="left", padx=2)

        ctk.CTkLabel(box, text="Key:", font=("SF Pro", 11)).pack(side="left", padx=(6, 2))
        self.entry_key = ctk.CTkEntry(box, width=170, show="*", font=("SF Pro", 11))
        self.entry_key.insert(0, DEFAULT_API_KEY)
        self.entry_key.pack(side="left", padx=2)

        self.btn_test_api = ctk.CTkButton(box, text="测试连接", width=70, height=26, command=self._test_api_connection)
        self.btn_test_api.pack(side="left", padx=6)

        self.lbl_api_status = ctk.CTkLabel(box, text="● 已预留", text_color="#10b981", font=("SF Pro", 11, "bold"))
        self.lbl_api_status.pack(side="left", padx=2)

    # ================== 并发设置组件 ==================

    def _build_concurrency_row(self, parent, prefix: str):
        """并发与稳定性设置行（两个模式共用一套控件样式）。"""
        row = ctk.CTkFrame(parent, corner_radius=8, fg_color="#0f172a")
        row.pack(fill="x", padx=10, pady=3)

        ctk.CTkLabel(row, text="⚡ 并发与稳定性:", font=("SF Pro", 12, "bold"), text_color="#f59e0b").pack(side="left", padx=(8, 6))

        ctk.CTkLabel(row, text="并发数", font=("SF Pro", 11)).pack(side="left", padx=(4, 2))
        combo_workers = ctk.CTkComboBox(
            row,
            values=["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "12", "16"],
            width=70,
            font=("SF Pro", 11)
        )
        combo_workers.set(str(DEFAULT_WORKERS))
        combo_workers.pack(side="left", padx=2)

        ctk.CTkLabel(row, text="最大重试", font=("SF Pro", 11)).pack(side="left", padx=(8, 2))
        combo_retry = ctk.CTkComboBox(row, values=["0", "1", "2", "3", "4", "5"], width=60, font=("SF Pro", 11))
        combo_retry.set("3")
        combo_retry.pack(side="left", padx=2)

        var_adaptive = tk.BooleanVar(value=True)
        chk_adaptive = ctk.CTkCheckBox(
            row,
            text="遇限流自动降速",
            variable=var_adaptive,
            font=("SF Pro", 11),
            text_color="#38bdf8"
        )
        chk_adaptive.pack(side="left", padx=8)

        var_compensate = tk.BooleanVar(value=True)
        chk_compensate = ctk.CTkCheckBox(
            row,
            text="结束后失败自动补偿重跑",
            variable=var_compensate,
            font=("SF Pro", 11),
            text_color="#34d399"
        )
        chk_compensate.pack(side="left", padx=8)

        lbl_runtime = ctk.CTkLabel(
            row,
            text="运行状态: 待命 | 当前并发 - | 重试 0 | 失败 0",
            font=("SF Pro", 11),
            text_color="#94a3b8"
        )
        lbl_runtime.pack(side="left", padx=10)

        setattr(self, f"combo_workers_{prefix}", combo_workers)
        setattr(self, f"combo_retry_{prefix}", combo_retry)
        setattr(self, f"var_adaptive_{prefix}", var_adaptive)
        setattr(self, f"var_compensate_{prefix}", var_compensate)
        setattr(self, f"lbl_runtime_{prefix}", lbl_runtime)
        return row

    def _get_runner_config(self, prefix: str, total_items: int) -> RunnerConfig:
        try:
            workers = int(getattr(self, f"combo_workers_{prefix}").get().strip())
        except Exception:
            workers = DEFAULT_WORKERS
        # 小批量没必要开大并发，自动收敛
        workers = min(workers, max(1, suggest_workers(total_items)) if total_items <= 10 else workers)

        try:
            retries = int(getattr(self, f"combo_retry_{prefix}").get().strip())
        except Exception:
            retries = 3

        return RunnerConfig(
            workers=workers,
            max_retries=retries,
            adaptive=bool(getattr(self, f"var_adaptive_{prefix}").get()),
            min_workers=2 if workers >= 2 else 1,
            enable_compensation=bool(getattr(self, f"var_compensate_{prefix}").get()),
            compensation_workers=max(1, min(2, workers)),
        )

    def _update_runtime_label(self, prefix: str, text: str, color: str = "#94a3b8"):
        lbl = getattr(self, f"lbl_runtime_{prefix}", None)
        if lbl is not None:
            lbl.configure(text=text, text_color=color)

    def _build_panel_p1(self):
        self.frame_p1 = ctk.CTkFrame(self.panel_container, fg_color="transparent")

        # 路径选择行
        g = ctk.CTkFrame(self.frame_p1, fg_color="transparent")
        g.pack(fill="x", padx=10, pady=2)
        g.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(g, text="相同任务Excel:", font=("SF Pro", 12, "bold")).grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.entry_excel_p1 = ctk.CTkEntry(g, placeholder_text="选择本地哈希筛出的相同照片 Excel 清单")
        self.entry_excel_p1.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        ctk.CTkButton(g, text="选择 Excel", width=90, command=self._select_excel_p1).grid(row=0, column=2, padx=5, pady=2)

        ctk.CTkLabel(g, text="照片根目录:", font=("SF Pro", 12, "bold")).grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.entry_photos_p1 = ctk.CTkEntry(g, placeholder_text="若图片在其它文件夹可在此选择 (默认自动优先读取表格内路径及同级目录)")
        self.entry_photos_p1.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        ctk.CTkButton(g, text="选择目录", width=90, command=self._select_photos_p1).grid(row=1, column=2, padx=5, pady=2)

        # 抽样参数与开关行
        row_params = ctk.CTkFrame(self.frame_p1, fg_color="transparent")
        row_params.pack(fill="x", padx=10, pady=2)

        ctk.CTkLabel(row_params, text="抽检方式:", font=("SF Pro", 12, "bold")).pack(side="left", padx=(5, 4))
        self.combo_sample_mode_p1 = ctk.CTkComboBox(
            row_params,
            values=[
                "科学推荐抽样 (AQL 工业质检标准)",
                "统计学 95% 置信度抽样 (容差5%)",
                "按比例抽样 (10%)",
                "自定义固定数量",
                "全量核验 (100%)"
            ],
            width=230,
            command=self._on_sample_mode_change_p1
        )
        self.combo_sample_mode_p1.pack(side="left", padx=4)

        ctk.CTkLabel(row_params, text="抽检组数:", font=("SF Pro", 12, "bold")).pack(side="left", padx=(8, 4))
        self.entry_sample_n_p1 = ctk.CTkEntry(row_params, width=65)
        self.entry_sample_n_p1.insert(0, "0")
        self.entry_sample_n_p1.pack(side="left", padx=4)

        # 开关 1：仅抽检本地已有照片
        self.var_only_ready = tk.BooleanVar(value=True)
        self.chk_only_ready = ctk.CTkCheckBox(
            row_params,
            text="仅抽检本地存在照片",
            variable=self.var_only_ready,
            command=self._on_only_ready_toggled,
            font=("SF Pro", 11, "bold"),
            text_color="#10b981"
        )
        self.chk_only_ready.pack(side="left", padx=8)

        # 开关 2：物理遮蔽水印开关 (现为自适应带高)
        self.var_mask_watermark = tk.BooleanVar(value=True)
        self.chk_mask_watermark = ctk.CTkCheckBox(
            row_params,
            text="🛡️ 前置物理遮蔽水印(自适应带高)",
            variable=self.var_mask_watermark,
            font=("SF Pro", 11, "bold"),
            text_color="#38bdf8"
        )
        self.chk_mask_watermark.pack(side="left", padx=8)

        # 展开/折叠比对提示词工作台按钮
        self.btn_toggle_prompt_p1 = ctk.CTkButton(
            row_params,
            text="⚙️ 查看/调整AI多维提示词",
            width=150,
            height=24,
            fg_color="#475569",
            hover_color="#334155",
            font=("SF Pro", 11),
            command=self._toggle_prompt_p1
        )
        self.btn_toggle_prompt_p1.pack(side="left", padx=8)

        self.lbl_stats_p1 = ctk.CTkLabel(row_params, text="总任务: 0 组 | 本地就绪: 0 对", text_color="#9ca3af", font=("SF Pro", 11))
        self.lbl_stats_p1.pack(side="left", padx=6)

        # 并发与稳定性设置行
        self._build_concurrency_row(self.frame_p1, "p1")

        # 模式一的前端提示词可视化卡片 (支持展开/折叠/实时编辑)
        self.frame_prompt_card_p1 = ctk.CTkFrame(self.frame_p1, corner_radius=8, fg_color="#1e293b")
        p1_bar = ctk.CTkFrame(self.frame_prompt_card_p1, fg_color="transparent")
        p1_bar.pack(fill="x", padx=8, pady=(4, 2))
        ctk.CTkLabel(p1_bar, text="【AI 多维证据链比对提示词】(机位透视/光影反光/瞬态摆放物/固有微观细节/店招独立解耦，可在此直接修改):", font=("SF Pro", 11, "bold"), text_color="#60a5fa").pack(side="left")
        ctk.CTkButton(p1_bar, text="恢复默认提示词", width=95, height=20, font=("SF Pro", 10), command=self._reset_prompt_p1).pack(side="right")

        self.txt_prompt_p1 = ctk.CTkTextbox(self.frame_prompt_card_p1, height=85, font=("SF Pro", 11), wrap="word")
        self.txt_prompt_p1.pack(fill="x", padx=8, pady=(2, 6))
        self.txt_prompt_p1.insert("1.0", DEFAULT_COMPARE_PROMPT)

        # 操作工具条 (独立一行，空间宽裕，按钮全部清晰露出)
        row_btns = ctk.CTkFrame(self.frame_p1, fg_color="transparent")
        row_btns.pack(fill="x", padx=10, pady=(3, 5))

        self.lbl_action_hint = ctk.CTkLabel(row_btns, text="状态提示：就绪，点击右侧按钮开始", font=("SF Pro", 11), text_color="#64748b")
        self.lbl_action_hint.pack(side="left", padx=5)

        self.btn_export_copy_p1 = ctk.CTkButton(
            row_btns,
            text="📋 生成原表追加副本",
            fg_color="#0d9488",
            hover_color="#0f766e",
            width=150,
            height=30,
            font=("SF Pro", 12, "bold"),
            command=self._export_annotated_copy_p1
        )
        self.btn_export_copy_p1.pack(side="right", padx=5)

        self.btn_export_p1 = ctk.CTkButton(
            row_btns,
            text="导出抽检清单",
            fg_color="#10b981",
            hover_color="#059669",
            width=110,
            height=30,
            font=("SF Pro", 12),
            command=self._export_p1
        )
        self.btn_export_p1.pack(side="right", padx=5)

        self.btn_retry_failed_p1 = ctk.CTkButton(
            row_btns,
            text="🔁 重跑失败项",
            fg_color="#f59e0b",
            hover_color="#d97706",
            width=110,
            height=30,
            font=("SF Pro", 12, "bold"),
            state="disabled",
            command=self._retry_failed_p1
        )
        self.btn_retry_failed_p1.pack(side="right", padx=5)

        self.btn_stop_p1 = ctk.CTkButton(
            row_btns,
            text="⏹ 停止",
            fg_color="#ef4444",
            hover_color="#dc2626",
            width=75,
            height=30,
            font=("SF Pro", 12, "bold"),
            state="disabled",
            command=self._stop_running
        )
        self.btn_stop_p1.pack(side="right", padx=5)

        self.btn_start_p1 = ctk.CTkButton(
            row_btns,
            text="▶ 开始对比抽检",
            fg_color="#3b82f6",
            hover_color="#2563eb",
            width=125,
            height=30,
            font=("SF Pro", 12, "bold"),
            command=self._start_checking_p1
        )
        self.btn_start_p1.pack(side="right", padx=5)

    def _toggle_prompt_p1(self):
        if self.show_prompt_p1:
            self.frame_prompt_card_p1.pack_forget()
            self.btn_toggle_prompt_p1.configure(text="⚙️ 查看/调整AI多维提示词")
            self.show_prompt_p1 = False
        else:
            self.frame_prompt_card_p1.pack(fill="x", padx=10, pady=3, after=self.frame_p1.winfo_children()[2])
            self.btn_toggle_prompt_p1.configure(text="▲ 收起提示词面板")
            self.show_prompt_p1 = True

    def _reset_prompt_p1(self):
        self.txt_prompt_p1.delete("1.0", "end")
        self.txt_prompt_p1.insert("1.0", DEFAULT_COMPARE_PROMPT)
        messagebox.showinfo("提示", "比对提示词已恢复为官方多维证据链标准版本！")

    def _build_panel_p2(self):
        self.frame_p2 = ctk.CTkFrame(self.panel_container, fg_color="transparent")

        prompt_card = ctk.CTkFrame(self.frame_p2, corner_radius=8)
        prompt_card.pack(fill="x", padx=10, pady=4)

        p_header = ctk.CTkFrame(prompt_card, fg_color="transparent")
        p_header.pack(fill="x", padx=8, pady=(4, 2))

        ctk.CTkLabel(p_header, text="门头照质检提示词设置 (支持实时调整修改):", font=("SF Pro", 12, "bold")).pack(side="left")
        ctk.CTkButton(p_header, text="恢复默认提示词", width=95, height=22, command=self._reset_prompt_p2).pack(side="right", padx=4)

        self.txt_prompt_p2 = ctk.CTkTextbox(prompt_card, height=85, font=("SF Pro", 11), wrap="word")
        self.txt_prompt_p2.pack(fill="x", padx=8, pady=2)
        self.txt_prompt_p2.insert("1.0", DEFAULT_STOREFRONT_PROMPT)

        debug_bar = ctk.CTkFrame(prompt_card, fg_color="transparent")
        debug_bar.pack(fill="x", padx=8, pady=(2, 6))

        ctk.CTkLabel(debug_bar, text="单图验证调试:", font=("SF Pro", 11, "bold"), text_color="#60a5fa").pack(side="left")
        self.entry_single_test_img = ctk.CTkEntry(debug_bar, width=380, placeholder_text="选择一张本地照片单图调试提示词...")
        self.entry_single_test_img.pack(side="left", padx=6)

        ctk.CTkButton(debug_bar, text="选择单图", width=80, height=24, command=self._select_single_test_img).pack(side="left", padx=4)
        ctk.CTkButton(debug_bar, text="⚡ 即时验证提示词", width=110, height=24, fg_color="#8b5cf6", hover_color="#7c3aed", command=self._test_single_image_prompt).pack(side="left", padx=4)

        batch_bar = ctk.CTkFrame(self.frame_p2, fg_color="transparent")
        batch_bar.pack(fill="x", padx=10, pady=3)
        batch_bar.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(batch_bar, text="批量照片目录:", font=("SF Pro", 12, "bold")).grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.entry_photos_p2 = ctk.CTkEntry(batch_bar, placeholder_text="选择包含待检门头照的文件夹 (自动扫描全部图片)")
        self.entry_photos_p2.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        ctk.CTkButton(batch_bar, text="选择目录", width=90, command=self._select_photos_p2).grid(row=0, column=2, padx=5, pady=2)

        # 并发与稳定性设置行
        self._build_concurrency_row(self.frame_p2, "p2")

        op_bar = ctk.CTkFrame(self.frame_p2, fg_color="transparent")
        op_bar.pack(fill="x", padx=10, pady=(3, 6))

        ctk.CTkLabel(op_bar, text="抽检方式:", font=("SF Pro", 12, "bold")).pack(side="left", padx=(5, 4))
        self.combo_sample_mode_p2 = ctk.CTkComboBox(
            op_bar,
            values=[
                "科学推荐抽样 (AQL 工业质检标准)",
                "统计学 95% 置信度抽样",
                "全量核验 (100%)",
                "按比例抽样 (10%)",
                "自定义固定数量"
            ],
            width=230,
            command=self._on_sample_mode_change_p2
        )
        self.combo_sample_mode_p2.pack(side="left", padx=4)

        ctk.CTkLabel(op_bar, text="抽检数:", font=("SF Pro", 12, "bold")).pack(side="left", padx=(8, 4))
        self.entry_sample_n_p2 = ctk.CTkEntry(op_bar, width=65)
        self.entry_sample_n_p2.insert(0, "0")
        self.entry_sample_n_p2.pack(side="left", padx=4)

        self.lbl_stats_p2 = ctk.CTkLabel(op_bar, text="已扫描照片: 0 张", text_color="#9ca3af", font=("SF Pro", 11))
        self.lbl_stats_p2.pack(side="left", padx=8)

        self.btn_export_p2 = ctk.CTkButton(op_bar, text="导出质检 Excel", fg_color="#10b981", hover_color="#059669", width=110, height=28, command=self._export_p2)
        self.btn_export_p2.pack(side="right", padx=4)

        self.btn_stop_p2 = ctk.CTkButton(op_bar, text="⏹ 停止", fg_color="#ef4444", hover_color="#dc2626", width=65, height=28, state="disabled", command=self._stop_running)
        self.btn_stop_p2.pack(side="right", padx=4)

        self.btn_start_p2 = ctk.CTkButton(op_bar, text="▶ 开始门头质检", fg_color="#3b82f6", hover_color="#2563eb", width=115, height=28, command=self._start_checking_p2)
        self.btn_start_p2.pack(side="right", padx=4)

    def _build_left_table(self):
        left_box = ctk.CTkFrame(self.work_split, fg_color="transparent")
        left_box.grid(row=0, column=0, sticky="nsew", padx=10, pady=8)
        left_box.grid_rowconfigure(1, weight=1)
        left_box.grid_columnconfigure(0, weight=1)

        self.lbl_table_title = ctk.CTkLabel(left_box, text="任务抽检清单 (单击任意一行查看多维证据链核验)", font=("SF Pro", 12, "bold"))
        self.lbl_table_title.grid(row=0, column=0, sticky="w", pady=(0, 4))

        self.tree = ttk.Treeview(left_box, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(left_box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.grid(row=1, column=0, sticky="nsew")
        vsb.grid(row=1, column=1, sticky="ns")

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        # 行状态颜色：失败/重试后成功一眼可辨
        self.tree.tag_configure("row_failed", foreground="#ef4444")
        self.tree.tag_configure("row_retried", foreground="#f59e0b")

    def _setup_tree_columns(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        if self.current_mode == "相同照片比对核验":
            # 完整展示店招与多维业务定性字段
            cols = ("id", "group", "task_a", "task_b", "verdict", "same_photo", "sign_1", "sign_2", "sign_match", "similarity", "run_status")
            self.tree.configure(columns=cols)
            self.tree.heading("id", text="序号")
            self.tree.heading("group", text="组标识")
            self.tree.heading("task_a", text="任务A")
            self.tree.heading("task_b", text="任务B")
            self.tree.heading("verdict", text="AI建议结论")
            self.tree.heading("same_photo", text="同一底片")
            self.tree.heading("sign_1", text="图1店招")
            self.tree.heading("sign_2", text="图2店招")
            self.tree.heading("sign_match", text="店招比对")
            self.tree.heading("similarity", text="场景重合")
            self.tree.heading("run_status", text="执行状态")

            self.tree.column("id", width=40, anchor="center")
            self.tree.column("group", width=60, anchor="center")
            self.tree.column("task_a", width=85, anchor="w")
            self.tree.column("task_b", width=85, anchor="w")
            self.tree.column("verdict", width=115, anchor="center")
            self.tree.column("same_photo", width=60, anchor="center")
            self.tree.column("sign_1", width=100, anchor="w")
            self.tree.column("sign_2", width=100, anchor="w")
            self.tree.column("sign_match", width=70, anchor="center")
            self.tree.column("similarity", width=58, anchor="center")
            self.tree.column("run_status", width=90, anchor="center")
        else:
            cols = ("id", "task_id", "photo_name", "is_real", "risk_type", "store_name", "confidence", "run_status")
            self.tree.configure(columns=cols)
            self.tree.heading("id", text="序号")
            self.tree.heading("task_id", text="任务ID")
            self.tree.heading("photo_name", text="照片文件名")
            self.tree.heading("is_real", text="真实门头")
            self.tree.heading("risk_type", text="风险类型")
            self.tree.heading("store_name", text="识别店名")
            self.tree.heading("confidence", text="置信度")
            self.tree.heading("run_status", text="执行状态")

            self.tree.column("id", width=45, anchor="center")
            self.tree.column("task_id", width=85, anchor="w")
            self.tree.column("photo_name", width=140, anchor="w")
            self.tree.column("is_real", width=75, anchor="center")
            self.tree.column("risk_type", width=85, anchor="center")
            self.tree.column("store_name", width=100, anchor="w")
            self.tree.column("confidence", width=60, anchor="center")
            self.tree.column("run_status", width=90, anchor="center")

    def _build_right_viewer(self):
        right_box = ctk.CTkFrame(self.work_split, fg_color="transparent")
        right_box.grid(row=0, column=1, sticky="nsew", padx=10, pady=8)
        right_box.grid_rowconfigure(1, weight=1)
        right_box.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(right_box, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        self.lbl_viewer_title = ctk.CTkLabel(header, text="核验大图与 AI 多维证据链看板", font=("SF Pro", 12, "bold"))
        self.lbl_viewer_title.pack(side="left")

        self.btn_open_native = ctk.CTkButton(header, text="打开原图", width=75, height=22, state="disabled", command=self._open_native_photos)
        self.btn_open_native.pack(side="right")

        card = ctk.CTkFrame(right_box, corner_radius=8)
        card.grid(row=1, column=0, sticky="nsew")
        card.grid_rowconfigure(0, weight=1) # 图像区
        card.grid_rowconfigure(1, weight=1) # 多维证据看板与文本
        card.grid_columnconfigure(0, weight=1)
        card.grid_columnconfigure(1, weight=1)

        self.f_img_a = ctk.CTkFrame(card, fg_color="#1e293b", corner_radius=6)
        self.f_img_a.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        self.lbl_t_a = ctk.CTkLabel(self.f_img_a, text="图片 A / 门头照", font=("SF Pro", 11, "bold"), text_color="#94a3b8")
        self.lbl_t_a.pack(side="top", pady=2)
        self.lbl_img_a = ctk.CTkLabel(self.f_img_a, text="暂无图片", text_color="#64748b")
        self.lbl_img_a.pack(expand=True, fill="both", padx=4, pady=4)

        self.f_img_b = ctk.CTkFrame(card, fg_color="#1e293b", corner_radius=6)
        self.f_img_b.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        self.lbl_t_b = ctk.CTkLabel(self.f_img_b, text="对比照片 B", font=("SF Pro", 11, "bold"), text_color="#94a3b8")
        self.lbl_t_b.pack(side="top", pady=2)
        self.lbl_img_b = ctk.CTkLabel(self.f_img_b, text="暂无图片", text_color="#64748b")
        self.lbl_img_b.pack(expand=True, fill="both", padx=4, pady=4)

        # 下方：多维证据链结构化面板与详细理由
        detail_box = ctk.CTkFrame(card, fg_color="transparent")
        detail_box.grid(row=1, column=0, columnspan=2, padx=6, pady=(0, 6), sticky="nsew")
        detail_box.grid_rowconfigure(2, weight=1)
        detail_box.grid_columnconfigure(0, weight=1)

        self.lbl_verdict = ctk.CTkLabel(detail_box, text="AI 建议结论: 就绪", font=("SF Pro", 12, "bold"), text_color="#9ca3af", anchor="w")
        self.lbl_verdict.grid(row=0, column=0, sticky="w", pady=(1, 2))

        # 多维证据链指标标签条 (店招、场景重合度、机位透视)
        self.lbl_multi_metrics = ctk.CTkLabel(
            detail_box,
            text="【多维证据指标】 🏷️ 店招匹配: 待检测 | 📐 场景物理重合度: - | 🛡️ 水印状态: 已物理遮蔽",
            font=("SF Pro", 11),
            text_color="#38bdf8",
            anchor="w"
        )
        self.lbl_multi_metrics.grid(row=1, column=0, sticky="w", pady=(0, 3))

        self.txt_reason = ctk.CTkTextbox(detail_box, wrap="word", font=("SF Pro", 12))
        self.txt_reason.grid(row=2, column=0, sticky="nsew")
        self.txt_reason.insert("1.0", "请在左侧列表中点击任意记录，此处将展示 AI 基于【多维证据链 (店招/透视几何/光影反光/瞬态陈列/微观纹理)】输出的深度质检报告。")
        self.txt_reason.configure(state="disabled")

        self.current_preview_paths = (None, None)

    def _build_bottom_bar(self):
        self.progress_bar = ctk.CTkProgressBar(self.bottom_frame, width=300)
        self.progress_bar.pack(side="left", padx=15)
        self.progress_bar.set(0)

        self.lbl_progress = ctk.CTkLabel(self.bottom_frame, text="就绪", font=("SF Pro", 12))
        self.lbl_progress.pack(side="left", padx=5)

        self.lbl_runner_state = ctk.CTkLabel(
            self.bottom_frame,
            text="并发: - | 重试: 0 | 失败: 0",
            font=("SF Pro", 11),
            text_color="#f59e0b"
        )
        self.lbl_runner_state.pack(side="left", padx=12)

        self.lbl_summary_stats = ctk.CTkLabel(
            self.bottom_frame,
            text="抽检总数: 0 | 同一底片: 0 | 正常/误判: 0",
            font=("SF Pro", 12, "bold"),
            text_color="#60a5fa"
        )
        self.lbl_summary_stats.pack(side="right", padx=15)

    def _on_mode_switched(self, mode_name):
        self._switch_mode(mode_name)

    def _switch_mode(self, mode_name):
        self.current_mode = mode_name
        if mode_name == "相同照片比对核验":
            self.frame_p2.pack_forget()
            self.frame_p1.pack(fill="x")
            self.f_img_b.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
            self.lbl_t_a.configure(text="照片 A")
            self.lbl_table_title.configure(text="相同照片抽检清单 (多维证据链比对)")
        else:
            self.frame_p1.pack_forget()
            self.frame_p2.pack(fill="x")
            self.f_img_b.grid_remove()
            self.f_img_a.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
            self.lbl_t_a.configure(text="商户门头照全景视图")
            self.lbl_table_title.configure(text="门头照质检抽检清单 (单图真实性质检)")

        self._setup_tree_columns()
        self.lbl_verdict.configure(text="AI 建议结论: 模式已切换为 " + mode_name, text_color="#9ca3af")
        self.progress_bar.set(0)
        self.lbl_progress.configure(text="就绪")

    def _get_client(self) -> ArkVisionClient:
        return ArkVisionClient(
            api_url=self.entry_url.get().strip(),
            model=self.entry_model.get().strip(),
            api_key=self.entry_key.get().strip()
        )

    def _test_api_connection(self):
        self.lbl_api_status.configure(text="测试中...", text_color="#f59e0b")
        self.btn_test_api.configure(state="disabled")

        def _w():
            client = self._get_client()
            ok, msg = client.test_connection()
            self.msg_queue.put(("api_test_result", (ok, msg)))

        threading.Thread(target=_w, daemon=True).start()

    # ================== 模式一：相同照片逻辑 ==================

    def _select_excel_p1(self):
        path = filedialog.askopenfilename(
            title="选择绍兴任务 Excel 文件",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if not path:
            return
        self.excel_path_p1 = path
        self.entry_excel_p1.delete(0, "end")
        self.entry_excel_p1.insert(0, path)

        excel_dir = os.path.dirname(path)
        if not self.photo_dir_p1:
            self.photo_dir_p1 = excel_dir
            self.entry_photos_p1.delete(0, "end")
            self.entry_photos_p1.insert(0, excel_dir)
            self.photo_indexer_p1 = PhotoIndex(excel_dir)

        self._reload_data_p1()

    def _select_photos_p1(self):
        path = filedialog.askdirectory(title="选择照片根目录文件夹")
        if not path:
            return
        self.photo_dir_p1 = path
        self.entry_photos_p1.delete(0, "end")
        self.entry_photos_p1.insert(0, path)

        self.lbl_stats_p1.configure(text="正在扫描索引照片...")
        def _w():
            idx = PhotoIndex(self.photo_dir_p1)
            self.msg_queue.put(("index_p1_done", idx))

        threading.Thread(target=_w, daemon=True).start()

    def _reload_data_p1(self):
        if not self.excel_path_p1 or not os.path.exists(self.excel_path_p1):
            return
        try:
            parser = ExcelTaskParser(self.excel_path_p1)
            self.all_pairs = parser.parse_pairs(photo_indexer=self.photo_indexer_p1)
            self.ready_pairs = [p for p in self.all_pairs if p.get("is_ready")]
            self._update_sampling_p1()
        except Exception as e:
            messagebox.showerror("读取错误", str(e))

    def _update_sampling_p1(self):
        total_all = len(self.all_pairs)
        total_ready = len(self.ready_pairs)

        use_pool = self.ready_pairs if self.var_only_ready.get() else self.all_pairs
        target_total = len(use_pool)

        mode = self.combo_sample_mode_p1.get()
        if "AQL" in mode:
            calc_n = calculate_aql_sample_size(target_total)
        elif "95%" in mode:
            calc_n = calculate_confidence_sample_size(target_total, confidence=0.95, margin_error=0.05, p=0.95)
        elif "10%" in mode:
            calc_n = max(1, int(target_total * 0.1)) if target_total > 0 else 0
        elif "全量" in mode:
            calc_n = target_total
        else:
            try:
                calc_n = int(self.entry_sample_n_p1.get().strip())
            except Exception:
                calc_n = min(50, target_total)

        self.entry_sample_n_p1.delete(0, "end")
        self.entry_sample_n_p1.insert(0, str(calc_n))

        missing_n = total_all - total_ready
        missing_tip = f" | 待挂载/缺失: {missing_n}对" if missing_n > 0 else ""
        self.lbl_stats_p1.configure(
            text=f"总任务: {total_all} 组 | 本地就绪: {total_ready} 对{missing_tip}"
        )

    def _on_only_ready_toggled(self):
        self._update_sampling_p1()

    def _on_sample_mode_change_p1(self, choice):
        self._update_sampling_p1()

    def _compare_task_fn(self, client, prompt, do_mask):
        """生成单对比对任务（供并发引擎调用，必须是线程安全的纯函数）。"""
        def _task(pair: Dict[str, Any]) -> Dict[str, Any]:
            pa, pb = pair.get("path_a"), pair.get("path_b")
            if not pa or not pb:
                return {
                    "is_same_photo": None,
                    "confidence": 0.0,
                    "verdict": "【照片缺失】",
                    "has_signboard": "未知",
                    "signboard_text_1": "-",
                    "signboard_text_2": "-",
                    "signboard_match": "-",
                    "scene_similarity": 0.0,
                    "reason": f"未找到本地照片 (A: {pa or '缺'}, B: {pb or '缺'})",
                    "call_status": "photo_missing",
                    "error": True
                }
            return client.compare_images(
                img_path_a=pa,
                img_path_b=pb,
                task_a=pair.get("task_a", ""),
                task_b=pair.get("task_b", ""),
                custom_prompt=prompt,
                mask_watermark=do_mask,
                mask_mode="adaptive",
                use_cache=True
            )
        return _task

    def _launch_run_p1(self, pairs: List[Dict[str, Any]], resume: bool = False):
        """启动（或重跑）一批并发比对任务。"""
        if not pairs:
            messagebox.showinfo("提示", "没有需要执行的任务。")
            return

        if not resume:
            self.results_p1 = []
            for item in self.tree.get_children():
                self.tree.delete(item)
            for p in pairs:
                self.tree.insert("", "end", iid=f"p1_{p['id']}", values=(
                    p["id"], p.get("group", ""), p.get("task_a", ""), p.get("task_b", ""),
                    "等待质检...", "-", "-", "-", "-", "-", "排队中"
                ))
        else:
            # 重跑失败项：先把旧的失败记录从结果里移除
            retry_ids = {p["id"] for p in pairs}
            self.results_p1 = [r for r in self.results_p1 if r.get("id") not in retry_ids]
            for p in pairs:
                iid = f"p1_{p['id']}"
                if self.tree.exists(iid):
                    vals = list(self.tree.item(iid, "values"))
                    vals[-1] = "重跑中"
                    self.tree.item(iid, values=tuple(vals), tags=("row_retried",))

        self.is_running = True
        self.stop_requested = False
        self.run_started_at = datetime.now()
        self.btn_start_p1.configure(state="disabled")
        self.btn_retry_failed_p1.configure(state="disabled")
        self.btn_stop_p1.configure(state="normal")
        self.progress_bar.set(0)

        client = self._get_client()
        prompt = self.txt_prompt_p1.get("1.0", "end").strip()
        do_mask = self.var_mask_watermark.get()
        cfg = self._get_runner_config("p1", len(pairs))
        task_fn = self._compare_task_fn(client, prompt, do_mask)

        self.runtime_stats = {"workers": cfg.workers, "retries": 0, "failed": 0, "throttled": 0}
        self._update_runtime_label(
            "p1", f"运行状态: 执行中 | 并发 {cfg.workers} | 重试 0 | 失败 0", "#f59e0b"
        )

        def _on_result(done, total, item, res, meta):
            record = {
                **item,
                **(res or {}),
                "check_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "重试次数": meta.get("retries", 0),
                "执行阶段": meta.get("phase", "main"),
            }
            self.results_p1.append(record)
            self.msg_queue.put(("item_p1_done", (done, total, record, meta)))

        def _worker():
            runner = ConcurrentRunner(
                config=cfg,
                on_result=_on_result,
                on_event=lambda e: self.msg_queue.put(("runner_event", e)),
                should_stop=lambda: self.stop_requested,
            )
            summary = runner.run(pairs, task_fn)
            self.msg_queue.put(("all_p1_done", summary))

        threading.Thread(target=_worker, daemon=True).start()

    def _start_checking_p1(self):
        if not self.all_pairs:
            messagebox.showwarning("提示", "请先选择相同照片任务 Excel！")
            return

        pool = self.ready_pairs if self.var_only_ready.get() else self.all_pairs
        if not pool:
            messagebox.showwarning("提示", "当前没有可用于抽检的有效任务对！请检查照片是否在本地。")
            return

        try:
            sample_n = int(self.entry_sample_n_p1.get().strip())
        except ValueError:
            messagebox.showerror("错误", "请输入有效的抽检数量！")
            return

        if sample_n <= 0:
            messagebox.showwarning("提示", "抽检数量须大于 0！")
            return

        self.sampled_pairs = stratified_sample(pool, min(sample_n, len(pool)))
        self._launch_run_p1(self.sampled_pairs, resume=False)

    def _retry_failed_p1(self):
        """手动补偿：只重跑上一轮仍失败的条目。"""
        failed_ids = {
            r.get("id") for r in self.results_p1
            if r.get("error") and str(r.get("call_status", "")) != "photo_missing"
        }
        if not failed_ids:
            messagebox.showinfo("提示", "没有需要重跑的失败项（照片缺失类重跑也无效，请先补齐照片）。")
            return
        pairs = [p for p in self.sampled_pairs if p.get("id") in failed_ids]
        self._launch_run_p1(pairs, resume=True)

    def _export_p1(self):
        if not self.results_p1:
            messagebox.showinfo("提示", "暂无对比抽检结果可导出。")
            return
        out = filedialog.asksaveasfilename(title="保存抽检清单报告", defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if not out:
            return
        try:
            export_results_to_excel(self.results_p1, out)
            messagebox.showinfo("导出成功", f"报告已导出至:\n{out}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def _export_annotated_copy_p1(self):
        if not self.excel_path_p1 or not os.path.exists(self.excel_path_p1):
            messagebox.showwarning("提示", "请先选择并加载原始 Excel 文件！")
            return

        if not self.results_p1:
            messagebox.showwarning("提示", "尚未执行 AI 抽检或暂无抽检结果！请先点击【开始对比抽检】。")
            return

        dir_name = os.path.dirname(self.excel_path_p1)
        base_name = os.path.splitext(os.path.basename(self.excel_path_p1))[0]
        suggested_out = os.path.join(dir_name, f"{base_name}_追加结论体系副本.xlsx")

        out = filedialog.asksaveasfilename(
            title="保存带有结论体系的原表副本",
            initialfile=os.path.basename(suggested_out),
            initialdir=dir_name,
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not out:
            return

        self.lbl_progress.configure(text="正在生成原表追加副本，请稍候...")

        def _w():
            try:
                export_annotated_original_excel_p1(self.excel_path_p1, self.results_p1, out)
                self.msg_queue.put(("copy_p1_done", out))
            except Exception as e:
                self.msg_queue.put(("copy_p1_error", str(e)))

        threading.Thread(target=_w, daemon=True).start()

    # ================== 模式二：真实门头照逻辑 ==================

    def _reset_prompt_p2(self):
        self.txt_prompt_p2.delete("1.0", "end")
        self.txt_prompt_p2.insert("1.0", DEFAULT_STOREFRONT_PROMPT)
        messagebox.showinfo("提示", "已恢复为官方标准门头质检提示词！")

    def _select_single_test_img(self):
        path = filedialog.askopenfilename(
            title="选择一张待测试的门头照片",
            filetypes=[("Image Files", "*.jpg *.jpeg *.png *.webp *.bmp"), ("All Files", "*.*")]
        )
        if not path:
            return
        self.single_test_img_path = path
        self.entry_single_test_img.delete(0, "end")
        self.entry_single_test_img.insert(0, path)

    def _test_single_image_prompt(self):
        img_p = self.entry_single_test_img.get().strip()
        if not img_p or not os.path.exists(img_p):
            messagebox.showwarning("提示", "请先选择一张有效的本地单图进行测试验证！")
            return

        prompt = self.txt_prompt_p2.get("1.0", "end").strip()
        self.lbl_verdict.configure(text="正在调用 AI 验证单图提示词...", text_color="#f59e0b")
        client = self._get_client()

        def _w():
            res = client.inspect_storefront_image(img_p, custom_prompt=prompt)
            self.msg_queue.put(("single_test_done", (img_p, res)))

        threading.Thread(target=_w, daemon=True).start()

    def _select_photos_p2(self):
        path = filedialog.askdirectory(title="选择包含待检门头照的文件夹")
        if not path:
            return
        self.photo_dir_p2 = path
        self.entry_photos_p2.delete(0, "end")
        self.entry_photos_p2.insert(0, path)

        self.lbl_stats_p2.configure(text="正在扫描照片...")
        def _w():
            items = collect_images_from_dir(self.photo_dir_p2)
            self.msg_queue.put(("collect_p2_done", items))

        threading.Thread(target=_w, daemon=True).start()

    def _update_sampling_p2(self):
        total = len(self.all_storefront_items)
        mode = self.combo_sample_mode_p2.get()

        if "AQL" in mode:
            calc_n = calculate_aql_sample_size(total)
        elif "95%" in mode:
            calc_n = calculate_confidence_sample_size(total, confidence=0.95, margin_error=0.05, p=0.95)
        elif "全量" in mode:
            calc_n = total
        elif "10%" in mode:
            calc_n = max(1, int(total * 0.1)) if total > 0 else 0
        else:
            try:
                calc_n = int(self.entry_sample_n_p2.get().strip())
            except Exception:
                calc_n = min(30, total)

        self.entry_sample_n_p2.delete(0, "end")
        self.entry_sample_n_p2.insert(0, str(calc_n))
        self.lbl_stats_p2.configure(text=f"已扫描待检门头照: {total} 张")

    def _on_sample_mode_change_p2(self, choice):
        self._update_sampling_p2()

    def _start_checking_p2(self):
        if not self.all_storefront_items:
            messagebox.showwarning("提示", "请先选择批量门头照目录！")
            return

        try:
            sample_n = int(self.entry_sample_n_p2.get().strip())
        except ValueError:
            messagebox.showerror("错误", "请输入合法的抽检数量！")
            return

        if sample_n <= 0:
            messagebox.showwarning("提示", "抽检数量须大于 0！")
            return

        self.sampled_storefront_items = stratified_sample(self.all_storefront_items, min(sample_n, len(self.all_storefront_items)))
        self.results_p2 = []

        for item in self.tree.get_children():
            self.tree.delete(item)

        for it in self.sampled_storefront_items:
            self.tree.insert("", "end", iid=f"p2_{it['id']}", values=(
                it["id"], it.get("task_id", ""), it.get("photo_name", ""),
                "等待中", "-", "-", "-", "排队中"
            ))

        self.is_running = True
        self.stop_requested = False
        self.run_started_at = datetime.now()
        self.btn_start_p2.configure(state="disabled")
        self.btn_stop_p2.configure(state="normal")
        self.progress_bar.set(0)

        client = self._get_client()
        prompt = self.txt_prompt_p2.get("1.0", "end").strip()
        cfg = self._get_runner_config("p2", len(self.sampled_storefront_items))

        self.runtime_stats = {"workers": cfg.workers, "retries": 0, "failed": 0, "throttled": 0}
        self._update_runtime_label(
            "p2", f"运行状态: 执行中 | 并发 {cfg.workers} | 重试 0 | 失败 0", "#f59e0b"
        )

        def _task(it: Dict[str, Any]) -> Dict[str, Any]:
            img_path = it.get("path")
            if not img_path or not os.path.exists(img_path):
                return {
                    "is_real_storefront": None,
                    "confidence": 0.0,
                    "store_name": "缺失",
                    "risk_type": "照片缺失",
                    "reason": f"未定位到本地图片: {img_path}",
                    "call_status": "photo_missing",
                    "error": True
                }
            return client.inspect_storefront_image(img_path, custom_prompt=prompt)

        def _on_result(done, total, item, res, meta):
            record = {
                **item,
                **(res or {}),
                "重试次数": meta.get("retries", 0),
                "执行阶段": meta.get("phase", "main"),
            }
            self.results_p2.append(record)
            self.msg_queue.put(("item_p2_done", (done, total, record, meta)))

        def _worker():
            runner = ConcurrentRunner(
                config=cfg,
                on_result=_on_result,
                on_event=lambda e: self.msg_queue.put(("runner_event", e)),
                should_stop=lambda: self.stop_requested,
            )
            summary = runner.run(self.sampled_storefront_items, _task)
            self.msg_queue.put(("all_p2_done", summary))

        threading.Thread(target=_worker, daemon=True).start()

    def _export_p2(self):
        if not self.results_p2:
            messagebox.showinfo("提示", "暂无门头照质检结果可导出。")
            return
        out = filedialog.asksaveasfilename(title="保存门头照质检报告", defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if not out:
            return
        try:
            export_storefront_results_to_excel(self.results_p2, out)
            messagebox.showinfo("导出成功", f"门头照质检报告已成功导出至:\n{out}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def _stop_running(self):
        self.stop_requested = True
        self.lbl_progress.configure(text="正在终止（等待已发出的请求回收）...")

    # ================== 运行报告 ==================

    def _format_summary(self, summary: Dict[str, Any]) -> str:
        total = summary.get("total", 0)
        ok = summary.get("succeeded", 0)
        failed = summary.get("failed", 0)
        elapsed = summary.get("elapsed_sec", 0.0) or 0.0
        speed = (ok / elapsed) if elapsed > 0 else 0.0
        lines = [
            f"总条数: {total}",
            f"成功: {ok} | 失败: {failed}",
            f"发生重试的条目: {summary.get('retried_items', 0)} （累计重试 {summary.get('total_retries', 0)} 次）",
            f"补偿轮: 重跑 {summary.get('compensated', 0)} 条，救回 {summary.get('compensated_ok', 0)} 条",
            f"限流降速事件: {summary.get('throttle_events', 0)} 次",
            f"并发: 起始 {summary.get('start_workers', 0)} → 结束 {summary.get('final_workers', 0)}",
            f"耗时: {elapsed:.1f} 秒，平均 {speed:.2f} 条/秒",
        ]
        if summary.get("stopped"):
            lines.append("注意：本次运行被手动终止，未完成部分可再次点击开始或重跑失败项。")
        if failed:
            lines.append("失败条目已保留在列表中（红色），可点击【🔁 重跑失败项】补偿；已成功的结果有缓存，不会重复计费。")
        return "\n".join(lines)

    # ================== 消息循环处理 ==================

    def _process_queue(self):
        try:
            while True:
                m_type, data = self.msg_queue.get_nowait()

                if m_type == "api_test_result":
                    ok, msg = data
                    self.btn_test_api.configure(state="normal")
                    if ok:
                        self.lbl_api_status.configure(text="● 正常", text_color="#10b981")
                        messagebox.showinfo("API 连接成功", msg)
                    else:
                        self.lbl_api_status.configure(text="● 失败", text_color="#ef4444")
                        messagebox.showerror("API 验证失败", msg)

                elif m_type == "index_p1_done":
                    self.photo_indexer_p1 = data
                    self._reload_data_p1()
                    cnt = len(self.photo_indexer_p1.by_filename)
                    messagebox.showinfo("索引完成", f"已成功索引 {cnt} 张照片文件！")

                elif m_type == "runner_event":
                    self._handle_runner_event(data)

                elif m_type == "item_p1_done":
                    cur, total, rec, meta = data
                    iid = f"p1_{rec['id']}"

                    verdict_s = rec.get("verdict", "完成")
                    same_val = rec.get("is_same_photo", rec.get("is_same"))
                    is_same_s = "-" if same_val is None else ("是" if same_val else "否")
                    sign_1_s = str(rec.get("signboard_text_1", "-"))[:15]
                    sign_2_s = str(rec.get("signboard_text_2", "-"))[:15]
                    match_s = rec.get("signboard_match", "-")
                    sim_s = f"{rec.get('scene_similarity', 0.0):.2f}"

                    retries = meta.get("retries", 0)
                    if not meta.get("ok"):
                        run_s = f"失败({meta.get('status', '')})"[:14]
                        tag = "row_failed"
                    elif retries > 0:
                        run_s = f"重试{retries}次后成功"
                        tag = "row_retried"
                    else:
                        run_s = "成功"
                        tag = ""

                    if self.tree.exists(iid):
                        self.tree.item(iid, values=(
                            rec["id"], rec.get("group", ""), rec.get("task_a", ""), rec.get("task_b", ""),
                            verdict_s, is_same_s, sign_1_s, sign_2_s, match_s, sim_s, run_s
                        ), tags=((tag,) if tag else ()))
                        self.tree.see(iid)

                    self.progress_bar.set(cur / max(1, total))
                    self.lbl_progress.configure(text=f"并发抽检进度: {cur}/{total}")

                    if not meta.get("ok"):
                        self.runtime_stats["failed"] += 1
                    self.runtime_stats["retries"] += retries
                    self.runtime_stats["workers"] = meta.get("workers_now", self.runtime_stats["workers"])
                    self._refresh_runner_state("p1")

                    same_n = sum(1 for r in self.results_p1 if r.get("is_same_photo", r.get("is_same")))
                    other_n = len(self.results_p1) - same_n
                    self.lbl_summary_stats.configure(
                        text=f"已检: {len(self.results_p1)} | AI建议同一底片: {same_n} | 其他: {other_n}"
                    )

                elif m_type == "all_p1_done":
                    summary = data or {}
                    self.is_running = False
                    self.btn_start_p1.configure(state="normal")
                    self.btn_stop_p1.configure(state="disabled")
                    has_failed = bool(summary.get("failed", 0))
                    self.btn_retry_failed_p1.configure(state="normal" if has_failed else "disabled")
                    self.lbl_progress.configure(text="本轮并发抽检结束")
                    self._update_runtime_label(
                        "p1",
                        f"运行状态: 完成 | 并发 {summary.get('final_workers', '-')} | 重试 {summary.get('total_retries', 0)} | 失败 {summary.get('failed', 0)}",
                        "#34d399" if not has_failed else "#ef4444"
                    )
                    messagebox.showinfo("运行报告", self._format_summary(summary))

                elif m_type == "copy_p1_done":
                    out_path = data
                    self.lbl_progress.configure(text="原表追加副本生成成功！")
                    messagebox.showinfo("生成原表副本成功", f"已在原表右侧追加结论体系与审计字段，并保存为副本文件:\n\n{out_path}")

                elif m_type == "copy_p1_error":
                    err_msg = data
                    self.lbl_progress.configure(text="生成副本失败")
                    messagebox.showerror("生成副本失败", f"处理过程中出错:\n{err_msg}")

                elif m_type == "collect_p2_done":
                    self.all_storefront_items = data
                    self._update_sampling_p2()
                    messagebox.showinfo("扫描完成", f"在指定目录中共找到 {len(self.all_storefront_items)} 张待检图片！")

                elif m_type == "single_test_done":
                    img_p, res = data
                    self.current_preview_paths = (img_p, None)
                    self.btn_open_native.configure(state="normal")
                    self._render_thumbnail(img_p, self.lbl_img_a, self.lbl_t_a, "单图验证照片")

                    is_real = res.get("is_real_storefront", False)
                    risk = res.get("risk_type", "未知")
                    store = res.get("store_name", "无法识别")
                    conf = res.get("confidence", 0.0)
                    reason = res.get("reason", "")

                    verdict_str = f"【单图测试】AI建议真实门头: {'✅ 是' if is_real else '❌ 否'} | 风险: {risk} | 店名: {store} (置信度仅参考: {conf:.2f})"
                    self.lbl_verdict.configure(text=verdict_str, text_color="#10b981" if is_real else "#ef4444")
                    self.lbl_multi_metrics.configure(text=f"【门头指标】 店名: {store} | 风险类别: {risk} | 置信度(仅参考): {conf:.2f}")

                    self.txt_reason.configure(state="normal")
                    self.txt_reason.delete("1.0", "end")
                    detail_text = f"【单图即时验证报告】:\n\n"
                    detail_text += f"● AI 建议是否真实门头照: {'真实门头' if is_real else '非真实门头 / 待核实'}\n"
                    detail_text += f"● 风险类型分类: {risk}\n"
                    detail_text += f"● 识别门头招牌: {store}\n"
                    detail_text += f"● AI置信度(仅参考): {conf:.2f}\n"
                    detail_text += f"● 详细判定依据:\n{reason}\n\n"
                    detail_text += f"测试照片绝对路径: {img_p}"
                    self.txt_reason.insert("1.0", detail_text)
                    self.txt_reason.configure(state="disabled")

                elif m_type == "item_p2_done":
                    cur, total, rec, meta = data
                    iid = f"p2_{rec['id']}"
                    real_val = rec.get("is_real_storefront")
                    is_real_s = "异常" if rec.get("error") else ("真实" if real_val else "待核实")
                    risk_s = rec.get("risk_type", "-")
                    store_s = rec.get("store_name", "-")
                    conf_s = f"{rec.get('confidence', 0.0):.2f}"

                    retries = meta.get("retries", 0)
                    if not meta.get("ok"):
                        run_s = f"失败({meta.get('status', '')})"[:14]
                        tag = "row_failed"
                    elif retries > 0:
                        run_s = f"重试{retries}次后成功"
                        tag = "row_retried"
                    else:
                        run_s = "成功"
                        tag = ""

                    if self.tree.exists(iid):
                        self.tree.item(iid, values=(
                            rec["id"], rec.get("task_id", ""), rec.get("photo_name", ""),
                            is_real_s, risk_s, store_s, conf_s, run_s
                        ), tags=((tag,) if tag else ()))
                        self.tree.see(iid)

                    self.progress_bar.set(cur / max(1, total))
                    self.lbl_progress.configure(text=f"并发质检进度: {cur}/{total}")

                    if not meta.get("ok"):
                        self.runtime_stats["failed"] += 1
                    self.runtime_stats["retries"] += retries
                    self.runtime_stats["workers"] = meta.get("workers_now", self.runtime_stats["workers"])
                    self._refresh_runner_state("p2")

                    real_n = sum(1 for r in self.results_p2 if r.get("is_real_storefront"))
                    other_n = len(self.results_p2) - real_n
                    self.lbl_summary_stats.configure(
                        text=f"已质检: {len(self.results_p2)} | AI建议真实: {real_n} | 待核实/异常: {other_n}"
                    )

                elif m_type == "all_p2_done":
                    summary = data or {}
                    self.is_running = False
                    self.btn_start_p2.configure(state="normal")
                    self.btn_stop_p2.configure(state="disabled")
                    self.lbl_progress.configure(text="门头照并发质检结束")
                    self._update_runtime_label(
                        "p2",
                        f"运行状态: 完成 | 并发 {summary.get('final_workers', '-')} | 重试 {summary.get('total_retries', 0)} | 失败 {summary.get('failed', 0)}",
                        "#34d399" if not summary.get("failed") else "#ef4444"
                    )
                    messagebox.showinfo("运行报告", self._format_summary(summary))

        except queue.Empty:
            pass
        self.after(100, self._process_queue)

    def _handle_runner_event(self, event: Dict[str, Any]):
        etype = event.get("type")
        prefix = "p1" if self.current_mode == "相同照片比对核验" else "p2"

        if etype == "throttled":
            self.runtime_stats["throttled"] += 1
            self.runtime_stats["workers"] = event.get("workers", self.runtime_stats["workers"])
            self.lbl_progress.configure(
                text=f"检测到限流/服务端压力，已自动降并发至 {event.get('workers')}"
            )
            self._refresh_runner_state(prefix)

        elif etype == "retry":
            self.lbl_progress.configure(
                text=f"第 {event.get('index')} 条重试中 ({event.get('attempt')}/{event.get('max_attempts')})：{event.get('status', '')}"
            )

        elif etype == "compensation_start":
            self.lbl_progress.configure(
                text=f"进入补偿轮：{event.get('count')} 条失败项将以 {event.get('workers')} 并发重跑"
            )
            self._update_runtime_label(prefix, f"运行状态: 补偿轮 | 并发 {event.get('workers')} ", "#f59e0b")

    def _refresh_runner_state(self, prefix: str):
        s = self.runtime_stats
        text = f"并发: {s.get('workers', '-')} | 重试: {s.get('retries', 0)} | 失败: {s.get('failed', 0)}"
        if s.get("throttled"):
            text += f" | 降速: {s['throttled']}次"
        self.lbl_runner_state.configure(text=text)
        self._update_runtime_label(
            prefix,
            f"运行状态: 执行中 | 当前并发 {s.get('workers', '-')} | 重试 {s.get('retries', 0)} | 失败 {s.get('failed', 0)}",
            "#f59e0b"
        )

    def _on_tree_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]

        if self.current_mode == "相同照片比对核验":
            item_id = int(iid.replace("p1_", ""))
            rec = next((r for r in self.results_p1 if r.get("id") == item_id), None)
            if not rec:
                rec = next((r for r in self.sampled_pairs if r.get("id") == item_id), None)
            if not rec:
                return

            pa, pb = rec.get("path_a"), rec.get("path_b")
            self.current_preview_paths = (pa, pb)
            self.btn_open_native.configure(state="normal" if pa or pb else "disabled")

            self._render_thumbnail(pa, self.lbl_img_a, self.lbl_t_a, f"照片A: {rec.get('photo_a', '')}")
            self._render_thumbnail(pb, self.lbl_img_b, self.lbl_t_b, f"照片B: {rec.get('photo_b', '')}")

            is_same = rec.get("is_same_photo", rec.get("is_same"))
            verdict = rec.get("verdict", "待抽检")
            conf = rec.get("confidence", 0.0)
            sign_match = rec.get("signboard_match", "未比对")
            sim_score = rec.get("scene_similarity", 0.0)
            reason = rec.get("reason", "等待开始抽检...")

            color = "#ef4444" if is_same else "#10b981"
            self.lbl_verdict.configure(text=f"AI 建议结论: {verdict} (置信度仅参考: {conf:.2f})", text_color=color)

            # 多维结构化看板更新
            s1 = rec.get