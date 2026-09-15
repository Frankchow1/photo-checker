"""界面构建：顶部栏、并发设置行、两个业务面板、左侧表格、底部状态栏。

从 app.py 抽出，避免单文件过大；通过 Mixin 混入 PhotoCheckerApp。
并发参数（并发数 / 最大重试 / 遇限流降速 / 失败补偿）都在这里定义，
两个模式各一套，互不干扰。
"""

import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk

from ai_client import (
    DEFAULT_API_URL,
    DEFAULT_MODEL,
    DEFAULT_API_KEY,
    DEFAULT_COMPARE_PROMPT,
    DEFAULT_STOREFRONT_PROMPT,
)
from concurrent_runner import RunnerConfig, DEFAULT_WORKERS, suggest_workers


class PanelsMixin:
    # ================== 顶部栏 ==================

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

        ttk.Separator(box, orient="vertical").pack(side="left", fill="y", padx=12, pady=4)

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

    # ================== 并发设置 ==================

    def _build_concurrency_row(self, parent, prefix: str):
        row = ctk.CTkFrame(parent, corner_radius=8, fg_color="#0f172a")
        row.pack(fill="x", padx=10, pady=3)

        ctk.CTkLabel(row, text="⚡ 并发与稳定性:", font=("SF Pro", 12, "bold"), text_color="#f59e0b").pack(side="left", padx=(8, 6))

        ctk.CTkLabel(row, text="并发数", font=("SF Pro", 11)).pack(side="left", padx=(4, 2))
        combo_workers = ctk.CTkComboBox(
            row, values=["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "12", "16"],
            width=70, font=("SF Pro", 11)
        )
        combo_workers.set(str(DEFAULT_WORKERS))
        combo_workers.pack(side="left", padx=2)

        ctk.CTkLabel(row, text="最大重试", font=("SF Pro", 11)).pack(side="left", padx=(8, 2))
        combo_retry = ctk.CTkComboBox(row, values=["0", "1", "2", "3", "4", "5"], width=60, font=("SF Pro", 11))
        combo_retry.set("3")
        combo_retry.pack(side="left", padx=2)

        var_adaptive = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            row, text="遇限流自动降速", variable=var_adaptive,
            font=("SF Pro", 11), text_color="#38bdf8"
        ).pack(side="left", padx=8)

        var_compensate = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            row, text="结束后失败自动补偿重跑", variable=var_compensate,
            font=("SF Pro", 11), text_color="#34d399"
        ).pack(side="left", padx=8)

        lbl_runtime = ctk.CTkLabel(
            row, text="运行状态: 待命 | 当前并发 - | 重试 0 | 失败 0",
            font=("SF Pro", 11), text_color="#94a3b8"
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
        if total_items <= 10:
            workers = min(workers, suggest_workers(total_items))

        try:
            retries = int(getattr(self, f"combo_retry_{prefix}").get().strip())
        except Exception:
            retries = 3

        return RunnerConfig(
            workers=max(1, workers),
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

    # ================== 模式一面板 ==================

    def _build_panel_p1(self):
        self.frame_p1 = ctk.CTkFrame(self.panel_container, fg_color="transparent")

        g = ctk.CTkFrame(self.frame_p1, fg_color="transparent")
        g.pack(fill="x", padx=10, pady=2)
        g.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(g, text="相同任务Excel:", font=("SF Pro", 12, "bold")).grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.entry_excel_p1 = ctk.CTkEntry(g, placeholder_text="选择本地哈希筛出的相同照片 Excel 清单")
        self.entry_excel_p1.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        ctk.CTkButton(g, text="选择 Excel", width=90, command=self._select_excel_p1).grid(row=0, column=2, padx=5, pady=2)

        ctk.CTkLabel(g, text="照片根目录:", font=("SF Pro", 12, "bold")).grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.entry_photos_p1 = ctk.CTkEntry(g, placeholder_text="若图片在其它文件夹可在此选择")
        self.entry_photos_p1.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        ctk.CTkButton(g, text="选择目录", width=90, command=self._select_photos_p1).grid(row=1, column=2, padx=5, pady=2)

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

        self.var_only_ready = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            row_params, text="仅抽检本地存在照片", variable=self.var_only_ready,
            command=self._on_only_ready_toggled, font=("SF Pro", 11, "bold"), text_color="#10b981"
        ).pack(side="left", padx=8)

        self.var_mask_watermark = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            row_params, text="🛡️ 前置遮蔽水印(自适应带高)", variable=self.var_mask_watermark,
            font=("SF Pro", 11, "bold"), text_color="#38bdf8"
        ).pack(side="left", padx=8)

        self.btn_toggle_prompt_p1 = ctk.CTkButton(
            row_params, text="⚙️ 查看/调整AI多维提示词", width=150, height=24,
            fg_color="#475569", hover_color="#334155", font=("SF Pro", 11), command=self._toggle_prompt_p1
        )
        self.btn_toggle_prompt_p1.pack(side="left", padx=8)

        self.lbl_stats_p1 = ctk.CTkLabel(
            row_params, text="总任务: 0 组 | 本地就绪: 0 对",
            text_color="#9ca3af", font=("SF Pro", 11)
        )
        self.lbl_stats_p1.pack(side="left", padx=6)

        self._build_concurrency_row(self.frame_p1, "p1")

        self.frame_prompt_card_p1 = ctk.CTkFrame(self.frame_p1, corner_radius=8, fg_color="#1e293b")
        p1_bar = ctk.CTkFrame(self.frame_prompt_card_p1, fg_color="transparent")
        p1_bar.pack(fill="x", padx=8, pady=(4, 2))
        ctk.CTkLabel(
            p1_bar,
            text="【AI 多维证据链比对提示词】(机位透视/光影反光/瞬态摆放物/固有微观细节/店招独立解耦):",
            font=("SF Pro", 11, "bold"), text_color="#60a5fa"
        ).pack(side="left")
        ctk.CTkButton(
            p1_bar, text="恢复默认提示词", width=95, height=20,
            font=("SF Pro", 10), command=self._reset_prompt_p1
        ).pack(side="right")

        self.txt_prompt_p1 = ctk.CTkTextbox(self.frame_prompt_card_p1, height=85, font=("SF Pro", 11), wrap="word")
        self.txt_prompt_p1.pack(fill="x", padx=8, pady=(2, 6))
        self.txt_prompt_p1.insert("1.0", DEFAULT_COMPARE_PROMPT)

        row_btns = ctk.CTkFrame(self.frame_p1, fg_color="transparent")
        row_btns.pack(fill="x", padx=10, pady=(3, 5))

        self.lbl_action_hint = ctk.CTkLabel(
            row_btns, text="状态提示：就绪，点击右侧按钮开始",
            font=("SF Pro", 11), text_color="#64748b"
        )
        self.lbl_action_hint.pack(side="left", padx=5)

        self.btn_export_copy_p1 = ctk.CTkButton(
            row_btns, text="📋 生成原表追加副本", fg_color="#0d9488", hover_color="#0f766e",
            width=150, height=30, font=("SF Pro", 12, "bold"), command=self._export_annotated_copy_p1
        )
        self.btn_export_copy_p1.pack(side="right", padx=5)

        self.btn_export_p1 = ctk.CTkButton(
            row_btns, text="导出抽检清单", fg_color="#10b981", hover_color="#059669",
            width=110, height=30, font=("SF Pro", 12), command=self._export_p1
        )
        self.btn_export_p1.pack(side="right", padx=5)

        self.btn_retry_failed_p1 = ctk.CTkButton(
            row_btns, text="🔁 重跑失败项", fg_color="#f59e0b", hover_color="#d97706",
            width=110, height=30, font=("SF Pro", 12, "bold"), state="disabled", command=self._retry_failed_p1
        )
        self.btn_retry_failed_p1.pack(side="right", padx=5)

        self.btn_stop_p1 = ctk.CTkButton(
            row_btns, text="⏹ 停止", fg_color="#ef4444", hover_color="#dc2626",
            width=75, height=30, font=("SF Pro", 12, "bold"), state="disabled", command=self._stop_running
        )
        self.btn_stop_p1.pack(side="right", padx=5)

        self.btn_start_p1 = ctk.CTkButton(
            row_btns, text="▶ 开始对比抽检", fg_color="#3b82f6", hover_color="#2563eb",
            width=125, height=30, font=("SF Pro", 12, "bold"), command=self._start_checking_p1
        )
        self.btn_start_p1.pack(side="right", padx=5)

    def _toggle_prompt_p1(self):
        if self.show_prompt_p1:
            self.frame_prompt_card_p1.pack_forget()
            self.btn_toggle_prompt_p1.configure(text="⚙️ 查看/调整AI多维提示词")
            self.show_prompt_p1 = False
        else:
            self.frame_prompt_card_p1.pack(fill="x", padx=10, pady=3)
            self.btn_toggle_prompt_p1.configure(text="▲ 收起提示词面板")
            self.show_prompt_p1 = True

    def _reset_prompt_p1(self):
        self.txt_prompt_p1.delete("1.0", "end")
        self.txt_prompt_p1.insert("1.0", DEFAULT_COMPARE_PROMPT)
        messagebox.showinfo("提示", "比对提示词已恢复为标准多维证据链版本！")

    # ================== 模式二面板 ==================

    def _build_panel_p2(self):
        self.frame_p2 = ctk.CTkFrame(self.panel_container, fg_color="transparent")

        prompt_card = ctk.CTkFrame(self.frame_p2, corner_radius=8)
        prompt_card.pack(fill="x", padx=10, pady=4)

        p_header = ctk.CTkFrame(prompt_card, fg_color="transparent")
        p_header.pack(fill="x", padx=8, pady=(4, 2))
        ctk.CTkLabel(p_header, text="门头照质检提示词设置 (支持实时调整):", font=("SF Pro", 12, "bold")).pack(side="left")
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
        ctk.CTkButton(
            debug_bar, text="⚡ 即时验证提示词", width=110, height=24,
            fg_color="#8b5cf6", hover_color="#7c3aed", command=self._test_single_image_prompt
        ).pack(side="left", padx=4)

        batch_bar = ctk.CTkFrame(self.frame_p2, fg_color="transparent")
        batch_bar.pack(fill="x", padx=10, pady=3)
        batch_bar.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(batch_bar, text="批量照片目录:", font=("SF Pro", 12, "bold")).grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.entry_photos_p2 = ctk.CTkEntry(batch_bar, placeholder_text="选择包含待检门头照的文件夹")
        self.entry_photos_p2.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        ctk.CTkButton(batch_bar, text="选择目录", width=90, command=self._select_photos_p2).grid(row=0, column=2, padx=5, pady=2)

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

        self.btn_export_p2 = ctk.CTkButton(
            op_bar, text="导出质检 Excel", fg_color="#10b981", hover_color="#059669",
            width=110, height=28, command=self._export_p2
        )
        self.btn_export_p2.pack(side="right", padx=4)

        self.btn_stop_p2 = ctk.CTkButton(
            op_bar, text="⏹ 停止", fg_color="#ef4444", hover_color="#dc2626",
            width=65, height=28, state="disabled", command=self._stop_running
        )
        self.btn_stop_p2.pack(side="right", padx=4)

        self.btn_start_p2 = ctk.CTkButton(
            op_bar, text="▶ 开始门头质检", fg_color="#3b82f6", hover_color="#2563eb",
            width=115, height=28, command=self._start_checking_p2
        )
        self.btn_start_p2.pack(side="right", padx=4)

    def _reset_prompt_p2(self):
        self.txt_prompt_p2.delete("1.0", "end")
        self.txt_prompt_p2.insert("1.0", DEFAULT_STOREFRONT_PROMPT)
        messagebox.showinfo("提示", "已恢复为标准门头质检提示词！")

    # ================== 左侧表格与底栏 ==================

    def _build_left_table(self):
        left_box = ctk.CTkFrame(self.work_split, fg_color="transparent")
        left_box.grid(row=0, column=0, sticky="nsew", padx=10, pady=8)
        left_box.grid_rowconfigure(1, weight=1)
        left_box.grid_columnconfigure(0, weight=1)

        self.lbl_table_title = ctk.CTkLabel(
            left_box, text="任务抽检清单 (单击任意一行查看证据链)",
            font=("SF Pro", 12, "bold")
        )
        self.lbl_table_title.grid(row=0, column=0, sticky="w", pady=(0, 4))

        self.tree = ttk.Treeview(left_box, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(left_box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=1, column=0, sticky="nsew")
        vsb.grid(row=1, column=1, sticky="ns")
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        self.tree.tag_configure("row_failed", foreground="#ef4444")
        self.tree.tag_configure("row_retried", foreground="#f59e0b")

    def _setup_tree_columns(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        if self.current_mode == "相同照片比对核验":
            cols = ("id", "group", "task_a", "task_b", "verdict", "same_photo",
                    "sign_1", "sign_2", "sign_match", "similarity", "run_status")
            heads = ("序号", "组标识", "任务A", "任务B", "AI建议结论", "同一底片",
                     "图1店招", "图2店招", "店招比对", "场景重合", "执行状态")
            widths = (40, 60, 85, 85, 115, 60, 100, 100, 70, 58, 95)
            anchors = ("center", "center", "w", "w", "center", "center", "w", "w", "center", "center", "center")
        else:
            cols = ("id", "task_id", "photo_name", "is_real", "risk_type", "store_name", "confidence", "run_status")
            heads = ("序号", "任务ID", "照片文件名", "真实门头", "风险类型", "识别店名", "置信度", "执行状态")
            widths = (45, 85, 140, 75, 85, 100, 60, 95)
            anchors = ("center", "w", "w", "center", "center", "w", "center", "center")

        self.tree.configure(columns=cols)
        for c, h, w, a in zip(cols, heads, widths, anchors):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, anchor=a)

    def _build_bottom_bar(self):
        self.progress_bar = ctk.CTkProgressBar(self.bottom_frame, width=300)
        self.progress_bar.pack(side="left", padx=15)
        self.progress_bar.set(0)

        self.lbl_progress = ctk.CTkLabel(self.bottom_frame, text="就绪", font=("SF Pro", 12))
        self.lbl_progress.pack(side="left", padx=5)

        self.lbl_runner_state = ctk.CTkLabel(
            self.bottom_frame, text="并发: - | 重试: 0 | 失败: 0",
            font=("SF Pro", 11), text_color="#f59e0b"
        )
        self.lbl_runner_state.pack(side="left", padx=12)

        self.lbl_summary_stats = ctk.CTkLabel(
            self.bottom_frame, text="抽检总数: 0 | 同一底片: 0 | 其他: 0",
            font=("SF Pro", 12, "bold"), text_color="#60a5fa"
        )
        self.lbl_summary_stats.pack(side="right", padx=15)
