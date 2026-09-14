import os
import threading
import queue
from tkinter import filedialog, messagebox
from typing import List, Dict, Any, Optional
from datetime import datetime

import customtkinter as ctk

from ai_client import ArkVisionClient
from sampling import (
    calculate_aql_sample_size,
    calculate_confidence_sample_size,
    stratified_sample,
)
from concurrent_runner import ConcurrentRunner, DEFAULT_WORKERS
from ui_panels import PanelsMixin
from ui_viewer import ViewerMixin
from data_processor import (
    PhotoIndex,
    ExcelTaskParser,
    collect_images_from_dir,
    export_results_to_excel,
    export_storefront_results_to_excel,
    export_annotated_original_excel_p1,
)

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

MODE_P1 = "相同照片比对核验"
MODE_P2 = "真实门头照质检识别"


class PhotoCheckerApp(PanelsMixin, ViewerMixin, ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("绍兴照片智能 AI 质检系统（并发版）")
        self.geometry("1400x950")
        self.minsize(1180, 780)

        self.current_mode = MODE_P1

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
        self.all_storefront_items: List[Dict[str, Any]] = []
        self.sampled_storefront_items: List[Dict[str, Any]] = []
        self.results_p2: List[Dict[str, Any]] = []
        self.single_test_img_path = ""

        # 并发控制
        self.is_running = False
        self.stop_requested = False
        self.msg_queue: queue.Queue = queue.Queue()
        self.run_started_at: Optional[datetime] = None
        self.runtime_stats = {"workers": DEFAULT_WORKERS, "retries": 0, "failed": 0, "throttled": 0}

        self._build_ui()
        self.after(120, self._process_queue)

    # ================== 布局 ==================

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        self.top_frame = ctk.CTkFrame(self, corner_radius=10)
        self.top_frame.grid(row=0, column=0, padx=15, pady=(12, 5), sticky="ew")
        self._build_top_bar()

        self.panel_container = ctk.CTkFrame(self, corner_radius=10)
        self.panel_container.grid(row=1, column=0, padx=15, pady=4, sticky="ew")
        self._build_panel_p1()
        self._build_panel_p2()

        self.work_split = ctk.CTkFrame(self, corner_radius=10)
        self.work_split.grid(row=2, column=0, padx=15, pady=6, sticky="nsew")
        self.work_split.grid_columnconfigure(0, weight=3)
        self.work_split.grid_columnconfigure(1, weight=2)
        self.work_split.grid_rowconfigure(0, weight=1)
        self._build_left_table()
        self._build_right_viewer()

        self.bottom_frame = ctk.CTkFrame(self, corner_radius=10, height=45)
        self.bottom_frame.grid(row=3, column=0, padx=15, pady=(4, 12), sticky="ew")
        self._build_bottom_bar()

        self._switch_mode(MODE_P1)

    def _on_mode_switched(self, mode_name):
        self._switch_mode(mode_name)

    def _switch_mode(self, mode_name):
        if self.is_running:
            messagebox.showwarning("提示", "执行中不建议切换模式，请先停止。")
            self.seg_mode.set(self.current_mode)
            return

        self.current_mode = mode_name
        if mode_name == MODE_P1:
            self.frame_p2.pack_forget()
            self.frame_p1.pack(fill="x")
            self.f_img_a.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
            self.f_img_b.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
            self.lbl_t_a.configure(text="照片 A")
            self.lbl_table_title.configure(text="相同照片抽检清单（多维证据链比对）")
        else:
            self.frame_p1.pack_forget()
            self.frame_p2.pack(fill="x")
            self.f_img_b.grid_remove()
            self.f_img_a.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
            self.lbl_t_a.configure(text="商户门头照全景视图")
            self.lbl_table_title.configure(text="门头照质检抽检清单（单图真实性质检）")

        self._setup_tree_columns()
        self.lbl_verdict.configure(text=f"AI 建议结论: 已切换为 {mode_name}", text_color="#9ca3af")
        self.progress_bar.set(0)
        self.lbl_progress.configure(text="就绪")

    # ================== API ==================

    def _get_client(self) -> ArkVisionClient:
        return ArkVisionClient(
            api_url=self.entry_url.get().strip(),
            model=self.entry_model.get().strip(),
            api_key=self.entry_key.get().strip(),
        )

    def _test_api_connection(self):
        self.lbl_api_status.configure(text="测试中...", text_color="#f59e0b")
        self.btn_test_api.configure(state="disabled")

        def _w():
            ok, msg = self._get_client().test_connection()
            self.msg_queue.put(("api_test_result", (ok, msg)))

        threading.Thread(target=_w, daemon=True).start()

    def _stop_running(self):
        self.stop_requested = True
        self.lbl_progress.configure(text="正在停止（等待已在飞的请求收尾）...")

    # ================== 模式一：数据 ==================

    def _select_excel_p1(self):
        path = filedialog.askopenfilename(
            title="选择任务 Excel 文件",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")],
        )
        if not path:
            return
        self.excel_path_p1 = path
        self.entry_excel_p1.delete(0, "end")
        self.entry_excel_p1.insert(0, path)

        if not self.photo_dir_p1:
            excel_dir = os.path.dirname(path)
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
            self.msg_queue.put(("index_p1_done", PhotoIndex(self.photo_dir_p1)))

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
        pool = self.ready_pairs if self.var_only_ready.get() else self.all_pairs
        target = len(pool)

        mode = self.combo_sample_mode_p1.get()
        if "AQL" in mode:
            calc_n = calculate_aql_sample_size(target)
        elif "95%" in mode:
            calc_n = calculate_confidence_sample_size(target, confidence=0.95, margin_error=0.05, p=0.95)
        elif "10%" in mode:
            calc_n = max(1, int(target * 0.1)) if target > 0 else 0
        elif "全量" in mode:
            calc_n = target
        else:
            try:
                calc_n = int(self.entry_sample_n_p1.get().strip())
            except Exception:
                calc_n = min(50, target)

        self.entry_sample_n_p1.delete(0, "end")
        self.entry_sample_n_p1.insert(0, str(calc_n))

        missing = total_all - total_ready
        tip = f" | 待挂载/缺失: {missing} 对" if missing > 0 else ""
        self.lbl_stats_p1.configure(text=f"总任务: {total_all} 组 | 本地就绪: {total_ready} 对{tip}")

    def _on_only_ready_toggled(self):
        self._update_sampling_p1()

    def _on_sample_mode_change_p1(self, choice=None):
        self._update_sampling_p1()

    # ================== 模式一：并发执行 ==================

    def _compare_task_fn(self, client, prompt, do_mask):
        """返回线程安全的单对比对函数（不触碰任何 tkinter 控件）。"""

        def _task(pair: Dict[str, Any]) -> Dict[str, Any]:
            pa, pb = pair.get("path_a"), pair.get("path_b")
            if not pa or not pb:
                return {
                    "is_same_photo": None,
                    "confidence": 0.0,
                    "verdict": "【照片缺失】",
                    "signboard_text_1": "-",
                    "signboard_text_2": "-",
                    "signboard_match": "-",
                    "scene_similarity": 0.0,
                    "reason": f"未找到本地照片（A: {pa or '缺'}, B: {pb or '缺'}）",
                    "call_status": "photo_missing",
                    "error": True,
                }
            return client.compare_images(
                img_path_a=pa,
                img_path_b=pb,
                task_a=pair.get("task_a", ""),
                task_b=pair.get("task_b", ""),
                custom_prompt=prompt,
                mask_watermark=do_mask,
                mask_mode="adaptive",
                use_cache=True,
            )

        return _task

    def _launch_run_p1(self, pairs: List[Dict[str, Any]], resume: bool = False):
        if not pairs:
            messagebox.showinfo("提示", "没有需要执行的任务。")
            return
        if self.is_running:
            messagebox.showwarning("提示", "已有任务在执行，请先停止或等待完成。")
            return

        if not resume:
            self.results_p1 = []
            for item in self.tree.get_children():
                self.tree.delete(item)
            for p in pairs:
                self.tree.insert(
                    "", "end", iid=f"p1_{p['id']}",
                    values=(p["id"], p.get("group", ""), p.get("task_a", ""), p.get("task_b", ""),
                            "等待质检...", "-", "-", "-", "-", "-", "排队中"),
                )
        else:
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
        self._update_runtime_label("p1", f"运行状态: 执行中 | 当前并发 {cfg.workers} | 重试 0 | 失败 0", "#f59e0b")

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
            messagebox.showwarning("提示", "当前没有可用于抽检的有效任务对！")
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
        failed_ids = {
            r.get("id") for r in self.results_p1
            if r.get("error") and str(r.get("call_status", "")) != "photo_missing"
        }
        if not failed_ids:
            messagebox.showinfo("提示", "没有可重跑的失败项（照片缺失类重跑无效，请先补齐照片）。")
            return
        self._launch_run_p1([p for p in self.sampled_pairs if p.get("id") in failed_ids], resume=True)

    def _export_p1(self):
        if not self.results_p1:
            messagebox.showinfo("提示", "暂无对比抽检结果可导出。")
            return
        out = filedialog.asksaveasfilename(
            title="保存抽检清单报告", defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
        )
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
            messagebox.showwarning("提示", "尚未执行 AI 抽检或暂无结果！")
            return

        base = os.path.splitext(os.path.basename(self.excel_path_p1))[0]
        out = filedialog.asksaveasfilename(
            title="保存带结论体系的原表副本",
            initialfile=f"{base}_追加结论体系副本.xlsx",
            initialdir=os.path.dirname(self.excel_path_p1),
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
        )
        if not out:
            return
        try:
            export_annotated_original_excel_p1(self.excel_path_p1, self.results_p1, out)
            messagebox.showinfo("导出成功", f"已生成原表追加副本:\n{out}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    # ================== 模式二：数据 ==================

    def _select_photos_p2(self):
        path = filedialog.askdirectory(title="选择待检门头照目录")
        if not path:
            return
        self.photo_dir_p2 = path
        self.entry_photos_p2.delete(0, "end")
        self.entry_photos_p2.insert(0, path)
        self.lbl_stats_p2.configure(text="正在扫描照片...")

        def _w():
            self.msg_queue.put(("scan_p2_done", collect_images_from_dir(self.photo_dir_p2)))

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
                calc_n = min(50, total)

        self.entry_sample_n_p2.delete(0, "end")
        self.entry_sample_n_p2.insert(0, str(calc_n))

    def _on_sample_mode_change_p2(self, choice=None):
        self._update_sampling_p2()

    def _select_single_test_img(self):
        path = filedialog.askopenfilename(
            title="选择单张照片",
            filetypes=[("Images", "*.jpg *.jpeg *.png *.webp"), ("All Files", "*.*")],
        )
        if not path:
            return
        self.single_test_img_path = path
        self.entry_single_test_img.delete(0, "end")
        self.entry_single_test_img.insert(0, path)

    def _test_single_image_prompt(self):
        if not self.single_test_img_path or not os.path.exists(self.single_test_img_path):
            messagebox.showwarning("提示", "请先选择一张本地照片！")
            return

        client = self._get_client()
        prompt = self.txt_prompt_p2.get("1.0", "end").strip()
        path = self.single_test_img_path

        def _w():
            try:
                res = client.check_storefront(img_path=path, custom_prompt=prompt, use_cache=False)
            except Exception as e:
                res = {"error": True, "call_status": "request_error", "reason": str(e)}
            self.msg_queue.put(("single_test_done", res))

        threading.Thread(target=_w, daemon=True).start()

    def _show_single_test_result(self, res: Dict[str, Any]):
        self.txt_reason.delete("1.0", "end")
        self.txt_reason.insert("1.0", str(res.get("reason", res)))
        self.lbl_verdict.configure(
            text=f"AI 建议结论（单图调试）: {res.get('verdict', '-')}",
            text_color="#f59e0b" if res.get("error") else "#10b981",
        )
        self._render_thumbnail(self.lbl_img_a, self.single_test_img_path)

    # ================== 模式二：并发执行 ==================

    def _start_checking_p2(self):
        if not self.all_storefront_items:
            messagebox.showwarning("提示", "请先选择待检照片目录！")
            return
        if self.is_running:
            messagebox.showwarning("提示", "已有任务在执行，请先停止或等待完成。")
            return

        try:
            sample_n = int(self.entry_sample_n_p2.get().strip())
        except ValueError:
            messagebox.showerror("错误", "请输入有效的抽检数量！")
            return
        if sample_n <= 0:
            messagebox.showwarning("提示", "抽检数量须大于 0！")
            return

        items = stratified_sample(self.all_storefront_items, min(sample_n, len(self.all_storefront_items)))
        self.sampled_storefront_items = items
        self.results_p2 = []
        for item in self.tree.get_children():
            self.tree.delete(item)
        for it in items:
            self.tree.insert(
                "", "end", iid=f"p2_{it['id']}",
                values=(it["id"], it.get("task_id", ""), it.get("photo_name", ""), "-", "-", "-", "-", "排队中"),
            )

        self.is_running = True
        self.stop_requested = False
        self.btn_start_p2.configure(state="disabled")
        self.btn_stop_p2.configure(state="normal")
        self.progress_bar.set(0)

        client = self._get_client()
        prompt = self.txt_prompt_p2.get("1.0", "end").strip()
        cfg = self._get_runner_config("p2", len(items))

        self.runtime_stats = {"workers": cfg.workers, "retries": 0, "failed": 0, "throttled": 0}
        self._update_runtime_label("p2", f"运行状态: 执行中 | 当前并发 {cfg.workers} | 重试 0 | 失败 0", "#f59e0b")

        def _task(item: Dict[str, Any]) -> Dict[str, Any]:
            path = item.get("path")
            if not path or not os.path.exists(path):
                return {
                    "is_real_storefront": None,
                    "risk_type": "-",
                    "store_name": "-",
                    "confidence": 0.0,
                    "verdict": "【照片缺失】",
                    "reason": f"本地未找到照片: {path}",
                    "call_status": "photo_missing",
                    "error": True,
                }
            return client.check_storefront(img_path=path, custom_prompt=prompt, use_cache=True)

        def _on_result(done, total, item, res, meta):
            record = {
                **item,
                **(res or {}),
                "check_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
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
            summary = runner.run(items, _task)
            self.msg_queue.put(("all_p2_done", summary))

        threading.Thread(target=_worker, daemon=True).start()

    def _export_p2(self):
        if not self.results_p2:
            messagebox.showinfo("提示", "暂无门头质检结果可导出。")
            return
        out = filedialog.asksaveasfilename(
            title="保存门头质检报告", defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
        )
        if not out:
            return
        try:
            export_storefront_results_to_excel(self.results_p2, out)
            messagebox.showinfo("导出成功", f"报告已导出至:\n{out}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))


if __name__ == "__main__":
    app = PhotoCheckerApp()
    app.mainloop()
