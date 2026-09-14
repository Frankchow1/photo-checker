"""右侧证据看板与主线程消息循环。

所有并发线程都只往 self.msg_queue 投递消息，由这里的 _process_queue 在
主线程统一刷界面，避免 tkinter 跨线程操作崩溃。

消息协议：
    api_test_result   -> (ok, msg)
    index_p1_done     -> PhotoIndex
    scan_p2_done      -> [items]
    single_test_done  -> result dict
    item_p1_done      -> (done, total, record, meta)
    item_p2_done      -> (done, total, record, meta)
    runner_event      -> {"type": start/retry/throttled/compensation_start/done, ...}
    all_p1_done       -> RunSummary
    all_p2_done       -> RunSummary
"""

import os
import queue
import platform
import subprocess
from tkinter import messagebox
import customtkinter as ctk

try:
    from PIL import Image
except ImportError:
    Image = None


class ViewerMixin:
    # ================== 右侧看板 ==================

    def _build_right_viewer(self):
        right_box = ctk.CTkFrame(self.work_split, fg_color="transparent")
        right_box.grid(row=0, column=1, sticky="nsew", padx=(0, 10), pady=8)
        right_box.grid_rowconfigure(1, weight=1)
        right_box.grid_columnconfigure(0, weight=1)

        self.lbl_verdict = ctk.CTkLabel(
            right_box, text="AI 建议结论: 等待选择记录",
            font=("SF Pro", 14, "bold"), text_color="#9ca3af", anchor="w"
        )
        self.lbl_verdict.grid(row=0, column=0, sticky="ew", pady=(0, 4))

        img_area = ctk.CTkFrame(right_box, corner_radius=8)
        img_area.grid(row=1, column=0, sticky="nsew")
        img_area.grid_rowconfigure(0, weight=1)
        img_area.grid_columnconfigure(0, weight=1)
        img_area.grid_columnconfigure(1, weight=1)

        self.f_img_a = ctk.CTkFrame(img_area, corner_radius=6)
        self.f_img_a.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        self.lbl_t_a = ctk.CTkLabel(self.f_img_a, text="照片 A", font=("SF Pro", 11, "bold"))
        self.lbl_t_a.pack(pady=(4, 0))
        self.lbl_img_a = ctk.CTkLabel(self.f_img_a, text="暂无图片", font=("SF Pro", 11), text_color="#64748b")
        self.lbl_img_a.pack(expand=True, fill="both", padx=4, pady=4)

        self.f_img_b = ctk.CTkFrame(img_area, corner_radius=6)
        self.f_img_b.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        ctk.CTkLabel(self.f_img_b, text="照片 B", font=("SF Pro", 11, "bold")).pack(pady=(4, 0))
        self.lbl_img_b = ctk.CTkLabel(self.f_img_b, text="暂无图片", font=("SF Pro", 11), text_color="#64748b")
        self.lbl_img_b.pack(expand=True, fill="both", padx=4, pady=4)

        self.lbl_multi_metrics = ctk.CTkLabel(
            right_box, text="店招比对: - | 场景重合: - | 调用状态: -",
            font=("SF Pro", 11), text_color="#60a5fa", anchor="w", justify="left", wraplength=520
        )
        self.lbl_multi_metrics.grid(row=2, column=0, sticky="ew", pady=(6, 2))

        self.txt_reason = ctk.CTkTextbox(right_box, height=130, font=("SF Pro", 11), wrap="word")
        self.txt_reason.grid(row=3, column=0, sticky="ew", pady=(2, 4))

        btn_row = ctk.CTkFrame(right_box, fg_color="transparent")
        btn_row.grid(row=4, column=0, sticky="ew")
        ctk.CTkButton(
            btn_row, text="🔍 用系统看图打开原图", height=26, width=160,
            fg_color="#475569", hover_color="#334155", command=self._open_native_photos
        ).pack(side="right", padx=2)

    # ================== 消息循环 ==================

    def _process_queue(self):
        try:
            while True:
                msg_type, payload = self.msg_queue.get_nowait()

                if msg_type == "api_test_result":
                    ok, msg = payload
                    self.lbl_api_status.configure(
                        text="● 连接正常" if ok else "● 连接失败",
                        text_color="#10b981" if ok else "#ef4444"
                    )
                    self.btn_test_api.configure(state="normal")
                    if not ok:
                        messagebox.showerror("API 连接失败", str(msg))

                elif msg_type == "index_p1_done":
                    self.photo_indexer_p1 = payload
                    self._reload_data_p1()

                elif msg_type == "scan_p2_done":
                    self.all_storefront_items = payload
                    self.lbl_stats_p2.configure(text=f"已扫描照片: {len(payload)} 张")
                    self._update_sampling_p2()

                elif msg_type == "single_test_done":
                    self._show_single_test_result(payload)

                elif msg_type in ("item_p1_done", "item_p2_done"):
                    cur, total, rec, meta = payload
                    self._update_row(msg_type, rec, meta)
                    self.progress_bar.set(cur / total if total else 0)
                    self.lbl_progress.configure(text=f"进度 {cur}/{total}")
                    self._refresh_summary_stats()

                elif msg_type == "runner_event":
                    self._handle_runner_event(payload)

                elif msg_type in ("all_p1_done", "all_p2_done"):
                    self._on_run_finished(msg_type, payload)

        except queue.Empty:
            pass
        finally:
            self.after(120, self._process_queue)

    # ================== 行更新 ==================

    def _update_row(self, msg_type, rec, meta):
        retries = int(meta.get("retries", 0) or 0)
        failed = bool(rec.get("error"))
        if failed:
            status_text = f"失败({rec.get('call_status', 'error')})"
            tag = "row_failed"
        elif retries > 0:
            status_text = f"成功(重试{retries}次)"
            tag = "row_retried"
        else:
            status_text = "成功"
            tag = ""

        if msg_type == "item_p1_done":
            iid = f"p1_{rec.get('id')}"
            values = (
                rec.get("id", ""), rec.get("group", ""), rec.get("task_a", ""), rec.get("task_b", ""),
                rec.get("verdict", "-"),
                "是" if rec.get("is_same_photo") is True else ("否" if rec.get("is_same_photo") is False else "-"),
                rec.get("signboard_text_1", "-"), rec.get("signboard_text_2", "-"),
                rec.get("signboard_match", "-"), rec.get("scene_similarity", "-"),
                status_text,
            )
        else:
            iid = f"p2_{rec.get('id')}"
            values = (
                rec.get("id", ""), rec.get("task_id", ""), rec.get("photo_name", ""),
                "是" if rec.get("is_real_storefront") is True else ("否" if rec.get("is_real_storefront") is False else "-"),
                rec.get("risk_type", "-"), rec.get("store_name", "-"),
                rec.get("confidence", "-"), status_text,
            )

        if self.tree.exists(iid):
            self.tree.item(iid, values=values, tags=(tag,) if tag else ())
        else:
            self.tree.insert("", "end", iid=iid, values=values, tags=(tag,) if tag else ())

    def _refresh_summary_stats(self):
        if self.current_mode == "相同照片比对核验":
            total = len(self.results_p1)
            same = sum(1 for r in self.results_p1 if r.get("is_same_photo") is True)
            self.lbl_summary_stats.configure(text=f"抽检总数: {total} | 同一底片: {same} | 其他: {total - same}")
        else:
            total = len(self.results_p2)
            bad = sum(1 for r in self.results_p2 if r.get("is_real_storefront") is False)
            self.lbl_summary_stats.configure(text=f"抽检总数: {total} | 疑似异常: {bad} | 正常: {total - bad}")

    # ================== 并发状态 ==================

    def _handle_runner_event(self, event):
        etype = event.get("type")
        if etype == "start":
            self.runtime_stats["workers"] = event.get("workers", self.runtime_stats.get("workers"))
        elif etype == "retry":
            self.runtime_stats["retries"] = self.runtime_stats.get("retries", 0) + 1
        elif etype == "throttled":
            self.runtime_stats["throttled"] = self.runtime_stats.get("throttled", 0) + 1
            if event.get("workers"):
                self.runtime_stats["workers"] = event["workers"]
        elif etype == "compensation_start":
            self.lbl_progress.configure(text=f"补偿重跑 {event.get('count', 0)} 项...")
        elif etype == "done":
            if event.get("workers"):
                self.runtime_stats["workers"] = event["workers"]

        self.runtime_stats["failed"] = sum(
            1 for r in (self.results_p1 if self.current_mode == "相同照片比对核验" else self.results_p2)
            if r.get("error")
        )
        self._refresh_runner_state()

    def _refresh_runner_state(self):
        s = self.runtime_stats
        text = (
            f"并发: {s.get('workers', '-')} | 重试: {s.get('retries', 0)} "
            f"| 失败: {s.get('failed', 0)} | 限流: {s.get('throttled', 0)}"
        )
        self.lbl_runner_state.configure(text=text)
        prefix = "p1" if self.current_mode == "相同照片比对核验" else "p2"
        color = "#ef4444" if s.get("failed") else ("#f59e0b" if self.is_running else "#34d399")
        self._update_runtime_label(
            prefix,
            f"运行状态: {'执行中' if self.is_running else '已结束'} | 当前并发 {s.get('workers', '-')} "
            f"| 重试 {s.get('retries', 0)} | 失败 {s.get('failed', 0)}",
            color
        )

    def _format_summary(self, summary) -> str:
        return (
            f"总任务: {summary.total}\n"
            f"成功: {summary.succeeded}｜失败: {summary.failed}\n"
            f"发生重试的项: {summary.retried_items}｜重试总次数: {summary.total_retries}\n"
            f"补偿重跑: {summary.compensated} 项，救回 {summary.compensated_ok} 项\n"
            f"并发: 起始 {summary.start_workers} → 结束 {summary.final_workers}｜限流降速 {summary.throttle_events} 次\n"
            f"耗时: {summary.elapsed_sec:.1f} 秒" + ("\n（已手动停止）" if summary.stopped else "")
        )

    def _on_run_finished(self, msg_type, summary):
        self.is_running = False
        self.stop_requested = False
        self.progress_bar.set(1)
        self.lbl_progress.configure(text="已完成")

        if msg_type == "all_p1_done":
            self.btn_start_p1.configure(state="normal")
            self.btn_stop_p1.configure(state="disabled")
            self.btn_retry_failed_p1.configure(state="normal" if summary.failed else "disabled")
        else:
            self.btn_start_p2.configure(state="normal")
            self.btn_stop_p2.configure(state="disabled")

        self.runtime_stats["failed"] = summary.failed
        self.runtime_stats["workers"] = summary.final_workers
        self._refresh_runner_state()
        messagebox.showinfo("运行报告", self._format_summary(summary))

    # ================== 图片预览 ==================

    def _on_tree_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]

        source = self.results_p1 if iid.startswith("p1_") else self.results_p2
        try:
            rec_id = int(iid.split("_", 1)[1])
        except ValueError:
            return
        rec = next((r for r in source if r.get("id") == rec_id), None)
        if not rec:
            self.lbl_verdict.configure(text="AI 建议结论: 该行尚未质检", text_color="#9ca3af")
            return

        verdict = rec.get("verdict", "-")
        color = "#ef4444" if rec.get("is_same_photo") is True or rec.get("is_real_storefront") is False else "#10b981"
        if rec.get("error"):
            color = "#f59e0b"
        self.lbl_verdict.configure(text=f"AI 建议结论: {verdict}（仅供人工终审参考）", text_color=color)

        if iid.startswith("p1_"):
            metrics = (
                f"店招比对: {rec.get('signboard_match', '-')} | 场景重合: {rec.get('scene_similarity', '-')} "
                f"| 置信度(仅参考): {rec.get('confidence', '-')}\n"
                f"调用状态: {rec.get('call_status', '-')} | 重试次数: {rec.get('重试次数', 0)} "
                f"| 模型: {rec.get('model_version', '-')}"
            )
            self._render_thumbnail(self.lbl_img_a, rec.get("path_a"))
            self._render_thumbnail(self.lbl_img_b, rec.get("path_b"))
        else:
            metrics = (
                f"风险类型: {rec.get('risk_type', '-')} | 识别店名: {rec.get('store_name', '-')} "
                f"| 置信度(仅参考): {rec.get('confidence', '-')}\n"
                f"调用状态: {rec.get('call_status', '-')} | 重试次数: {rec.get('重试次数', 0)}"
            )
            self._render_thumbnail(self.lbl_img_a, rec.get("path"))

        self.lbl_multi_metrics.configure(text=metrics)
        self.txt_reason.delete("1.0", "end")
        self.txt_reason.insert("1.0", str(rec.get("reason", "")))

    def _render_thumbnail(self, label_widget, path):
        if not path or not os.path.exists(path) or Image is None:
            label_widget.configure(image=None, text="暂无图片")
            return
        try:
            img = Image.open(path)
            w, h = img.size
            max_w, max_h = 330, 300
            scale = min(max_w / w, max_h / h, 1.0)
            size = (max(1, int(w * scale)), max(1, int(h * scale)))
            ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=size)
            label_widget.configure(image=ctk_img, text="")
            label_widget.image = ctk_img
        except Exception as e:
            label_widget.configure(image=None, text=f"加载失败: {e}")

    def _open_native_photos(self):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        source = self.results_p1 if iid.startswith("p1_") else self.results_p2
        try:
            rec_id = int(iid.split("_", 1)[1])
        except ValueError:
            return
        rec = next((r for r in source if r.get("id") == rec_id), None)
        if not rec:
            return

        paths = [p for p in (rec.get("path_a"), rec.get("path_b"), rec.get("path")) if p and os.path.exists(p)]
        for p in paths:
            try:
                if platform.system() == "Darwin":
                    subprocess.run(["open", p], check=False)
                elif platform.system() == "Windows":
                    os.startfile(p)  # type: ignore[attr-defined]
                else:
                    subprocess.run(["xdg-open", p], check=False)
            except Exception:
                pass
