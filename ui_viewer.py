"""右侧核验看板 + 消息循环 + 运行报告。

从 app.py 抽出，避免单文件过大；通过 Mixin 混入 PhotoCheckerApp。
并发模式下，所有 UI 更新都只在主线程的 _process_queue 里做，
工作线程只往 msg_queue 投消息，不直接碰 tkinter 控件。
"""

import os
import queue
import sys
from typing import Any, Dict, Optional
from tkinter import messagebox
from PIL import Image
import customtkinter as ctk


class ViewerMixin:
    # ================== 右侧看板 ==================

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
        card.grid_rowconfigure(0, weight=1)
        card.grid_rowconfigure(1, weight=1)
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

        detail_box = ctk.CTkFrame(card, fg_color="transparent")
        detail_box.grid(row=1, column=0, columnspan=2, padx=6, pady=(0, 6), sticky="nsew")
        detail_box.grid_rowconfigure(2, weight=1)
        detail_box.grid_columnconfigure(0, weight=1)

        self.lbl_verdict = ctk.CTkLabel(detail_box, text="AI 建议结论: 就绪", font=("SF Pro", 12, "bold"), text_color="#9ca3af", anchor="w")
        self.lbl_verdict.grid(row=0, column=0, sticky="w", pady=(1, 2))

        self.lbl_multi_metrics = ctk.CTkLabel(
            detail_box,
            text="【多维证据指标】 🏷️ 店招匹配: 待检测 | 📐 场景物理重合度: - | 🛡️ 水印状态: 自适应遮蔽",
            font=("SF Pro", 11),
            text_color="#38bdf8",
            anchor="w"
        )
        self.lbl_multi_metrics.grid(row=1, column=0, sticky="w", pady=(0, 3))

        self.txt_reason = ctk.CTkTextbox(detail_box, wrap="word", font=("SF Pro", 12))
        self.txt_reason.grid(row=2, column=0, sticky="nsew")
        self.txt_reason.insert("1.0", "请在左侧列表中点击任意记录，此处展示 AI 基于多维证据链输出的建议，以及本条的执行状态（重试次数、调用状态、遮蔽比例）。")
        self.txt_reason.configure(state="disabled")

        self.current_preview_paths = (None, None)

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
            f"发生重试的条目: {summary.get('retried_items', 0)}（累计重试 {summary.get('total_retries', 0)} 次）",
            f"补偿轮: 重跑 {summary.get('compensated', 0)} 条，救回 {summary.get('compensated_ok', 0)} 条",
            f"限流降速事件: {summary.get('throttle_events', 0)} 次",
            f"并发: 起始 {summary.get('start_workers', 0)} → 结束 {summary.get('final_workers', 0)}",
            f"耗时: {elapsed:.1f} 秒，平均 {speed:.2f} 条/秒",
        ]
        if summary.get("stopped"):
            lines.append("注意：本次运行被手动终止，未完成部分可再次开始或重跑失败项。")
        if failed:
            lines.append("失败条目在列表中标红，可点【🔁 重跑失败项】补偿；已成功的结果有缓存，不会重复计费。")
        return "\n".join(lines)

    def _handle_runner_event(self, event: Dict[str, Any]):
        etype = event.get("type")
        prefix = "p1" if self.current_mode == "相同照片比对核验" else "p2"

        if etype == "start":
            self.runtime_stats["workers"] = event.get("workers", self.runtime_stats.get("workers"))
            self._refresh_runner_state(prefix)

        elif etype == "throttled":
            self.runtime_stats["throttled"] = self.runtime_stats.get("throttled", 0) + 1
            self.runtime_stats["workers"] = event.get("workers", self.runtime_stats.get("workers"))
            self.lbl_progress.configure(text=f"检测到限流/服务端压力，已自动降并发至 {event.get('workers')}")
            self._refresh_runner_state(prefix)

        elif etype == "retry":
            self.lbl_progress.configure(
                text=f"第 {event.get('index')} 条重试中 ({event.get('attempt')}/{event.get('max_attempts')})：{event.get('status', '')}"
            )

        elif etype == "compensation_start":
            self.lbl_progress.configure(
                text=f"进入补偿轮：{event.get('count')} 条失败项将以 {event.get('workers')} 并发重跑"
            )
            self._update_runtime_label(prefix, f"运行状态: 补偿轮 | 并发 {event.get('workers')}", "#f59e0b")

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

    @staticmethod
    def _run_status_text(meta: Dict[str, Any]):
        retries = meta.get("retries", 0)
        if not meta.get("ok"):
            return f"失败({meta.get('status', '')})"[:16], "row_failed"
        if retries > 0:
            return f"重试{retries}次后成功", "row_retried"
        return "成功", ""

    # ================== 消息循环 ==================

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
                    messagebox.showinfo("索引完成", f"已成功索引 {len(self.photo_indexer_p1.by_filename)} 张照片文件！")

                elif m_type == "runner_event":
                    self._handle_runner_event(data)

                elif m_type == "item_p1_done":
                    cur, total, rec, meta = data
                    iid = f"p1_{rec['id']}"
                    same_val = rec.get("is_same_photo", rec.get("is_same"))
                    is_same_s = "-" if same_val is None else ("是" if same_val else "否")
                    run_s, tag = self._run_status_text(meta)

                    if self.tree.exists(iid):
                        self.tree.item(iid, values=(
                            rec["id"], rec.get("group", ""), rec.get("task_a", ""), rec.get("task_b", ""),
                            rec.get("verdict", "完成"), is_same_s,
                            str(rec.get("signboard_text_1", "-"))[:15],
                            str(rec.get("signboard_text_2", "-"))[:15],
                            rec.get("signboard_match", "-"),
                            f"{rec.get('scene_similarity', 0.0):.2f}",
                            run_s
                        ), tags=((tag,) if tag else ()))
                        self.tree.see(iid)

                    self.progress_bar.set(cur / max(1, total))
                    self.lbl_progress.configure(text=f"并发抽检进度: {cur}/{total}")

                    if not meta.get("ok"):
                        self.runtime_stats["failed"] = self.runtime_stats.get("failed", 0) + 1
                    self.runtime_stats["retries"] = self.runtime_stats.get("retries", 0) + meta.get("retries", 0)
                    self.runtime_stats["workers"] = meta.get("workers_now", self.runtime_stats.get("workers"))
                    self._refresh_runner_state("p1")

                    same_n = sum(1 for r in self.results_p1 if r.get("is_same_photo", r.get("is_same")))
                    self.lbl_summary_stats.configure(
                        text=f"已检: {len(self.results_p1)} | AI建议同一底片: {same_n} | 其他: {len(self.results_p1) - same_n}"
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
                        "#ef4444" if has_failed else "#34d399"
                    )
                    messagebox.showinfo("运行报告", self._format_summary(summary))

                elif m_type == "copy_p1_done":
                    self.lbl_progress.configure(text="原表追加副本生成成功！")
                    messagebox.showinfo("生成原表副本成功", f"已在原表右侧追加结论体系与审计字段:\n\n{data}")

                elif m_type == "copy_p1_error":
                    self.lbl_progress.configure(text="生成副本失败")
                    messagebox.showerror("生成副本失败", f"处理过程中出错:\n{data}")

                elif m_type == "collect_p2_done":
                    self.all_storefront_items = data
                    self._update_sampling_p2()
                    messagebox.showinfo("扫描完成", f"共找到 {len(self.all_storefront_items)} 张待检图片！")

                elif m_type == "single_test_done":
                    img_p, res = data
                    self.current_preview_paths = (img_p, None)
                    self.btn_open_native.configure(state="normal")
                    self._render_thumbnail(img_p, self.lbl_img_a, self.lbl_t_a, "单图验证照片")

                    is_real = res.get("is_real_storefront", False)
                    risk = res.get("risk_type", "未知")
                    store = res.get("store_name", "无法识别")
                    conf = res.get("confidence", 0.0)

                    self.lbl_verdict.configure(
                        text=f"【单图测试】AI建议: {'真实门头' if is_real else '非真实/待核实'} | 风险: {risk} | 店名: {store}",
                        text_color="#10b981" if is_real else "#ef4444"
                    )
                    self.lbl_multi_metrics.configure(
                        text=f"【门头指标】 店名: {store} | 风险类别: {risk} | 置信度(仅参考): {conf:.2f} | 调用状态: {res.get('call_status', '-')}"
                    )
                    self.txt_reason.configure(state="normal")
                    self.txt_reason.delete("1.0", "end")
                    self.txt_reason.insert("1.0", (
                        "【单图即时验证报告】\n\n"
                        f"● AI 建议: {'真实门头' if is_real else '非真实门头 / 待核实'}\n"
                        f"● 风险类型: {risk}\n"
                        f"● 识别招牌: {store}\n"
                        f"● 置信度(仅参考): {conf:.2f}\n"
                        f"● 模型版本: {res.get('model_version', '-')} | prompt版本: {res.get('prompt_version', '-')}\n"
                        f"● 判定依据:\n{res.get('reason', '')}\n\n"
                        f"照片路径: {img_p}"
                    ))
                    self.txt_reason.configure(state="disabled")

                elif m_type == "item_p2_done":
                    cur, total, rec, meta = data
                    iid = f"p2_{rec['id']}"
                    is_real_s = "异常" if rec.get("error") else ("真实" if rec.get("is_real_storefront") else "待核实")
                    run_s, tag = self._run_status_text(meta)

                    if self.tree.exists(iid):
                        self.tree.item(iid, values=(
                            rec["id"], rec.get("task_id", ""), rec.get("photo_name", ""),
                            is_real_s, rec.get("risk_type", "-"), rec.get("store_name", "-"),
                            f"{rec.get('confidence', 0.0):.2f}", run_s
                        ), tags=((tag,) if tag else ()))
                        self.tree.see(iid)

                    self.progress_bar.set(cur / max(1, total))
                    self.lbl_progress.configure(text=f"并发质检进度: {cur}/{total}")

                    if not meta.get("ok"):
                        self.runtime_stats["failed"] = self.runtime_stats.get("failed", 0) + 1
                    self.runtime_stats["retries"] = self.runtime_stats.get("retries", 0) + meta.get("retries", 0)
                    self.runtime_stats["workers"] = meta.get("workers_now", self.runtime_stats.get("workers"))
                    self._refresh_runner_state("p2")

                    real_n = sum(1 for r in self.results_p2 if r.get("is_real_storefront"))
                    self.lbl_summary_stats.configure(
                        text=f"已质检: {len(self.results_p2)} | AI建议真实: {real_n} | 待核实/异常: {len(self.results_p2) - real_n}"
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
                        "#ef4444" if summary.get("failed") else "#34d399"
                    )
                    messagebox.showinfo("运行报告", self._format_summary(summary))

        except queue.Empty:
            pass
        self.after(100, self._process_queue)

    # ================== 选中行详情 ==================

    def _on_tree_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]

        if self.current_mode == "相同照片比对核验":
            item_id = int(iid.replace("p1_", ""))
            rec = next((r for r in self.results_p1 if r.get("id") == item_id), None) \
                or next((r for r in self.sampled_pairs if r.get("id") == item_id), None)
            if not rec:
                return

            pa, pb = rec.get("path_a"), rec.get("path_b")
            self.current_preview_paths = (pa, pb)
            self.btn_open_native.configure(state="normal" if pa or pb else "disabled")
            self._render_thumbnail(pa, self.lbl_img_a, self.lbl_t_a, f"照片A: {rec.get('photo_a', '')}")
            self._render_thumbnail(pb, self.lbl_img_b, self.lbl_t_b, f"照片B: {rec.get('photo_b', '')}")

            is_same = rec.get("is_same_photo", rec.get("is_same"))
            conf = rec.get("confidence", 0.0)
            color = "#ef4444" if is_same else "#10b981"
            self.lbl_verdict.configure(
                text=f"AI 建议结论: {rec.get('verdict', '待抽检')} （置信度仅参考: {conf:.2f}，最终结论以结论等级为准）",
                text_color=color
            )
            self.lbl_multi_metrics.configure(text=(
                f"【多维证据】 店招1: [{rec.get('signboard_text_1', '-')}] | 店招2: [{rec.get('signboard_text_2', '-')}] | "
                f"匹配: {rec.get('signboard_match', '-')} | 场景重合: {rec.get('scene_similarity', 0.0):.2f} | "
                f"遮蔽比例: {rec.get('mask_pct_used', '-')} | 调用: {rec.get('call_status', '-')}"
            ))

            self.txt_reason.configure(state="normal")
            self.txt_reason.delete("1.0", "end")
            self.txt_reason.insert("1.0", (
                f"【AI 多维证据链建议依据】:\n{rec.get('reason', '等待开始抽检...')}\n\n"
                f"● 执行状态: {rec.get('call_status', '-')} | 重试次数: {rec.get('重试次数', 0)} | 阶段: {rec.get('执行阶段', '-')}\n"
                f"● 模型版本: {rec.get('model_version', '-')} | prompt版本: {rec.get('prompt_version', '-')}\n"
                f"● 图1任务: {rec.get('task_a')} | 路径: {pa or '未找到'}\n"
                f"● 图2任务: {rec.get('task_b')} | 路径: {pb or '未找到'}"
            ))
            self.txt_reason.configure(state="disabled")

        else:
            item_id = int(iid.replace("p2_", ""))
            rec = next((r for r in self.results_p2 if r.get("id") == item_id), None) \
                or next((r for r in self.sampled_storefront_items if r.get("id") == item_id), None)
            if not rec:
                return

            p = rec.get("path")
            self.current_preview_paths = (p, None)
            self.btn_open_native.configure(state="normal" if p else "disabled")
            self._render_thumbnail(p, self.lbl_img_a, self.lbl_t_a, f"门头照: {rec.get('photo_name', '')}")

            is_real = rec.get("is_real_storefront")
            conf = rec.get("confidence", 0.0)
            self.lbl_verdict.configure(
                text=f"AI 建议: {'真实门头' if is_real else '非真实门头/待核实'} | 风险: {rec.get('risk_type', '-')} | 店名: {rec.get('store_name', '-')}",
                text_color="#10b981" if is_real else "#ef4444"
            )
            self.lbl_multi_metrics.configure(
                text=f"【门头指标】 店名: {rec.get('store_name', '-')} | 风险: {rec.get('risk_type', '-')} | 置信度(仅参考): {conf:.2f} | 调用: {rec.get('call_status', '-')}"
            )
            self.txt_reason.configure(state="normal")
            self.txt_reason.delete("1.0", "end")
            self.txt_reason.insert("1.0", (
                "【门头照质检结果】\n"
                f"● AI 建议: {'真实门头' if is_real else '非真实门头 / 待核实'}\n"
                f"● 风险类型: {rec.get('risk_type', '-')}\n"
                f"● 识别招牌: {rec.get('store_name', '-')}\n"
                f"● 置信度(仅参考): {conf:.2f}\n"
                f"● 执行状态: {rec.get('call_status', '-')} | 重试次数: {rec.get('重试次数', 0)}\n"
                f"● 判定理由:\n{rec.get('reason', '等待开始质检...')}\n\n"
                f"图片路径: {p or '未找到'}"
            ))
            self.txt_reason.configure(state="disabled")

    def _render_thumbnail(self, img_path: Optional[str], target_lbl, title_lbl, title_text: str):
        title_lbl.configure(text=title_text[:30] + ("..." if len(title_text) > 30 else ""))
        if not img_path or not os.path.exists(img_path):
            target_lbl.configure(image=None, text="未找到图片文件")
            return
        try:
            pil_img = Image.open(img_path)
            max_w = 540 if self.current_mode != "相同照片比对核验" else 260
            pil_img.thumbnail((max_w, 240), Image.Resampling.LANCZOS)
            ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=pil_img.size)
            target_lbl.configure(image=ctk_img, text="")
        except Exception as e:
            target_lbl.configure(image=None, text=f"加载失败: {str(e)}")

    def _open_native_photos(self):
        pa, pb = self.current_preview_paths
        for p in (pa, pb):
            if p and os.path.exists(p):
                if sys.platform == "darwin":
                    os.system(f'open "{p}"')
                elif sys.platform.startswith("win"):
                    os.system(f'start "" "{p}"')
                else:
                    os.system(f'xdg-open "{p}"')
