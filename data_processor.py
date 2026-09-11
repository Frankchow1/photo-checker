import os
import re
from typing import List, Dict, Any, Tuple, Optional
import pandas as pd
from datetime import datetime

SUPPORTED_IMG_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".jfif", ".tif", ".tiff"}

class PhotoIndex:
    """照片文件快速索引器"""
    def __init__(self, root_dir: str):
        self.root_dir = root_dir
        self.by_filename: Dict[str, str] = {}
        self.by_stem: Dict[str, str] = {}
        self._build_index()

    def _build_index(self):
        if not self.root_dir or not os.path.exists(self.root_dir):
            return
        for root, _, files in os.walk(self.root_dir):
            for file in files:
                ext = os.path.splitext(file)[1].lower()
                if ext in SUPPORTED_IMG_EXTS:
                    full_path = os.path.join(root, file)
                    lower_name = file.lower()
                    self.by_filename[lower_name] = full_path
                    stem = os.path.splitext(lower_name)[0]
                    self.by_stem[stem] = full_path

    def find_photo(self, photo_str: str) -> Optional[str]:
        if not photo_str:
            return None
        photo_str = str(photo_str).strip()
        
        # 1. 绝对路径且真实存在
        if os.path.isabs(photo_str) and os.path.exists(photo_str):
            return photo_str

        # 2. 相对路径拼装
        if self.root_dir:
            rel_candidate = os.path.join(self.root_dir, photo_str)
            if os.path.exists(rel_candidate):
                return rel_candidate

        # 3. 提取文件名在当前索引库中匹配
        base_name = os.path.basename(photo_str).lower()
        if base_name in self.by_filename:
            return self.by_filename[base_name]

        # 4. 去除后缀匹配
        stem_name = os.path.splitext(base_name)[0]
        if stem_name in self.by_stem:
            return self.by_stem[stem_name]

        # 5. 补充后缀匹配
        for ext in SUPPORTED_IMG_EXTS:
            test_name = f"{stem_name}{ext}"
            if test_name in self.by_filename:
                return self.by_filename[test_name]

        return None


def collect_images_from_dir(folder_path: str) -> List[Dict[str, Any]]:
    items = []
    if not folder_path or not os.path.exists(folder_path):
        return items

    for root, _, files in os.walk(folder_path):
        for f in sorted(files):
            ext = os.path.splitext(f)[1].lower()
            if ext in SUPPORTED_IMG_EXTS:
                full_p = os.path.join(root, f)
                items.append({
                    "id": len(items) + 1,
                    "task_id": f"IMG_{len(items)+1:04d}",
                    "photo_name": f,
                    "path": full_p
                })
    return items


class ExcelTaskParser:
    def __init__(self, excel_path: str):
        self.excel_path = excel_path
        self.df: Optional[pd.DataFrame] = None
        self.columns: List[str] = []
        self._load_excel()

    def _load_excel(self):
        try:
            self.df = pd.read_excel(self.excel_path)
            self.df.columns = [str(c).strip() for c in self.df.columns]
            self.columns = list(self.df.columns)
        except Exception as e:
            raise RuntimeError(f"读取 Excel 失败: {str(e)}")

    def guess_columns(self) -> Dict[str, Optional[str]]:
        mapping = {
            "task_a": None,
            "task_b": None,
            "photo_a": None,
            "photo_b": None,
            "path_a_col": None,
            "path_b_col": None,
            "group_id": None,
            "task_col": None,
            "photo_col": None
        }
        lower_cols = {c.lower(): c for c in self.columns}

        for k, v in lower_cols.items():
            if k in ["task_no_1", "任务号1", "任务1", "task1", "task_a"]:
                mapping["task_a"] = v
            elif k in ["task_no_2", "任务号2", "任务2", "task2", "task_b"]:
                mapping["task_b"] = v
            elif k in ["文件名1", "照片1", "图1", "photo1", "image1", "pic1", "photo_a"]:
                mapping["photo_a"] = v
            elif k in ["文件名2", "照片2", "图2", "photo2", "image2", "pic2", "photo_b"]:
                mapping["photo_b"] = v
            elif k in ["完整路径1", "路径1", "path1", "path_a"]:
                mapping["path_a_col"] = v
            elif k in ["完整路径2", "路径2", "path2", "path_b"]:
                mapping["path_b_col"] = v
            elif k in ["候选组", "组号", "group", "group_id", "重复组", "hash"]:
                mapping["group_id"] = v

        for k, v in lower_cols.items():
            if not mapping["task_a"] and any(w in k for w in ["任务1", "任务a", "task1", "task_a", "task_no_1"]):
                mapping["task_a"] = v
            elif not mapping["task_b"] and any(w in k for w in ["任务2", "任务b", "task2", "task_b", "task_no_2"]):
                mapping["task_b"] = v
            elif not mapping["photo_a"] and any(w in k for w in ["照片1", "图1", "photo1", "文件名1"]):
                mapping["photo_a"] = v
            elif not mapping["photo_b"] and any(w in k for w in ["照片2", "图2", "photo2", "文件名2"]):
                mapping["photo_b"] = v
            elif not mapping["group_id"] and any(w in k for w in ["候选组", "组", "group", "hash"]):
                mapping["group_id"] = v
            elif not mapping["task_col"] and any(w in k for w in ["任务号", "任务编号", "task_no", "task"]):
                mapping["task_col"] = v
            elif not mapping["photo_col"] and any(w in k for w in ["照片", "图片", "文件名", "photo", "image"]):
                mapping["photo_col"] = v

        return mapping

    def parse_pairs(self, mode: str = "auto", photo_indexer: Optional[PhotoIndex] = None, custom_mapping: Optional[Dict[str, str]] = None) -> List[Dict[str, Any]]:
        if self.df is None or self.df.empty:
            return []

        guessed = self.guess_columns()
        if custom_mapping:
            guessed.update({k: v for k, v in custom_mapping.items() if v})

        pairs: List[Dict[str, Any]] = []
        is_paired_row = False
        if mode == "paired" or (mode == "auto" and (guessed["task_a"] and guessed["task_b"] or guessed["photo_a"] and guessed["photo_b"])):
            is_paired_row = True

        if is_paired_row:
            col_t_a = guessed["task_a"] or guessed["task_col"]
            col_t_b = guessed["task_b"]
            col_p_a = guessed["photo_a"] or guessed["photo_col"]
            col_p_b = guessed["photo_b"]
            col_path_a = guessed.get("path_a_col")
            col_path_b = guessed.get("path_b_col")

            for idx, row in self.df.iterrows():
                val_t_a = str(row[col_t_a]).strip() if col_t_a and pd.notna(row[col_t_a]) else f"Task_{idx}_A"
                val_t_b = str(row[col_t_b]).strip() if col_t_b and pd.notna(row[col_t_b]) else f"Task_{idx}_B"
                val_p_a = str(row[col_p_a]).strip() if col_p_a and pd.notna(row[col_p_a]) else ""
                val_p_b = str(row[col_p_b]).strip() if col_p_b and pd.notna(row[col_p_b]) else ""

                path_a = None
                if col_path_a and pd.notna(row[col_path_a]):
                    cand_a = str(row[col_path_a]).strip()
                    if os.path.exists(cand_a):
                        path_a = cand_a

                path_b = None
                if col_path_b and pd.notna(row[col_path_b]):
                    cand_b = str(row[col_path_b]).strip()
                    if os.path.exists(cand_b):
                        path_b = cand_b

                if not path_a and photo_indexer:
                    path_a = photo_indexer.find_photo(val_p_a)
                if not path_b and photo_indexer:
                    path_b = photo_indexer.find_photo(val_p_b)

                is_ready = bool(path_a and path_b)

                pairs.append({
                    "id": len(pairs) + 1,
                    "group": str(row[guessed["group_id"]]) if guessed["group_id"] and pd.notna(row[guessed["group_id"]]) else f"Group_{idx+1}",
                    "task_a": val_t_a,
                    "photo_a": val_p_a,
                    "path_a": path_a,
                    "task_b": val_t_b,
                    "photo_b": val_p_b,
                    "path_b": path_b,
                    "is_ready": is_ready,
                    "raw_index": idx
                })
        else:
            group_col = guessed["group_id"]
            task_col = guessed["task_col"] or (self.columns[0] if self.columns else None)
            photo_col = guessed["photo_col"] or (self.columns[1] if len(self.columns) > 1 else None)

            if group_col and group_col in self.df.columns:
                grouped = self.df.groupby(group_col)
                for g_val, sub_df in grouped:
                    if len(sub_df) >= 2:
                        records = sub_df.to_dict("records")
                        for i in range(len(records) - 1):
                            r1 = records[i]
                            r2 = records[i+1]
                            val_t_a = str(r1.get(task_col, f"T_{i}")).strip()
                            val_t_b = str(r2.get(task_col, f"T_{i+1}")).strip()
                            val_p_a = str(r1.get(photo_col, "")).strip()
                            val_p_b = str(r2.get(photo_col, "")).strip()

                            path_a = photo_indexer.find_photo(val_p_a) if photo_indexer else None
                            path_b = photo_indexer.find_photo(val_p_b) if photo_indexer else None

                            pairs.append({
                                "id": len(pairs) + 1,
                                "group": str(g_val),
                                "task_a": val_t_a,
                                "photo_a": val_p_a,
                                "path_a": path_a,
                                "task_b": val_t_b,
                                "photo_b": val_p_b,
                                "path_b": path_b,
                                "is_ready": bool(path_a and path_b),
                                "raw_index": None
                            })
            else:
                for i in range(0, len(self.df) - 1, 2):
                    r1 = self.df.iloc[i]
                    r2 = self.df.iloc[i+1]
                    val_t_a = str(r1[task_col]).strip() if task_col else f"T_{i}"
                    val_t_b = str(r2[task_col]).strip() if task_col else f"T_{i+1}"
                    val_p_a = str(r1[photo_col]).strip() if photo_col else ""
                    val_p_b = str(r2[photo_col]).strip() if photo_col else ""

                    path_a = photo_indexer.find_photo(val_p_a) if photo_indexer else None
                    path_b = photo_indexer.find_photo(val_p_b) if photo_indexer else None

                    pairs.append({
                        "id": len(pairs) + 1,
                        "group": f"相邻对_{i//2 + 1}",
                        "task_a": val_t_a,
                        "photo_a": val_p_a,
                        "path_a": path_a,
                        "task_b": val_t_b,
                        "photo_b": val_p_b,
                        "path_b": path_b,
                        "is_ready": bool(path_a and path_b),
                        "raw_index": i
                    })

        return pairs


def export_results_to_excel(results: List[Dict[str, Any]], output_path: str):
    """导出相同照片多维证据链抽检结果至独立 Excel 报告"""
    data = []
    for r in results:
        is_same = r.get("is_same_photo", r.get("is_same", False))
        data.append({
            "抽检序号": r.get("id"),
            "分组/组号": r.get("group", ""),
            "任务A编号": r.get("task_a", ""),
            "任务B编号": r.get("task_b", ""),
            "AI定性结论": r.get("verdict", "未定性"),
            "AI判定是否同一照片": "是 (相同底片)" if is_same else "否 (不同底片)",
            "图1店招文字": r.get("signboard_text_1", "无店招"),
            "图2店招文字": r.get("signboard_text_2", "无店招"),
            "店招比对结论": r.get("signboard_match", "无店招不适用"),
            "物理场景重合度": f"{r.get('scene_similarity', 0.0):.2f}",
            "AI置信度": f"{r.get('confidence', 0.0):.2f}",
            "多维证据详细依据": r.get("reason", ""),
            "照片A名称": r.get("photo_a", ""),
            "照片A路径": r.get("path_a", ""),
            "照片B名称": r.get("photo_b", ""),
            "照片B路径": r.get("path_b", ""),
            "核验时间": r.get("check_time", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        })

    df_out = pd.DataFrame(data)
    total = len(results)
    same_count = sum(1 for r in results if r.get("is_same_photo", r.get("is_same", False)))
    diff_count = total - same_count
    accuracy_rate = (same_count / total * 100) if total > 0 else 0

    summary_data = [
        {"指标": "抽检总对数", "数值": total},
        {"指标": "AI判定为同一照片 (违规复用)", "数值": same_count},
        {"指标": "AI判定为不同照片 (重拍/误判)", "数值": diff_count},
        {"指标": "重复违规吻合率", "数值": f"{accuracy_rate:.1f}%"},
        {"指标": "报告生成时间", "数值": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    ]
    df_summary = pd.DataFrame(summary_data)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_summary.to_excel(writer, sheet_name="抽检统计概览", index=False)
        df_out.to_excel(writer, sheet_name="抽检详细记录", index=False)


def export_annotated_original_excel_p1(original_excel_path: str, results: List[Dict[str, Any]], output_path: str) -> str:
    """在原始 Excel 文件副本中追加完整的多维证据链与店招 AI 抽检结果列"""
    if not os.path.exists(original_excel_path):
        raise FileNotFoundError(f"原 Excel 文件不存在: {original_excel_path}")

    df = pd.read_excel(original_excel_path)

    # 初始化新列
    df["AI抽检状态"] = "未抽检"
    df["AI定性结论"] = ""
    df["AI是否同一照片"] = ""
    df["图1店招文字"] = ""
    df["图2店招文字"] = ""
    df["店招比对结论"] = ""
    df["物理场景重合度"] = ""
    df["AI置信度"] = ""
    df["多维证据详细依据"] = ""
    df["AI核验时间"] = ""

    results_by_idx = {}
    for r in results:
        idx = r.get("raw_index")
        if idx is not None and 0 <= idx < len(df):
            results_by_idx[idx] = r

    for idx, r in results_by_idx.items():
        df.loc[idx, "AI抽检状态"] = "已抽检"
        df.loc[idx, "AI定性结论"] = r.get("verdict", "")
        is_same = r.get("is_same_photo", r.get("is_same", False))
        df.loc[idx, "AI是否同一照片"] = "是 (同一底片)" if is_same else "否 (不同底片)"
        df.loc[idx, "图1店招文字"] = r.get("signboard_text_1", "无店招")
        df.loc[idx, "图2店招文字"] = r.get("signboard_text_2", "无店招")
        df.loc[idx, "店招比对结论"] = r.get("signboard_match", "无店招不适用")
        df.loc[idx, "物理场景重合度"] = f"{r.get('scene_similarity', 0.0):.2f}"
        df.loc[idx, "AI置信度"] = f"{r.get('confidence', 0.0):.2f}"
        df.loc[idx, "多维证据详细依据"] = r.get("reason", "")
        df.loc[idx, "AI核验时间"] = r.get("check_time", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="数据明细(含AI多维质检)", index=False)

    return output_path


def export_storefront_results_to_excel(results: List[Dict[str, Any]], output_path: str):
    """导出真实门头照质检结果至 Excel"""
    data = []
    for r in results:
        data.append({
            "序号": r.get("id"),
            "任务/标识号": r.get("task_id", ""),
            "照片名称": r.get("photo_name", ""),
            "照片完整路径": r.get("path", ""),
            "是否真实门头": "真实门头照" if r.get("is_real_storefront") else "非真实门头/违规",
            "风险/异常类型": r.get("risk_type", "未知"),
            "识别门头店名": r.get("store_name", ""),
            "判定置信度": f"{r.get('confidence', 0.0):.2f}",
            "AI详细核验原因": r.get("reason", ""),
            "质检时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })

    df_out = pd.DataFrame(data)
    total = len(results)
    real_count = sum(1 for r in results if r.get("is_real_storefront"))
    fake_count = total - real_count
    real_rate = (real_count / total * 100) if total > 0 else 0

    summary_data = [
        {"指标": "质检照片总数", "数值": total},
        {"指标": "真实门头合格数", "数值": real_count},
        {"指标": "非真实/异常数", "数值": fake_count},
        {"指标": "门头合规率", "数值": f"{real_rate:.1f}%"},
        {"指标": "报告生成时间", "数值": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    ]
    df_summary = pd.DataFrame(summary_data)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_summary.to_excel(writer, sheet_name="门头性质检概览", index=False)
        df_out.to_excel(writer, sheet_name="门头照质检明细", index=False)


def export_annotated_original_excel_p2(original_excel_path: str, results: List[Dict[str, Any]], output_path: str) -> str:
    """在原始 Excel 文件副本中追加真实门头照单图质检结果列"""
    if not os.path.exists(original_excel_path):
        raise FileNotFoundError(f"原 Excel 文件不存在: {original_excel_path}")

    df = pd.read_excel(original_excel_path)
    df["AI门头质检状态"] = "未质检"
    df["AI是否真实门头"] = ""
    df["AI门头风险类型"] = ""
    df["AI识别店招名称"] = ""
    df["AI置信度"] = ""
    df["AI核验详细原因"] = ""
    df["AI质检时间"] = ""

    results_by_idx = {}
    for r in results:
        idx = r.get("raw_index")
        if idx is not None and 0 <= idx < len(df):
            results_by_idx[idx] = r

    for idx, r in results_by_idx.items():
        df.loc[idx, "AI门头质检状态"] = "已质检"
        df.loc[idx, "AI是否真实门头"] = "真实门头照" if r.get("is_real_storefront") else "非真实门头/违规"
        df.loc[idx, "AI门头风险类型"] = r.get("risk_type", "未知")
        df.loc[idx, "AI识别店招名称"] = r.get("store_name", "")
        df.loc[idx, "AI置信度"] = f"{r.get('confidence', 0.0):.2f}"
        df.loc[idx, "AI核验详细原因"] = r.get("reason", "")
        df.loc[idx, "AI质检时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="明细(含真实门头质检)", index=False)

    return output_path

