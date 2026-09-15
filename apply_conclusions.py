"""对已有候选/抽检结果 xlsx 批量套用结论体系 v2.1，并导出可运营的三张清单。

用法：

    python3 apply_conclusions.py --input 照片相似候选_100k_7天窗口_地址一致性判断.xlsx

    # 同店同一天连拍也计入问题队列（类型C）
    python3 apply_conclusions.py --input xxx.xlsx --count-same-day-burst

    # 跨期口径收紧到 30 天
    python3 apply_conclusions.py --input xxx.xlsx --window-days 30

输出（默认与输入同目录）：
    1. <原名>_结论分级.xlsx      原表全列保持不变，右侧追加结论体系列 + 概览页
    2. <原名>_确认清单.xlsx      结论等级=确认，按问题类型分表（类型B 优先）
    3. <原名>_高度疑似清单.xlsx  结论等级=高度疑似，交人工复核
    4. <原名>_误报回流清单.xlsx  结论等级=算法误报，用于回流校准阈值

说明：本脚本不调用大模型。表中若已有 AI 抽检列则纳入判定，没有则按
"未抽检"处理，确定性算法（pHash + ORB）仍可独立给出"确认"级结论。
"""

import argparse
import os
from typing import Any, Dict, List, Optional

import pandas as pd

from conclusion_engine import (
    CONCLUSION_COLUMNS,
    LEVEL_CONFIRMED,
    LEVEL_FALSE_POSITIVE,
    LEVEL_SUSPECT,
    TYPE_A,
    TYPE_B,
    TYPE_C,
    ConclusionConfig,
    build_conclusion_columns,
    guess_business_columns,
)

# 表中若已存在这些 AI 列，则复用为 AI 判据
_AI_COLUMN_ALIASES = {
    "verdict": ["AI建议结论", "AI判定结论", "AI定性结论"],
    "is_same_photo": ["AI是否同一底片", "AI是否同一照片", "AI判定是否相同"],
    "confidence": ["AI置信度(仅参考)", "AI置信度"],
    "signboard_text_1": ["AI店招1", "图1店招文字"],
    "signboard_text_2": ["AI店招2", "图2店招文字"],
    "signboard_match": ["AI店招比对", "店招比对结论"],
    "scene_similarity": ["AI场景重合度", "物理场景重合度"],
    "reason": ["AI证据理由", "多维证据详细依据", "AI详细分析原因"],
    "model_version": ["模型版本"],
    "prompt_version": ["prompt版本"],
    "mask_pct_used": ["遮蔽比例"],
    "call_status": ["调用状态"],
    "check_time": ["AI核验时间"],
    "status": ["AI抽检状态"],
}


def _resolve_ai_columns(columns: List[str]) -> Dict[str, str]:
    found: Dict[str, str] = {}
    for key, aliases in _AI_COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in columns:
                found[key] = alias
                break
    return found


def _row_ai(row: Dict[str, Any], ai_cols: Dict[str, str]) -> Optional[Dict[str, Any]]:
    """把原表里已有的 AI 结果列还原成 ai 结果字典；未抽检返回 None。"""
    if not ai_cols:
        return None

    status_col = ai_cols.get("status")
    if status_col and str(row.get(status_col, "")).strip() in ("未抽检", "", "nan"):
        return None

    verdict = str(row.get(ai_cols.get("verdict", ""), "") or "").strip()
    same_raw = str(row.get(ai_cols.get("is_same_photo", ""), "") or "").strip()
    if not verdict and not same_raw:
        return None

    def num(key: str) -> Optional[float]:
        col = ai_cols.get(key)
        if not col:
            return None
        try:
            return float(str(row.get(col, "")).strip())
        except Exception:
            return None

    is_same = None
    if same_raw:
        is_same = same_raw.startswith("是") or same_raw in ("相同", "True", "true")

    ai: Dict[str, Any] = {
        "verdict": verdict,
        "is_same_photo": is_same,
        "confidence": num("confidence"),
        "scene_similarity": num("scene_similarity"),
        "error": False,
    }
    for key in (
        "signboard_text_1",
        "signboard_text_2",
        "signboard_match",
        "reason",
        "model_version",
        "prompt_version",
        "mask_pct_used",
        "call_status",
        "check_time",
    ):
        col = ai_cols.get(key)
        if col:
            val = row.get(col)
            ai[key] = "" if val is None else str(val)

    if "异常" in verdict or str(ai.get("call_status", "")).startswith("http"):
        ai["error"] = True
    return ai


def apply_conclusions(df: pd.DataFrame, config: ConclusionConfig) -> pd.DataFrame:
    columns = [str(c) for c in df.columns]
    mapping = guess_business_columns(columns)
    ai_cols = _resolve_ai_columns(columns)

    # 避免与已有同名列冲突：先移除旧的结论列
    df = df.drop(columns=[c for c in CONCLUSION_COLUMNS if c in df.columns], errors="ignore")

    records: List[Dict[str, Any]] = []
    for _, row in df.iterrows():
        row_dict = row.to_dict()
        records.append(
            build_conclusion_columns(
                row_dict,
                ai=_row_ai(row_dict, ai_cols),
                mapping=mapping,
                config=config,
            )
        )

    add_df = pd.DataFrame(records, columns=CONCLUSION_COLUMNS, index=df.index)
    return pd.concat([df, add_df], axis=1)


def _write_list(df: pd.DataFrame, path: str, split_by_type: bool = False) -> int:
    if df.empty:
        return 0
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        if split_by_type:
            # 类型B 严重度更高，放在前面
            for ptype, sheet in (
                (TYPE_B, "类型B_跨主体套用"),
                (TYPE_A, "类型A_跨期复用"),
                (TYPE_C, "类型C_同次连拍"),
            ):
                sub = df[df["问题类型"] == ptype]
                if not sub.empty:
                    sub.to_excel(writer, sheet_name=sheet, index=False)
            rest = df[~df["问题类型"].isin([TYPE_A, TYPE_B, TYPE_C])]
            if not rest.empty:
                rest.to_excel(writer, sheet_name="其他", index=False)
        else:
            df.to_excel(writer, sheet_name="清单", index=False)
    return len(df)


def main() -> None:
    parser = argparse.ArgumentParser(description="套用结论体系 v2.1 并导出三张运营清单")
    parser.add_argument("--input", required=True, help="输入 xlsx 路径")
    parser.add_argument("--sheet", default=0, help="工作表名或序号，默认第一个")
    parser.add_argument("--outdir", default="", help="输出目录，默认与输入同目录")
    parser.add_argument(
        "--window-days",
        type=float,
        default=7.0,
        help="同商户/同地址视为同一作业周期的天数窗口，默认 7",
    )
    parser.add_argument(
        "--count-same-day-burst",
        action="store_true",
        help="把窗口内的同店重复计为类型C（同次连拍分摊），默认不计入问题队列",
    )
    parser.add_argument("--phash-confirm-max", type=float, default=0.0)
    parser.add_argument("--orb-confirm-min-inliers", type=float, default=30.0)
    args = parser.parse_args()

    if not os.path.exists(args.input):
        raise SystemExit("输入文件不存在: %s" % args.input)

    config = ConclusionConfig(
        same_entity_window_days=args.window_days,
        count_same_day_burst=args.count_same_day_burst,
        phash_confirm_max=args.phash_confirm_max,
        orb_confirm_min_inliers=args.orb_confirm_min_inliers,
    )

    sheet: Any = args.sheet
    try:
        sheet = int(sheet)
    except Exception:
        pass

    print("读取: %s" % args.input)
    df = pd.read_excel(args.input, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]
    print("行数: %d，列数: %d" % (len(df), len(df.columns)))

    out_df = apply_conclusions(df, config)

    outdir = args.outdir or os.path.dirname(os.path.abspath(args.input))
    base = os.path.splitext(os.path.basename(args.input))[0]
    os.makedirs(outdir, exist_ok=True)

    counts = out_df["结论等级"].value_counts()
    type_counts = out_df["问题类型"].value_counts()
    summary = pd.DataFrame(
        [{"维度": "结论等级", "取值": k, "数量": int(v)} for k, v in counts.items()]
        + [{"维度": "问题类型", "取值": k, "数量": int(v)} for k, v in type_counts.items()]
        + [
            {"维度": "配置", "取值": "同周期窗口(天)", "数量": args.window_days},
            {
                "维度": "配置",
                "取值": "类型C是否计入问题队列",
                "数量": 1 if args.count_same_day_burst else 0,
            },
        ]
    )

    main_out = os.path.join(outdir, "%s_结论分级.xlsx" % base)
    with pd.ExcelWriter(main_out, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="结论概览", index=False)
        out_df.to_excel(writer, sheet_name="明细(含结论体系)", index=False)
    print("已输出: %s" % main_out)

    n1 = _write_list(
        out_df[out_df["结论等级"] == LEVEL_CONFIRMED],
        os.path.join(outdir, "%s_确认清单.xlsx" % base),
        split_by_type=True,
    )
    n2 = _write_list(
        out_df[out_df["结论等级"] == LEVEL_SUSPECT],
        os.path.join(outdir, "%s_高度疑似清单.xlsx" % base),
        split_by_type=True,
    )
    n3 = _write_list(
        out_df[out_df["结论等级"] == LEVEL_FALSE_POSITIVE],
        os.path.join(outdir, "%s_误报回流清单.xlsx" % base),
    )

    print("确认: %d 对 | 高度疑似: %d 对 | 算法误报: %d 对" % (n1, n2, n3))
    print("提示：所有等级都不是处罚结论，需在「人工终审结论」列完成终审后才可进入处理流程。")


if __name__ == "__main__":
    main()
