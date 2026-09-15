"""结论体系引擎 v2.1

设计原则（与业务评审结论一致）：

1. 「问题类型」与「结论等级」必须拆成两列。
   - 问题类型回答"是哪种作弊形态"：类型A 跨期复用旧图 / 类型B 跨主体套用 / 类型C 同次连拍分摊。
   - 结论等级回答"有多确定"：确认 / 高度疑似 / 待核实 / 正常 / 算法误报 / 未抽检。
2. 「确认」级别只能由确定性算法（pHash 距离 + ORB 内点）联合业务条件给出，
   大模型只能给出「高度疑似」，永远不单独定性为确认。
3. 大模型自报 confidence 未经校准，仅作参考列，不参与等级判定。
4. 结论 != 处罚。所有等级都保留「人工终审结论」空列，只有人工终审后才可进入处理流程。

所有阈值集中在 ConclusionConfig，便于按业务口径调整。
"""

import re
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

ENGINE_VERSION = "2.1.0"
RULE_VERSION = "conclusion-rules-2.1"

# ---------------- 问题类型 ----------------
TYPE_A = "类型A-跨期复用旧图"
TYPE_B = "类型B-跨主体套用"
TYPE_C = "类型C-同次连拍分摊"
TYPE_SAME_TASK = "无问题-同任务"
TYPE_SAME_PERIOD = "无问题-同店同周期"
TYPE_UNKNOWN = "待定-信息不足"

# ---------------- 结论等级 ----------------
LEVEL_CONFIRMED = "确认"
LEVEL_SUSPECT = "高度疑似"
LEVEL_PENDING = "待核实"
LEVEL_NORMAL = "正常"
LEVEL_FALSE_POSITIVE = "算法误报"
LEVEL_NOT_SAMPLED = "未抽检"

# 结论体系统一输出列（顺序即导出顺序）
CONCLUSION_COLUMNS: List[str] = [
    "问题类型",
    "结论等级",
    "结论依据",
    "机器判据",
    "AI抽检状态",
    "AI建议结论",
    "AI是否同一底片",
    "AI置信度(仅参考)",
    "AI店招1",
    "AI店招2",
    "AI店招比对",
    "AI场景重合度",
    "AI证据理由",
    "人工终审结论",
    "终审人",
    "终审时间",
    "模型版本",
    "prompt版本",
    "遮蔽比例",
    "调用状态",
    "AI核验时间",
    "判定规则版本",
]


class ConclusionConfig(object):
    """结论判定阈值集合。"""

    def __init__(
        self,
        same_entity_window_days: float = 7.0,
        count_same_day_burst: bool = False,
        phash_confirm_max: float = 0.0,
        orb_confirm_min_inliers: float = 30.0,
        phash_suspect_max: float = 8.0,
    ):
        # 同商户/同地址在该窗口内视为同一次作业周期，不算跨期复用
        self.same_entity_window_days = same_entity_window_days
        # 是否把"同一次到店连拍分摊到多个任务"计入问题队列（类型C）
        self.count_same_day_burst = count_same_day_burst
        # pHash 汉明距离 <= 该值才允许进入"确认"
        self.phash_confirm_max = phash_confirm_max
        # ORB 内点数 >= 该值才允许进入"确认"（缺失时不阻断）
        self.orb_confirm_min_inliers = orb_confirm_min_inliers
        # 超过该 pHash 距离的对，不应出现在候选池，出现则标记待核实
        self.phash_suspect_max = phash_suspect_max


DEFAULT_CONFIG = ConclusionConfig()

_SPACES = re.compile(r"[\s\u3000]+")
_EMPTY_TOKENS = {"", "nan", "none", "nat", "-", "null", "信息不足"}
_DT_FORMATS = (
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y/%m/%d %H:%M:%S",
    "%Y/%m/%d %H:%M",
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%Y%m%d",
)


def _txt(value: Any) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    if s.lower() in _EMPTY_TOKENS:
        return ""
    return _SPACES.sub("", s)


def _num(value: Any) -> Optional[float]:
    if value is None:
        return None
    s = str(value).strip()
    if s.lower() in _EMPTY_TOKENS:
        return None
    if s in ("True", "true", "是"):
        return None
    try:
        return float(s)
    except Exception:
        m = re.search(r"-?\d+(\.\d+)?", s)
        return float(m.group(0)) if m else None


def _dt(value: Any) -> Optional[datetime]:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    s = str(value).strip()
    if s.lower() in _EMPTY_TOKENS:
        return None
    s = s.replace("T", " ").split(".")[0]
    for fmt in _DT_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue
    return None


def _suffix_of(lower_name: str) -> Optional[str]:
    if re.search(r"(_2|2)$", lower_name) or "_2" in lower_name or "图2" in lower_name:
        return "2"
    if re.search(r"(_1|1)$", lower_name) or "_1" in lower_name or "图1" in lower_name:
        return "1"
    return None


def guess_business_columns(columns: List[str]) -> Dict[str, Optional[str]]:
    """从候选表列名中识别业务判据列（商户、地址、完成时间、哈希距离等）。"""
    mapping: Dict[str, Optional[str]] = {
        "task_1": None,
        "task_2": None,
        "merchant_1": None,
        "merchant_2": None,
        "address_1": None,
        "address_2": None,
        "finish_1": None,
        "finish_2": None,
        "phash": None,
        "orb": None,
        "addr_same": None,
        "day_diff": None,
        "candidate_type": None,
    }

    def put(key: str, col: str) -> None:
        if key in mapping and not mapping[key]:
            mapping[key] = col

    for col in columns:
        raw = str(col).strip()
        lc = raw.lower()
        sfx = _suffix_of(lc)

        if "地址是否一致" in raw:
            put("addr_same", raw)
            continue
        if "时间差" in raw:
            put("day_diff", raw)
            continue
        if "候选业务类型" in raw:
            put("candidate_type", raw)
            continue
        if ("phash" in lc or "汉明" in raw or "哈希距离" in raw) and "dhash" not in lc:
            put("phash", raw)
            continue
        if "orb" in lc and ("内点" in raw or "inlier" in lc):
            put("orb", raw)
            continue

        if not sfx:
            continue

        # 商户名称（排除商户编号，编号口径已废弃）
        if ("merchant" in lc or "商户" in raw or "店名" in raw) and not any(
            w in lc for w in ("_no", "no_", "编号", "商户号")
        ):
            put("merchant_" + sfx, raw)
            continue
        # 原始地址（排除地址工具产生的标准化/核心列）
        if ("address" in lc or "地址" in raw) and not any(
            w in raw for w in ("标准化", "核心", "置信度", "依据")
        ):
            put("address_" + sfx, raw)
            continue
        if any(w in raw for w in ("完成时间", "提交时间", "作业时间", "审核时间")) or "finish" in lc:
            put("finish_" + sfx, raw)
            continue
        if any(w in raw for w in ("任务号", "任务")) or "task" in lc:
            put("task_" + sfx, raw)
            continue

    return mapping


def extract_context(
    row: Dict[str, Any],
    mapping: Optional[Dict[str, Optional[str]]] = None,
    config: Optional[ConclusionConfig] = None,
) -> Dict[str, Any]:
    """把一行候选对抽取成结构化业务判据。

    注意：时间一律取结构化完成时间字段，绝不读取图内水印时间。
    """
    cfg = config or DEFAULT_CONFIG
    mp = mapping or guess_business_columns(list(row.keys()))

    def val(key: str) -> Any:
        col = mp.get(key)
        return row.get(col) if col else None

    task_1, task_2 = _txt(val("task_1")), _txt(val("task_2"))
    merchant_1, merchant_2 = _txt(val("merchant_1")), _txt(val("merchant_2"))
    address_1, address_2 = _txt(val("address_1")), _txt(val("address_2"))

    merchant_same: Optional[bool] = None
    if merchant_1 and merchant_2:
        merchant_same = merchant_1 == merchant_2

    # 优先采用地址一致性工具的结论，其次退回字符串相等
    address_same: Optional[bool] = None
    addr_flag = str(val("addr_same") or "").strip()
    if addr_flag:
        if addr_flag in ("一致", "疑似一致"):
            address_same = True
        elif addr_flag == "不一致":
            address_same = False
    elif address_1 and address_2:
        address_same = address_1 == address_2

    day_diff = _num(val("day_diff"))
    if day_diff is None:
        d1, d2 = _dt(val("finish_1")), _dt(val("finish_2"))
        if d1 and d2:
            day_diff = abs((d1 - d2).total_seconds()) / 86400.0
    if day_diff is not None:
        day_diff = abs(day_diff)

    return {
        "task_1": task_1,
        "task_2": task_2,
        "task_same": bool(task_1 and task_2 and task_1 == task_2),
        "merchant_1": merchant_1,
        "merchant_2": merchant_2,
        "merchant_same": merchant_same,
        "address_1": address_1,
        "address_2": address_2,
        "address_same": address_same,
        "day_diff": day_diff,
        "phash": _num(val("phash")),
        "orb": _num(val("orb")),
        "candidate_type": _txt(val("candidate_type")),
        "window_days": cfg.same_entity_window_days,
    }


def classify_problem_type(
    ctx: Dict[str, Any], config: Optional[ConclusionConfig] = None
) -> Tuple[str, str]:
    """判断问题类型，返回 (类型, 依据说明)。"""
    cfg = config or DEFAULT_CONFIG
    window = cfg.same_entity_window_days

    if ctx.get("task_same"):
        return TYPE_SAME_TASK, "同一任务内的两张照片，不纳入违规判定"

    same_entity = ctx.get("merchant_same") is True or ctx.get("address_same") is True
    if same_entity:
        basis = "同商户" if ctx.get("merchant_same") is True else "同地址"
        day_diff = ctx.get("day_diff")
        if day_diff is None:
            return TYPE_UNKNOWN, "%s，但缺少结构化完成时间，无法判断是否跨期" % basis
        if day_diff <= window:
            if cfg.count_same_day_burst:
                return TYPE_C, "%s且完成时间相差 %.1f 天（<=%.0f 天），疑似同一次到店连拍分摊到多个任务" % (
                    basis,
                    day_diff,
                    window,
                )
            return TYPE_SAME_PERIOD, "%s且完成时间相差 %.1f 天（<=%.0f 天），视为同一次作业周期内的正常重复" % (
                basis,
                day_diff,
                window,
            )
        return TYPE_A, "%s且完成时间相差 %.1f 天（>%.0f 天），属跨期复用旧图" % (basis, day_diff, window)

    if ctx.get("merchant_same") is False and ctx.get("address_same") is False:
        return TYPE_B, "商户名称不同且地址判定不一致，照片脱离了拍摄对象，属跨主体套用"

    if ctx.get("merchant_same") is False and ctx.get("address_same") is None:
        return TYPE_UNKNOWN, "商户不同，但地址一致性未判定，无法区分跨主体套用与同址档口"

    return TYPE_UNKNOWN, "商户与地址信息不足，无法归类"


def machine_evidence(
    ctx: Dict[str, Any], config: Optional[ConclusionConfig] = None
) -> Tuple[bool, str]:
    """确定性算法判据：只有它才能支撑"确认"级结论。"""
    cfg = config or DEFAULT_CONFIG
    phash, orb = ctx.get("phash"), ctx.get("orb")

    if phash is None:
        return False, "缺少 pHash 距离，无法给出确定性判据"

    orb_txt = "ORB内点=%.0f" % orb if orb is not None else "ORB未计算"
    if phash <= cfg.phash_confirm_max and (
        orb is None or orb >= cfg.orb_confirm_min_inliers
    ):
        return True, "pHash距离=%.0f（<=%.0f）且 %s，确定性算法判定同一底片" % (
            phash,
            cfg.phash_confirm_max,
            orb_txt,
        )
    return False, "pHash距离=%.0f，%s，未达确定性同一底片门槛" % (phash, orb_txt)


def _ai_state(ai: Optional[Dict[str, Any]]) -> str:
    if not ai:
        return "none"
    if ai.get("error"):
        return "error"
    if ai.get("is_same_photo") is None and ai.get("is_same") is None:
        return "error"
    return "ok"


def grade_conclusion(
    ctx: Dict[str, Any],
    problem_type: str,
    ai: Optional[Dict[str, Any]] = None,
    config: Optional[ConclusionConfig] = None,
) -> Tuple[str, str]:
    """判断结论等级，返回 (等级, 依据说明)。

    规则要点：
    - 「确认」只由确定性算法 + 业务条件给出，不采信模型自报置信度；
    - 大模型最高只能给到「高度疑似」；
    - 模型与机器判据冲突、调用失败、信息不足，一律「待核实」。
    """
    cfg = config or DEFAULT_CONFIG
    machine_ok, machine_reason = machine_evidence(ctx, cfg)
    state = _ai_state(ai)
    ai = ai or {}

    if problem_type in (TYPE_SAME_TASK, TYPE_SAME_PERIOD):
        return LEVEL_NORMAL, "业务条件已排除：%s" % problem_type

    is_problem_type = problem_type in (TYPE_A, TYPE_B, TYPE_C)

    if machine_ok and is_problem_type:
        return LEVEL_CONFIRMED, "%s；业务条件成立（%s）。建议按抽检比例复核而非全量人工" % (
            machine_reason,
            problem_type,
        )

    if state == "none":
        return LEVEL_NOT_SAMPLED, "尚未进入 AI 抽检队列；%s" % machine_reason
    if state == "error":
        return LEVEL_PENDING, "AI 调用失败或返回不合规（%s），必须人工核实" % (
            ai.get("call_status") or "unknown"
        )

    verdict = str(ai.get("verdict") or ai.get("suggested_verdict") or "")
    is_same = bool(ai.get("is_same_photo", ai.get("is_same")))

    if "误判" in verdict:
        return LEVEL_FALSE_POSITIVE, "AI 判定为不同商户或不同场景，应回流校准哈希阈值与地址规则"

    if is_same:
        if not is_problem_type:
            return LEVEL_PENDING, "AI 判定同一底片，但业务归类为「%s」，需人工确认归属" % problem_type
        extra = ""
        phash = ctx.get("phash")
        if phash is not None and phash > cfg.phash_suspect_max:
            extra = "；注意 pHash 距离 %.0f 超出候选阈值 %.0f" % (phash, cfg.phash_suspect_max)
        return LEVEL_SUSPECT, "AI 判定同一底片且业务条件成立（%s）。模型结论不作定性，须人工终审%s" % (
            problem_type,
            extra,
        )

    if "合规" in verdict:
        if machine_ok:
            return LEVEL_PENDING, "确定性算法判同一底片但 AI 判为现场重拍，两者冲突，需人工核实"
        return LEVEL_NORMAL, "AI 判定为现场真实重拍（存在机位/光影物理差异）"

    return LEVEL_PENDING, "AI 未给出明确结论，需人工核实"


def build_conclusion_columns(
    row: Dict[str, Any],
    ai: Optional[Dict[str, Any]] = None,
    mapping: Optional[Dict[str, Optional[str]]] = None,
    config: Optional[ConclusionConfig] = None,
) -> Dict[str, Any]:
    """生成一行的结论体系列（含可复现与审计字段）。"""
    cfg = config or DEFAULT_CONFIG
    ctx = extract_context(row, mapping, cfg)
    problem_type, type_reason = classify_problem_type(ctx, cfg)
    level, level_reason = grade_conclusion(ctx, problem_type, ai, cfg)
    machine_ok, machine_reason = machine_evidence(ctx, cfg)
    state = _ai_state(ai)
    ai = ai or {}

    confidence = ai.get("confidence")
    scene = ai.get("scene_similarity")

    return {
        "问题类型": problem_type,
        "结论等级": level,
        "结论依据": "%s；%s" % (type_reason, level_reason),
        "机器判据": machine_reason + ("（满足确认门槛）" if machine_ok else ""),
        "AI抽检状态": "未抽检" if state == "none" else "已抽检",
        "AI建议结论": ai.get("verdict", ai.get("suggested_verdict", "")) or "",
        "AI是否同一底片": ""
        if state == "none"
        else ("是" if bool(ai.get("is_same_photo", ai.get("is_same"))) else "否"),
        "AI置信度(仅参考)": "" if confidence is None else "%.2f" % float(confidence),
        "AI店招1": ai.get("signboard_text_1", "") or "",
        "AI店招2": ai.get("signboard_text_2", "") or "",
        "AI店招比对": ai.get("signboard_match", "") or "",
        "AI场景重合度": "" if scene is None else "%.2f" % float(scene),
        "AI证据理由": ai.get("reason", "") or "",
        "人工终审结论": "",
        "终审人": "",
        "终审时间": "",
        "模型版本": ai.get("model_version", ai.get("model", "")) or "",
        "prompt版本": ai.get("prompt_version", "") or "",
        "遮蔽比例": ai.get("mask_pct_used", "") or "",
        "调用状态": ai.get("call_status", "") or ("未调用" if state == "none" else ""),
        "AI核验时间": ai.get("check_time", "") or "",
        "判定规则版本": "%s/%s" % (RULE_VERSION, ENGINE_VERSION),
    }


def summarize(rows: List[Dict[str, Any]]) -> Dict[str, int]:
    """按问题类型与结论等级汇总。"""
    out: Dict[str, int] = {}
    for r in rows:
        for key in ("问题类型", "结论等级"):
            v = str(r.get(key, "")) or "(空)"
            out["%s:%s" % (key, v)] = out.get("%s:%s" % (key, v), 0) + 1
    return out
