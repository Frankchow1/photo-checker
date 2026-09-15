"""
规则引擎：把 AI 报告的原子事实（observations）映射为业务结论。

故意只支持有限算子，避免退化成一门小语言；复杂逻辑应写成场景专属钩子。

条件语法（可嵌套）：
    {"all": [cond, ...]}   全部成立
    {"any": [cond, ...]}   任一成立
    {"not": cond}          取反
    {"field": "...", "op": "contains_any", "value": [...]}   叶子条件

规则按数组顺序逐条匹配，**首条命中即返回**，因此需把「信息不足」
这类兵守规则写在前面。全部未命中时返回场景声明的 fallback。
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional

CONCLUSION_LEVELS = ("确认", "高度疑似", "正常", "待核实", "未抽检")


class RuleError(Exception):
    pass


def _to_list(value: Any) -> List[Any]:
    if value is None:
        return []
    if isinstance(value, (list, tuple, set)):
        return list(value)
    return [value]


def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (list, tuple)):
        return " ".join(_text(v) for v in value)
    return str(value).strip()


def _number(value: Any) -> Optional[float]:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _is_empty(value: Any) -> bool:
    return value is None or value == "" or value == [] or value == {}


def eval_leaf(cond: Dict[str, Any], obs: Dict[str, Any]) -> bool:
    """求值单条叶子条件。字符串比较统一去空格、忽略英文大小写。"""
    name = cond.get("field")
    if not name:
        raise RuleError(f"叶子条件缺少 field: {cond}")
    op = (cond.get("op") or "equals").strip()
    expected = cond.get("value")
    actual = obs.get(name)

    if op == "exists":
        return not _is_empty(actual)
    if op == "missing":
        return _is_empty(actual)
    if op == "is_true":
        return bool(actual) and _text(actual).lower() not in ("false", "0", "否")
    if op == "is_false":
        return (not bool(actual)) or _text(actual).lower() in ("false", "0", "否")

    actual_text = _text(actual).lower()
    expected_list = [_text(v).lower() for v in _to_list(expected)]

    if op == "equals":
        return actual_text == (expected_list[0] if expected_list else "")
    if op == "not_equals":
        return actual_text != (expected_list[0] if expected_list else "")
    if op == "in":
        return actual_text in expected_list
    if op == "not_in":
        return actual_text not in expected_list
    if op == "contains_any":
        return any(token and token in actual_text for token in expected_list)
    if op == "contains_all":
        return all(token and token in actual_text for token in expected_list)
    if op == "not_contains_any":
        return not any(token and token in actual_text for token in expected_list)
    if op == "regex_any":
        return any(re.search(pattern, actual_text) for pattern in expected_list if pattern)
    if op in ("gte", "lte", "gt", "lt"):
        left = _number(actual)
        right = _number(expected)
        if left is None or right is None:
            return False
        if op == "gte":
            return left >= right
        if op == "lte":
            return left <= right
        if op == "gt":
            return left > right
        return left < right

    raise RuleError(f"不支持的算子: {op}")


def eval_condition(cond: Any, obs: Dict[str, Any]) -> bool:
    """递归求值条件树。空条件视为恒真（用于兑底规则）。"""
    if cond is None or cond == {}:
        return True
    if not isinstance(cond, dict):
        raise RuleError(f"条件必须是对象: {cond!r}")
    if "all" in cond:
        return all(eval_condition(c, obs) for c in _to_list(cond["all"]))
    if "any" in cond:
        return any(eval_condition(c, obs) for c in _to_list(cond["any"]))
    if "not" in cond:
        return not eval_condition(cond["not"], obs)
    return eval_leaf(cond, obs)


def evaluate_rules(
    rules: List[Dict[str, Any]],
    obs: Dict[str, Any],
    fallback: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """逐条匹配，首条命中即返回。返回的 dict 直接并入统一结果模型。"""
    for rule in rules or []:
        rule_id = rule.get("id") or "unnamed"
        try:
            hit = eval_condition(rule.get("when"), obs)
        except RuleError as exc:
            raise RuleError(f"规则 {rule_id} 求值失败: {exc}") from exc
        if hit:
            outcome = dict(rule.get("then") or {})
            outcome["matched_rule"] = rule_id
            return outcome
    outcome = dict(fallback or {"problem_type": "待定", "conclusion_level": "待核实", "evidence": "未命中任何规则"})
    outcome["matched_rule"] = outcome.get("matched_rule", "__fallback__")
    return outcome


def validate_rules(rules: List[Dict[str, Any]]) -> List[str]:
    """静态体检：加载场景包时调用，提前暴露配置错误。"""
    errors: List[str] = []
    seen = set()
    probe = {"__probe__": ""}
    for index, rule in enumerate(rules or []):
        rule_id = rule.get("id")
        if not rule_id:
            errors.append(f"第 {index + 1} 条规则缺少 id")
        elif rule_id in seen:
            errors.append(f"规则 id 重复: {rule_id}")
        else:
            seen.add(rule_id)
        then = rule.get("then") or {}
        level = then.get("conclusion_level")
        if level and level not in CONCLUSION_LEVELS:
            errors.append(f"规则 {rule_id} 的结论等级非法: {level}")
        try:
            eval_condition(rule.get("when"), probe)
        except RuleError as exc:
            errors.append(str(exc))
    return errors


def collect_fields(cond: Any, into: Optional[set] = None) -> set:
    """收集条件树用到的字段名，用于校验规则引用的字段都在 schema 里。"""
    into = into if into is not None else set()
    if not isinstance(cond, dict):
        return into
    for key in ("all", "any"):
        if key in cond:
            for child in _to_list(cond[key]):
                collect_fields(child, into)
            return into
    if "not" in cond:
        return collect_fields(cond["not"], into)
    if "field" in cond:
        into.add(cond["field"])
    return into
