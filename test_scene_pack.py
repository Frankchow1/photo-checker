"""
场景包自测：不调用任何 API、不需照片，纯本地零成本验收。

用法：
    python test_scene_pack.py

校验内容：
  1. scenes/ 下所有场景包能正常加载（配置/规则/提示词体检）
  2. 提示词变量渲染正常，且不残留未替换占位符
  3. 建行场景的典型用例判定结果符合预期（包含难负例）
  4. 缓存 key 确实随提示词版本变化
"""

from __future__ import annotations

import sys

from scene_pack import load_all_scene_packs, load_scene_pack

CASES = [
    (
        "标准建行网点",
        {"signboard_text": "中国建设银行绍兴柯桥支行", "has_logo": True, "logo_main_color": "蓝",
         "venue_type": "银行网点", "has_branch_suffix": True, "visible_scope": "完整门头",
         "text_confidence": "高"},
        ("建行网点", "确认"),
    ),
    (
        "难负例-中国银行（同为蓝色）",
        {"signboard_text": "中国银行绍兴分行", "has_logo": True, "logo_main_color": "蓝",
         "venue_type": "银行网点", "visible_scope": "完整门头", "text_confidence": "高"},
        ("非目标品牌", "正常"),
    ),
    (
        "难负例-建行ATM自助区",
        {"signboard_text": "中国建设银行 24小时自助银行", "has_logo": True, "logo_main_color": "蓝",
         "venue_type": "ATM自助区", "visible_scope": "完整门头", "text_confidence": "高"},
        ("关联但非网点", "正常"),
    ),
    (
        "难负例-商铺贴建行广告牌",
        {"signboard_text": "建行分期先享后付", "has_logo": True, "logo_main_color": "蓝",
         "venue_type": "广告牌", "visible_scope": "完整门头", "text_confidence": "高"},
        ("关联但非网点", "正常"),
    ),
    (
        "信息不足-只拍到局部",
        {"signboard_text": "□□银行", "has_logo": False, "logo_main_color": "无法判断",
         "venue_type": "无法判断", "visible_scope": "局部", "text_confidence": "低"},
        ("信息不足", "待核实"),
    ),
    (
        "仅文字命中-场所不明",
        {"signboard_text": "建设银行自助服务点", "has_logo": False, "logo_main_color": "其他",
         "venue_type": "普通商铺", "visible_scope": "完整门头", "text_confidence": "高"},
        ("建行网点", "高度疑似"),
    ),
    (
        "无关商铺",
        {"signboard_text": "老王面馆", "has_logo": False, "logo_main_color": "红",
         "venue_type": "普通商铺", "visible_scope": "完整门头", "text_confidence": "高"},
        ("非目标品牌", "正常"),
    ),
]


def main() -> int:
    failures = []

    packs, problems = load_all_scene_packs("scenes")
    print(f"[1] 加载场景包: 成功 {len(packs)} 个")
    for pack in packs:
        pv = pack.prompt()
        flag = "已验证" if pack.is_verified else "未验证"
        print(f"    - {pack.scene_id} v{pack.version} [{pack.input_type}] prompt={pv.version}({pv.hash}) {flag}")
    for problem in problems:
        failures.append(f"场景包加载失败: {problem}")

    try:
        ccb = load_scene_pack("scenes/bank_branch_ccb")
    except Exception as exc:
        print(f"[FATAL] 建行场景包无法加载: {exc}")
        return 1

    text, pv = ccb.rendered_prompt()
    print(f"\n[2] 提示词渲染: {pv.describe()} 长度={len(text)}")
    if "{{" in text:
        failures.append("提示词渲染后仍残留未替换占位符 {{...}}")
    if "中国建设银行" not in text:
        failures.append("提示词未注入品牌名")

    print("\n[3] 规则判定用例:")
    for title, obs, expected in CASES:
        schema_errors = ccb.validate_observations(obs)
        outcome = ccb.evaluate(obs)
        got = (outcome.get("problem_type"), outcome.get("conclusion_level"))
        ok = got == expected and not schema_errors
        mark = "PASS" if ok else "FAIL"
        print(f"    [{mark}] {title}: {got[0]} / {got[1]}  <- {outcome.get('matched_rule')}")
        if schema_errors:
            print(f"           schema: {schema_errors}")
            failures.append(f"{title} schema 校验不通过: {schema_errors}")
        if got != expected:
            failures.append(f"{title} 预期 {expected} 实际 {got}")

    print("\n[4] 缓存 key 隔离:")
    key_a = ccb.cache_key("fingerprint-abc")
    key_b = ccb.cache_key("fingerprint-xyz")
    print(f"    同场景不同图片: {key_a[:12]} vs {key_b[:12]}")
    if key_a == key_b:
        failures.append("不同图片的缓存 key 碰撞")

    print("\n[5] 级联门禁:")
    gate_hit = ccb.ocr_gate("中国建设银行柯桥支行")
    gate_miss = ccb.ocr_gate("老王面馆")
    print(f"    命中关键词 -> pass_to_l2={gate_hit['pass_to_l2']} ({gate_hit['reason']})")
    print(f"    未命中     -> pass_to_l2={gate_miss['pass_to_l2']} ({gate_miss['reason']})")

    print("\n" + "=" * 60)
    if failures:
        print(f"失败 {len(failures)} 项:")
        for item in failures:
            print(f"  - {item}")
        return 1
    print("全部通过。场景包框架可用。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
