"""
场景包（Scene Pack）内核 —— P0 阶段。

设计原则（详见 Notion《多场景质检框架设计（Scene Pack 架构）》）：
  1. AI 只报客观事实，结论由规则层计算（见 rule_engine.py）。
  2. 内核不认识任何具体业务词，新增场景 = 新增一个配置目录。
  3. 级联分层（L0/L1/L2/L3）由场景包声明。
  4. 提示词通过 "场景@标签" 引用版本化文件（见 prompt_store.py）。
  5. 缓存 key 必须包含场景与提示词版本。
"""

from __future__ import annotations

import hashlib
import json
import os
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

from prompt_store import PromptVersion, resolve_prompt
from rule_engine import CONCLUSION_LEVELS, collect_fields, evaluate_rules, validate_rules

FRAMEWORK_VERSION = "0.1.0"
SCENES_ROOT = "scenes"
SCENE_CONFIG_FILENAME = "scene.json"
INPUT_TYPES = ("single_image", "image_pair", "image_with_fields")

# 统一结果模型：所有场景共用，导出层据此通用化
RESULT_COLUMNS: List[Tuple[str, str]] = [
    ("scene_id", "场景ID"),
    ("scene_version", "场景版本"),
    ("cascade_stage", "判定层"),
    ("problem_type", "问题类型"),
    ("conclusion_level", "结论等级"),
    ("evidence", "结论依据"),
    ("matched_rule", "命中规则"),
    ("confidence", "置信度(仅参考)"),
    ("prompt_version", "prompt版本"),
    ("prompt_hash", "prompt指纹"),
    ("model_version", "模型版本"),
    ("call_status", "调用状态"),
    ("retries", "重试次数"),
    ("checked_at", "核验时间"),
    ("manual_conclusion", "人工终审结论"),
    ("manual_reviewer", "终审人"),
    ("manual_reviewed_at", "终审时间"),
]


class SceneConfigError(Exception):
    pass


@dataclass
class ScenePack:
    scene_id: str
    name: str
    version: str
    input_type: str
    root_dir: str
    prompt_ref: str = ""
    variables: Dict[str, Any] = field(default_factory=dict)
    output_schema: List[Dict[str, Any]] = field(default_factory=list)
    rules: List[Dict[str, Any]] = field(default_factory=list)
    fallback: Dict[str, Any] = field(default_factory=dict)
    cascade: List[Dict[str, Any]] = field(default_factory=list)
    columns: List[Dict[str, Any]] = field(default_factory=list)
    defaults: Dict[str, Any] = field(default_factory=dict)
    samples: Dict[str, Any] = field(default_factory=dict)
    description: str = ""

    # ---------- 提示词 ----------
    def prompt(self) -> PromptVersion:
        return resolve_prompt(self.root_dir, self.prompt_ref, self.scene_id)

    def rendered_prompt(self) -> Tuple[str, PromptVersion]:
        """返回注入品牌包等变量后的提示词正文。"""
        pv = self.prompt()
        return pv.render(self.variables), pv

    # ---------- 输出校验 ----------
    def schema_fields(self) -> List[str]:
        return [f["name"] for f in self.output_schema if f.get("name")]

    def validate_observations(self, obs: Dict[str, Any]) -> List[str]:
        """校验 AI 返回的原子事实；返回错误列表，空列表代表通过。"""
        errors: List[str] = []
        if not isinstance(obs, dict):
            return ["observations 必须是对象"]
        for spec in self.output_schema:
            name = spec.get("name")
            if not name:
                continue
            value = obs.get(name)
            missing = value is None or value == ""
            if spec.get("required") and missing:
                errors.append(f"缺少必填字段: {name}")
                continue
            if missing:
                continue
            enum = spec.get("enum")
            if enum and str(value).strip() not in [str(e) for e in enum]:
                errors.append(f"字段 {name} 取值非法: {value}（应为 {enum}）")
        return errors

    # ---------- 结论 ----------
    def evaluate(self, obs: Dict[str, Any]) -> Dict[str, Any]:
        return evaluate_rules(self.rules, obs or {}, self.fallback)

    # ---------- 级联 ----------
    def layer(self, layer_name: str) -> Optional[Dict[str, Any]]:
        for item in self.cascade:
            if item.get("layer") == layer_name and item.get("enabled", True):
                return item
        return None

    def ocr_gate(self, ocr_text: str) -> Dict[str, Any]:
        """L1 文字门禁：决定这张图要不要花钱进 L2 大模型。

        未配置 L1 时一律放行，保证默认行为与改造前一致。
        """
        cfg = self.layer("L1_ocr")
        if not cfg:
            return {"pass_to_l2": True, "cascade_stage": "L2", "reason": "未启用L1门禁"}
        text = (ocr_text or "").lower()
        blockers = [str(k).lower() for k in cfg.get("blocking_keywords", []) if k]
        for token in blockers:
            if token and token in text:
                blocked = dict(cfg.get("on_block") or {})
                blocked.setdefault("problem_type", "正常")
                blocked.setdefault("conclusion_level", "正常")
                blocked.setdefault("evidence", f"OCR命中排除词: {token}")
                return {"pass_to_l2": False, "cascade_stage": "L1", "outcome": blocked, "reason": blocked["evidence"]}
        keywords = [str(k).lower() for k in cfg.get("keywords_any", []) if k]
        if not keywords:
            return {"pass_to_l2": True, "cascade_stage": "L2", "reason": "L1未配置关键词"}
        hit = next((k for k in keywords if k in text), "")
        if hit:
            return {"pass_to_l2": True, "cascade_stage": "L2", "reason": f"OCR命中关键词: {hit}"}
        miss = dict(cfg.get("on_miss") or {})
        miss.setdefault("problem_type", "正常")
        miss.setdefault("conclusion_level", "正常")
        miss.setdefault("evidence", "OCR未命中任何关键词，未进入大模型")
        return {
            "pass_to_l2": False,
            "cascade_stage": "L1",
            "outcome": miss,
            "reason": miss["evidence"],
            "audit_sample_rate": float(cfg.get("audit_sample_rate", 0.0) or 0.0),
        }

    # ---------- 缓存 ----------
    def cache_key(self, image_fingerprint: str, extra: str = "") -> str:
        """缓存 key 必须含场景与提示词版本，否则改了口径仍命中旧结果。"""
        pv = self.prompt()
        raw = "|".join([self.scene_id, self.version, pv.version, pv.hash, image_fingerprint or "", extra or ""])
        return hashlib.sha1(raw.encode("utf-8")).hexdigest()

    # ---------- 展示 ----------
    @property
    def is_verified(self) -> bool:
        """没有金标集的场景在 UI 上应标注「未验证」。"""
        golden = (self.samples or {}).get("golden_set")
        if not golden:
            return False
        return os.path.isdir(os.path.join(self.root_dir, golden))

    def display_label(self) -> str:
        suffix = "" if self.is_verified else "（未验证）"
        return f"{self.name}{suffix}"

    def result_template(self) -> Dict[str, Any]:
        pv = self.prompt()
        return {
            "scene_id": self.scene_id,
            "scene_version": self.version,
            "prompt_version": pv.version,
            "prompt_hash": pv.hash,
            "cascade_stage": "",
            "problem_type": "",
            "conclusion_level": "未抽检",
            "evidence": "",
            "matched_rule": "",
            "confidence": None,
            "model_version": "",
            "call_status": "",
            "retries": 0,
            "checked_at": "",
        }


def _read_json(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_scene_pack(scene_dir: str) -> ScenePack:
    """加载并体检一个场景包。配置错误在加载时就报，不拖到跑批量时才爆。"""
    config_path = os.path.join(scene_dir, SCENE_CONFIG_FILENAME)
    if not os.path.exists(config_path):
        raise SceneConfigError(f"缺少 {SCENE_CONFIG_FILENAME}: {scene_dir}")
    cfg = _read_json(config_path)

    scene_id = cfg.get("scene_id") or os.path.basename(scene_dir.rstrip(os.sep))
    input_type = cfg.get("input_type", "single_image")
    if input_type not in INPUT_TYPES:
        raise SceneConfigError(f"{scene_id}: input_type 非法 {input_type}（应为 {INPUT_TYPES}）")

    pack = ScenePack(
        scene_id=scene_id,
        name=cfg.get("name") or scene_id,
        version=str(cfg.get("version") or "1.0.0"),
        input_type=input_type,
        root_dir=scene_dir,
        prompt_ref=cfg.get("prompt_ref") or f"{scene_id}@production",
        variables=cfg.get("variables") or {},
        output_schema=cfg.get("output_schema") or [],
        rules=cfg.get("rules") or [],
        fallback=cfg.get("fallback") or {},
        cascade=cfg.get("cascade") or [],
        columns=cfg.get("columns") or [],
        defaults=cfg.get("defaults") or {},
        samples=cfg.get("samples") or {},
        description=cfg.get("description") or "",
    )

    errors = validate_rules(pack.rules)
    schema_names = set(pack.schema_fields())
    if schema_names:
        for rule in pack.rules:
            for used in collect_fields(rule.get("when")):
                if used not in schema_names:
                    errors.append(f"规则 {rule.get('id')} 引用了 schema 中不存在的字段: {used}")
    level = (pack.fallback or {}).get("conclusion_level")
    if level and level not in CONCLUSION_LEVELS:
        errors.append(f"fallback 结论等级非法: {level}")
    try:
        pack.prompt()
    except Exception as exc:
        errors.append(f"提示词解析失败: {exc}")
    if errors:
        raise SceneConfigError(f"场景包 {scene_id} 配置有误:\n  - " + "\n  - ".join(errors))
    return pack


def load_all_scene_packs(root: str = SCENES_ROOT, strict: bool = False) -> Tuple[List[ScenePack], List[str]]:
    """扫描 scenes/ 目录加载全部场景包，返回（可用场景, 错误信息）。"""
    packs: List[ScenePack] = []
    problems: List[str] = []
    if not os.path.isdir(root):
        return packs, [f"场景目录不存在: {root}"]
    for name in sorted(os.listdir(root)):
        scene_dir = os.path.join(root, name)
        if not os.path.isdir(scene_dir) or name.startswith("."):
            continue
        if not os.path.exists(os.path.join(scene_dir, SCENE_CONFIG_FILENAME)):
            continue
        try:
            packs.append(load_scene_pack(scene_dir))
        except Exception as exc:
            if strict:
                raise
            problems.append(str(exc))
    return packs, problems


def find_scene_pack(scene_id: str, root: str = SCENES_ROOT) -> ScenePack:
    return load_scene_pack(os.path.join(root, scene_id))
