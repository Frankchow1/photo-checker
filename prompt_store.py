"""
提示词版本化存储（Prompt Store）。

解决的问题：提示词若直接写死在场景配置或代码里，改一次覆盖一次，
导致历史结果无法还原口径、无法回滚、无法 A/B。

设计：版本文件 + 标签指针
    scenes/<scene_id>/prompts/
        labels.json    {"production": "v2", "staging": "v3"}
        v1.md
        v2.md

场景配置只写引用："prompt_ref": "bank_branch_ccb@production"

不依赖任何第三方服务，纯本地文件实现。
"""

from __future__ import annotations

import hashlib
import json
import os
import re
from dataclasses import dataclass
from typing import Dict, List, Optional

PROMPTS_DIRNAME = "prompts"
LABELS_FILENAME = "labels.json"
DEFAULT_LABEL = "production"
VERSION_PATTERN = re.compile(r"^v(\d+)$")


class PromptStoreError(Exception):
    pass


@dataclass
class PromptVersion:
    """一个具体的提示词版本。hash 用于缓存 key 与结果追溯。"""

    scene_id: str
    version: str
    label: Optional[str]
    text: str

    @property
    def hash(self) -> str:
        return hashlib.sha1(self.text.encode("utf-8")).hexdigest()[:12]

    def render(self, variables: Optional[Dict[str, object]] = None) -> str:
        """用 {{key}} 占位符注入变量（例如品牌包字段）。"""
        return render_template(self.text, variables or {})

    def describe(self) -> str:
        tag = f"@{self.label}" if self.label else ""
        return f"{self.scene_id}{tag} -> {self.version} ({self.hash})"


def render_template(text: str, variables: Dict[str, object]) -> str:
    """最小模板引擎：仅支持 {{key}} 与 {{key|join}}（列表按顿号连接）。"""

    def _replace(match: re.Match) -> str:
        raw = match.group(1).strip()
        key, _, modifier = raw.partition("|")
        key = key.strip()
        modifier = modifier.strip()
        if key not in variables:
            return match.group(0)
        value = variables[key]
        if isinstance(value, (list, tuple)):
            sep = "、" if modifier in ("", "join") else modifier
            return sep.join(str(v) for v in value)
        return str(value)

    return re.sub(r"\{\{([^{}]+)\}\}", _replace, text)


def prompts_dir(scene_dir: str) -> str:
    return os.path.join(scene_dir, PROMPTS_DIRNAME)


def _labels_path(scene_dir: str) -> str:
    return os.path.join(prompts_dir(scene_dir), LABELS_FILENAME)


def load_labels(scene_dir: str) -> Dict[str, str]:
    path = _labels_path(scene_dir)
    if not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        raise PromptStoreError(f"labels.json 格式错误: {path}")
    return {str(k): str(v) for k, v in data.items()}


def list_versions(scene_dir: str) -> List[str]:
    """返回按数字序排列的版本名，如 ['v1', 'v2', 'v10']。"""
    directory = prompts_dir(scene_dir)
    if not os.path.isdir(directory):
        return []
    versions = []
    for name in os.listdir(directory):
        if not name.endswith(".md"):
            continue
        stem = name[:-3]
        if VERSION_PATTERN.match(stem):
            versions.append(stem)
    return sorted(versions, key=lambda v: int(VERSION_PATTERN.match(v).group(1)))


def next_version(scene_dir: str) -> str:
    versions = list_versions(scene_dir)
    if not versions:
        return "v1"
    return "v" + str(int(VERSION_PATTERN.match(versions[-1]).group(1)) + 1)


def parse_ref(ref: str, default_scene_id: str = "") -> (str, str):
    """解析 'scene_id@label'、'@label'、'scene_id@v2'、'v2' 四种写法。"""
    ref = (ref or "").strip()
    if not ref:
        return default_scene_id, DEFAULT_LABEL
    if "@" in ref:
        scene_id, _, target = ref.partition("@")
        return (scene_id.strip() or default_scene_id), (target.strip() or DEFAULT_LABEL)
    if VERSION_PATTERN.match(ref):
        return default_scene_id, ref
    return ref, DEFAULT_LABEL


def resolve_prompt(scene_dir: str, ref: str, scene_id: str = "") -> PromptVersion:
    """把引用解析为具体版本内容。

    解析顺序：显式版本号 -> labels.json 中的标签 -> production 标签 -> 最新版本。
    """
    resolved_scene_id, target = parse_ref(ref, scene_id)
    resolved_scene_id = resolved_scene_id or scene_id
    labels = load_labels(scene_dir)
    label_used: Optional[str] = None

    if VERSION_PATTERN.match(target):
        version = target
    elif target in labels:
        version = labels[target]
        label_used = target
    elif DEFAULT_LABEL in labels:
        version = labels[DEFAULT_LABEL]
        label_used = DEFAULT_LABEL
    else:
        available = list_versions(scene_dir)
        if not available:
            raise PromptStoreError(f"场景 {resolved_scene_id} 没有任何提示词版本: {prompts_dir(scene_dir)}")
        version = available[-1]

    path = os.path.join(prompts_dir(scene_dir), version + ".md")
    if not os.path.exists(path):
        raise PromptStoreError(f"提示词版本文件不存在: {path}（引用 {ref}）")
    with open(path, "r", encoding="utf-8") as f:
        text = f.read().strip()
    if not text:
        raise PromptStoreError(f"提示词版本内容为空: {path}")
    return PromptVersion(scene_id=resolved_scene_id, version=version, label=label_used, text=text)


def write_version(scene_dir: str, text: str, version: str = "") -> PromptVersion:
    """写入新版本（供 AI 生成提示词 / 自动迭代优化使用），不自动改标签。"""
    directory = prompts_dir(scene_dir)
    os.makedirs(directory, exist_ok=True)
    version = version or next_version(scene_dir)
    if not VERSION_PATTERN.match(version):
        raise PromptStoreError(f"版本名必须形如 v1/v2: {version}")
    path = os.path.join(directory, version + ".md")
    if os.path.exists(path):
        raise PromptStoreError(f"版本已存在，不允许覆盖: {path}")
    with open(path, "w", encoding="utf-8") as f:
        f.write(text.strip() + "\n")
    return PromptVersion(scene_id=os.path.basename(scene_dir.rstrip(os.sep)), version=version, label=None, text=text.strip())


def set_label(scene_dir: str, label: str, version: str) -> Dict[str, str]:
    """把标签指向某个版本（上线 / 回滚就是改这一行）。"""
    if version not in list_versions(scene_dir):
        raise PromptStoreError(f"版本不存在: {version}")
    labels = load_labels(scene_dir)
    labels[label] = version
    directory = prompts_dir(scene_dir)
    os.makedirs(directory, exist_ok=True)
    with open(_labels_path(scene_dir), "w", encoding="utf-8") as f:
        json.dump(labels, f, ensure_ascii=False, indent=2)
        f.write("\n")
    return labels
