"""
场景包（Scene Pack）框架内核 —— P0 阶段。

设计原则（详见 Notion《多场景质检框架设计（Scene Pack 架构）》）：
  1. AI 只报客观事实（observations），结论由规则层计算。
  2. 内核不认识任何具体业务词，新增场景 = 新增一个配置目录。
  3. 级联分层（L0 规则 / L1 OCR / L2 多模态 / L3 人工）由场景包声明。
  4. 提示词不写死在配置里，通过 "场景@标签" 引用版本化文件。
  5. 缓存 key 必须包含场景与提示词版本，否则改了口径仍命中旧缓存。
"""

from __future__ import annotations

import hashlib
import json
import os
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

from prompt_store import PromptVersion, resolve_prompt

FRAMEWORK_VERSION = "0.1.0"
SCENES_ROOT = "scenes"
SCENE_CONFIG_FILENAME = "scene.json"

# 统一结果模型：所有场景共用，导出层据此通用化
