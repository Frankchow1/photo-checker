# 场景包（Scene Pack）框架 · P0

目标：新增一个识别需求（如「建行网点」「某连锁便利店」）**只改配置，不改代码**。

设计说明见 Notion《多场景质检框架设计（Scene Pack 架构）》。

## 文件构成

| 文件 | 职责 |
| --- | --- |
| `prompt_store.py` | 提示词版本化：版本文件 + 标签指针，支持回滚与 A/B |
| `rule_engine.py` | 规则引擎：把 AI 报的客观事实映射为问题类型 + 结论等级 |
| `scene_pack.py` | 场景包加载器、级联门禁、缓存 key、统一结果模型 |
| `scenes/<scene_id>/` | 具体场景包（唯一需要新增的地方） |
| `test_scene_pack.py` | 零成本自测，不调 API、不需照片 |

## 快速验收

```bash
python test_scene_pack.py
```

会依次校验：场景包能否加载、提示词变量是否渲染完全、七个典型用例（含四个难负例）判定是否符合预期、缓存 key 是否隔离、OCR 门禁是否生效。

## 核心原则：AI 只报事实，结论由规则算

不要问模型「这是不是建行网点」（会顺着问题幻觉，且换品牌就要重写），而是让它回答可观察事实：招牌文字、logo 主色、场所类型、可见范围、转写置信度……然后由 `rules` 计算结论。

带来三个直接好处：

1. 换品牌只改配置，不改代码
2. **改判定口径不用重跑 AI**：事实已缓存，调规则重算即可，零成本
3. 每行结果都能说清命中了哪条规则（`matched_rule`）

## 新增一个场景

```
scenes/<scene_id>/
    scene.json              配置主体
    prompts/
        labels.json         {"production": "v1"}
        v1.md               提示词正文
    samples/
        positive/           正例
        hard_negative/      难负例（差一点就判错的）
        golden/             金标集
```

直接复制 `scenes/bank_branch_ccb/` 改写最快。

<b>无 `samples/golden/` 目录的场景，`is_verified` 为 False，UI 上会标注「未验证」。</b>

### scene.json 关键字段

| 字段 | 说明 |
| --- | --- |
| `input_type` | `single_image` / `image_pair` / `image_with_fields` |
| `prompt_ref` | `场景id@标签`，不写死提示词内容 |
| `variables` | 品牌包变量，用 `{{key}}` 注入提示词 |
| `cascade` | 声明启用哪几层（L1_ocr / L2_vlm / L3_manual） |
| `output_schema` | 原子事实字段定义，用于校验与重试 |
| `rules` | 按数组顺序逐条匹配，**首条命中即返回** |
| `fallback` | 全部未命中时的兑底结论 |
| `columns` | 表格与导出列（`obs.xxx` 表示取原子事实字段） |
| `samples` | 样例集目录 |

### 规则条件语法

```json
{ "all": [ 条件, ... ] }
{ "any": [ 条件, ... ] }
{ "not": 条件 }
{ "field": "venue_type", "op": "equals", "value": "银行网点" }
```

支持算子（故意只有这些，避免退化成小语言）：
`equals` `not_equals` `in` `not_in` `contains_any` `contains_all` `not_contains_any` `regex_any` `exists` `missing` `is_true` `is_false` `gte` `lte` `gt` `lt`

字符串比较统一去空格、忽略英文大小写，所以 `value` 里的英文请写小写（如 `ccb`、`atm自助区`）。

### 规则顺序很重要

建行场景的顺序安排可直接参考：

1. `competitor_bank` —— 先排除同业，避免蓝色 logo 混淆
2. `ccb_atm_or_ad` —— ATM/广告牌单独成一类，不混入网点
3. `ccb_confirmed` —— 文字 + （logo 主色 或 场所类型）双条件
4. `low_confidence_or_partial` —— 信息不足兵守（放在确认之后，才不会把铁证降级）
5. `ccb_text_only` / `visual_only` —— 单条件只给「高度疑似」

## 级联门禁（控成本）

`L1_ocr` 当前默认 `enabled: false`（尚未接入 OCR），此时所有照片都走大模型，行为与改造前一致。

接入 OCR 后改为 `true`，效果：只有 OCR 读到银行相关文字的照片才进大模型，10 万张量级下预计只有 3%~8% 需要付费调用。

<b>`audit_sample_rate` 不要设为 0</b>：未命中项必须保留少量抽样进大模型，用于监控 OCR 漏检率。

## 提示词版本管理

```python
from prompt_store import write_version, set_label, list_versions

write_version("scenes/bank_branch_ccb", new_text)     # 写入 v2，不自动上线
set_label("scenes/bank_branch_ccb", "staging", "v2")  # 先跑金标集
set_label("scenes/bank_branch_ccb", "production", "v2")  # 确认更好再上线
```

回滚就是把 `production` 指回 `v1`，一行配置，不动提示词内容。

对外分发时，调口径只需替换 `prompts/` 下的文件，**不需重新打包程序**。

## 缓存 key

`cache_key()` 已把 `scene_id + scene_version + prompt_version + prompt_hash + 图片指纹` 全部纳入。

<b>改了提示词会自动失效旧缓存</b>，不会出现“改了口径但结果没变”这类最难排查的问题。

## 当前进度与待办

| 阶段 | 状态 |
| --- | --- |
| P0 框架内核（加载器/规则引擎/级联/版本化） | 已完成 |
| P2 建行场景包（配置 + 提示词 v1 + 自测用例） | 已完成（缺金标集） |
| P1 把现有 P2 真实门头照质检改写为场景包 | 待做 |
| P3 UI 场景下拉 + 动态渲染 | 待做 |
| P4 AI 生成提示词 | 待做（依赖金标集） |

### 已知局限

1. `variables` 目前只在**提示词渲染**时生效，`rules` 中仍需写字面值。换品牌时两处都要改；后续可支持规则内变量替换。
2. 尚未接入 OCR，`L1_ocr` 先置 false。
3. 尚未接入 GUI，当前只能通过自测脚本验证规则。
4. 建行场景缺 `samples/golden/`，按约定属于「未验证」状态，不应直接用于对外出结论。
