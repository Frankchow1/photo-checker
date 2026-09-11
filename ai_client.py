import os
import io
import json
import time
import base64
import hashlib
import urllib.request
import urllib.error
from typing import Tuple, Dict, Any, Optional
from PIL import Image, ImageDraw


def _load_local_config():
    cfg_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.local.json")
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


_local_cfg = _load_local_config()
DEFAULT_API_URL = os.environ.get("ARK_API_URL", _local_cfg.get("api_url", "https://ark.cn-beijing.volces.com/api/plan/v3"))
DEFAULT_MODEL = os.environ.get("ARK_MODEL", _local_cfg.get("model", "doubao-seed-2.0-lite"))
DEFAULT_API_KEY = os.environ.get("ARK_API_KEY", _local_cfg.get("api_key", ""))

# 可复现性相关的版本号：修改提示词内容时请同步递增，以便结果可追溯
COMPARE_PROMPT_VERSION = "compare-v2.1"
STOREFRONT_PROMPT_VERSION = "storefront-v2.1"
CLIENT_VERSION = "ai_client-2.1.0"

# 采样参数固定，保证同一对照片多次调用结果尽可能一致
FIXED_TEMPERATURE = 0.0
FIXED_TOP_P = 1.0

# 返回值必须包含的字段（缺字段即视为不合规，重试一次）
COMPARE_REQUIRED_KEYS = ("is_same_photo", "confidence", "verdict", "reason")
STOREFRONT_REQUIRED_KEYS = ("is_real_storefront", "confidence", "risk_type", "reason")

# 默认内置的多维证据链相同照片比对提示词
# 注意：模型输出只是【建议】，最终定性由 conclusion_engine 结合确定性算法与人工终审给出
DEFAULT_COMPARE_PROMPT = """你是一个专业的商户现场照片取证与真实性核验专家。
【重要前置说明】：
如果已启用物理去水印，照片底部的水印区域已在输入前用黑色色块遮蔽，你看到的是纯净的现场实拍图像。即使未遮蔽，外勤软件拍照必然叠加不同水印，【继绝不要关注水印，严禁以水印不同作为判定依据，也严禁读取水印中的时间作为判据】！
时间关系由系统结构化字段判定，不需要你推测。

【核验任务】：
综合多维物理证据，判断图1 与 图2 是否为【同一张原始拍摄照片（即同一底片复用）】。
你的结论是【建议】，不是最终处罚依据；不确定时请如实给出低置信度并说明原因，不要强行二选一。

【多维度综合取证依据】：
1. 【店招维度】（注：若拍摄的是室内/工位/陈列/无招牌门面，请标明无店招，店招缺失不影响其他维度的同一性判定）：
   - 图1与图2是否存在门头招牌/店名？若有，请原文转写招牌文字，不要自行补全或编造。
2. 【机位与透视几何】：拍摄机位、透视角度、画框边缘截断处是否完全一致？
3. 【光影与反光特征】：光源照射角度、阴影走向与长度、玻璃反光光斑是否完全重叠？
4. 【瞬态陈列与偶发动态物】：临时停放车辆、纸箱、海报、花篮等的相对位置与形态是否静态一致？
5. 【微观固有细节】：卷帘门划痕、墙体破损、瓷砖缝隙及特定污渍是否吻合？

【容差规则】：请忽略微小的数字压缩噪点或轻微色差，不要把 JPEG 压缩失真误判为物理差异。
【连锁与统一装修】：若两张图是高度相似但可辨认为不同店铺（连锁品牌、统一装修、商场档口），请判为【误判】。

请严格返回合法 JSON，不要输出 markdown 代码块以外的任何文字：
{
  "is_same_photo": true 或 false,
  "confidence": 0.0 到 1.0 之间的置信度数值,
  "has_signboard": "双方均有" / "图1有图2无" / "图1无图2有" / "双方均无店招",
  "signboard_text_1": "图1店招文字(无则写无店招)",
  "signboard_text_2": "图2店招文字(无则写无店招)",
  "signboard_match": "一致" / "不一致" / "无店招不适用",
  "scene_similarity": 0.0 到 1.0 之间的物理场景吻合度,
  "verdict": "【违规】完全相同底片(同一照片复用)" / "【合规】现场真实重拍(有物理角度/光影变化)" / "【误判】不同商户或场景",
  "reason": "综合店招对比、机位构图、光影反光、陈列细节，详细陈述判定依据"
}"""

# 默认内置的真实门头照识别提示词
DEFAULT_STOREFRONT_PROMPT = """你是一个专业的商户真实性审核与门头照质检专家。
请仔细分析这张照片，判断它是否为【真实的商户门头照】。
你的结论是【提示标签】，不直接作为处罚依据。

【审核规则】：
1. 真实门头标准：实体店铺的临街实景、商业街或商场内部真实门面，包含出入口、门面构造或经营场所外观。
2. 违规与造假类型：翻拍屏幕（摩尔纹、屏幕黑边）、翻拍纸质照片、非门头场景、严重遮挡或模糊、纯色块涂抹或 AI 伪造。
3. 【重要】现场玻璃门反光、逆光、雨天水气很容易被误认为翻拍屏幕。
   若无摩尔纹、无屏幕边框、无像素网格等确凿证据，请判为正常门头并降低置信度，不要仅凭反光就定性造假。

请严格输出合法的 JSON 格式，不要包含任何 markdown 符号或其它文字：
{
  "is_real_storefront": true 或 false,
  "confidence": 0.0 到 1.0 之间的置信度数值,
  "store_name": "识别提取出的门头招牌文字，若无法识别则填'无法识别'",
  "risk_type": "正常门头" / "翻拍屏幕" / "翻拍相片" / "非门头场景" / "招牌缺失" / "模糊或遮挡" / "疑似造假",
  "reason": "详细列出客观分析理由，包括你看到的具体证据"
}"""


class ArkVisionClient:
    def __init__(
        self,
        api_url: str = DEFAULT_API_URL,
        model: str = DEFAULT_MODEL,
        api_key: str = DEFAULT_API_KEY,
        cache_dir: Optional[str] = None,
    ):
        self.api_url = api_url.strip().rstrip("/")
        self.model = model.strip()
        self.api_key = api_key.strip()
        # 结果缓存：同一对照片 + 同一模型 + 同一提示词 不重复计费，断点续跑也靠它
        self.cache_dir = cache_dir or os.path.join(
            os.path.dirname(os.path.abspath(__file__)), ".ai_cache"
        )

    # ---------------- 基础能力 ----------------

    def _get_chat_url(self) -> str:
        if self.api_url.endswith("/chat/completions"):
            return self.api_url
        return f"{self.api_url}/chat/completions"

    def test_connection(self) -> Tuple[bool, str]:
        payload = {
            "model": self.model,
            "messages": [{"role": "user", "content": "ping"}],
            "max_tokens": 10,
        }
        try:
            req = urllib.request.Request(
                self._get_chat_url(),
                headers={
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {self.api_key}",
                },
                data=json.dumps(payload).encode("utf-8"),
            )
            with urllib.request.urlopen(req, timeout=10) as resp:
                if resp.status == 200:
                    return True, "连接成功！模型与 API Key 验证正常。"
                return False, f"连接返回状态码: {resp.status}"
        except urllib.error.HTTPError as e:
            err_body = e.read().decode("utf-8", errors="ignore")
            try:
                msg = json.loads(err_body).get("error", {}).get("message", err_body)
            except Exception:
                msg = err_body
            return False, f"HTTP错误 ({e.code}): {msg}"
        except Exception as e:
            return False, f"连接异常: {str(e)}"

    # ---------------- 水印遮蔽 ----------------

    @staticmethod
    def estimate_watermark_band(
        img: Image.Image,
        min_pct: float = 0.06,
        max_pct: float = 0.28,
        white_threshold: int = 225,
        row_hit_ratio: float = 0.004,
    ) -> Tuple[float, str]:
        """自适应估算底部水印带高度，只遮蔽最小必要区域。

        固定遮蔽 20% 在 3120x4160 竖图上约为 830px，可能把台阶、电瓶车、卷帘门下沿
        这些“偶发物铁证”一起遮掉；而模图或其他批次可能 20% 反而遮不完。
        因此改为按白色水印文字分布估算带高，并回传实际遮蔽比例供审计。
        """
        w, h = img.size
        try:
            gray = img.convert("L")
            start_y = int(h * (1.0 - max_pct))
            region = gray.crop((0, start_y, w, h))
            px = region.load()
            rw, rh = region.size
            step = max(1, rw // 240)
            sampled = max(1, len(range(0, rw, step)))

            seen = False
            gap = 0
            top_y = rh
            gap_limit = max(8, int(h * 0.01))
            for y in range(rh - 1, -1, -1):
                hits = 0
                for x in range(0, rw, step):
                    if px[x, y] >= white_threshold:
                        hits += 1
                if hits / sampled >= row_hit_ratio:
                    seen = True
                    gap = 0
                    top_y = y
                elif seen:
                    gap += 1
                    if gap >= gap_limit:
                        break

            if not seen:
                return min_pct, "未检出水印文字，回落最小遮蔽"

            band_px = (rh - top_y) + int(h * 0.015)
            pct = band_px / float(h)
            pct = max(min_pct, min(max_pct, pct))
            return pct, "自适应检出水印带高约 %dpx（图高 %dpx）" % (band_px, h)
        except Exception as e:
            return min_pct, "估算失败，回落最小遮蔽: %s" % str(e)

    @classmethod
    def encode_image_with_meta(
        cls,
        image_path: str,
        max_dimension: int = 1280,
        mask_watermark: bool = False,
        mask_height_pct: Optional[float] = None,
        mask_mode: str = "adaptive",
        min_mask_pct: float = 0.06,
        max_mask_pct: float = 0.28,
    ) -> Tuple[Optional[str], Dict[str, Any], Optional[str]]:
        """读图 -> 可选遮蔽 -> 缩放 -> base64，同时返回可审计的处理元信息。"""
        meta: Dict[str, Any] = {
            "mask_enabled": bool(mask_watermark),
            "mask_mode": mask_mode if mask_watermark else "none",
            "mask_pct": 0.0,
            "mask_note": "",
            "width": None,
            "height": None,
            "file_sha1": "",
        }

        if not os.path.exists(image_path):
            return None, meta, f"文件不存在: {image_path}"

        try:
            with open(image_path, "rb") as f:
                meta["file_sha1"] = hashlib.sha1(f.read()).hexdigest()
        except Exception:
            pass

        try:
            with Image.open(image_path) as img:
                if img.mode != "RGB":
                    img = img.convert("RGB")

                w, h = img.size
                meta["width"], meta["height"] = w, h

                if mask_watermark:
                    if mask_mode == "fixed":
                        pct = float(mask_height_pct if mask_height_pct else 0.20)
                        note = "固定比例遮蔽"
                    else:
                        pct, note = cls.estimate_watermark_band(
                            img, min_pct=min_mask_pct, max_pct=max_mask_pct
                        )
                        if mask_height_pct:
                            # 传入上限时不超过它，避免一下子遮掉太多画面
                            pct = min(pct, float(mask_height_pct))
                    pct = max(0.0, min(0.5, pct))
                    if pct > 0:
                        draw = ImageDraw.Draw(img)
                        draw.rectangle(
                            [0, int(h * (1.0 - pct)), w, h], fill="#000000"
                        )
                    meta["mask_pct"] = round(pct, 4)
                    meta["mask_note"] = note

                if w < 14 or h < 14:
                    w, h = max(w, 14), max(h, 14)
                    img = img.resize((w, h), Image.Resampling.LANCZOS)

                if max(w, h) > max_dimension:
                    scale = max_dimension / float(max(w, h))
                    img = img.resize(
                        (int(w * scale), int(h * scale)), Image.Resampling.LANCZOS
                    )

                buffer = io.BytesIO()
                img.save(buffer, format="JPEG", quality=85, optimize=True)
                raw = buffer.getvalue()
                meta["masked_sha1"] = hashlib.sha1(raw).hexdigest()
                b64_str = base64.b64encode(raw).decode("utf-8")
                return f"data:image/jpeg;base64,{b64_str}", meta, None
        except Exception as e:
            return None, meta, f"处理图片失败: {str(e)}"

    @classmethod
    def encode_image_to_base64(
        cls,
        image_path: str,
        max_dimension: int = 1280,
        mask_watermark: bool = False,
        mask_height_pct: float = 0.20,
    ) -> Tuple[Optional[str], Optional[str]]:
        """向后兼容的封装（返回两元组）。"""
        b64, _meta, err = cls.encode_image_with_meta(
            image_path,
            max_dimension=max_dimension,
            mask_watermark=mask_watermark,
            mask_height_pct=mask_height_pct,
            mask_mode="fixed",
        )
        return b64, err

    # ---------------- 缓存 ----------------

    def _cache_path(self, key: str) -> str:
        return os.path.join(self.cache_dir, key[:2], key + ".json")

    def _cache_get(self, key: str) -> Optional[Dict[str, Any]]:
        path = self._cache_path(key)
        if not os.path.exists(path):
            return None
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None

    def _cache_put(self, key: str, value: Dict[str, Any]) -> None:
        path = self._cache_path(key)
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "w", encoding="utf-8") as f:
                json.dump(value, f, ensure_ascii=False)
        except Exception:
            pass

    # ---------------- 比对 ----------------

    def _error_result(self, message: str, call_status: str) -> Dict[str, Any]:
        return {
            "is_same_photo": None,
            "confidence": 0.0,
            "verdict": "【异常】" + message[:20],
            "suggested_verdict": "【异常】" + message[:20],
            "has_signboard": "异常",
            "signboard_text_1": "-",
            "signboard_text_2": "-",
            "signboard_match": "异常",
            "scene_similarity": 0.0,
            "reason": message,
            "call_status": call_status,
            "model_version": self.model,
            "prompt_version": COMPARE_PROMPT_VERSION,
            "client_version": CLIENT_VERSION,
            "error": True,
        }

    def compare_images(
        self,
        img_path_a: str,
        img_path_b: str,
        task_a: str = "",
        task_b: str = "",
        custom_prompt: Optional[str] = None,
        mask_watermark: bool = True,
        mask_height_pct: float = 0.20,
        mask_mode: str = "adaptive",
        use_cache: bool = True,
        consistency_check: bool = False,
    ) -> Dict[str, Any]:
        """多维度比对两张照片。

        与 v1 的区别：
        - temperature 固定 0，结果可复现；
        - 自适应遮蔽并记录实际遮蔽比例；
        - 返回值做字段校验，不合规重试一次，失败落表而不静默当成“正常”；
        - 保留原始返回、模型/prompt 版本、图片哈希；
        - 结果写缓存，重跑幂等，支持断点续跑；
        - consistency_check=True 时重复调用一次，抽测结论稳定性。
        """
        b64_a, meta_a, err_a = self.encode_image_with_meta(
            img_path_a,
            mask_watermark=mask_watermark,
            mask_height_pct=mask_height_pct,
            mask_mode=mask_mode,
        )
        if err_a:
            return self._error_result(f"照片A加载失败: {err_a}", "image_error")

        b64_b, meta_b, err_b = self.encode_image_with_meta(
            img_path_b,
            mask_watermark=mask_watermark,
            mask_height_pct=mask_height_pct,
            mask_mode=mask_mode,
        )
        if err_b:
            return self._error_result(f"照片B加载失败: {err_b}", "image_error")

        prompt_tpl = (
            custom_prompt.strip()
            if custom_prompt and custom_prompt.strip()
            else DEFAULT_COMPARE_PROMPT
        )
        prompt = f"{prompt_tpl}\n\n当前核验任务编号参考: 图1任务号={task_a}, 图2任务号={task_b}。"
        prompt_hash = hashlib.sha1(prompt_tpl.encode("utf-8")).hexdigest()[:10]

        cache_key = hashlib.sha1(
            "|".join(
                [
                    self.model,
                    prompt_hash,
                    str(meta_a.get("masked_sha1", meta_a.get("file_sha1", ""))),
                    str(meta_b.get("masked_sha1", meta_b.get("file_sha1", ""))),
                ]
            ).encode("utf-8")
        ).hexdigest()

        mask_pct_used = "A=%.1f%% / B=%.1f%%" % (
            float(meta_a.get("mask_pct", 0.0)) * 100,
            float(meta_b.get("mask_pct", 0.0)) * 100,
        )

        def decorate(res: Dict[str, Any], call_status: str) -> Dict[str, Any]:
            res = dict(res)
            res["suggested_verdict"] = res.get("verdict", "")
            res["call_status"] = res.get("call_status") or call_status
            res["model_version"] = self.model
            res["prompt_version"] = COMPARE_PROMPT_VERSION
            res["prompt_hash"] = prompt_hash
            res["client_version"] = CLIENT_VERSION
            res["temperature"] = FIXED_TEMPERATURE
            res["mask_pct_used"] = mask_pct_used
            res["mask_note"] = "A: %s | B: %s" % (
                meta_a.get("mask_note", ""),
                meta_b.get("mask_note", ""),
            )
            res["photo_sha1_a"] = meta_a.get("file_sha1", "")
            res["photo_sha1_b"] = meta_b.get("file_sha1", "")
            res["masked_sha1_a"] = meta_a.get("masked_sha1", "")
            res["masked_sha1_b"] = meta_b.get("masked_sha1", "")
            res["cache_key"] = cache_key
            if "is_same" in res and res.get("is_same_photo") is None:
                res["is_same_photo"] = res["is_same"]
            return res

        if use_cache:
            cached = self._cache_get(cache_key)
            if cached:
                return decorate(cached, "cache_hit")

        payload = {
            "model": self.model,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url", "image_url": {"url": b64_a}},
                        {"type": "image_url", "image_url": {"url": b64_b}},
                    ],
                }
            ],
            "temperature": FIXED_TEMPERATURE,
            "top_p": FIXED_TOP_P,
            "max_tokens": 800,
        }

        res, status = self._send_vision_request(
            payload, default_key="is_same_photo", required_keys=COMPARE_REQUIRED_KEYS
        )

        if consistency_check and not res.get("error"):
            res2, _ = self._send_vision_request(
                payload,
                default_key="is_same_photo",
                required_keys=COMPARE_REQUIRED_KEYS,
            )
            same_again = res2.get("is_same_photo", res2.get("is_same"))
            res["consistency"] = (
                "一致"
                if same_again == res.get("is_same_photo", res.get("is_same"))
                else "不一致(结论翻转)"
            )

        out = decorate(res, status)
        if use_cache and not out.get("error"):
            self._cache_put(cache_key, res)
        return out

    def inspect_storefront_image(
        self, img_path: str, custom_prompt: Optional[str] = None
    ) -> Dict[str, Any]:
        """单图门头真实性识别。结论作为提示标签，不单独定性。"""
        b64_img, meta, err = self.encode_image_with_meta(img_path)
        if err:
            return {
                "is_real_storefront": None,
                "confidence": 0.0,
                "store_name": "无法识别",
                "risk_type": "读取失败",
                "reason": f"图片读取失败: {err}",
                "call_status": "image_error",
                "model_version": self.model,
                "prompt_version": STOREFRONT_PROMPT_VERSION,
                "error": True,
            }

        prompt = (
            custom_prompt.strip()
            if custom_prompt and custom_prompt.strip()
            else DEFAULT_STOREFRONT_PROMPT
        )

        payload = {
            "model": self.model,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url", "image_url": {"url": b64_img}},
                    ],
                }
            ],
            "temperature": FIXED_TEMPERATURE,
            "top_p": FIXED_TOP_P,
            "max_tokens": 600,
        }

        res, status = self._send_vision_request(
            payload,
            default_key="is_real_storefront",
            required_keys=STOREFRONT_REQUIRED_KEYS,
        )
        res["call_status"] = res.get("call_status") or status
        res["model_version"] = self.model
        res["prompt_version"] = STOREFRONT_PROMPT_VERSION
        res["client_version"] = CLIENT_VERSION
        res["photo_sha1"] = meta.get("file_sha1", "")
        return res

    # ---------------- 请求与校验 ----------------

    @staticmethod
    def _parse_content(content: str) -> Dict[str, Any]:
        content = content.strip()
        if content.startswith("```"):
            lines = content.splitlines()
            if lines and lines[0].startswith("```"):
                lines = lines[1:]
            if lines and lines[-1].startswith("```"):
                lines = lines[:-1]
            content = "\n".join(lines).strip()
        return json.loads(content)

    def _send_once(self, payload: dict) -> Tuple[Optional[Dict[str, Any]], str, str]:
        """发一次请求，返回 (parsed, raw_content, status)。"""
        try:
            req = urllib.request.Request(
                self._get_chat_url(),
                headers={
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {self.api_key}",
                },
                data=json.dumps(payload).encode("utf-8"),
            )
            with urllib.request.urlopen(req, timeout=45) as resp:
                res_data = json.loads(resp.read().decode("utf-8"))
                content = res_data["choices"][0]["message"]["content"].strip()
            try:
                return self._parse_content(content), content, "ok"
            except Exception:
                return None, content, "parse_error"
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="ignore")
            return None, body, "http_%d" % e.code
        except Exception as e:
            return None, str(e), "request_error"

    def _send_vision_request(
        self,
        payload: dict,
        default_key: str,
        required_keys: Tuple[str, ...] = (),
    ) -> Tuple[Dict[str, Any], str]:
        strict_hint = (
            "\n\n【格式约束】上一次输出不合规。请只输出一个合法 JSON 对象，"
            "包含全部必填字段，不要添加任何解释文字。"
        )

        attempts = []
        for attempt in range(2):
            body = payload
            if attempt == 1:
                body = json.loads(json.dumps(payload))
                try:
                    body["messages"][0]["content"][0]["text"] += strict_hint
                except Exception:
                    pass
                time.sleep(0.5)

            parsed, raw, status = self._send_once(body)
            attempts.append(status)

            if parsed is not None:
                missing = [k for k in required_keys if k not in parsed]
                if not missing:
                    parsed["error"] = False
                    parsed["raw_response"] = raw[:2000]
                    parsed["attempts"] = attempts
                    return parsed, ("retry_ok" if attempt else "ok")
                status = "schema_invalid(缺%s)" % ",".join(missing)
                attempts[-1] = status

            # http 4xx 不重试（除 429）
            if status.startswith("http_4") and not status.startswith("http_429"):
                break

        return (
            {
                default_key: None,
                "confidence": 0.0,
                "verdict": "【异常】调用或解析失败",
                "risk_type": "调用或解析失败",
                "reason": "尝试记录: %s；原始返回(截断): %s"
                % (" -> ".join(attempts), str(raw)[:500]),
                "raw_response": str(raw)[:2000],
                "attempts": attempts,
                "error": True,
            },
            attempts[-1],
        )
