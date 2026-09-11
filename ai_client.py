import os
import io
import json
import base64
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

# 默认内置的多维证据链相同照片比对提示词
DEFAULT_COMPARE_PROMPT = """你是一个专业的商户现场照片取证与真实性核验专家。
【重要前置说明】：
如果已启用物理去水印，照片底部的水印区域已在输入前用黑色色块完全物理遮蔽，你看到的是纯净的现场实拍图像。即使未遮蔽，外勤软件拍照必然叠加不同水印，【绝对严禁关注水印，严禁以水印不同作为判定不同的依据】！

【核验任务】：
综合多维物理证据，判断图1 与 图2 是否为【同一张原始拍摄照片（即同一底片复用/违规）】。

【多维度综合取证依据】：
1. 【店招维度】（注：若拍摄的是室内/工位/陈列/无招牌门面，请标明无店招，店招缺失不影响其他维度的同一性判定）：
   - 图1与图2是否存在门头招牌/店名？若有，文字是否一致？
2. 【机位与透视几何】：
   - 拍摄机位、透视角度、画框边缘截断处是否完全一致？（手持手机重新拍摄必然存在视角与相对遮挡的物理偏差）
3. 【光影与反光特征】：
   - 太阳/室内光源照射角度、阴影走向与长度、玻璃及金属表面的反光光斑是否完全重叠？
4. 【瞬态陈列与偶发动态物】：
   - 门口临时停放车辆、杂物纸箱、促销海报、花篮横幅等临时摆放物的相对位置和形态是否完全静态一致？
5. 【微观固有细节】：
   - 卷帘门划痕凹痕、墙体破损、瓷砖缝隙及特定污渍是否吻合？

【容差规则】：请忽略微小的数字压缩噪点或轻微色差，不要把 JPEG 压缩失真误判为物理差异。

请严格返回合法 JSON 格式：
{
  "is_same_photo": true 或 false,
  "confidence": 0.0 到 1.0 之间的置信度数值,
  "has_signboard": "双方均有" / "图1有图2无" / "图1无图2有" / "双方均无店招",
  "signboard_text_1": "图1店招文字(无则写无店招)",
  "signboard_text_2": "图2店招文字(无则写无店招)",
  "signboard_match": "一致" / "不一致" / "无店招不适用",
  "scene_similarity": 0.0 到 1.0 之间的物理场景吻合度,
  "verdict": "【违规】完全相同底片(同一照片复用)" / "【合规】现场真实重拍(有物理角度/光影变化)" / "【误判】不同商户或场景",
  "reason": "综合店招对比、机位构图、光影反光、陈列细节，详细阐述判定依据"
}"""

# 默认内置的真实门头照识别提示词
DEFAULT_STOREFRONT_PROMPT = """你是一个专业的商户真实性审核与门头照质检专家。
请仔细分析这张照片，严格判断它是否为【真实的商户门头照】。

【审核规则】：
1. 真实门头标准：必须为实体店铺的临街实景、商业街或商场内部真实门面，画面包含店铺出入口、门面构造或经营场所外观，并具有清晰可辨识的门头招牌（店名文字/Logo）。
2. 违规与造假类型（若存在以下任意情况，必须判定为非真实门头 is_real_storefront: false）：
   - 翻拍屏幕：翻拍电脑、手机或iPad屏幕（存在摩尔纹、反光光斑、屏幕黑边等）；
   - 翻拍相片：翻拍洗印照片或纸质宣传页；
   - 非门头场景：拍摄纯室内、局部货物商品、收银台、营业执照、住宅防盗门、施工白墙等非商户实体门面；
   - 严重遮挡或模糊：招牌被大面积遮挡、严重失焦无法识别主体；
   - 虚假造假：纯色块涂抹、AI生成伪造。

请严格输出合法的 JSON 格式，不要包含任何 markdown 符号或其它文字：
{
  "is_real_storefront": true 或 false,
  "confidence": 0.0 到 1.0 之间的置信度数值,
  "store_name": "识别提取出的门头招牌文字，若无法识别则填'无法识别'",
  "risk_type": "正常门头" / "翻拍屏幕" / "翻拍相片" / "非门头场景" / "招牌缺失" / "模糊或遮挡" / "疑似造假",
  "reason": "详细列出对照片拍摄环境、是否实景实拍、门头招牌文字及细节判定的客观分析理由"
}"""

class ArkVisionClient:
    def __init__(self, api_url: str = DEFAULT_API_URL, model: str = DEFAULT_MODEL, api_key: str = DEFAULT_API_KEY):
        self.api_url = api_url.strip().rstrip("/")
        self.model = model.strip()
        self.api_key = api_key.strip()

    def _get_chat_url(self) -> str:
        if self.api_url.endswith("/chat/completions"):
            return self.api_url
        return f"{self.api_url}/chat/completions"

    def test_connection(self) -> Tuple[bool, str]:
        payload = {
            "model": self.model,
            "messages": [
                {"role": "user", "content": "ping"}
            ],
            "max_tokens": 10
        }
        try:
            req = urllib.request.Request(
                self._get_chat_url(),
                headers={
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {self.api_key}"
                },
                data=json.dumps(payload).encode("utf-8")
            )
            with urllib.request.urlopen(req, timeout=10) as resp:
                if resp.status == 200:
                    return True, "连接成功！模型与 API Key 验证正常。"
                return False, f"连接返回状态码: {resp.status}"
        except urllib.error.HTTPError as e:
            err_body = e.read().decode("utf-8", errors="ignore")
            try:
                err_json = json.loads(err_body)
                msg = err_json.get("error", {}).get("message", err_body)
            except Exception:
                msg = err_body
            return False, f"HTTP错误 ({e.code}): {msg}"
        except Exception as e:
            return False, f"连接异常: {str(e)}"

    @staticmethod
    def encode_image_to_base64(
        image_path: str,
        max_dimension: int = 1280,
        mask_watermark: bool = False,
        mask_height_pct: float = 0.20
    ) -> Tuple[Optional[str], Optional[str]]:
        """读取本地图片，支持自适应物理遮蔽底部水印，适度等比缩放并转为 Base64"""
        if not os.path.exists(image_path):
            return None, f"文件不存在: {image_path}"

        try:
            with Image.open(image_path) as img:
                if img.mode in ("RGBA", "P"):
                    img = img.convert("RGB")
                elif img.mode != "RGB":
                    img = img.convert("RGB")

                w, h = img.size

                # 如果开启水印物理遮蔽，将底部 20% 区域使用纯黑色块完全覆盖
                if mask_watermark and mask_height_pct > 0:
                    draw = ImageDraw.Draw(img)
                    draw.rectangle([0, int(h * (1.0 - mask_height_pct)), w, h], fill="#000000")

                # 火山引擎要求最小 14x14
                if w < 14 or h < 14:
                    new_w = max(w, 14)
                    new_h = max(h, 14)
                    img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                    w, h = new_w, new_h

                # 等比缩放长边不超过 max_dimension
                if max(w, h) > max_dimension:
                    scale = max_dimension / max(w, h)
                    new_size = (int(w * scale), int(h * scale))
                    img = img.resize(new_size, Image.Resampling.LANCZOS)

                buffer = io.BytesIO()
                img.save(buffer, format="JPEG", quality=85, optimize=True)
                b64_str = base64.b64encode(buffer.getvalue()).decode("utf-8")
                return f"data:image/jpeg;base64,{b64_str}", None
        except Exception as e:
            return None, f"处理图片失败: {str(e)}"

    def compare_images(
        self,
        img_path_a: str,
        img_path_b: str,
        task_a: str = "",
        task_b: str = "",
        custom_prompt: Optional[str] = None,
        mask_watermark: bool = True,
        mask_height_pct: float = 0.20
    ) -> Dict[str, Any]:
        """多维度综合对比两张照片（支持前置物理去水印遮蔽与自定义提示词）"""
        b64_a, err_a = self.encode_image_to_base64(
            img_path_a,
            mask_watermark=mask_watermark,
            mask_height_pct=mask_height_pct
        )
        if err_a:
            return {
                "is_same_photo": False,
                "confidence": 0.0,
                "verdict": "【异常】读取失败",
                "has_signboard": "异常",
                "signboard_text_1": "-",
                "signboard_text_2": "-",
                "signboard_match": "异常",
                "scene_similarity": 0.0,
                "reason": f"照片A加载失败: {err_a}",
                "error": True
            }

        b64_b, err_b = self.encode_image_to_base64(
            img_path_b,
            mask_watermark=mask_watermark,
            mask_height_pct=mask_height_pct
        )
        if err_b:
            return {
                "is_same_photo": False,
                "confidence": 0.0,
                "verdict": "【异常】读取失败",
                "has_signboard": "异常",
                "signboard_text_1": "-",
                "signboard_text_2": "-",
                "signboard_match": "异常",
                "scene_similarity": 0.0,
                "reason": f"照片B加载失败: {err_b}",
                "error": True
            }

        prompt_tpl = custom_prompt.strip() if custom_prompt and custom_prompt.strip() else DEFAULT_COMPARE_PROMPT
        prompt = f"{prompt_tpl}\n\n当前核验任务编号参考: 图1任务号={task_a}, 图2任务号={task_b}。"

        payload = {
            "model": self.model,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url", "image_url": {"url": b64_a}},
                        {"type": "image_url", "image_url": {"url": b64_b}}
                    ]
                }
            ],
            "temperature": 0.1,
            "max_tokens": 800
        }

        res = self._send_vision_request(payload, default_key="is_same_photo")
        # 字段规范化回落
        if "is_same" in res and "is_same_photo" not in res:
            res["is_same_photo"] = res["is_same"]
        return res

    def inspect_storefront_image(self, img_path: str, custom_prompt: Optional[str] = None) -> Dict[str, Any]:
        """识别单张照片是否为真实的商户门头照（支持自定义提示词）"""
        b64_img, err = self.encode_image_to_base64(img_path)
        if err:
            return {
                "is_real_storefront": False,
                "confidence": 0.0,
                "store_name": "无法识别",
                "risk_type": "读取失败",
                "reason": f"图片读取失败: {err}",
                "error": True
            }

        prompt = custom_prompt.strip() if custom_prompt and custom_prompt.strip() else DEFAULT_STOREFRONT_PROMPT

        payload = {
            "model": self.model,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url", "image_url": {"url": b64_img}}
                    ]
                }
            ],
            "temperature": 0.1,
            "max_tokens": 600
        }

        res = self._send_vision_request(payload, default_key="is_real_storefront")
        return res

    def _send_vision_request(self, payload: dict, default_key: str) -> Dict[str, Any]:
        try:
            req = urllib.request.Request(
                self._get_chat_url(),
                headers={
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {self.api_key}"
                },
                data=json.dumps(payload).encode("utf-8")
            )
            with urllib.request.urlopen(req, timeout=45) as resp:
                res_data = json.loads(resp.read().decode("utf-8"))
                content = res_data["choices"][0]["message"]["content"].strip()
                
                # 清洗 markdown 代码块
                if content.startswith("```"):
                    lines = content.splitlines()
                    if lines[0].startswith("```"):
                        lines = lines[1:]
                    if lines and lines[-1].startswith("```"):
                        lines = lines[:-1]
                    content = "\n".join(lines).strip()

                parsed = json.loads(content)
                parsed["error"] = False
                return parsed
        except urllib.error.HTTPError as e:
            err_body = e.read().decode("utf-8", errors="ignore")
            return {
                default_key: False,
                "confidence": 0.0,
                "verdict": "【异常】API调用失败",
                "risk_type": "API调用失败",
                "reason": f"HTTP {e.code}: {err_body}",
                "error": True
            }
        except Exception as e:
            return {
                default_key: False,
                "confidence": 0.0,
                "verdict": "【异常】解析失败",
                "risk_type": "解析失败",
                "reason": f"请求或解析失败: {str(e)}",
                "error": True
            }
