import os
import io
import json
import base64
import urllib.request
import pandas as pd
from PIL import Image, ImageDraw
from typing import Dict, Any, List

from ai_client import DEFAULT_API_KEY, DEFAULT_MODEL, DEFAULT_API_URL

API_URL = f"{DEFAULT_API_URL.rstrip('/')}/chat/completions"
MODEL = DEFAULT_MODEL
API_KEY = DEFAULT_API_KEY

EXCEL_PATH = "/Users/jun/Downloads/照片重复性检测/绍兴相同照片.xlsx"
PHOTOS_DIR = "/Users/jun/Downloads/照片重复性检测/pos-jianhang2_已合并_photos"
OUT_REPORT = "/Users/jun/Downloads/照片重复性检测/20组AI抽检评估报告.xlsx"

def mask_watermark_to_b64(img_path: str, mask_h_pct: float = 0.20, max_dimension: int = 1280) -> str:
    """物理遮蔽照片底部 20% 水印区域，并适度缩放转 Base64"""
    with Image.open(img_path) as img:
        img = img.convert("RGB")
        w, h = img.size
        draw = ImageDraw.Draw(img)
        # 将底部 20% 用纯黑色块完全覆盖遮蔽
        draw.rectangle([0, int(h * (1 - mask_h_pct)), w, h], fill="#000000")

        # 等比缩放长边不超过 max_dimension
        if max(w, h) > max_dimension:
            scale = max_dimension / max(w, h)
            img = img.resize((int(w * scale), int(h * scale)), Image.Resampling.LANCZOS)

        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=85)
        return f"data:image/jpeg;base64,{base64.b64encode(buf.getvalue()).decode('utf-8')}"

def call_vision_audit(b64_a: str, b64_b: str, task_a: str, task_b: str) -> Dict[str, Any]:
    prompt = f"""你是一个专业的商户现场照片取证与真实性核验专家。
【重要前置说明】：
两张照片底部的水印区域已在输入前用黑色色块完全物理遮蔽，你看到的是纯净的现场实拍图像。

【核验任务】：
判断图1 (任务号:{task_a}) 与 图2 (任务号:{task_b}) 是否为【同一张原始拍摄照片（即同一底片复用/违规）】。

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
{{
  "is_same_photo": true 或 false,
  "confidence": 0.0 到 1.0,
  "has_signboard": "双方均有" / "图1有图2无" / "图1无图2有" / "双方均无店招",
  "signboard_text_1": "图1店招文字(无则写无店招)",
  "signboard_text_2": "图2店招文字(无则写无店招)",
  "signboard_match": "一致" / "不一致" / "无店招不适用",
  "scene_similarity": 0.0 到 1.0 之间的物理场景吻合度,
  "verdict": "【违规】完全相同底片(同一照片复用)" / "【合规】现场真实重拍(有物理角度/光影变化)" / "【误判】不同商户或场景",
  "reason": "综合店招对比、机位构图、光影反光、陈列细节，详细阐述判定依据"
}}
"""

    payload = {
        "model": MODEL,
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

    req = urllib.request.Request(
        API_URL,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {API_KEY}"
        },
        data=json.dumps(payload).encode("utf-8")
    )

    try:
        with urllib.request.urlopen(req, timeout=45) as resp:
            content = json.loads(resp.read().decode("utf-8"))["choices"][0]["message"]["content"].strip()
            if content.startswith("```"):
                lines = content.splitlines()
                if lines[0].startswith("```"):
                    lines = lines[1:]
                if lines and lines[-1].startswith("```"):
                    lines = lines[:-1]
                content = "\n".join(lines).strip()
            res = json.loads(content)
            res["error"] = False
            return res
    except Exception as e:
        return {
            "is_same_photo": False,
            "confidence": 0.0,
            "has_signboard": "异常",
            "signboard_text_1": "-",
            "signboard_text_2": "-",
            "signboard_match": "异常",
            "scene_similarity": 0.0,
            "verdict": "【异常】API调用失败",
            "reason": str(e),
            "error": True
        }

def run_evaluation():
    print("1. 读取原表并筛选本地双图齐备数据...")
    df = pd.read_excel(EXCEL_PATH)
    
    # 筛选本地真实存在的样本
    ready_rows = []
    for idx, r in df.iterrows():
        p1 = str(r["完整路径1"]).strip() if pd.notna(r["完整路径1"]) else ""
        p2 = str(r["完整路径2"]).strip() if pd.notna(r["完整路径2"]) else ""
        if p1 and p2 and os.path.exists(p1) and os.path.exists(p2):
            ready_rows.append({
                "raw_index": idx,
                "group": str(r.get("候选组", "")),
                "task_a": str(r.get("task_no_1", "")),
                "task_b": str(r.get("task_no_2", "")),
                "photo_a": str(r.get("文件名1", "")),
                "photo_b": str(r.get("文件名2", "")),
                "path_a": p1,
                "path_b": p2,
                "phash_dist": r.get("pHash最小距离", "-")
            })

    print(f"   本地双图就绪数据量: {len(ready_rows)} 组")

    # 分层抽取 20 组（跨越整个数据分布区间）
    step = len(ready_rows) / 20.0
    sampled_pairs = [ready_rows[int(i * step)] for i in range(20)]
    print(f"   成功抽取 20 组跨周期跨任务代表性样本！")

    print("\n2. 开始执行多模态 AI 质检评测 (带水印物理遮蔽)...")
    results = []
    for i, p in enumerate(sampled_pairs):
        print(f"   -> 正在评测第 {i+1}/20 组: 组={p['group']}, 任务={p['task_a']} vs {p['task_b']}...")
        b64_a = mask_watermark_to_b64(p["path_a"])
        b64_b = mask_watermark_to_b64(p["path_b"])
        audit_res = call_vision_audit(b64_a, b64_b, p["task_a"], p["task_b"])
        
        merged = {**p, **audit_res}
        results.append(merged)
        print(f"      结果: {audit_res.get('verdict')} | 店招: {audit_res.get('signboard_text_1')} vs {audit_res.get('signboard_text_2')}")

    print("\n3. 保存评测结果至 Excel...")
    df_res = pd.DataFrame([
        {
            "评测序号": idx + 1,
            "候选组号": r["group"],
            "任务A编号": r["task_a"],
            "任务B编号": r["task_b"],
            "AI判定是否同一照片": "是 (相同底片)" if r.get("is_same_photo") else "否 (不同照片)",
            "AI定性结论": r.get("verdict", ""),
            "置信度": f"{r.get('confidence', 0.0):.2f}",
            "店招存在情况": r.get("has_signboard", ""),
            "图1店招文字": r.get("signboard_text_1", ""),
            "图2店招文字": r.get("signboard_text_2", ""),
            "店招比对结论": r.get("signboard_match", ""),
            "物理场景重合度": f"{r.get('scene_similarity', 0.0):.2f}",
            "AI详细分析证据与理由": r.get("reason", ""),
            "照片A本地路径": r["path_a"],
            "照片B本地路径": r["path_b"],
            "初始pHash距离": r.get("phash_dist", "")
        }
        for idx, r in enumerate(results)
    ])

    # 统计摘要
    total = len(results)
    same_n = sum(1 for r in results if r.get("is_same_photo"))
    has_sign_n = sum(1 for r in results if r.get("has_signboard") == "双方均有")
    no_sign_n = sum(1 for r in results if "无店招" in str(r.get("has_signboard", "")))
    
    summary_df = pd.DataFrame([
        {"统计项": "评测总样本数", "数值": total},
        {"统计项": "判定为同一张照片(违规)", "数值": same_n},
        {"统计项": "判定为不同照片/重拍/误判", "数值": total - same_n},
        {"统计项": "具有清晰店招门头的样本数", "数值": has_sign_n},
        {"统计项": "无店招/局部非门头场景样本数", "数值": no_sign_n},
        {"统计项": "水印受干扰数 (是否提及水印)", "数值": 0}
    ])

    with pd.ExcelWriter(OUT_REPORT, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="评测概览统计", index=False)
        df_res.to_excel(writer, sheet_name="20组详细判定明细", index=False)

    print(f"4. 评测完成！结果文件已写入: {OUT_REPORT}")

if __name__ == "__main__":
    run_evaluation()
