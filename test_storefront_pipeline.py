import os
from PIL import Image, ImageDraw
from ai_client import ArkVisionClient, DEFAULT_STOREFRONT_PROMPT
from data_processor import collect_images_from_dir, export_storefront_results_to_excel

base_dir = "/Users/jun/.gemini/antigravity/scratch/shaoxing_photo_checker"
storefront_test_dir = os.path.join(base_dir, "test_storefront_photos")
os.makedirs(storefront_test_dir, exist_ok=True)

# 1. 制造一张符合商户实体门头特征的测试图
img1 = Image.new("RGB", (640, 480), color="#cbd5e1")
d1 = ImageDraw.Draw(img1)
# 门头招牌
d1.rectangle([40, 30, 600, 130], fill="#dc2626")
d1.text((100, 65), "Shaoxing Traditional Rice Wine Store", fill="#fef08a")
# 店面大门
d1.rectangle([120, 140, 320, 460], fill="#1e293b")
d1.rectangle([360, 180, 560, 380], fill="#38bdf8")
img1.save(os.path.join(storefront_test_dir, "storefront_sample_01.jpg"), "JPEG")

# 2. 制造一张非门头图 (例如纯室内商品特写)
img2 = Image.new("RGB", (640, 480), color="#1e1e1e")
d2 = ImageDraw.Draw(img2)
d2.ellipse([200, 150, 440, 390], fill="gold")
d2.text((220, 260), "Only Goods Specimen", fill="black")
img2.save(os.path.join(storefront_test_dir, "non_storefront_sample_02.jpg"), "JPEG")

print("1. 扫描测试目录中的门头照...")
items = collect_images_from_dir(storefront_test_dir)
print("   扫描到图片数量:", len(items))

print("2. 调用火山引擎进行门头照识别测试 (使用预设提示词)...")
client = ArkVisionClient()
results = []
for it in items:
    res = client.inspect_storefront_image(it["path"], custom_prompt=DEFAULT_STOREFRONT_PROMPT)
    merged = {**it, **res}
    results.append(merged)
    print(f"   [识别结果] {it['photo_name']}:")
    print(f"     是否真实门头: {merged.get('is_real_storefront')}")
    print(f"     风险分类: {merged.get('risk_type')}")
    print(f"     提取店名: {merged.get('store_name')}")
    print(f"     分析理由: {merged.get('reason', '')[:65]}...")

print("3. 测试导出门头质检 Excel 报表...")
out_excel = os.path.join(base_dir, "门头照质检报告_测试.xlsx")
export_storefront_results_to_excel(results, out_excel)
print("   导出完成:", os.path.exists(out_excel), "文件大小:", os.path.getsize(out_excel), "字节")
