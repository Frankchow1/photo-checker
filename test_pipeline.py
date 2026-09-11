import os
from data_processor import PhotoIndex, ExcelTaskParser, export_results_to_excel
from ai_client import ArkVisionClient

base_dir = "/Users/jun/.gemini/antigravity/scratch/shaoxing_photo_checker"
excel_file = os.path.join(base_dir, "绍兴相同照片哈希初筛清单_测试样例.xlsx")
photos_dir = os.path.join(base_dir, "test_photos")

print("1. 正在索引照片目录...")
indexer = PhotoIndex(photos_dir)
print("   已索引照片数量:", len(indexer.by_filename))

print("2. 正在解析 Excel 任务...")
parser = ExcelTaskParser(excel_file)
pairs = parser.parse_pairs(photo_indexer=indexer)
print("   解析出对比组数:", len(pairs))
for p in pairs:
    ta = p["task_a"]
    tb = p["task_b"]
    pa = bool(p["path_a"])
    pb = bool(p["path_b"])
    print(f"   - 任务对: {ta} vs {tb}, 路径A就绪: {pa}, 路径B就绪: {pb}")

print("3. 调用火山引擎 Vision 进行真实比对测试...")
client = ArkVisionClient()
results = []
for p in pairs:
    res = client.compare_images(
        img_path_a=p["path_a"],
        img_path_b=p["path_b"],
        task_a=p["task_a"],
        task_b=p["task_b"]
    )
    merged = {**p, **res}
    results.append(merged)
    is_same = merged.get("is_same")
    conf = merged.get("confidence")
    reason = merged.get("reason", "")[:80]
    print(f"   [比对结果] {p['task_a']} vs {p['task_b']}: 判定相同={is_same}, 置信度={conf}")
    print(f"     原因摘要: {reason}...")

print("4. 测试导出 Excel 结果报告...")
out_excel = os.path.join(base_dir, "绍兴抽检结果导出报告_测试.xlsx")
export_results_to_excel(results, out_excel)
print("   导出成功:", os.path.exists(out_excel), "文件大小:", os.path.getsize(out_excel), "字节")
