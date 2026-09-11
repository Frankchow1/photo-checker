import os
import pandas as pd
from PIL import Image, ImageDraw

base_dir = "/Users/jun/.gemini/antigravity/scratch/shaoxing_photo_checker"
photos_dir = os.path.join(base_dir, "test_photos")
os.makedirs(photos_dir, exist_ok=True)

# 1. 生成几张不同场景的测试图片 (高清晰度 400x300，避免过于简单)
def create_test_image(filename, color, shape_type, text):
    img = Image.new("RGB", (400, 300), color=color)
    draw = ImageDraw.Draw(img)
    if shape_type == "circle":
        draw.ellipse([80, 50, 320, 250], fill="white", outline="yellow", width=4)
    elif shape_type == "rectangle":
        draw.rectangle([60, 40, 340, 260], fill="lightblue", outline="navy", width=4)
    elif shape_type == "triangle":
        draw.polygon([(200, 40), (80, 260), (320, 260)], fill="orange", outline="brown")
    
    # 模拟拍摄水印或文字
    draw.text((20, 20), text, fill="black")
    path = os.path.join(photos_dir, filename)
    img.save(path, "JPEG")
    return path

# 组1：完全相同的照片 (task_001 与 task_002 相同)
create_test_image("SX_TASK_001.jpg", "green", "circle", "Shaoxing Check Point A")
create_test_image("SX_TASK_002.jpg", "green", "circle", "Shaoxing Check Point A")

# 组2：微小裁剪/高度相似 (task_003 与 task_004)
create_test_image("SX_TASK_003.jpg", "darkblue", "rectangle", "Shaoxing Bridge #1")
create_test_image("SX_TASK_004.jpg", "darkblue", "rectangle", "Shaoxing Bridge #1 (Crop)")

# 组3：明显不同 (哈希算法误判案例，task_005 与 task_006)
create_test_image("SX_TASK_005.jpg", "darkred", "circle", "Red Circle Gate")
create_test_image("SX_TASK_006.jpg", "purple", "triangle", "Purple Triangle Tower")

# 2. 生成模拟绍兴相同照片 Excel 表格
excel_rows = [
    {
        "重复组号": "GROUP_01",
        "原任务号": "SX_2026_0911_001",
        "原照片名称": "SX_TASK_001.jpg",
        "对比任务号": "SX_2026_0911_002",
        "对比照片名称": "SX_TASK_002.jpg",
        "初始哈希距离": 0
    },
    {
        "重复组号": "GROUP_02",
        "原任务号": "SX_2026_0911_003",
        "原照片名称": "SX_TASK_003.jpg",
        "对比任务号": "SX_2026_0911_004",
        "对比照片名称": "SX_TASK_004.jpg",
        "初始哈希距离": 2
    },
    {
        "重复组号": "GROUP_03",
        "原任务号": "SX_2026_0911_005",
        "原照片名称": "SX_TASK_005.jpg",
        "对比任务号": "SX_2026_0911_006",
        "对比照片名称": "SX_TASK_006.jpg",
        "初始哈希距离": 4
    }
]

df = pd.DataFrame(excel_rows)
excel_path = os.path.join(base_dir, "绍兴相同照片哈希初筛清单_测试样例.xlsx")
df.to_excel(excel_path, index=False)
print(f"测试数据生成完成:\nExcel: {excel_path}\n照片目录: {photos_dir}")
