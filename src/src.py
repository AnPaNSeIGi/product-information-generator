# =============================================================================
# 图片预处理工具（重命名、缩放、加水印）
# 用于配合 pig.py（乐天商品上架自动化系统）
# =============================================================================

import os
import shutil
import re
from PIL import Image
from datetime import datetime   # 新增，用于生成时间戳

# =============================================================================
# ================= 动态路径（自动定位项目根目录） =================
# =============================================================================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)

# =============================================================================
# ================= 用户配置区（请根据实际情况修改） =================
# =============================================================================

# 源目录：存放商品子文件夹（每个子文件夹放一个商品的原始图片）
SOURCE_DIR = os.path.join(PROJECT_ROOT, "data", "source")

# 备份目录（处理前自动备份到此），设为 None 则不备份
BACKUP_DIR = os.path.join(PROJECT_ROOT, "data", "backup")

# 支持的图片扩展名
IMAGE_EXTS = ('.jpg', '.jpeg', '.png', '.gif', '.bmp')

# 子文件夹名称正则：数字 + 点 + 任意内容（仅用于识别需要处理的文件夹，不再用于命名）
FOLDER_PATTERN = re.compile(r'^(\d+)\..+')

# 水印配置（如果不需要水印，设为空列表 []）
# 只使用第一个有效水印配置，直接覆盖原图。
WATERMARK_CONFIGS = [
    {
        "path": os.path.join(PROJECT_ROOT, "data", "watermark.png"),
        "scale": 0.3,            # 水印宽度占图片宽度的比例
        "opacity": 0.7           # 透明度（0.0~1.0）
    },
]

# 缩放到统一宽度（像素），None 表示不缩放
TARGET_WIDTH = 800

# =============================================================================
# ================= 以下代码请勿修改 =================
# =============================================================================

if BACKUP_DIR:
    os.makedirs(BACKUP_DIR, exist_ok=True)

def backup_folder(src, dst):
    """备份整个文件夹"""
    if not dst:
        return
    if os.path.exists(dst):
        shutil.rmtree(dst)
    shutil.copytree(src, dst)
    print(f"✅ 备份完成：{dst}")

def rename_files_in_subfolder(folder_path, y_value=None):
    """
    对单个子文件夹内的图片进行重命名（按时间戳+序号）
    y_value: 保留参数但不再使用（兼容旧调用）
    """
    files = [f for f in os.listdir(folder_path) 
             if os.path.isfile(os.path.join(folder_path, f)) and f.lower().endswith(IMAGE_EXTS)]
    if not files:
        return

    files.sort()  # 自然排序

    # 生成当前时间戳（精确到秒），保证该文件夹内所有图片共享同一时间戳
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    for idx, filename in enumerate(files, start=1):
        file_path = os.path.join(folder_path, filename)
        _, ext = os.path.splitext(filename)
        ext = ext.lower()

        # 新文件名：时间戳_三位序号.扩展名
        new_name = f"{timestamp}_{idx:03d}{ext}"
        new_path = os.path.join(folder_path, new_name)

        # 防冲突（极少数情况）
        counter = 1
        while os.path.exists(new_path):
            base, ext2 = os.path.splitext(new_name)
            new_path = os.path.join(folder_path, f"{base}_{counter}{ext2}")
            counter += 1

        os.rename(file_path, new_path)
        print(f"  重命名：{filename} → {os.path.basename(new_path)}")

def adjust_opacity(watermark, opacity):
    """
    调整 RGBA 图片的透明度（兼容新旧 Pillow 版本）
    """
    if opacity >= 1:
        return watermark
    r, g, b, a = watermark.split()
    a = a.point(lambda i: i * opacity)
    return Image.merge('RGBA', (r, g, b, a))

def process_images(folder_path):
    """
    对子文件夹内所有图片：缩放，如有水印则添加水印（仅第一个配置），
    直接覆盖原图（备份已存在，安全）。
    """
    # 获取所有图片
    files = [f for f in os.listdir(folder_path) 
             if os.path.isfile(os.path.join(folder_path, f)) and f.lower().endswith(IMAGE_EXTS)]

    # 检查水印配置（只取第一个有效配置）
    use_watermark = False
    watermark_cfg = None
    if WATERMARK_CONFIGS:
        cfg = WATERMARK_CONFIGS[0]
        if os.path.isfile(cfg["path"]):
            use_watermark = True
            watermark_cfg = cfg
        else:
            print(f"⚠️ 水印文件不存在：{cfg['path']}，仅缩放不加水印")

    for filename in files:
        file_path = os.path.join(folder_path, filename)
        try:
            with Image.open(file_path) as img:
                # 缩放
                w, h = img.size
                if TARGET_WIDTH:
                    new_h = int(h * TARGET_WIDTH / w)
                    img_resized = img.resize((TARGET_WIDTH, new_h), Image.LANCZOS)
                else:
                    img_resized = img.copy()

                if use_watermark:
                    # 加载水印
                    watermark = Image.open(watermark_cfg["path"])
                    if watermark.mode != 'RGBA':
                        watermark = watermark.convert('RGBA')

                    # 调整透明度
                    opacity = watermark_cfg.get("opacity")
                    if opacity is not None and 0 <= opacity < 1:
                        watermark = adjust_opacity(watermark, opacity)

                    # 缩放水印
                    scale = watermark_cfg.get("scale")
                    if scale is not None and 0 < scale < 1:
                        wm_w, wm_h = watermark.size
                        new_wm_w = int(img_resized.width * scale)
                        new_wm_h = int(wm_h * new_wm_w / wm_w)
                        watermark = watermark.resize((new_wm_w, new_wm_h), Image.LANCZOS)

                    # 复制背景并转为 RGBA
                    img_work = img_resized.copy()
                    if img_work.mode != 'RGBA':
                        img_work = img_work.convert('RGBA')

                    # 粘贴水印（右下角）
                    wm_w, wm_h = watermark.size
                    pos_x = max(0, img_work.width - wm_w)
                    pos_y = max(0, img_work.height - wm_h)
                    img_work.paste(watermark, (pos_x, pos_y), watermark)

                    # 根据原格式转换色彩模式
                    ext = os.path.splitext(filename)[1].lower()
                    if ext in ('.jpg', '.jpeg', '.bmp'):
                        img_work = img_work.convert('RGB')

                    # 直接覆盖原图
                    img_work.save(file_path)
                    print(f"  处理图片：{filename} (添加水印，覆盖原图)")
                else:
                    # 无水印，仅缩放，覆盖原图
                    img_resized.save(file_path)
                    print(f"  处理图片：{filename} (仅缩放)")

        except Exception as e:
            print(f"  处理失败 {filename}：{e}")

def main():
    print("=" * 60)
    print("🖼️ 图片预处理工具（重命名、缩放、加水印）")
    print("=" * 60)

    if not os.path.isdir(SOURCE_DIR):
        print(f"❌ 源文件夹不存在：{SOURCE_DIR}")
        return

    if BACKUP_DIR:
        backup_folder(SOURCE_DIR, BACKUP_DIR)

    print("\n========== 开始重命名 ==========")
    for item in os.listdir(SOURCE_DIR):
        sub_path = os.path.join(SOURCE_DIR, item)
        if not os.path.isdir(sub_path):
            continue
        match = FOLDER_PATTERN.match(item)
        if match:
            y = match.group(1)  # 仅用于打印信息，不再用于命名
            print(f"\n📁 处理子文件夹：{item} (y={y})")
            rename_files_in_subfolder(sub_path)  # 不再传递 y
        else:
            print(f"\n⏩ 跳过非标准文件夹（不符合 数字.名称）：{item}")

    print("\n========== 开始处理图片 ==========")
    for item in os.listdir(SOURCE_DIR):
        sub_path = os.path.join(SOURCE_DIR, item)
        if not os.path.isdir(sub_path):
            continue
        if FOLDER_PATTERN.match(item):
            print(f"\n📁 处理子文件夹：{item}")
            process_images(sub_path)

    print("\n✅ 全部操作完成！")

if __name__ == "__main__":
    main()