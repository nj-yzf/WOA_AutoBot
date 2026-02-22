import cv2
import numpy as np
import os
from adb_controller import AdbController

# === 配置区 ===
# 请确保这里和 main_adb.py 里的 ROI 定义完全一致
ROI_GLOBAL = (566, 92, 82, 22)  # 右上角 (x, y, w, h)
ROI_TASK = (262, 670, 80, 25)  # 任务详情页 (x, y, w, h)


def process_and_save(img, name):
    if img is None: return

    # 1. 放大 4 倍 (与脚本逻辑一致)
    h, w = img.shape[:2]
    scaled = cv2.resize(img, (w * 4, h * 4), interpolation=cv2.INTER_CUBIC)

    # 2. 转灰度
    gray = cv2.cvtColor(scaled, cv2.COLOR_BGR2GRAY)

    # 3. 二值化 (阈值 170)
    _, binary = cv2.threshold(gray, 170, 255, cv2.THRESH_BINARY)

    # 4. 保存
    filename = f"template_source_{name}.png"
    cv2.imwrite(filename, binary)
    print(f"✅ 已生成素材图: {filename}")
    print(f"   -> 请打开它，裁剪出 0-9 和 /，替换 icon/digits/{name}/ 下的文件")


if __name__ == "__main__":
    print(">>> 正在初始化 ADB...")

    # 1. 扫描设备
    devices = AdbController.scan_devices()

    if not devices:
        print("❌ 未发现任何设备，请检查模拟器是否启动！")
        exit()

    print(f"✅ 发现设备: {devices[0]}")

    # 2. 连接第一个设备
    adb = AdbController(target_device=devices[0])

    print(">>> 正在获取屏幕...")
    screen = adb.get_screenshot()
    # ...

    if screen is None:
        print("❌ 截图失败")
        exit()

    # 生成右上角的素材
    print(">>> 处理右上角区域...")
    x, y, w, h = ROI_GLOBAL
    crop_global = screen[y:y + h, x:x + w]
    process_and_save(crop_global, "global")

    # 生成任务详情页的素材
    print(">>> 处理任务需求区域...")
    x, y, w, h = ROI_TASK
    crop_task = screen[y:y + h, x:x + w]
    process_and_save(crop_task, "task")

    print("\n完成！现在你的根目录下有两个巨大的黑白图片。")