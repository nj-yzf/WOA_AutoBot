from adb_controller import AdbController
import cv2

adb = AdbController()
print("正在获取 ADB 截图...")
img = adb.get_screenshot()

if img is not None:
    cv2.imwrite("adb_screen_full.png", img)
    print("✅ 截图成功！已保存为 adb_screen_full.png")
    print("请打开这张图片，从中截取你的小图标，并覆盖到 icon 文件夹中。")
else:
    print("❌ 截图失败")