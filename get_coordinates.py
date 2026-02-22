import cv2
from adb_controller import AdbController

# 初始化 ADB
adb = AdbController()
print("正在截取当前屏幕，请稍候...")

# 获取截图
img = adb.get_screenshot()

if img is None:
    print("❌ 截图失败，请检查模拟器连接")
    exit()

# 保存为文件
filename = "adb_coordinate_map.png"
cv2.imwrite(filename, img)

print(f"✅ 截图成功！已保存为: {filename}")
print("-" * 40)
print("【操作指南】")
print(f"1. 请打开项目文件夹下的 {filename} 图片")
print("2. 推荐使用 Windows 自带的【画图】(Paint) 打开")
print("3. 鼠标移动到你想点击的位置")
print("4. 【画图】软件的左下角会直接显示当前鼠标的坐标 (例如: 1050, 480px)")
print("5. 那个就是你要的 ADB 坐标！")
print("-" * 40)