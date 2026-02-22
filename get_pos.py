# 把这段另存为 get_pos.py 运行，把鼠标放到列表位置就能看到坐标
import pyautogui, time
while True:
    print(pyautogui.position())
    time.sleep(1)