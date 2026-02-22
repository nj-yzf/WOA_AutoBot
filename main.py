import pyautogui
import time
import random

# --- 1. 基础配置与工具函数 ---
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3
ICON_PATH = 'icon/'

# 【新增】全局变量：记录上一次出现地勤人员不足的时间戳
LAST_STAFF_SHORTAGE_TIME = 0


def random_sleep(min_s, max_s):
    time.sleep(random.uniform(min_s, max_s))


def close_window():
    """双击指定坐标关闭弹窗"""
    print(">>> [操作] 双击 (1000, 450) 关闭窗口")
    pyautogui.doubleClick(1000, 450)
    time.sleep(1)


def safe_locate(image_name, confidence=0.8):
    """安全查找函数"""
    try:
        return pyautogui.locateOnScreen(ICON_PATH + image_name, confidence=confidence)
    except pyautogui.ImageNotFoundException:
        return None
    except Exception:
        return None


def find_and_click(image_name, confidence=0.8, wait=0.5):
    """查找并点击图片的通用函数"""
    try:
        loc = pyautogui.locateOnScreen(ICON_PATH + image_name, confidence=confidence)
        if loc:
            pyautogui.click(pyautogui.center(loc))
            time.sleep(wait)
            return True
    except pyautogui.ImageNotFoundException:
        pass
    return False


# --- 2. 四大业务处理模块 ---

# 【类别1】处理进场
def handle_approach_task():
    print(">>> [任务] 处理进场飞机...")
    time.sleep(1.5)

    if find_and_click('landing_permitted.png', wait=1.0):
        print("   -> (初始) 允许降落")
        return True

    vacant = safe_locate('stand_vacant.png')
    if vacant:
        print("   -> 正在分配机位...")
        pyautogui.click(pyautogui.center(vacant))
        time.sleep(0.5)

        if find_and_click('stand_confirm.png', wait=0.5):
            print("   -> 确认分配完成")
            print("   -> 检查后续：是否可降落？")
            time.sleep(1.5)

            if find_and_click('landing_permitted.png', wait=1.0):
                print("   -> (分配后) 顺手点击了允许降落！")
                return True
            else:
                print("   -> 未出现降落按钮，关闭窗口")
                close_window()
                return True
        else:
            print("❌ 找不到确认按钮")
            close_window()
            return False

    if safe_locate('select_stand_unavailable.png'):
        print("⚠️ 暂时没有空闲机位")

    close_window()
    return False


# 【类别2】处理滑行
def handle_taxiing_task():
    print(">>> [任务] 处理滑行飞机...")
    time.sleep(1.5)

    if find_and_click('cross_runway.png', wait=0.2):
        print("   -> 穿越跑道完成")
        return True

    print("⚠️ 未找到 Cross Runway 按钮")
    close_window()
    return False


# 【类别3】处理停机位服务 (核心修改部分)
def handle_stand_task():
    global LAST_STAFF_SHORTAGE_TIME  # 声明使用全局变量
    print(">>> [任务] 处理停机位服务...")
    time.sleep(0.8)

    # 0. 容错：Cross Runway
    if find_and_click('cross_runway.png', wait=0.2):
        print("   -> [容错] 原来是滑行飞机！")
        return True

    while True:
        avail_staff = self.check_global_staff()
        if avail_staff is None: avail_staff = 999
        if avail_staff == 0:
            # ... (人数为0处理逻辑不变) ...
            self.in_staff_shortage_mode = True
            self.last_known_available_staff = 0
            self.close_window()
            return False

        # 带重试的读取
        cost_text = ""
        for i in range(3):
            cost_text = self.ocr.recognize_number(self.REGION_TASK_COST, mode='task')
            # 只要读到了含斜杠的文本，就说明界面加载出来了，不用重试了
            if cost_text and '/' in cost_text:
                break
            time.sleep(0.5)

        # 解析数字
        required_cost = self.ocr.parse_cost(cost_text)

        # === 分支 A: 进错门了 (根本没有斜杠) ===
        # 如果 cost_text 是空的，或者乱码里没有斜杠，说明这压根不是 Stand 界面
        if not cost_text or '/' not in cost_text:
            self.log(f"⚠️ 界面异常 (OCR: '{cost_text}')，可能是误入其他任务")
            self.close_window()
            # 既然进错了，可能是图标识别漂移了，强制休息一下防止死循环
            time.sleep(1.0)
            # 标记跳过，下次试试别的
            self.stand_skip_index += 1
            return False

        # === 分支 B: 进对门了，但数字看不清 (有斜杠，但解析失败) ===
        # 例如识别成 0/0，或者 0/
        if required_cost is None:
            self.log(f"⚠️ 数字识别异常 (OCR: '{cost_text}')，但确认在Stand界面，执行盲做")

            # 盲做前刷新人数
            recheck_avail = self.check_global_staff()
            if recheck_avail is not None: avail_staff = recheck_avail

            if self._perform_stand_action_sequence():
                self.log("   -> 盲做成功")
                self.stand_skip_index = 0
                return True
            else:
                # 盲做失败处理 (同之前)
                if self.in_staff_shortage_mode:
                    if self._try_get_bonus_staff():
                        self.stand_skip_index = 0
                        return True
                self.stand_skip_index += 1
                return False

        # === 分支 C: 识别成功，正常算账 (逻辑不变) ===
        self.log(f"   -> 需求: {required_cost}+1 | 可用: {avail_staff}")

        if avail_staff >= (required_cost + 1):
            if self._perform_stand_action_sequence():
                self.log("   -> 服务启动成功")
                self.stand_skip_index = 0
                return True
            else:
                pass

                # 分支 D: 人力不足
        self.log(f"🛑 人力不足 (缺 {required_cost + 1 - avail_staff})")
        self.close_window()
        if self._try_get_bonus_staff():
            self.stand_skip_index = 0
            return True
        self.stand_skip_index += 1
        return False


# 【类别4】处理起飞
def handle_takeoff_task():
    print(">>> [任务] 处理起飞飞机...")
    time.sleep(1.5)

    if find_and_click('cross_runway.png', wait=1.0):
        print("   -> [容错] 原来是滑行飞机！")
        return True

    if find_and_click('get_award_1.png', wait=1.0):
        print("   -> 领取奖励 1/2")
        if find_and_click('get_award_2.png', wait=1.0):
            print("   -> 领取奖励 2/2")
            if find_and_click('push_back.png', wait=1.5):
                print("   -> 推开 (奖励后)")
                return True
        close_window()
        return False

    action_buttons = ['push_back.png', 'wait.png', 'takeoff_by_gliding.png', 'takeoff.png']
    for btn_name in action_buttons:
        if find_and_click(btn_name, wait=1.5):
            print(f"   -> 执行动作: {btn_name}")
            return True

    print("⚠️ 未找到起飞相关按钮")
    close_window()
    return False


# --- 3. 核心调度逻辑 ---

def scan_and_process():
    global LAST_STAFF_SHORTAGE_TIME  # 声明使用全局变量
    print("-" * 40)

    # 检查是否处于地勤冷却期
    time_since_shortage = time.time() - LAST_STAFF_SHORTAGE_TIME
    is_staff_cooldown = time_since_shortage < 30

    if is_staff_cooldown:
        remaining = 30 - int(time_since_shortage)
        print(f"🛑 地勤人员休息中 (剩余 {remaining} 秒)，跳过停机位任务")
    else:
        print("正在扫描右侧列表...")

    # 任务映射表
    task_map = [
        ('pending_approach.png', handle_approach_task, 0.8),
        ('pending_taxiing.png', handle_taxiing_task, 0.9),
        ('pending_stand.png', handle_stand_task, 0.8),
        ('pending_takeoff.png', handle_takeoff_task, 0.8)
    ]

    all_tasks = []

    for img_name, handler_func, conf_level in task_map:
        # 如果处于冷却期，并且当前扫描的是停机位任务，直接跳过
        if is_staff_cooldown and img_name == 'pending_stand.png':
            continue

        try:
            found = list(pyautogui.locateAllOnScreen(ICON_PATH + img_name, confidence=conf_level, grayscale=True))
            for box in found:
                all_tasks.append({
                    'y': box.top,
                    'box': box,
                    'handler': handler_func,
                    'name': img_name
                })
        except pyautogui.ImageNotFoundException:
            pass
        except Exception:
            pass

    if not all_tasks:
        return False

    all_tasks.sort(key=lambda t: t['y'])
    top_task = all_tasks[0]

    print(f"发现 {len(all_tasks)} 个任务，优先处理: {top_task['name']}")
    pyautogui.click(pyautogui.center(top_task['box']))

    success = top_task['handler']()
    return success


# --- 4. 主循环 ---

print(">>> WAO 全自动机场控制台已启动")
time.sleep(2)

try:
    while True:
        did_work = scan_and_process()
        if did_work:
            time.sleep(1.0)
        else:
            print(".", end="", flush=True)
            time.sleep(3.0)

except KeyboardInterrupt:
    print("\n>>> 脚本已停止")