import cv2
import numpy as np
import time
import random
import threading
import sys
import os
import gc
import traceback
from adb_controller import AdbController


def get_resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        if hasattr(sys, '_MEIPASS'):
            base = sys._MEIPASS
        else:
            base = os.path.dirname(sys.executable)
        return os.path.join(base, relative_path)
    base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


class StopSignal(Exception):
    pass


class SimpleOCR:
    def __init__(self, adb_controller, icon_path):
        self.adb = adb_controller
        self.root_path = os.path.join(icon_path, "digits")
        self.SCALE_FACTOR = 4
        self.templates_global = {}
        self.templates_task = {}
        self._load_templates("global", self.templates_global)
        self._load_templates("task", self.templates_task)

    def _process_image(self, img):
        if img is None or img.size == 0: return None
        h, w = img.shape[:2]
        try:
            scaled_img = cv2.resize(img, (w * self.SCALE_FACTOR, h * self.SCALE_FACTOR), interpolation=cv2.INTER_LINEAR)
            if len(scaled_img.shape) == 3:
                gray = cv2.cvtColor(scaled_img, cv2.COLOR_BGR2GRAY)
            else:
                gray = scaled_img
            _, binary = cv2.threshold(gray, 170, 255, cv2.THRESH_BINARY)
            return binary
        except Exception:
            return None

    def _load_templates(self, sub_folder, target_dict):
        folder_path = os.path.join(self.root_path, sub_folder) + os.sep
        if not os.path.exists(folder_path): return
        chars = [str(i) for i in range(10)] + ['slash']
        for char in chars:
            path = folder_path + char + ".png"
            if os.path.exists(path):
                img = cv2.imdecode(np.fromfile(path, dtype=np.uint8), cv2.IMREAD_GRAYSCALE)
                target_dict[char] = img

    def recognize_number(self, region, mode='global', screen_image=None):
        x, y, w, h = region
        if screen_image is not None:
            full_screen = screen_image
        else:
            full_screen = self.adb.get_screenshot()
        if full_screen is None: return None

        if y < 0 or x < 0 or y + h > full_screen.shape[0] or x + w > full_screen.shape[1]:
            return None

        crop_img = full_screen[y:y + h, x:x + w]
        processed_crop = self._process_image(crop_img)
        if processed_crop is None: return None

        templates = self.templates_global if mode == 'global' else self.templates_task
        if not templates: return None
        matches = []
        threshold = 0.7
        for char, template in templates.items():
            if template.shape[0] > processed_crop.shape[0] or template.shape[1] > processed_crop.shape[1]:
                continue
            res = cv2.matchTemplate(processed_crop, template, cv2.TM_CCOEFF_NORMED)
            loc = np.where(res >= threshold)
            t_w = template.shape[1]
            for pt in zip(*loc[::-1]):
                score = res[pt[1], pt[0]]
                matches.append({'x': pt[0], 'char': '/' if char == 'slash' else char, 'score': score, 'width': t_w})
        if not matches: return None
        matches.sort(key=lambda k: k['x'])
        final_results = []
        while len(matches) > 0:
            curr = matches.pop(0)
            keep_curr = True
            if final_results:
                last = final_results[-1]
                is_one_slash = (curr['char'] == '1' and last['char'] == '/') or \
                               (curr['char'] == '/' and last['char'] == '1')
                if is_one_slash:
                    final_results.append(curr)
                    continue
                start = max(last['x'], curr['x'])
                end = min(last['x'] + last['width'], curr['x'] + curr['width'])
                overlap = max(0, end - start)
                min_width = min(last['width'], curr['width'])
                if overlap > min_width * 0.4:
                    if curr['score'] > last['score']:
                        final_results.pop()
                        final_results.append(curr)
                    keep_curr = False
            if keep_curr:
                final_results.append(curr)
        result_str = "".join([m['char'] for m in final_results])
        return result_str

    def parse_staff_count(self, text):
        try:
            if not text or '/' not in text: return None
            clean = "".join([c for c in text if c.isdigit() or c == '/'])
            parts = clean.split('/')
            if len(parts) < 2: return None
            used = int(parts[0])
            total = int(parts[1])
            avail = total - used
            if avail < 0: return None
            return used, total, avail
        except:
            return None

    def parse_cost(self, text):
        try:
            if not text or '/' not in text: return None
            clean = "".join([c for c in text if c.isdigit() or c == '/'])
            parts = clean.split('/')
            if len(parts) < 2: return None
            cost_str = parts[1]
            if not cost_str: return None
            cost = int(cost_str)
            if cost == 0: return 10
            if cost < 0 or cost > 25: return None
            return cost
        except:
            return None


class WoaBot:
    def _check_running(self):
        if not self.running:
            raise StopSignal()

    def __init__(self, log_callback=None, config_callback=None):
        self.last_staff_log_time = 0
        self.config_callback = config_callback
        self.adb = None
        self.target_device = None
        self.running = False
        self.log_callback = log_callback
        self.icon_path = get_resource_path('icon') + os.sep
        self.last_staff_shortage_time = 0
        self.CLOSE_X = 1153
        self.CLOSE_Y = 181
        self.ocr = None
        self.stand_skip_index = 0
        self.in_staff_shortage_mode = False
        self.enable_bonus_staff = False
        self.last_bonus_staff_time = 0
        self.BONUS_COOLDOWN = 2 * 60 * 60
        self.REGION_GLOBAL_STAFF = (573, 92, 640 - 573, 112 - 92)
        self.REGION_TASK_COST = (270, 670, 330 - 270, 695 - 670)
        self.REGION_GREEN_DOT = (405, 517, 201, 109)
        self.next_bonus_retry_time = 0
        self.doing_task_forbidden_until = 0
        self.next_list_refresh_time = 0
        self.enable_vehicle_buy = False
        self.enable_speed_mode = False
        self.enable_skip_staff = False
        self.enable_delay_bribe = False
        self.enable_random_task = False
        self.control_method = "minitouch"
        self.screenshot_method = "nemu_ipc"

        self.slide_min_duration = 250
        self.slide_max_duration = 500
        self.REGION_MAIN_ANCHOR = (30, 30, 55, 45)
        self.REGION_REWARD_RECOVERY = (308, 428, 1007, 311)
        self.REWARD_FLOW_BUTTONS = ['get_award_1.png', 'get_award_2.png', 'get_award_3.png', 'get_award_4.png',
                                    'push_back.png', 'taxi_to_runway.png', 'start_general.png']
        self.last_seen_main_interface_time = time.time()
        self.STUCK_TIMEOUT = 15.0
        self.auto_delay_count = 0
        self.TOWER_CHECK_POINTS = [(656, 809), (634, 831), (634, 809), (656, 830)]
        # BGR: 红(需延时) 绿(无需延时) 灰(塔台关闭)
        self.TOWER_RED_BGR = (110, 112, 251)
        self.TOWER_GREEN_BGR = (153, 219, 94)
        self.TOWER_OFF_COLOR = (128, 111, 94)
        self.COLOR_LIGHT = (203, 191, 179)
        self.COLOR_DARK = (101, 85, 70)
        self.FILTER_MENU_BTN = (1537, 37)
        self.FILTER_CHECK_POINTS_MODE1 = [
            ((1542, 190), True), ((1535, 118), True), ((1540, 261), True),
            ((1533, 331), True), ((1537, 403), True), ((1542, 474), False)
        ]
        self.FILTER_CHECK_POINTS_MODE2 = [
            ((1542, 190), False), ((1535, 118), True), ((1540, 261), True),
            ((1533, 331), True), ((1537, 403), True), ((1542, 474), False)
        ]
        self.enable_cancel_stand_filter = False
        self.REGION_STATUS_TITLE = (20, 320, 190, 250)
        self.LIST_ROI_X = 1312
        self.LIST_ROI_W = 60
        self.LIST_ROI_H = 900
        self.REGION_BOTTOM_ROI = (20, 750, 340, 130)
        self.REGION_VACANT_ROI = (390, 690, 730, 150)

        # 防卡死相关
        self.consecutive_timeout_count = 0
        self.last_recovery_time = 0  # 冷却时间
        self.last_window_close_time = time.time()

        self.last_checked_avail_staff = -1
        self.last_read_success = False
        self.thinking_mode = 0
        self.thinking_range = (0, 0)

        self.ICON_ROIS = {
            'cross_runway.png': self.REGION_BOTTOM_ROI,
            'get_award_1.png': self.REGION_BOTTOM_ROI,
            'landing_permitted.png': self.REGION_BOTTOM_ROI,
            'landing_prohibited.png': self.REGION_BOTTOM_ROI,
            'push_back.png': self.REGION_BOTTOM_ROI,
            'stand_confirm.png': self.REGION_BOTTOM_ROI,
            'start_ground_support.png': self.REGION_BOTTOM_ROI,
            'start_ice.png': self.REGION_BOTTOM_ROI,
            'takeoff.png': self.REGION_BOTTOM_ROI,
            'takeoff_by_gliding.png': self.REGION_BOTTOM_ROI,
            'taxi_to_runway.png': self.REGION_BOTTOM_ROI,
            'start_general.png': self.REGION_BOTTOM_ROI,
            'wait.png': self.REGION_BOTTOM_ROI,
            'go_repair.png': self.REGION_BOTTOM_ROI,
            'start_repair.png': self.REGION_BOTTOM_ROI,
            'ground_support_done.png': self.REGION_BOTTOM_ROI,
            'stand_vacant.png': self.REGION_VACANT_ROI,
            'green_dot.png': self.REGION_GREEN_DOT
        }

        self.task_templates = {}
        task_files = [
            'pending_ice.png', 'pending_repair.png', 'pending_doing.png',
            'pending_approach.png', 'pending_taxiing.png', 'pending_takeoff.png',
            'pending_stand.png'
        ]
        for tf in task_files:
            p = self.icon_path + tf
            if os.path.exists(p):
                self.task_templates[tf] = cv2.imread(p)

    def set_random_task_mode(self, enabled, log_change=True):
        if self.enable_random_task == enabled:
            return
        self.enable_random_task = enabled
        if log_change:
            self.log(f">>> [配置] 随机任务选择: {'已开启' if enabled else '已关闭'}")

    def set_cancel_stand_filter_when_tower_off(self, enabled):
        if self.enable_cancel_stand_filter == enabled:
            return
        self.enable_cancel_stand_filter = enabled
        self.log(f">>> [配置] 塔台关闭时取消停机位筛选: {'已开启' if enabled else '已关闭'}")


    def _color_diff(self, a, b):
        return sum(abs(int(a[i]) - int(b[i])) for i in range(3))

    def _is_pixel_light(self, screen, x, y):
        try:
            b, g, r = screen[y, x]
            return self._color_diff((b, g, r), self.COLOR_LIGHT) < 80
        except:
            return False

    def _is_pixel_dark(self, screen, x, y):
        try:
            b, g, r = screen[y, x]
            return self._color_diff((b, g, r), self.COLOR_DARK) < 80
        except:
            return False

    def _is_tower_off(self, screen):
        """四个检测点全部为灰色才表示塔台关闭"""
        tb, tg, tr = self.TOWER_OFF_COLOR
        for (x, y) in self.TOWER_CHECK_POINTS:
            try:
                b, g, r = screen[y, x]
                if self._color_diff((b, g, r), (tb, tg, tr)) > 70:
                    return False
            except:
                return False
        return True

    def _is_point_red(self, b, g, r):
        """检测点是否为红色（需延时）：R 主导，与绿/灰区分"""
        rb, rg, rr = self.TOWER_RED_BGR
        diff_red = self._color_diff((b, g, r), (rb, rg, rr))
        diff_green = self._color_diff((b, g, r), self.TOWER_GREEN_BGR)
        diff_gray = self._color_diff((b, g, r), self.TOWER_OFF_COLOR)
        return diff_red < 90 and diff_red <= diff_green and diff_red <= diff_gray

    def _matches_filter_mode(self, screen, points_config):
        for (x, y), want_light in points_config:
            if want_light:
                if not self._is_pixel_light(screen, x, y):
                    return False
            else:
                if not self._is_pixel_dark(screen, x, y):
                    return False
        return True

    def _click_filter_point(self, x, y):
        self.adb.click(x, y, random_offset=5)

    def _periodic_15s_check(self, force_initial_filter_check=False):
        if not hasattr(self, 'last_periodic_check_time'):
            self.last_periodic_check_time = 0
        now = time.time()
        if not force_initial_filter_check and now - self.last_periodic_check_time < 15.0:
            return
        self.last_periodic_check_time = now

        # 1. 检测主界面
        if self.safe_locate('main_interface.png', region=self.REGION_MAIN_ANCHOR, confidence=0.8):
            self.last_seen_main_interface_time = time.time()

        # 2. 中间领奖区防卡死：不论是否在主界面，均遍历寻找领奖按钮并点击直至恢复正常
        rx, ry, rw, rh = self.REGION_REWARD_RECOVERY
        for _ in range(10):
            screen = self.adb.get_screenshot()
            if screen is None:
                break
            roi = screen[ry:ry + rh, rx:rx + rw]
            clicked = False
            for btn in self.REWARD_FLOW_BUTTONS:
                res = self.adb.locate_image(self.icon_path + btn, confidence=0.65, screen_image=roi)
                if res:
                    self.log(f"🚨 [15s周期检测] 在领奖区域内发现 {btn}，点击恢复...")
                    self.adb.click(res[0] + rx, res[1] + ry, random_offset=3)
                    self.sleep(1.0)
                    clicked = True
                    break
            if not clicked:
                break

        # 3. 筛选状态检查 (仅在确认在主界面时执行)
        if not self.safe_locate('main_interface.png', region=self.REGION_MAIN_ANCHOR, confidence=0.8):
            return
        screen = self.adb.get_screenshot()
        if screen is None:
            return

        mx, my = self.FILTER_MENU_BTN
        if self._is_pixel_light(screen, mx, my):
            self.log("📋 [筛选] 展开筛选菜单...")
            self._click_filter_point(mx, my)
            self.sleep(0.5)
            for _ in range(5):
                screen = self.adb.get_screenshot()
                if screen is None:
                    break
                if self._is_pixel_dark(screen, mx, my):
                    break
                self._click_filter_point(mx, my)
                self.sleep(0.3)
            screen = self.adb.get_screenshot()
            if screen is None:
                return

        is_mode1 = self._matches_filter_mode(screen, self.FILTER_CHECK_POINTS_MODE1)
        is_mode2 = self._matches_filter_mode(screen, self.FILTER_CHECK_POINTS_MODE2)

        def apply_mode(points_config):
            for (x, y), want_light in points_config:
                screen = self.adb.get_screenshot()
                if screen is None:
                    break
                is_light = self._is_pixel_light(screen, x, y)
                if (want_light and not is_light) or (not want_light and is_light):
                    self._click_filter_point(x, y)
                    self.sleep(0.2)

        need_mode1_only = self.enable_cancel_stand_filter and self._is_tower_off(screen)
        if need_mode1_only and not is_mode1:
            self.log("📋 [筛选] 切换至仅待处理... (塔台已关闭)")
            apply_mode(self.FILTER_CHECK_POINTS_MODE1)
            return

        # 模式1或模式2均可接受；启动时若已是模式2也无需切换
        if is_mode1 or is_mode2:
            return
        self.log("📋 [筛选] 状态异常，默认切换至仅待处理...")
        apply_mode(self.FILTER_CHECK_POINTS_MODE1)

    def _nemu_ipc_debug_save_mismatch(self, nemu_img, adb_img):
        """nemu_ipc 与 ADB 截图不匹配时保存对比图，便于排查"""
        try:
            import datetime
            debug_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "nemu_ipc_debug")
            os.makedirs(debug_dir, exist_ok=True)
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            p_n = os.path.join(debug_dir, f"mismatch_nemu_{ts}.png")
            p_a = os.path.join(debug_dir, f"mismatch_adb_{ts}.png")
            cv2.imwrite(p_n, nemu_img)
            cv2.imwrite(p_a, adb_img)
            self.log(f"📋 [调试] 已保存对比图: {p_n} / {p_a}")
        except Exception as e:
            self.log(f"📋 [调试] 保存对比图失败: {e}")

    def _droidcast_raw_debug_save_mismatch(self, droidcast_img, adb_img):
        """DroidCast_raw 与 ADB 截图不匹配时保存对比图，便于排查"""
        try:
            import datetime
            debug_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "droidcast_raw_debug")
            os.makedirs(debug_dir, exist_ok=True)
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            p_d = os.path.join(debug_dir, f"mismatch_droidcast_{ts}.png")
            p_a = os.path.join(debug_dir, f"mismatch_adb_{ts}.png")
            cv2.imwrite(p_d, droidcast_img)
            cv2.imwrite(p_a, adb_img)
            self.log(f"📋 [调试] 已保存对比图: {p_d} / {p_a}")
        except Exception as e:
            self.log(f"📋 [调试] 保存对比图失败: {e}")

    def _run_pending_detection(self, list_roi_img):
        """按行识别：每行只保留该行内置信度最高的类型，避免跨行竞争导致相似图标误判。"""
        task_defs = [
            ('pending_ice.png', self.handle_ice_task, 0.8, 'ice'),
            ('pending_repair.png', self.handle_repair_task, 0.8, 'repair'),
            ('pending_doing.png', self.handle_vehicle_check_task, 0.85, 'doing'),
            ('pending_approach.png', self.handle_approach_task, 0.8, 'approach'),
            ('pending_taxiing.png', self.handle_taxiing_task, 0.8, 'taxiing'),
            ('pending_takeoff.png', self.handle_takeoff_task, 0.8, 'takeoff'),
            ('pending_stand.png', self.handle_stand_task, 0.8, 'stand')
        ]
        ROW_HEIGHT = 24

        all_matches = []
        for img_name, handler, conf, t_type in task_defs:
            found = self._fast_locate_all(list_roi_img, img_name, confidence=conf)
            for item in found:
                rel_cx, rel_cy = item['center']
                abs_cx = rel_cx + self.LIST_ROI_X
                abs_cy = rel_cy
                type_for_logic = 'stand' if t_type == 'stand' else ('doing' if t_type == 'doing' else 'other')
                all_matches.append({
                    'y': rel_cy, 'center': (abs_cx, abs_cy), 'handler': handler, 'name': img_name,
                    'score': item['score'], 'type': type_for_logic, 'raw_type': t_type
                })

        def _same_row(y1, y2):
            return abs(y1 - y2) <= ROW_HEIGHT

        final_tasks = []
        used = [False] * len(all_matches)
        for i, det in enumerate(all_matches):
            if used[i]:
                continue
            cy = det['y']
            same_row = [all_matches[j] for j in range(len(all_matches)) if not used[j] and _same_row(all_matches[j]['y'], cy)]
            if len(same_row) <= 1:
                final_tasks.append(det)
                used[i] = True
                for j, d in enumerate(all_matches):
                    if d is det or (not used[j] and _same_row(d['y'], cy)):
                        used[j] = True
                continue
            same_row.sort(key=lambda x: x['score'], reverse=True)
            best = same_row[0]
            final_tasks.append(best)
            for j, d in enumerate(all_matches):
                if not used[j] and _same_row(d['y'], cy):
                    used[j] = True

        final_tasks.sort(key=lambda t: t['y'])
        raw_detections = all_matches
        return raw_detections, final_tasks

    def _fast_locate_all(self, screen_roi, template_name, confidence=0.8):
        if template_name not in self.task_templates:
            return []

        template = self.task_templates[template_name]
        if template is None: return []

        try:
            res = cv2.matchTemplate(screen_roi, template, cv2.TM_CCOEFF_NORMED)
        except:
            return []

        h, w = template.shape[:2]
        loc = np.where(res >= confidence)
        found_items = []
        for pt in zip(*loc[::-1]):
            score = res[pt[1], pt[0]]
            is_duplicate = False
            for item in found_items:
                if abs(pt[0] - item['box'][0]) < 10 and abs(pt[1] - item['box'][1]) < 10:
                    if score > item['score']:
                        item['score'] = score
                        item['box'] = (pt[0], pt[1], w, h)
                        item['center'] = (pt[0] + w // 2, pt[1] + h // 2)
                    is_duplicate = True
                    break
            if not is_duplicate:
                found_items.append({
                    'box': (pt[0], pt[1], w, h),
                    'center': (pt[0] + w // 2, pt[1] + h // 2),
                    'score': score
                })
        return found_items

    def set_thinking_time_mode(self, mode_index, log_change=True):
        mode_index = int(mode_index)
        if hasattr(self, 'thinking_mode') and self.thinking_mode == mode_index:
            return
        self.thinking_mode = mode_index
        if self.thinking_mode == 1:
            self.thinking_range = (0.1, 0.4)
            desc = "短 (0.1s-0.4s)"
        elif self.thinking_mode == 2:
            self.thinking_range = (0.3, 1.0)
            desc = "中 (0.3s-1.0s)"
        elif self.thinking_mode == 3:
            self.thinking_range = (0.8, 2.0)
            desc = "长 (0.8s-2.0s)"
        else:
            self.thinking_range = (0, 0)
            desc = "关闭"
        prev = getattr(self, '_last_thinking_desc', None)
        if prev == desc:
            if self.adb:
                self.adb.set_thinking_strategy(*self.thinking_range)
            return
        self._last_thinking_desc = desc
        if self.adb:
            self.adb.set_thinking_strategy(*self.thinking_range)
        if log_change:
            self.log(f">>> [配置] 思考时间: {desc}")

    def set_bonus_staff_feature(self, enabled):
        if self.enable_bonus_staff == enabled: return
        self.enable_bonus_staff = enabled
        self.log(f">>> [配置] 自动领取地勤: {'已开启' if enabled else '已关闭'}")

    def set_vehicle_buy_feature(self, enabled):
        if self.enable_vehicle_buy == enabled: return
        self.enable_vehicle_buy = enabled
        self.log(f">>> [配置] 自动购买车辆: {'已开启' if enabled else '已关闭'}")

    def set_speed_mode(self, enabled):
        if self.enable_speed_mode == enabled: return
        self.enable_speed_mode = enabled
        self.log(f">>> [配置] 跳过二次校验: {'已开启' if enabled else '已关闭'}")

    def set_skip_staff_verify(self, enabled):
        if self.enable_skip_staff == enabled: return
        self.enable_skip_staff = enabled
        self.log(f">>> [配置] 跳过地勤验证: {'已开启' if enabled else '已关闭'}")

    def set_auto_delay(self, count):
        count = int(count)
        self.auto_delay_count = count

    def set_delay_bribe(self, enabled):
        if self.enable_delay_bribe == enabled: return
        self.enable_delay_bribe = enabled
        self.log(f">>> [配置] 延误飞机贿赂: {'已开启' if enabled else '已关闭'}")

    def set_slide_duration_range(self, min_d, max_d, log_change=True):
        min_d = int(min_d)
        max_d = int(max_d)
        if hasattr(self, 'slide_min_duration') and hasattr(self, 'slide_max_duration'):
            if self.slide_min_duration == min_d and self.slide_max_duration == max_d:
                return
        self.slide_min_duration = min_d
        self.slide_max_duration = max_d
        if log_change:
            self.log(f">>> [配置] 滑块随机耗时: {self.slide_min_duration}ms - {self.slide_max_duration}ms")

    def set_device(self, device_serial):
        self.target_device = device_serial

    def set_control_method(self, method):
        m = (method or "adb").lower()
        valid = ("adb", "minitouch", "uiautomator2")
        if m not in valid:
            m = "adb"
        if self.control_method != m:
            self.control_method = m

    def set_screenshot_method(self, method):
        m = (method or "adb").lower()
        if m not in ("adb", "nemu_ipc", "uiautomator2", "droidcast_raw"):
            m = "adb"
        if self.screenshot_method != m:
            self.screenshot_method = m

    def log(self, message):
        if not message or not str(message).strip():
            return
        print(message)
        if self.log_callback:
            self.log_callback(message)

        # 【核心修正】智能防卡死逻辑
        # 1. 过滤掉恢复日志本身，防止递归触发
        if "防卡死" in message: return

        # 2. 统计警告次数
        if "⚠️" in message or "超时" in message:
            self.consecutive_timeout_count += 1
        elif "✅" in message or "成功" in message:
            self.consecutive_timeout_count = 0

        # 3. 触发条件：连续3次警告 + 冷却时间已过
        if self.consecutive_timeout_count > 3:
            if time.time() - self.last_recovery_time > 10:
                self.log("🚨 [防卡死] 检测到连续多次卡顿，尝试紧急寻找领奖图标...")
                self.consecutive_timeout_count = 0
                self.last_recovery_time = time.time()
                self._attempt_emergency_reward_recovery()
            else:
                self.consecutive_timeout_count = 0  # 冷却中，暂时重置

    def _attempt_emergency_reward_recovery(self):
        # 全屏搜索领奖图标
        targets = ['get_award_1.png', 'get_award_2.png', 'get_award_3.png', 'get_award_4.png']
        # 尝试循环检测3次，确保如果点到第1步能接着点第2步
        for _ in range(3):
            clicked = False
            for t in targets:
                # 使用 region=None 进行全屏搜索，降低一点阈值以防图标变灰或变暗
                res = self.safe_locate(t, confidence=0.65, region=None)
                if res:
                    self.log(f"   -> 🚨 紧急恢复：点击 {t}")
                    self.adb.click(res[0], res[1])
                    self.sleep(1.5)  # 点击后多等一会儿
                    clicked = True
                    break
            if not clicked:
                break

    def wait_and_click(self, image_name, timeout=3.0, click_wait=0.2, confidence=0.8, random_offset=5):
        self._check_running()
        start_time = time.time()
        use_roi = False
        roi_x, roi_y, roi_w, roi_h = 0, 0, 0, 0
        if image_name in self.ICON_ROIS:
            use_roi = True
            roi_x, roi_y, roi_w, roi_h = self.ICON_ROIS[image_name]

        while time.time() - start_time < timeout:
            self._check_running()
            screen = self.adb.get_screenshot()
            if screen is None:
                time.sleep(0.1)
                continue
            if use_roi:
                search_img = screen[roi_y:roi_y + roi_h, roi_x:roi_x + roi_w]
                offset_x, offset_y = roi_x, roi_y
            else:
                search_img = screen
                offset_x, offset_y = 0, 0

            result = self.adb.locate_image(self.icon_path + image_name, confidence=confidence, screen_image=search_img)
            if result:
                self._check_running()
                real_x = result[0] + offset_x
                real_y = result[1] + offset_y
                self.adb.click(real_x, real_y, random_offset=random_offset)
                if click_wait > 0: self.sleep(click_wait)
                return True
            time.sleep(0.1)
        return False

    def start(self):
        if self.running: return
        self.stand_skip_index = 0
        self.in_staff_shortage_mode = False
        self.last_checked_avail_staff = -1
        self.last_window_close_time = time.time()
        self.auto_delay_count = 0
        if self.config_callback:
            self.config_callback("auto_delay_count", 0)
        # 初始化计数器
        self.consecutive_timeout_count = 0
        self.last_recovery_time = 0

        if not self.target_device:
            self.log("❌ 未选择设备！")
            return
        self.running = True
        self.log(f">>> 连接设备: {self.target_device} ...")
        try:
            self.adb = AdbController(
                target_device=self.target_device,
                control_method=self.control_method,
                screenshot_method=self.screenshot_method,
            )
            self.adb.set_thinking_strategy(*self.thinking_range)
            self.ocr = SimpleOCR(self.adb, self.icon_path)
            self.log("✅ OCR 模块已加载")
            test_img = self.adb.get_screenshot()
            if test_img is None:
                self.log("❌ 连接成功但无法获取画面！")
                self.running = False
                try:
                    if self.adb:
                        self.adb.close()
                except Exception:
                    pass
                return
            h, w = test_img.shape[:2]
            if w != 1600 or h != 900:
                self.log(f"🛑 分辨率错误：{w}x{h} (必须 1600x900)")
                self.running = False
                try:
                    if self.adb:
                        self.adb.close()
                except Exception:
                    pass
                return
            self.log(f"✅ 画面正常，脚本启动")
            ctrl_map = {"adb": "ADB", "minitouch": "minitouch", "uiautomator2": "uiautomator2"}
            ctrl = ctrl_map.get(self.adb.control_method, "ADB")
            shot = self.adb.screenshot_method if self.adb.screenshot_method != "adb" else "ADB"
            self.log(f">>> [模式] 触控: {ctrl}, 截图: {shot}")
            thread = threading.Thread(target=self._main_loop)
            thread.daemon = True
            thread.start()
        except Exception as e:
            self.log(f"❌ 启动失败: {e}")
            self.running = False
            try:
                if hasattr(self, 'adb') and self.adb:
                    self.adb.close()
            except Exception:
                pass

    def stop(self):
        self.running = False
        self.log(">>> 正在停止脚本...")
        self.next_bonus_retry_time = 0
        try:
            if hasattr(self, 'adb') and self.adb:
                self.adb.close()
        except Exception:
            pass

    def _main_loop(self):
        self.log("[DEBUG] 主循环线程已启动")
        self.sleep(1.0)
        self.last_periodic_check_time = 0
        self._periodic_15s_check(force_initial_filter_check=True)
        idle_count = 0
        gc_counter = 0
        while self.running:
            try:
                did_work = self.scan_and_process()
                if did_work:
                    self.sleep(0.05)
                    idle_count = 0
                else:
                    self.sleep(0.5)
                    idle_count += 1
                    if idle_count >= 3:
                        self.close_window()
                        idle_count = 0
                gc_counter += 1
                if gc_counter > 200:
                    gc.collect()
                    gc_counter = 0
            except StopSignal:
                self.log(">>> [系统] 停止指令，终止...")
                break
            except Exception as e:
                traceback.print_exc()
                self.log(f"❌ 运行出错: {e}")
                try:
                    self.sleep(3.0)
                except:
                    break
        self.log(">>> 脚本已完全停止")
        try:
            if hasattr(self, 'adb') and self.adb:
                self.adb.close()
        except Exception:
            pass

    def random_sleep(self, min_s, max_s):
        self._check_running()
        self.sleep(random.uniform(min_s, max_s))
        self._check_running()

    def sleep(self, seconds):
        end_time = time.time() + seconds
        while time.time() < end_time:
            self._check_running()
            remaining = end_time - time.time()
            sleep_time = min(0.1, remaining)
            if sleep_time > 0: time.sleep(sleep_time)

    def close_window(self):
        self.adb.double_click(self.CLOSE_X, self.CLOSE_Y, random_offset=30)
        self.last_window_close_time = time.time()
        self.sleep(0.1)

    def find_and_click(self, image_name, confidence=0.8, wait=0.5, random_offset=5):
        self._check_running()
        screen = self.adb.get_screenshot()
        if screen is None: return False
        search_img = screen
        offset_x, offset_y = 0, 0
        if image_name in self.ICON_ROIS:
            roi = self.ICON_ROIS[image_name]
            x, y, w, h = roi
            search_img = screen[y:y + h, x:x + w]
            offset_x, offset_y = x, y
        result = self.adb.locate_image(self.icon_path + image_name, confidence=confidence, screen_image=search_img)
        if result:
            self._check_running()
            real_x = result[0] + offset_x
            real_y = result[1] + offset_y
            self.adb.click(real_x, real_y, random_offset=random_offset)
            self.sleep(wait)
            return True
        return False

    def safe_locate(self, image_name, confidence=0.8, region=None):
        self._check_running()
        if region is None:
            return self.adb.locate_image(self.icon_path + image_name, confidence=confidence)
        screen = self.adb.get_screenshot()
        if screen is None: return None
        x, y, w, h = region
        x = max(0, int(x));
        y = max(0, int(y))
        search_img = screen[y:y + h, x:x + w]
        result = self.adb.locate_image(self.icon_path + image_name, confidence=confidence, screen_image=search_img)
        if result:
            return (result[0] + x, result[1] + y)
        return None

    def check_global_staff(self, screen_image=None):
        text = self.ocr.recognize_number(self.REGION_GLOBAL_STAFF, mode='global', screen_image=screen_image)
        if not text: return None
        result = self.ocr.parse_staff_count(text)
        if result: return result[2]
        return None

    def _verify_and_redirect(self, expected_status_img):
        if self.enable_speed_mode: return True
        status_map = [
            ('status_stand.png', self.handle_stand_task),
            ('status_takeoff.png', self.handle_takeoff_task),
            ('status_taxiing.png', self.handle_taxiing_task),
            ('status_approach.png', self.handle_approach_task),
            ('status_ice.png', self.handle_ice_task),
            ('status_doing.png', self.handle_vehicle_check_task)
        ]

        def _check_expected(roi_img):
            return self.adb.locate_image(self.icon_path + expected_status_img, confidence=0.7, screen_image=roi_img)

        def _find_any_status(roi_img):
            for img, _ in status_map:
                if self.adb.locate_image(self.icon_path + img, confidence=0.7, screen_image=roi_img):
                    return img
            return None

        x, y, w, h = self.REGION_STATUS_TITLE
        for attempt in range(2):
            full_screen = self.adb.get_screenshot()
            if full_screen is None: continue
            roi_image = full_screen[y:y + h, x:x + w]
            if _check_expected(roi_image):
                return True
            if attempt == 0:
                self.sleep(0.15)

        self.log(f"   -> 状态校验不匹配 ({expected_status_img})，尝试纠错...")
        full_screen = self.adb.get_screenshot()
        if full_screen is None: return False
        roi_image = full_screen[y:y + h, x:x + w]
        found = _find_any_status(roi_image)
        if found:
            for img, handler in status_map:
                if img == found:
                    if img == 'status_doing.png' and time.time() <= self.doing_task_forbidden_until:
                        self.log("   -> ⏳ Doing 任务尚在6秒冷却中，关闭窗口")
                        self.close_window()
                        return False
                    self.log(f"   -> ↪️ 自动跳转至: {img}")
                    try:
                        handler()
                    except TypeError:
                        handler(None)
                    return False
        self.sleep(0.2)
        full_screen = self.adb.get_screenshot()
        if full_screen is not None:
            roi_image = full_screen[y:y + h, x:x + w]
            if _check_expected(roi_image):
                return True
            found = _find_any_status(roi_image)
            if found:
                for img, handler in status_map:
                    if img == found:
                        if img == 'status_doing.png' and time.time() <= self.doing_task_forbidden_until:
                            self.log("   -> ⏳ Doing 任务尚在6秒冷却中，关闭窗口")
                            self.close_window()
                            return False
                        self.log(f"   -> ↪️ 自动跳转至: {img}")
                        try:
                            handler()
                        except TypeError:
                            handler(None)
                        return False
        self.log("   -> 未知状态，退出")
        self.close_window()
        return False

    def _update_staff_tracker(self, val):
        if val is None:
            if self.last_read_success:
                self.log(f"⚠️ [状态监测] 可用地勤读取失败")
                self.last_read_success = False
            return
        if not self.last_read_success:
            self.log(f"📊 [状态监测] 读取恢复: {val}")
            self.last_read_success = True
            self.last_checked_avail_staff = val
        elif val != self.last_checked_avail_staff:
            if self.last_checked_avail_staff == -1:
                self.log(f"📊 [状态监测] 当前可用地勤: {val}")
            else:
                self.log(f"📊 [状态监测] 可用地勤: {self.last_checked_avail_staff} -> {val}")
            self.last_checked_avail_staff = val

    def _check_and_perform_auto_delay(self):
        if self.auto_delay_count <= 0: return False

        if time.time() < self.doing_task_forbidden_until:
            return False

        screen = self.adb.get_screenshot()
        if screen is None: return False
        if self._is_tower_off(screen):
            return False
        is_triggered = False
        for (x, y) in self.TOWER_CHECK_POINTS:
            try:
                b, g, r = screen[y, x]
                if self._is_point_red(int(b), int(g), int(r)):
                    is_triggered = True
                    break
            except:
                pass
        if is_triggered:
            self.log(f"🚨 [最高优] 监测到自动塔台红灯...")
            self.adb.click(646, 822)
            self.sleep(0.8)
            self.adb.click(362, 785)
            self.sleep(1.0)
            confirm_success = False
            for i in range(3):
                res = self.adb.locate_image(self.icon_path + 'delay.png', confidence=0.8)
                if res:
                    self.log(f"   -> 发现确认按钮，第 {i + 1} 次点击...")
                    self.adb.click(res[0], res[1])
                    self.sleep(0.5)
                    if not self.adb.locate_image(self.icon_path + 'delay.png', confidence=0.8):
                        confirm_success = True
                        break
                else:
                    self.sleep(0.5)
            if confirm_success:
                self.log("   -> ✅ 延时确认成功")
                self.auto_delay_count -= 1
                if self.config_callback: self.config_callback("auto_delay_count", self.auto_delay_count)
            else:
                self.log("   -> ⚠️ 未能确认延时")
            self.sleep(0.5)
            if not self.wait_and_click('back.png', timeout=4.0, click_wait=0.5, random_offset=2):
                self.close_window()
            return True
        return False

    def handle_vehicle_check_task(self, target_pos=None):
        self.sleep(0.2)
        self.log(">>> [任务] 检查 Doing 状态...")
        red_warn = self.safe_locate('red_warning.png', confidence=0.75)
        if red_warn:
            if self.enable_vehicle_buy:
                self.log("   -> 🚨 发现车辆不足，准备购买")
                self.adb.click(red_warn[0], red_warn[1], random_offset=1)
                self.sleep(0.5)
                if self.wait_and_click('buy_vehicle.png', timeout=2.0, click_wait=1.0):
                    if self.wait_and_click('buy_vehicle_confirm.png', timeout=3.0, click_wait=0.5):
                        self.log("   -> ✅ 购买确认成功")
                    elif self.wait_and_click('back.png', timeout=3.0, click_wait=0.5, random_offset=2):
                        self.log("   -> 🛑 金钱不足，取消购买")
                        self.enable_vehicle_buy = False
                        if self.config_callback: self.config_callback("vehicle_buy", False)
                else:
                    self.log("   -> 未找到购买按钮")
            else:
                self.log("   -> 🚨 发现车辆不足，忽略")
            self.close_window()
            return True
        screen = self.adb.get_screenshot()
        if screen is not None:
            bx, by, bw, bh = self.REGION_BOTTOM_ROI
            roi_img = screen[by:by + bh, bx:bx + bw]
            res_done = self.adb.locate_image(self.icon_path + 'ground_support_done.png', confidence=0.8,
                                             screen_image=roi_img)
            if res_done:
                self.log("   -> 🕒 发现延误/完成飞机")
                if self.enable_delay_bribe:
                    agent_loc = self.safe_locate('stand_agent_false.png')
                    if agent_loc:
                        self.log("   -> [贿赂] 点击服务代理...")
                        self.adb.click(agent_loc[0], agent_loc[1])
                        if self.wait_and_click('stand_agent_true.png', timeout=2.0, click_wait=0):
                            self.log("   -> [贿赂] 代理已激活")
                self.adb.click(res_done[0] + bx, res_done[1] + by)
                self.log("   -> ✅ 点击结束服务")
                self.sleep(0.5)
                return True
        self.log("   -> 未发现可操作项目")
        self.close_window()
        return True

    def handle_approach_task(self, target_pos=None):
        self.sleep(0.2)
        if not self._verify_and_redirect('status_approach.png'): return True
        self.log(">>> [任务] 处理进场...")
        start_time = time.time()
        while time.time() - start_time < 2.0:
            self._check_running()
            if self.find_and_click('landing_permitted.png', wait=0):
                return True
            screen = self.adb.get_screenshot()
            if screen is None: continue
            vx, vy, vw, vh = self.REGION_VACANT_ROI
            vacant_roi = screen[vy:vy + vh, vx:vx + vw]
            res_vacant = self.adb.locate_image(self.icon_path + 'stand_vacant.png', confidence=0.8,
                                               screen_image=vacant_roi)
            if res_vacant:
                self.log("   -> 分配机位")
                self.adb.click(res_vacant[0] + vx, res_vacant[1] + vy)
                self.sleep(0.1)
                if self.wait_and_click('stand_confirm.png', timeout=1.5, click_wait=0):
                    w_start = time.time()
                    while time.time() - w_start < 2.5:
                        self._check_running()
                        if self.find_and_click('landing_permitted.png', wait=0):
                            return True

                        check_screen = self.adb.get_screenshot()
                        if check_screen is not None:
                            bx, by, bw, bh = self.REGION_BOTTOM_ROI
                            bottom_roi = check_screen[by:by + bh, bx:bx + bw]
                            if self.adb.locate_image(self.icon_path + 'landing_prohibited.png', confidence=0.8,
                                                     screen_image=bottom_roi):
                                return True

                        time.sleep(0.1)
                    return True
                else:
                    self.log("❌ 找不到确认按钮")
                    self.close_window()
                    return False
            time.sleep(0.1)
        self.log("⚠️ 进场超时")
        self.close_window()
        self.sleep(1.0)
        return False

    def handle_taxiing_task(self, target_pos=None):
        self.sleep(0.2)
        if not self._verify_and_redirect('status_taxiing.png'): return True
        self.log(">>> [任务] 处理跑道穿越...")
        if self.wait_and_click('cross_runway.png', timeout=3.0, click_wait=0): return True
        self.log("⚠️ 未找到按钮")
        self.close_window()
        return False

    def handle_takeoff_task(self, target_pos=None):
        self.sleep(0.2)
        if not self._verify_and_redirect('status_takeoff.png'): return True
        self.log(">>> [任务] 处理离场...")
        start_time = time.time()
        action_buttons = ['push_back.png', 'taxi_to_runway.png', 'wait.png', 'takeoff_by_gliding.png', 'takeoff.png',
                          'get_award_1.png', 'get_award_4.png', 'start_general.png']
        sm = getattr(self.adb, 'screenshot_method', 'adb')
        scan_timeout = 1.0 if sm in ('nemu_ipc', 'uiautomator2', 'droidcast_raw') else 5.0
        while time.time() - start_time < scan_timeout:
            self._check_running()
            screen = self.adb.get_screenshot()
            if screen is None: continue
            bx, by, bw, bh = self.REGION_BOTTOM_ROI
            roi_img = screen[by:by + bh, bx:bx + bw]
            for btn in action_buttons:
                res = self.adb.locate_image(self.icon_path + btn, confidence=0.8, screen_image=roi_img)
                if res:
                    x = res[0] + bx
                    y = res[1] + by
                    if btn == 'get_award_1.png' or btn == 'get_award_4.png':
                        self.log("   -> 🎁 发现领奖图标，进入流程")
                        self.adb.click(x, y)
                        got_step2 = False
                        t2_start = time.time()
                        while time.time() - t2_start < 15.0:
                            self._check_running()
                            if self.find_and_click('get_award_2.png', wait=0.5) or \
                                    self.find_and_click('get_award_3.png', wait=0.5) or \
                                    self.find_and_click('get_award_4.png', wait=0.5):
                                got_step2 = True
                                break
                            time.sleep(0.1)
                        if not got_step2:
                            self.log("🛑 领奖流程卡死")
                            return False
                        self.log("   -> 领奖确认，等待后续...")
                        t3_start = time.time()
                        while time.time() - t3_start < 15.0:
                            self._check_running()
                            s3_screen = self.adb.get_screenshot()
                            if s3_screen is None: continue
                            s3_roi = s3_screen[by:by + bh, bx:bx + bw]
                            for final_btn in ['push_back.png', 'taxi_to_runway.png', 'start_general.png']:
                                res_final = self.adb.locate_image(self.icon_path + final_btn, confidence=0.7,
                                                                  screen_image=s3_roi)
                                if res_final:
                                    self.sleep(1.5)
                                    self.adb.click(res_final[0] + bx, res_final[1] + by)
                                    self.log("   -> ✅ 离场动作执行完毕")
                                    return True

                            self.sleep(2.0)
                            if self.safe_locate('green_dot.png', region=self.REGION_GREEN_DOT):
                                self.log("   -> ⚠️ 检测到绿点，跳转至地勤分配...")
                                self.sleep(0.5)
                                return self.handle_stand_task()
                            else:
                                self.log("   -> ℹ️ 未检测到绿点，判定为塔台已接管")
                                return True

                            time.sleep(0.1)
                        return True
                    self.adb.click(x, y)
                    self.sleep(0.5)
                    return True
            time.sleep(0.1)
        self.log("⚠️ 离场任务扫描超时")
        self.close_window()
        return False

    def handle_stand_task(self, target_pos=None):
        self.sleep(0.1)
        if not self._verify_and_redirect('status_stand.png'): return True
        self.log(">>> [任务] 处理停机位队列...")
        while True:
            avail_staff = None
            is_read_success = False
            for _ in range(3):
                val = self.check_global_staff()
                if val is not None and val < 900:
                    avail_staff = val
                    is_read_success = True
                    self._update_staff_tracker(val)
                    break
                time.sleep(0.2)
            if avail_staff is None:
                self.log("⚠️ 无法读取地勤人数，尝试盲做")
                self._update_staff_tracker(None)
                avail_staff = 999
                is_read_success = False
            cost_text = self.ocr.recognize_number(self.REGION_TASK_COST, mode='task')
            required_cost = self.ocr.parse_cost(cost_text)
            if required_cost is None:
                self.log(f"⚠️ 读取花费失败，盲做")
                if self._perform_stand_action_sequence(force_verify=not is_read_success):
                    self.log("   -> 盲做成功")
                    self.stand_skip_index = 0
                    self.doing_task_forbidden_until = time.time() + 6.0
                    self.next_list_refresh_time = time.time() + 6.0
                    return True
                else:
                    if self.in_staff_shortage_mode:
                        self.log("🛑 盲做触发人员不足")
                        self.close_window()
                        if self._try_get_bonus_staff():
                            self.stand_skip_index = 0
                            return True
                    self.stand_skip_index += 1
                    return False
            self.log(f"   -> 需求: {required_cost}+1 | 可用: {avail_staff}")

            if avail_staff >= (required_cost + 1):
                if self.safe_locate('green_dot.png', region=self.REGION_GREEN_DOT):
                    self.log("   -> ✅ 地勤人员充足，开始分配")
                    if self._perform_stand_action_sequence(force_verify=not is_read_success):
                        self.log("   -> 地勤保障开始成功")
                        self.stand_skip_index = 0
                        self.doing_task_forbidden_until = time.time() + 6.0
                        self.next_list_refresh_time = time.time() + 6.0
                        return True
                    else:
                        if not self.in_staff_shortage_mode:
                            self.log("   -> 操作未成功，跳过本架")
                            self.stand_skip_index += 1
                            return False
                else:
                    self.log("   -> ⚠️ 人员充足但未检测到绿点，跳过")
                    self.close_window()
                    return False

            self.log(f"🛑 人力不足 (缺 {required_cost + 1 - avail_staff})")
            self.close_window()
            if self._try_get_bonus_staff():
                self.log("   -> 领取成功，重试")
                self.stand_skip_index = 0
                return True
            self.stand_skip_index += 1
            self.in_staff_shortage_mode = True
            self.last_known_available_staff = avail_staff
            return False

    def handle_ice_task(self, target_pos=None):
        self.sleep(0.2)
        if not self._verify_and_redirect('status_ice.png'): return True
        self.log(">>> [任务] 处理除冰...")
        if self.wait_and_click('start_ice.png', timeout=3.0, click_wait=0.5):
            self.log("   -> 除冰开始")
            return True
        self.log("⚠️ 未找到除冰按钮")
        self.close_window()
        return False

    def handle_repair_task(self, target_pos=None):
        self.sleep(0.2)
        self.log(">>> [任务] 处理维修/维护...")
        action_buttons = ['go_repair.png', 'start_repair.png', 'start_general.png']
        start_time = time.time()
        while time.time() - start_time < 3.0:
            self._check_running()
            screen = self.adb.get_screenshot()
            if screen is None: continue
            bx, by, bw, bh = self.REGION_BOTTOM_ROI
            roi_img = screen[by:by + bh, bx:bx + bw]
            for btn in action_buttons:
                res = self.adb.locate_image(self.icon_path + btn, confidence=0.8, screen_image=roi_img)
                if res:
                    abs_x = res[0] + bx
                    abs_y = res[1] + by
                    self.adb.click(abs_x, abs_y)
                    self.log(f"   -> ✅ 维护开始 ({btn})")
                    self.sleep(0.5)
                    return True
            time.sleep(0.1)
        self.log("⚠️ 未找到维修按钮")
        self.close_window()
        return False

    def _try_get_bonus_staff(self):
        if not self.enable_bonus_staff: return False
        now = time.time()
        if now < self.next_bonus_retry_time: return False
        self.log(">>> [福利] 开始领取流程...")
        if self.find_and_click('top_ground_staff.png', wait=0.8):
            self.adb.click(800, 580)
            self.sleep(2.0)
            if self.wait_and_click('get_staff.png', timeout=2.0, click_wait=1.0, confidence=0.7):
                self.log("   -> ✅ 领取成功！")
                self.sleep(2.0)
                self.next_bonus_retry_time = time.time() + (2 * 60 * 60)
                # back.png 偏移量减小
                self.find_and_click('back.png', wait=0.5, random_offset=2)
                self.sleep(3.0)
                return True
            else:
                self.log("   -> 未找到领取按钮")
                self.next_bonus_retry_time = time.time() + (15 * 60)
                # back.png 偏移量减小
                self.find_and_click('back.png', wait=0.5, random_offset=2)
                self.sleep(3.0)
                return True
        self.log("❌ 未找到顶部地勤图标")
        self.next_bonus_retry_time = time.time() + (15 * 60)
        return False

    def _perform_stand_action_sequence(self, force_verify=False):
        if self.slide_min_duration >= self.slide_max_duration:
            rand_duration = self.slide_min_duration
        else:
            rand_duration = random.randint(self.slide_min_duration, self.slide_max_duration)

        start_x = int(random.gauss(494, 5))
        start_y = int(random.gauss(574, 5))
        end_x = random.randint(800, 900)
        end_y = int(random.gauss(574, 5))

        self.log(f"   -> [动作] 拟人滑块: ({start_x},{start_y})->({end_x},{end_y}) 耗时:{rand_duration}ms")
        self.adb.swipe(start_x, start_y, end_x, end_y, duration_ms=rand_duration)

        self.sleep(0.1)
        agent_x = int(random.gauss(803, 5))
        agent_y = int(random.gauss(646, 5))
        self.adb.click(agent_x, agent_y)

        self.sleep(0.3)
        if self.safe_locate('insufficient_ground_staff.png', confidence=0.7):
            self.log("🛑 警告：人员不足 (盲操作后检测)")
            self.in_staff_shortage_mode = True
            self.last_known_available_staff = 0
            self.close_window()
            return False

        should_skip = self.enable_skip_staff and (not force_verify)
        if should_skip:
            self.log("   -> [极速] 跳过地勤验证")
        else:
            if self.enable_skip_staff and force_verify:
                self.log("   -> [安全] 地勤人数未知，强制执行验证")
            check_x, check_y = 63, 546
            b, g, r = self.adb.get_pixel_color(check_x, check_y)
            target_b, target_g, target_r = 153, 220, 96
            diff = abs(b - target_b) + abs(g - target_g) + abs(r - target_r)
            is_green = diff < 100

            if is_green:
                self.log("   -> [颜色检测] ✅ 通过")
            else:
                self.log(f"   -> [颜色检测] ❌ 失败 (diff={diff})")

            is_success_icon = self.safe_locate('stand_agent_true.png', confidence=0.85)
            if is_success_icon:
                self.log("   -> [图标检测] ✅ 通过")
            else:
                self.log("   -> [图标检测] ❌ 失败")
            if not (is_green and is_success_icon):
                self.log("🛑 验证失败：颜色或图标未通过")
                if self.safe_locate('insufficient_ground_staff.png', confidence=0.7):
                    self.log("   -> 原因：发现了人员不足警告")
                    self.in_staff_shortage_mode = True
                    self.last_known_available_staff = 0
                self.close_window()
                return False

        if self.wait_and_click('start_ground_support.png', timeout=2.0, click_wait=0):
            self.sleep(0.5)
            return True
        else:
            if self.safe_locate('insufficient_ground_staff.png', confidence=0.7):
                self.log("🛑 警告：人员不足 (寻找开始按钮时)")
                self.in_staff_shortage_mode = True
                self.last_known_available_staff = 0
            self.close_window()
            return False

    def _check_and_recover_interface(self):
        if self.safe_locate('main_interface.png', region=self.REGION_MAIN_ANCHOR, confidence=0.8):
            self.last_seen_main_interface_time = time.time()
            return
        elapsed = time.time() - self.last_seen_main_interface_time
        if elapsed > self.STUCK_TIMEOUT:
            self.log(f"🚨 [防卡死] 未检测到主界面已 {int(elapsed)}秒，尝试强行返回...")

            # First Start 恢复逻辑
            if self.find_and_click('first_start_1.png', wait=0.5):
                self.log("   -> 尝试 First Start 恢复流程...")
                fs_start = time.time()
                found_step2 = False
                while time.time() - fs_start < 5.0:
                    if self.find_and_click('first_start_2.png', wait=0.5):
                        found_step2 = True
                        break
                    time.sleep(0.5)

                if found_step2:
                    self.log("   -> 正在等待返回主界面...")
                    wait_main = time.time()
                    while time.time() - wait_main < 20.0:
                        if self.safe_locate('main_interface.png', region=self.REGION_MAIN_ANCHOR, confidence=0.8):
                            self.last_seen_main_interface_time = time.time()
                            self.log("   -> ✅ 恢复成功")
                            return
                        time.sleep(1.0)

            if self.wait_and_click('back.png', timeout=1.0, click_wait=0.5, random_offset=2):
                self.log("   -> 点击了 Back 按钮")
                self.last_seen_main_interface_time = time.time()
                return
            if self.wait_and_click('cancel.png', timeout=1.0, click_wait=0.5):
                self.log("   -> 点击了 Cancel 按钮")
                self.last_seen_main_interface_time = time.time()
                return
            self.log("   -> 尝试盲点关闭区域")
            self.close_window()
            self.last_seen_main_interface_time = time.time() - (self.STUCK_TIMEOUT - 5)

    def scan_and_process(self):
        if self.auto_delay_count > 0:
            if time.time() - self.last_window_close_time > 80:
                self.log("🛡️ [防遮挡] 强制关窗以检测塔台...")
                self.close_window()
                self.sleep(2.5)
                return True

        if self._check_and_perform_auto_delay(): return True
        if not hasattr(self, 'last_staff_check_time'): self.last_staff_check_time = 0
        now = time.time()
        prev_staff = self.last_checked_avail_staff
        staff_this_round = None
        if now - self.last_staff_check_time > 3.0:
            staff_this_round = self.check_global_staff()
            self._update_staff_tracker(staff_this_round)
            self.last_staff_check_time = now

        if self.in_staff_shortage_mode or self.stand_skip_index > 0:
            current_avail = staff_this_round if staff_this_round is not None else self.check_global_staff()
            if current_avail is not None:
                is_blind_recovery = (prev_staff >= 900) and (current_avail < 900)
                is_changed = (prev_staff != -1) and (current_avail != prev_staff)
                is_zero_recovery = (prev_staff == 0) and (current_avail > 0)
                is_safe_amount = current_avail >= 15
                if is_safe_amount or is_changed or is_blind_recovery or is_zero_recovery:
                    if is_changed: self.log(f"✅ 地勤变化 ({prev_staff}->{current_avail})，恢复")
                    self.in_staff_shortage_mode = False
                    self.stand_skip_index = 0
                self.last_checked_avail_staff = current_avail

        self._check_and_recover_interface()
        self._periodic_15s_check()

        current_screen = self.adb.get_screenshot()
        if current_screen is None:
            return False
        lx, ly, lw, lh = self.LIST_ROI_X, 0, self.LIST_ROI_W, self.LIST_ROI_H
        list_roi_img = current_screen[ly:ly + lh, lx:lx + lw]
        raw_detections, final_tasks = self._run_pending_detection(list_roi_img)
        if len(final_tasks) == 0 and self.adb.screenshot_method == "nemu_ipc":
            adb_screen = self.adb.get_screenshot(force_method="adb")
            if adb_screen is not None:
                list_roi_adb = adb_screen[ly:ly + lh, lx:lx + lw]
                _, final_tasks = self._run_pending_detection(list_roi_adb)
                if len(final_tasks) > 0:
                    if not getattr(self, '_nemu_ipc_fallback_logged', False):
                        self._nemu_ipc_fallback_logged = True
                        self._nemu_ipc_debug_save_mismatch(current_screen, adb_screen)
                        self.log(f"📋 [检测] nemu_ipc 与模板不匹配，已切换 ADB。请设 NEMU_IPC_DEBUG=1 并重启以生成调试截图，或尝试 NEMU_IPC_PIXEL_FORMAT=bgra / NEMU_IPC_FLIP=0")
                    self.adb.set_screenshot_method("adb")
        if len(final_tasks) == 0 and self.adb.screenshot_method == "droidcast_raw":
            adb_screen = self.adb.get_screenshot(force_method="adb")
            if adb_screen is not None:
                list_roi_adb = adb_screen[ly:ly + lh, lx:lx + lw]
                _, final_tasks = self._run_pending_detection(list_roi_adb)
                if len(final_tasks) > 0:
                    if not getattr(self, '_droidcast_raw_fallback_logged', False):
                        self._droidcast_raw_fallback_logged = True
                        self._droidcast_raw_debug_save_mismatch(current_screen, adb_screen)
                        self.log(f"📋 [检测] DroidCast_raw 与模板不匹配，已切换 ADB。请查看 droidcast_raw_debug/ 对比图排查。")
                    self.adb.set_screenshot_method("adb")

        doing_tasks = [d for d in final_tasks if d['type'] == 'doing']
        for det in doing_tasks:
            if time.time() <= self.doing_task_forbidden_until:
                continue
            self.log(f"⚡ [高优] 发现 Doing 任务 (分数: {det['score']:.2f})")
            self.adb.click(det['center'][0] + 60, det['center'][1], random_offset=3)
            try:
                return det['handler']()
            except:
                return det['handler'](None)

        valid_candidates = []
        skips_left = self.stand_skip_index
        for t in final_tasks:
            if t['type'] == 'doing': continue
            if t['type'] == 'stand':
                if self.in_staff_shortage_mode: continue
                if skips_left > 0:
                    skips_left -= 1
                    continue
            valid_candidates.append(t)

        if not valid_candidates:
            if not getattr(self, '_log_detect_once', False):
                self._log_detect_once = True
                n_doing = len(doing_tasks)
                n_total = len(final_tasks)
                self.log(f"📋 [检测] 任务数={n_total}, Doing={n_doing}, 有效候选=0 -> 关闭窗口")
            self.close_window()
            return False

        selected_task = None
        if self.enable_random_task and len(valid_candidates) > 1:
            top_k = 3
            if random.random() < 0.8:
                pool = valid_candidates[:top_k]
                selected_task = random.choice(pool)
            else:
                pool = valid_candidates[top_k:]
                if not pool:
                    pool = valid_candidates[:top_k]
                selected_task = random.choice(pool)
        else:
            selected_task = valid_candidates[0]

        time_until_refresh = self.next_list_refresh_time - time.time()
        if 0 < time_until_refresh < 0.8:
            self.sleep(time_until_refresh + 0.5)
            return self.scan_and_process()

        self.log(f"识别结果: {selected_task['name']} (分数: {selected_task['score']:.2f})")
        self.adb.click(selected_task['center'][0] + 60, selected_task['center'][1], random_offset=3)
        try:
            return selected_task['handler']()
        except:
            return selected_task['handler'](None)