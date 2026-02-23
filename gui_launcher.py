import sys
import os
import threading
import queue
import json
import datetime
import collections
import traceback
import tkinter as tk
import ctypes
import subprocess
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from tkinter.constants import BOTH, END, LEFT, RIGHT, TOP, X, Y

# === 引入现代 UI 库 ===
import ttkbootstrap as ttkb  # type: ignore[import-untyped]
from ttkbootstrap.constants import *  # type: ignore[import-untyped]  # noqa: F401, F403
from ttkbootstrap.tooltip import ToolTip  # type: ignore[import-untyped]

# === 引入 PIL 以修复图标显示 ===
from PIL import Image, ImageTk

# 引入后端逻辑
from adb_controller import set_custom_adb_path, AdbController, CURRENT_ADB_PATH, close_all_and_kill_server, get_woa_debug_dir

# MuMu 常用 ADB 端口（部分机型如 MuMu12+Vulkan 需用 MuMu 自带 adb 才能正常点击）
_MUMU_PORTS = {16384, 16385, 16416, 16448, 7555, 5555}


# 资源路径获取（兼容 PyInstaller 与 Nuitka）
def get_resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        if hasattr(sys, '_MEIPASS'):
            base = sys._MEIPASS
        else:
            base = os.path.dirname(sys.executable)
        return os.path.join(base, relative_path)
    base_path = os.path.dirname(os.path.abspath(__file__))
    p1 = os.path.join(base_path, relative_path)
    if os.path.exists(p1):
        return p1
    if hasattr(sys, 'executable'):
        exe_path = os.path.dirname(sys.executable)
        p2 = os.path.join(exe_path, relative_path)
        if os.path.exists(p2):
            return p2
    p3 = os.path.join(os.getcwd(), relative_path)
    if os.path.exists(p3):
        return p3
    return p1


_ICON_DIR = "icon"


CONFIG_FILE = "config.json"

LOCAL_VERSION = "v1.2.4.1"
_GITEE_RAW_URL = "https://gitee.com/shuang-nagi/WOA_AutoBot/raw/master/{}"
_GITEE_API_URL = "https://gitee.com/api/v5/repos/shuang-nagi/WOA_AutoBot/contents/{}?ref=master"
_GITEE_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
             "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36")

def _parse_version(ver_str):
    """将 'v1.2.3' / '1.2.3.4' 解析为整数元组 (1,2,3) / (1,2,3,4)，失败返回 ()"""
    import re
    m = re.findall(r'\d+', (ver_str or "").strip())
    return tuple(int(x) for x in m) if m else ()

def _remote_is_newer(remote_ver, local_ver=LOCAL_VERSION):
    """仅当在线版本严格大于本地版本时返回 True"""
    r, l = _parse_version(remote_ver), _parse_version(local_ver)
    return r > l if r and l else False

_gitee_ssl_ctx_cache = None

def _fetch_gitee_text(filename, timeout=10):
    """从 Gitee 获取文本文件，多策略回退，针对国内网络优化"""
    import urllib.request
    import base64
    import ssl
    global _gitee_ssl_ctx_cache

    headers = {
        "User-Agent": _GITEE_UA,
        "Referer": "https://gitee.com/shuang-nagi/WOA_AutoBot",
        "Accept": "text/plain, application/json, */*",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Connection": "close",
    }

    def _try_raw(ctx=None):
        req = urllib.request.Request(_GITEE_RAW_URL.format(filename), headers=headers)
        kw = {"timeout": timeout}
        if ctx is not None:
            kw["context"] = ctx
        with urllib.request.urlopen(req, **kw) as resp:
            return resp.read().decode("utf-8", errors="replace").strip()

    def _try_api(ctx=None):
        req = urllib.request.Request(_GITEE_API_URL.format(filename), headers=headers)
        kw = {"timeout": timeout}
        if ctx is not None:
            kw["context"] = ctx
        with urllib.request.urlopen(req, **kw) as resp:
            data = json.loads(resp.read())
            if isinstance(data, dict) and "content" in data:
                return base64.b64decode(data["content"]).decode("utf-8", errors="replace").strip()
        return None

    for attempt in (_try_raw, _try_api):
        try:
            result = attempt()
            if result:
                return result
        except Exception:
            pass

    if _gitee_ssl_ctx_cache is None:
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        _gitee_ssl_ctx_cache = ctx
    for attempt in (_try_raw, _try_api):
        try:
            result = attempt(_gitee_ssl_ctx_cache)
            if result:
                return result
        except Exception:
            pass

    return None

def _write_crash_report(exc_type, exc_value, exc_traceback):
    """写入崩溃报告文件，返回文件路径。任何阶段出错都不抛异常。"""
    crash_log_path = None
    try:
        from adb_controller import get_woa_debug_dir
        debug_dir = get_woa_debug_dir()
        os.makedirs(debug_dir, exist_ok=True)
        crash_log_path = os.path.join(debug_dir, f"crash_report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
    except Exception:
        try:
            crash_log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"crash_report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        except Exception:
            crash_log_path = f"crash_report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

    last_logs = ""
    try:
        if hasattr(sys.stdout, "log_buffer"):
            last_logs = "\n".join(list(sys.stdout.log_buffer))
        elif hasattr(sys.stdout, "stream") and hasattr(sys.stdout.stream, "log_buffer"):
            last_logs = "\n".join(list(sys.stdout.stream.log_buffer))
    except Exception:
        pass

    try:
        with open(crash_log_path, "w", encoding="utf-8") as f:
            f.write("=== WOA AutoBot CRASH REPORT ===\n")
            f.write(f"Time: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Thread: {threading.current_thread().name}\n\n")
            f.write("--- EXCEPTION STACK TRACE ---\n")
            traceback.print_exception(exc_type, exc_value, exc_traceback, file=f)
            f.write("\n--- LAST PRESERVED LOGS ---\n")
            f.write(last_logs if last_logs else "(No logs preserved in buffer)")
            f.write("\n\n=== END REPORT ===\n")
    except Exception:
        crash_log_path = None
    return crash_log_path


def handle_exception(exc_type, exc_value, exc_traceback):
    """全局未捕获异常处理，生成崩溃日志"""
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    crash_log_path = _write_crash_report(exc_type, exc_value, exc_traceback)

    try:
        if crash_log_path:
            print(f"\n🛑 [严重错误] 脚本发生异常退出，详细日志已保存至: {crash_log_path}")
        traceback.print_exception(exc_type, exc_value, exc_traceback)
    except Exception:
        pass

    is_main = (threading.current_thread() is threading.main_thread())
    if is_main:
        try:
            messagebox.showerror("程序崩溃", f"脚本发生严重错误，已保存详细日志到: {crash_log_path}")
        except Exception:
            pass


def _thread_excepthook(args):
    """Python 3.8+ 子线程未捕获异常兜底"""
    if args.exc_type is SystemExit:
        return
    handle_exception(args.exc_type, args.exc_value, args.exc_traceback)


sys.excepthook = handle_exception
threading.excepthook = _thread_excepthook

try:
    import faulthandler as _fh
    _crash_fd = None
    try:
        from adb_controller import get_woa_debug_dir as _get_dbg
        _dbg_dir = _get_dbg()
        os.makedirs(_dbg_dir, exist_ok=True)
        _crash_fd = open(os.path.join(_dbg_dir, "crash_segfault.log"), "a", encoding="utf-8")
    except Exception:
        pass
    _fh.enable(file=_crash_fd if _crash_fd else sys.stderr, all_threads=True)
except Exception:
    pass


# === 增强型日志重定向器 ===
class MultiTextRedirector(object):
    def __init__(self, widgets=None, tag="stdout"):
        if widgets is None:
            widgets = []
        self.widgets = widgets
        self.tag = tag
        self.log_buffer = collections.deque(maxlen=200)  # 保留最近200条日志以便发生错误时导出

    def add_widget(self, widget):
        if widget not in self.widgets:
            self.widgets.append(widget)
            self._setup_tags(widget)

    def _setup_tags(self, widget):
        widget.tag_config("time", foreground="#999999", font=("Consolas", 8))
        widget.tag_config("normal", foreground="#333333")
        widget.tag_config("success", foreground="#75b798")
        widget.tag_config("error", foreground="#ea868f")
        widget.tag_config("highlight", foreground="#fd7e14")
        widget.tag_config("method", foreground="#c9a227")
        widget.tag_config("update", foreground="#e63946", font=("Microsoft YaHei UI", 10, "bold"))

    def write(self, str_val):
        if "-> 执行动作:" in str_val: return
        if str_val == "\n":
            self._insert_to_all("\n", "normal")
            return

        now_str = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-4]
        time_prefix = f"[{now_str}] "

        tag = "normal"
        if "[版本更新]" in str_val:
            tag = "update"
        elif any(x in str_val for x in ["✅", "成功", "恢复", "通过"]):
            tag = "success"
        elif any(x in str_val for x in ["🛑", "❌", "错误", "失败", "严重", "卡死"]):
            tag = "error"
        elif any(x in str_val for x in ["⚠️", "警告", "注意", "跳过", "超时"]):
            tag = "highlight"
        elif any(x in str_val for x in ["[模式]", "触控方案", "截图方案", "触控:", "截图:"]):
            tag = "method"

        self.log_buffer.append(f"{time_prefix}{str_val}")
        self._insert_to_all(time_prefix, "time", str_val, tag)

    def _insert_to_all(self, txt1, tag1, txt2=None, tag2=None):
        for w in self.widgets:
            try:
                if not w.winfo_exists(): continue
                w.configure(state="normal")
                w.insert("end", txt1, (tag1,))
                if txt2: w.insert("end", txt2, (tag2,))

                # 日志长度控制：超过1000行自动删除最旧的
                try:
                    if int(w.index('end-1c').split('.')[0]) > 1000:
                        w.delete("1.0", "2.0")
                except Exception:
                    pass

                w.see("end")
                w.configure(state="disabled")
            except Exception:
                pass

    def flush(self):
        pass


class TeeToFile:
    """调试模式下将日志同时输出到控件和文件"""
    def __init__(self, stream, filepath):
        self.stream = stream
        self.filepath = filepath
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        self._file = open(filepath, "w", encoding="utf-8")
        self._file.write(f"=== WOA AutoBot 调试日志 {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===\n\n")

    def write(self, s):
        self.stream.write(s)
        try:
            ts = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-4]
            self._file.write(f"[{ts}] {s}")
            self._file.flush()
        except Exception:
            pass

    def flush(self):
        self.stream.flush()
        try:
            self._file.flush()
        except Exception:
            pass

    def close(self):
        try:
            self._file.close()
        except Exception:
            pass


class Application(ttkb.Window):
    def __init__(self):
        try:
            myappid = 'woabot.launcher.v1.2.4.1'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except:
            pass

        super().__init__(themename="cosmo")

        self.style.colors.success = "#75b798"
        self.style.colors.danger = "#ea868f"
        self.style.colors.primary = "#89b0ae"
        self.style.colors.info = "#9cbfdd"

        self.title("WOA AutoBot v1.2.4.1")
        self.geometry("680x850")
        self.last_geometry = "680x850"
        self.is_mini_mode = False

        self.config = self.load_config()
        self.var_bonus_staff = tk.BooleanVar(value=self.config.get("bonus_staff", False))
        self.var_vehicle_buy = tk.BooleanVar(value=self.config.get("vehicle_buy", False))
        self.var_speed_mode = tk.BooleanVar(value=self.config.get("speed_mode", False))
        self.var_skip_staff = tk.BooleanVar(value=self.config.get("skip_staff", False))
        self.var_delay_bribe = tk.BooleanVar(value=self.config.get("delay_bribe", False))
        self.var_delay_count = tk.StringVar(value=str(self.config.get("auto_delay_count", 0)))
        self.var_random_task = tk.BooleanVar(value=self.config.get("random_task_order", True))
        self.var_cancel_stand_filter = tk.BooleanVar(value=self.config.get("cancel_stand_filter", True))
        self.var_mini_top = tk.BooleanVar(value=False)

        if self.config.get("adb_path"):
            set_custom_adb_path(self.config["adb_path"])

        self.bot = None
        self.log_queue = queue.Queue()
        self.queue_check_interval = 100

        self.redirector = MultiTextRedirector()
        self._log_tee = None
        if os.environ.get("WOA_DEBUG", "").strip().lower() in ("1", "true", "yes"):
            try:
                debug_dir = get_woa_debug_dir()
                log_path = os.path.join(debug_dir, f"log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
                self._log_tee = TeeToFile(self.redirector, log_path)
                sys.stdout = self._log_tee
            except Exception:
                sys.stdout = self.redirector
        else:
            sys.stdout = self.redirector

        self.container_main = ttkb.Frame(self)
        self.container_mini = ttkb.Frame(self)

        self.setup_main_ui()
        self.setup_mini_ui()

        self.container_main.pack(fill=BOTH, expand=True)
        self.after(self.queue_check_interval, self.process_log_queue)

        def _emit_notice():
            m1 = "此脚本为开源免费项目，如您是从任何渠道，例如淘宝、闲鱼、拼多多购买的，请立即退款并举报！"
            m2 = "获取更新和反馈问题请加入QQ群1067076460。"
            print(m1)
            print(m2)
            orig = getattr(sys, "__stdout__", None)
            if orig and getattr(sys, "stdout", None) is not orig:
                try:
                    orig.write(m1 + "\n")
                    orig.write(m2 + "\n")
                    orig.flush()
                except Exception:
                    pass
        self.after(100, _emit_notice)

        def _online_checks():
            threading.Thread(target=self._check_version_and_popup, daemon=True).start()
        self.after(800, _online_checks)

        self.after(500, self.setup_window_icon)
        self.bind("<Map>", self._on_window_map)
        self._icon_loaded = False

        self.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _on_closing(self):
        """关闭窗口时停止脚本并清理资源，避免进程残留"""
        self.stop_bot()
        try:
            close_all_and_kill_server()
        except Exception:
            pass
        try:
            if getattr(self, "_log_tee", None):
                self._log_tee.close()
            sys.stdout = sys.__stdout__
        except Exception:
            pass
        self.destroy()
        try:
            self.quit()
        except Exception:
            pass

    def _on_window_map(self, event):
        if not self._icon_loaded and event.widget == self:
            self.setup_window_icon()
            self._icon_loaded = True

    def setup_window_icon(self):
        try:
            icon_rel = os.path.join(_ICON_DIR, "app.ico")
            icon_path = get_resource_path(icon_rel)
            if not os.path.exists(icon_path): return
            try:
                self.iconbitmap(default=icon_path)
            except:
                pass
            try:
                with open(icon_path, "rb") as f:
                    img = Image.open(f)
                    img.load()
                if hasattr(Image, 'Resampling'):
                    resample = Image.Resampling.LANCZOS
                else:
                    resample = Image.LANCZOS
                img16 = ImageTk.PhotoImage(img.resize((16, 16), resample))
                img32 = ImageTk.PhotoImage(img.resize((32, 32), resample))
                img48 = ImageTk.PhotoImage(img.resize((48, 48), resample))
                img64 = ImageTk.PhotoImage(img.resize((64, 64), resample))
                self.wm_iconphoto(True, img64, img48, img32, img16)
                self._icon_refs = [img16, img32, img48, img64]
            except:
                pass
        except:
            pass

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                return {}
        return {}

    def save_config(self):
        self.config["bonus_staff"] = self.var_bonus_staff.get()
        self.config["vehicle_buy"] = self.var_vehicle_buy.get()
        self.config["speed_mode"] = self.var_speed_mode.get()
        self.config["skip_staff"] = self.var_skip_staff.get()
        self.config["delay_bribe"] = self.var_delay_bribe.get()
        self.config["random_task_order"] = self.var_random_task.get()
        try:
            self.config["auto_delay_count"] = int(self.var_delay_count.get())
        except:
            self.config["auto_delay_count"] = 0
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=4)
        except Exception as e:
            print(f"配置保存失败: {e}")

    def create_info_icon(self, parent, text):
        lbl = ttkb.Label(parent, text="ⓘ", font=("Segoe UI Symbol", 10), bootstyle="secondary", cursor="hand2")
        ToolTip(lbl, text=text, bootstyle="secondary-inverse")
        return lbl

    def toggle_mode(self):
        if self.is_mini_mode:
            self.container_mini.pack_forget()
            self.geometry(self.last_geometry)
            self.attributes('-topmost', False)
            self.overrideredirect(False)
            self.container_main.pack(fill=BOTH, expand=True)
            self.is_mini_mode = False
        else:
            self.last_geometry = self.geometry()
            self.container_main.pack_forget()
            self.geometry("320x180")
            self.attributes('-topmost', self.var_mini_top.get())
            self.container_mini.pack(fill=BOTH, expand=True)
            self.is_mini_mode = True

    def toggle_mini_top_state(self):
        if self.is_mini_mode:
            self.attributes('-topmost', self.var_mini_top.get())

    def setup_mini_ui(self):
        pad = 5
        top_row = ttkb.Frame(self.container_mini)
        top_row.pack(fill=X, padx=pad, pady=(pad, 0))
        ttkb.Label(top_row, text="WOA Mini", font=("Arial", 9, "bold"), bootstyle="secondary").pack(side=LEFT)
        ttkb.Button(top_row, text="还原", bootstyle="outline-warning", command=self.toggle_mode, padding=(5, 0)).pack(
            side=RIGHT)
        cb_top = ttkb.Checkbutton(top_row, text="置顶", variable=self.var_mini_top, bootstyle="toolbutton-secondary",
                                  command=self.toggle_mini_top_state)
        cb_top.pack(side=RIGHT, padx=5)
        ctl_row = ttkb.Frame(self.container_mini)
        ctl_row.pack(fill=X, padx=pad, pady=2)
        self.btn_mini_start = ttkb.Button(ctl_row, text="▶", bootstyle="success", width=4, command=self.start_bot)
        self.btn_mini_start.pack(side=LEFT, padx=(0, 2), fill=X, expand=True)
        self.btn_mini_stop = ttkb.Button(ctl_row, text="■", bootstyle="danger", width=4, state="disabled",
                                         command=self.stop_bot)
        self.btn_mini_stop.pack(side=LEFT, padx=(2, 0), fill=X, expand=True)
        log_frame = ttkb.Frame(self.container_mini)
        log_frame.pack(fill=BOTH, expand=True, padx=pad, pady=pad)
        self.txt_mini_log = tk.Text(log_frame, state="disabled", font=("Consolas", 8), bg="#f8f9fa", relief="flat",
                                    height=4)
        self.txt_mini_log.pack(fill=BOTH, expand=True)
        self.redirector.add_widget(self.txt_mini_log)

    def setup_main_ui(self):
        main_pad = 15
        top_bar = ttkb.Frame(self.container_main, padding=main_pad)
        top_bar.pack(fill=X)
        lf_dev = ttkb.Labelframe(top_bar, text="设备连接", padding=10, bootstyle="primary")
        lf_dev.pack(side=LEFT, fill=X, expand=True)
        self.combo_devices = ttkb.Combobox(lf_dev, state="readonly", width=18)
        self.combo_devices.pack(side=LEFT, fill=X, expand=True, padx=(0, 5))
        self.btn_scan = ttkb.Button(lf_dev, text="智能扫描", bootstyle="outline-primary", command=self.refresh_devices)
        self.btn_scan.pack(side=LEFT)
        f_right = ttkb.Frame(top_bar)
        f_right.pack(side=RIGHT, padx=(15, 0), fill=Y)
        f_help_wrap = ttkb.Frame(f_right)
        f_help_wrap.pack(side=TOP, fill=X, pady=(5, 2))
        self.btn_help = ttkb.Button(f_help_wrap, text="📖 使用说明", bootstyle="outline-info", command=self.open_help_window)
        self.btn_help.pack(fill=X)
        self._help_badge = None
        ttkb.Button(f_right, text="⚙ 设置", bootstyle="outline-secondary", command=self.open_settings_window).pack(
            side=TOP, fill=X, pady=2)
        ttkb.Button(f_right, text="⤢ 小窗", bootstyle="outline-warning", command=self.toggle_mode).pack(side=TOP,
                                                                                                        fill=X,
                                                                                                        pady=(2, 0))
        lf_func = ttkb.Labelframe(self.container_main, text="功能配置", padding=main_pad, bootstyle="default")
        lf_func.pack(fill=X, padx=main_pad)

        def add_switch(parent, text, var, help_txt):
            f = ttkb.Frame(parent)
            f.pack(side=LEFT, padx=(0, 15), pady=8)
            ttkb.Checkbutton(f, text=text, variable=var, bootstyle="success-round-toggle",
                             command=self.sync_all_configs_to_bot).pack(side=LEFT)
            self.create_info_icon(f, help_txt).pack(side=LEFT, padx=5)

        f_row0 = ttkb.Frame(lf_func)
        f_row0.pack(fill=X)
        add_switch(f_row0, "自动领取地勤", self.var_bonus_staff,
                   "地勤不足时尝试领取免费地勤 (2小时冷却)；\n无法自动看广告，每次重新开始运行会重置冷却。")
        add_switch(f_row0, "自动购买地勤车辆", self.var_vehicle_buy, "地勤车辆不足时自动购买地勤车辆。\n（实验性功能）")
        add_switch(f_row0, "延误飞机贿赂", self.var_delay_bribe,
                   "处理延误飞机时，是否贿赂代理；\n开启后会消耗一定的银飞机。")

        f_row1 = ttkb.Frame(lf_func)
        f_row1.pack(fill=X)
        add_switch(f_row1, "塔台关闭时取消停机位筛选", self.var_cancel_stand_filter,
                   "开启后，塔台关闭时，脚本会强制取消停机位飞机的筛选，处理全部的待处理飞机。")

        f_tower = ttkb.Frame(f_row1)
        f_tower.pack(side=LEFT, padx=(15, 0), pady=8)
        ttkb.Label(f_tower, text="自动延时塔台:").pack(side=LEFT)
        ttkb.Entry(f_tower, textvariable=self.var_delay_count, width=4).pack(side=LEFT, padx=3)
        ttkb.Label(f_tower, text="次").pack(side=LEFT)
        ttkb.Button(f_tower, text="确认", bootstyle="outline-success", width=4, padding=0,
                    command=self.on_confirm_tower_delay).pack(side=LEFT, padx=5)
        self.create_info_icon(f_tower,
                              "填0表示功能关闭，最大值144；\n使用前请手动开启塔台，目前仅支持四个控制器全开，并设置好延时界面；\n你设置的延时时间是多久，脚本一次就延时多久，脚本不会主动修改。\n（实验性功能）").pack(
            side=LEFT)

        ctl_frame = ttkb.Frame(self.container_main, padding=main_pad)
        ctl_frame.pack(fill=X)
        self.btn_main_start = ttkb.Button(ctl_frame, text="▶ 启动脚本", bootstyle="success", command=self.start_bot)
        self.btn_main_start.pack(side=LEFT, fill=X, expand=True, padx=5)
        self.btn_main_stop = ttkb.Button(ctl_frame, text="⏹ 停止运行", bootstyle="danger", state="disabled",
                                         command=self.stop_bot)
        self.btn_main_stop.pack(side=LEFT, fill=X, expand=True, padx=5)
        log_group = ttkb.Labelframe(self.container_main, text="运行日志", padding=5)
        log_group.pack(fill=BOTH, expand=True, padx=main_pad, pady=(0, main_pad))
        self.txt_main_log = ScrolledText(log_group, state="disabled", font=("Consolas", 9))
        self.txt_main_log.pack(fill=BOTH, expand=True)
        self.redirector.add_widget(self.txt_main_log)
        # 先显示 UI，延迟执行首次扫描（避免启动卡顿）
        self.after(100, self._do_initial_scan)

    def _do_initial_scan(self):
        """后台线程执行首次设备扫描，完成后更新 UI"""
        def _worker():
            try:
                devs = AdbController.scan_devices(debug=True)
            except Exception as e:
                print(f">>> [扫描异常] {e}")
                devs = []
            self.after(0, lambda: self._apply_scan_result(devs))

        self.btn_scan.configure(text="扫描中...", state="disabled")
        for btn in [self.btn_main_start, self.btn_mini_start]:
            btn.configure(state="disabled")
        self.update_idletasks()
        t = threading.Thread(target=_worker)
        t.daemon = True
        t.start()

    def _apply_scan_result(self, devs):
        """在主线程更新扫描结果"""
        try:
            self.btn_scan.configure(text="智能扫描", state="normal")
            if not (getattr(self, 'bot', None) and self.bot.running):
                for btn in [self.btn_main_start, self.btn_mini_start]:
                    btn.configure(state="normal", text="▶ 启动脚本")
            self.combo_devices['values'] = devs
            if devs:
                self.combo_devices.current(0)
                print(f">>> 扫描完成: 发现 {len(devs)} 台设备")
            else:
                print(">>> 扫描完成: 未发现设备")
        except Exception:
            pass

    def _center_toplevel_on_parent(self, win):
        """将子窗口居中于主窗口"""
        self.update_idletasks()
        pw, ph = self.winfo_width(), self.winfo_height()
        px, py = self.winfo_rootx(), self.winfo_rooty()
        if pw < 100 or ph < 100:
            g = self.geometry()
            if "x" in g:
                parts = g.split("+")[0].split("x")
                if len(parts) == 2:
                    pw, ph = int(parts[0] or 680), int(parts[1] or 850)
        win.update_idletasks()
        w, h = win.winfo_reqwidth(), win.winfo_reqheight()
        if w <= 1 or h <= 1:
            g = win.geometry()
            if "x" in g:
                parts = g.split("+")[0].split("x")
                if len(parts) == 2:
                    w, h = int(parts[0] or 400), int(parts[1] or 400)
        x = px + max(0, (pw - w) // 2)
        y = py + max(0, (ph - h) // 2)
        win.geometry(f"+{x}+{y}")

    _LOCAL_HELP_CONTENT = """
· 下载失败，您看到的使用说明是离线版本！
· 如您的网络没有问题，请确认脚本的获取来源是否正常！

【声明】
- 此脚本为开源免费项目，如您是从任何渠道，例如淘宝、闲鱼、拼多多购买的，请立即退款并举报！
- 获取更新和反馈问题请加入QQ群1067076460。
- 项目开源地址：https://github.com/nj-yzf/WOA_AutoBot
- 如遇任何问题或bug，请在QQ群内或github上进行反馈。
- 脚本尚不稳定，如果造成账号内游戏币损失，本人概不负责！使用辅助工具有风险，请自行评估，如造成账号封禁，与作者无关！

【环境配置】
1. 仅支持在Windows系统上使用的安卓模拟器，本脚本专为MuMu模拟器优化，强烈推荐使用MuMu模拟器，模拟器分辨率必须设置为 1600x900。
2. Mumu模拟器默认地址为127.0.0.1:16384（其他模拟器或多开，请到模拟器设置内查看），并且自备加速器，保证网络通畅。
3. 请优先连接127.0.0.1:16384，127.0.0.1:16416之类的端口，尽量不要连接127.0.0.1:5555，emulator-5554之类的端口。
4. 使用MuMu模拟器时，请在设备设置中关闭“网络桥接模式”，关闭“后台挂机时保活运行”选项。
5. - 如模拟器连接遇到问题，请首先尝试手动指定ADB路径。
   - 如nemu_ipc方案无法启用，请首先尝试手动指定MuMu安装路径（指定到例如D:\Program Files\MuMuPlayer即可，不要指定到MuMuPlayer\nx_main文件夹）。
   - 如遇到未知问题，请尝试切换模拟器渲染模式为DirectX。

【使用须知】
1. 游戏语言：必须设置为[简体中文]。
2. 请勿与脚本同时操作！手动操作前请先停止运行。
3. 脚本使用双击空白处的方式关闭窗口，默认是窗口右上角附近的位置，如您发现脚本会误触飞机，请调整挂机视角，或将视角拉到最近并置于在空白处。
4. 机位分配只会点第一个，如果不希望C型机停DEF的机位等情况，需要手动筛选机位停机类型，并且与时刻表功能不兼容，请把时刻表重置。

【功能说明】
1. 推荐使用nemu_ipc + minitouch（默认，Mumu专用，且不支持Mumu国际版）或uiautomator2 + minitouch的方案。脚本运行速度主要取决于[截图方案]，运行速度如下：nemu_ipc > uiautomator2 >> droidcast_raw >= ADB。
2. 使用高速方案（如nemu_ipc或uiautomator2）时，由于速度很快，出错会增多，非常不建议关闭“跳过二次校验”和“跳过地勤分配验证”开关。 
3. 脚本运行时必须保持游戏右侧筛选选项中，仅筛选出带有黄色感叹号的待处理飞机。但您无需担心！脚本可以自动检测并调整筛选状态。
4. 使用“自动延时塔台”功能前，请保证您已开启塔台（且目前仅支持四个控制器全开），并设置好带有[延时]按钮的界面，脚本不会主动调整。

【已知问题和缺陷】
1. 脚本本身支持多开，但测试并不充分，多开很可能存在未知问题。若脚本正在运行时，开启（或关闭）第二个脚本或类似软件（如ALAS），会导致脚本运行中断，请注意，尝试停止后再重新运行。
2. 脚本很有可能被杀毒软件误杀，如您遇到类似问题，请关闭杀毒软件。
3. 脚本处理[需要维护]的飞机时，暂无法应对绿币不足的情况，请您根据机队规模，预留充足的绿币。
4. 任何情况下脚本目前都没有滑动右侧任务列表的能力。
5. 地勤不足时，脚本只会在可用地勤数量发生变化时尝试恢复分配，暂无法根据不同机型的需求智能分配。
6. 脚本无法设置起降飞机的比例，如您需要处理的飞机很多，请配合塔台使用。
"""

    def _check_version_and_popup(self):
        """后台线程：检查版本更新 + 首次弹窗"""
        remote_ver = _fetch_gitee_text("version.txt", timeout=8)
        self.after(0, lambda r=remote_ver: self._log_version_check(r))
        if remote_ver and _remote_is_newer(remote_ver):
            self.after(0, lambda v=remote_ver.strip(): self._show_update_notice(v))

        if not self.config.get("popup_shown"):
            popup_text = _fetch_gitee_text("popup.txt", timeout=5)
            if popup_text:
                self.after(0, lambda t=popup_text: self._show_first_run_popup(t))
            else:
                self.after(0, lambda: print("[联网] popup.txt 获取失败，跳过首次弹窗"))
                self.after(0, self._show_first_run_popup_offline)

        import hashlib
        instr_text = _fetch_gitee_text("instruction.txt", timeout=5)
        if instr_text:
            new_hash = hashlib.md5(instr_text.encode("utf-8")).hexdigest()
            old_hash = self.config.get("instruction_hash", "")
            if new_hash != old_hash:
                self._cached_instruction_text = instr_text
                self._cached_instruction_hash = new_hash
                self.after(0, self._show_help_badge)

    def _log_version_check(self, remote_ver):
        """在日志中输出本地版本与在线版本（无论是否一致）"""
        online = (remote_ver and remote_ver.strip()) or "获取失败"
        print(f"[版本检查] 本地版本：{LOCAL_VERSION}，在线版本：{online}")

    def _show_update_notice(self, remote_ver):
        print(f"⬆️ [版本更新] 发现新版本 {remote_ver}（当前 {LOCAL_VERSION}），请前往 QQ群 或 GitHub 获取更新！")
        print(f"⬆️ [版本更新] 下载地址：https://github.com/nj-yzf/WOA_AutoBot")

    def _show_first_run_popup(self, text):
        win = ttkb.Toplevel(self)
        win.title("通知")
        win.geometry("520x420")
        win.transient(self)
        win.grab_set()
        container = ttkb.Frame(win, padding=15)
        container.pack(fill=BOTH, expand=True)
        text_area = tk.Text(container, font=("Microsoft YaHei UI", 10), wrap="word",
                            bg="white", fg="#333333", relief="flat")
        text_area.pack(side=LEFT, fill=BOTH, expand=True)
        scroll = ttkb.Scrollbar(container, command=text_area.yview)
        scroll.pack(side=RIGHT, fill=Y)
        text_area.config(yscrollcommand=scroll.set)
        text_area.insert("end", text)
        text_area.configure(state="disabled")
        def _on_close():
            self.config["popup_shown"] = True
            self.save_config()
            win.destroy()
        win.protocol("WM_DELETE_WINDOW", _on_close)
        btn_frame = ttkb.Frame(win, padding=(15, 0, 15, 10))
        btn_frame.pack(fill=X)
        ttkb.Button(btn_frame, text="我知道了", bootstyle="primary", command=_on_close).pack(side=RIGHT)
        self._center_toplevel_on_parent(win)

    _FIRST_RUN_OFFLINE_POPUP = """网络连接失败，如您的网络没有问题；
请确认脚本的获取来源是否正常，确认您使用的是否为最新版本。

请注意：此脚本为开源免费项目！
如您是从任何渠道，例如淘宝、闲鱼、拼多多购买的，请立即退款并举报！
使用前请阅读窗口右上角【使用说明】

[此弹窗是一次性的，关闭后不会再出现]"""

    def _show_first_run_popup_offline(self):
        win = ttkb.Toplevel(self)
        win.title("通知")
        win.geometry("520x420")
        win.transient(self)
        win.grab_set()
        container = ttkb.Frame(win, padding=15)
        container.pack(fill=BOTH, expand=True)
        text_area = tk.Text(container, font=("Microsoft YaHei UI", 10), wrap="word",
                            bg="white", fg="#333333", relief="flat")
        text_area.pack(side=LEFT, fill=BOTH, expand=True)
        scroll = ttkb.Scrollbar(container, command=text_area.yview)
        scroll.pack(side=RIGHT, fill=Y)
        text_area.config(yscrollcommand=scroll.set)
        text_area.insert("end", self._FIRST_RUN_OFFLINE_POPUP)
        text_area.configure(state="disabled")
        def _on_close():
            self.config["popup_shown"] = True
            self.save_config()
            win.destroy()
        win.protocol("WM_DELETE_WINDOW", _on_close)
        btn_frame = ttkb.Frame(win, padding=(15, 0, 15, 10))
        btn_frame.pack(fill=X)
        ttkb.Button(btn_frame, text="我知道了", bootstyle="primary", command=_on_close).pack(side=RIGHT)
        self._center_toplevel_on_parent(win)

    def _show_help_badge(self):
        if self._help_badge is not None:
            return
        parent = self.btn_help.master
        try:
            bg = ttkb.Style().lookup("TFrame", "background") or "#ffffff"
        except Exception:
            bg = "#ffffff"
        dot = tk.Canvas(parent, width=10, height=10, highlightthickness=0, bd=0, bg=bg)
        dot.create_oval(1, 1, 9, 9, fill="#e63946", outline="#e63946")
        dot.place(relx=1.0, rely=0.0, anchor="ne", x=-2, y=2)
        self._help_badge = dot

    def _hide_help_badge(self):
        if self._help_badge is not None:
            self._help_badge.destroy()
            self._help_badge = None

    def open_help_window(self):
        win = ttkb.Toplevel(self)
        win.title("使用说明")
        win.geometry("900x800")
        container = ttkb.Frame(win, padding=15)
        container.pack(fill=BOTH, expand=True)
        text_area = tk.Text(container, font=("Microsoft YaHei UI", 10), wrap="word", bg="white", fg="#333333",
                            relief="flat")
        text_area.pack(side=LEFT, fill=BOTH, expand=True)
        scroll = ttkb.Scrollbar(container, command=text_area.yview)
        scroll.pack(side=RIGHT, fill=Y)
        text_area.config(yscrollcommand=scroll.set)
        text_area.insert("end", "正在加载...\n")
        text_area.configure(state="disabled")
        self._center_toplevel_on_parent(win)

        def _load():
            import hashlib
            cached = getattr(self, "_cached_instruction_text", None)
            if cached:
                content = cached
                new_hash = getattr(self, "_cached_instruction_hash", "")
            else:
                content = _fetch_gitee_text("instruction.txt", timeout=6)
                new_hash = hashlib.md5(content.encode("utf-8")).hexdigest() if content else ""
            failed = not content
            if not content:
                content = self._LOCAL_HELP_CONTENT
            def _fill():
                if failed:
                    print("[联网] instruction.txt 获取失败，已显示本地使用说明")
                text_area.configure(state="normal")
                text_area.delete("1.0", "end")
                text_area.insert("end", content)
                text_area.configure(state="disabled")
                if new_hash:
                    self.config["instruction_hash"] = new_hash
                    self.save_config()
                self._cached_instruction_text = None
                self._cached_instruction_hash = None
                self._hide_help_badge()
            self.after(0, _fill)
        threading.Thread(target=_load, daemon=True).start()

    def open_settings_window(self):
        if hasattr(self, 'settings_win') and self.settings_win.winfo_exists():
            self.settings_win.lift()
            return
        win = ttkb.Toplevel(self)
        self.settings_win = win
        win.title("高级设置")
        win.geometry("540x820")
        win.transient(self)
        win.grab_set()
        body = ttkb.Frame(win, padding=20)
        body.pack(fill=BOTH, expand=True)

        ttkb.Label(body, text="手动连接", font=("bold")).pack(anchor="w")
        f_manual = ttkb.Frame(body);
        f_manual.pack(fill=X, pady=5)
        e_manual_ip = ttkb.Entry(f_manual)
        e_manual_ip.pack(side=LEFT, fill=X, expand=True, padx=(0, 5))

        def run_manual_connect():
            ip = e_manual_ip.get().strip()
            if ip:
                print(f">>> 尝试手动连接: {ip}")
                try:
                    subprocess.run([CURRENT_ADB_PATH, "connect", ip], timeout=5, creationflags=0x08000000)
                    self.refresh_devices()
                except Exception as e:
                    print(f"❌ 连接失败: {e}")

        # 【颜色统一】手动连接按钮 -> 绿色
        ttkb.Button(f_manual, text="连接", bootstyle="success", command=run_manual_connect).pack(side=LEFT)

        ttkb.Separator(body).pack(fill=X, pady=10)
        ttkb.Label(body, text="ADB 路径", font=("bold")).pack(anchor="w")
        f_adb = ttkb.Frame(body);
        f_adb.pack(fill=X, pady=5)
        e_adb = ttkb.Entry(f_adb)
        e_adb.pack(side=LEFT, fill=X, expand=True, padx=(0, 5))
        if self.config.get("adb_path"): e_adb.insert(0, self.config["adb_path"])

        def browse():
            p = filedialog.askopenfilename(parent=win, filetypes=[("EXE", "*.exe")])
            if p: e_adb.delete(0, END); e_adb.insert(0, p)
            win.lift()

        # 【颜色统一】浏览路径按钮 -> 绿色边框
        ttkb.Button(f_adb, text="...", bootstyle="outline-success", command=browse).pack(side=LEFT)

        ttkb.Label(body, text="MuMu 安装路径", font=("bold")).pack(anchor="w")
        f_mumu = ttkb.Frame(body)
        f_mumu.pack(fill=X, pady=5)
        e_mumu = ttkb.Entry(f_mumu)
        e_mumu.pack(side=LEFT, fill=X, expand=True, padx=(0, 5))
        mp = self.config.get("mumu_path", "")
        if mp:
            e_mumu.insert(0, mp)

        def browse_mumu():
            p = filedialog.askdirectory(parent=win, title="选择 MuMu 安装目录")
            if p:
                e_mumu.delete(0, END)
                e_mumu.insert(0, p)
            win.lift()

        ttkb.Button(f_mumu, text="...", bootstyle="outline-success", command=browse_mumu).pack(side=LEFT)
        ToolTip(e_mumu, text="此路径仅用于 nemu_ipc 截图，非必需。留空则自动检测；自动检测成功时会回填此处。", bootstyle="info")

        ttkb.Separator(body).pack(fill=X, pady=10)
        ttkb.Label(body, text="触控方式（影响运行速度，但总体不明显）", font=("bold")).pack(anchor="w")
        f_ctrl = ttkb.Frame(body)
        f_ctrl.pack(fill=X, pady=5)
        ctrl_values = ("ADB", "minitouch", "uiautomator2")
        ctrl_method = ttkb.Combobox(f_ctrl, values=ctrl_values, state="readonly", width=16)
        ctrl_method.pack(side=LEFT, padx=(0, 5))
        cm = self.config.get("control_method", "minitouch").lower()
        idx = next((i for i, v in enumerate(ctrl_values) if v.lower() == cm), 0)
        ctrl_method.current(idx)
        tip = ("选择点击/滑动时使用的方案，速度越快脚本反应越灵敏。\n\n"
               "• ADB：系统自带方式，兼容性最好但较慢。\n"
               "• minitouch：速度最快。\n"
               "• uiautomator2：速度较快，滑动速度慢。")
        self.create_info_icon(f_ctrl, tip).pack(side=LEFT, padx=5)

        ttkb.Label(body, text="截图方式（对整体运行速度影响最大）", font=("bold")).pack(anchor="w")
        f_scshot = ttkb.Frame(body)
        f_scshot.pack(fill=X, pady=5)
        scshot_method = ttkb.Combobox(f_scshot, values=("ADB", "nemu_ipc", "uiautomator2", "DroidCast_raw"), state="readonly", width=16)
        scshot_method.pack(side=LEFT, padx=(0, 5))
        sm = self.config.get("screenshot_method", "nemu_ipc")
        if sm == "nemu_ipc":
            scshot_method.current(1)
        elif sm == "uiautomator2":
            scshot_method.current(2)
        elif sm == "droidcast_raw":
            scshot_method.current(3)
        else:
            scshot_method.current(0)
        self.create_info_icon(f_scshot,
            "选择获取屏幕画面的方式；截图越慢，整体运行越慢，建议优先选快的。\n\n"
            "• nemu_ipc：仅 MuMu 模拟器可用，速度极快。\n"
            "• uiautomator2：速度次快。\n"
            "• DroidCast_raw：速度较慢。\n"
            "• ADB：系统自带方式，兼容性最好但最慢。").pack(side=LEFT, padx=5)

        ttkb.Separator(body).pack(fill=X, pady=15)
        ttkb.Label(body, text="速度优化（风险选项）", font=("bold")).pack(anchor="w")
        f_filter = ttkb.Frame(body)
        f_filter.pack(fill=X, pady=5)
        ttkb.Checkbutton(f_filter, text="跳过二次校验", variable=self.var_speed_mode,
                         bootstyle="success-round-toggle").pack(side=LEFT)
        self.create_info_icon(f_filter,
                              "跳过对于飞机类型的二次校验；\n风险较低，运行速度提升轻微。").pack(side=LEFT, padx=5)
        f_skip = ttkb.Frame(body)
        f_skip.pack(fill=X, pady=5)
        ttkb.Checkbutton(f_skip, text="跳过地勤分配验证", variable=self.var_skip_staff,
                         bootstyle="success-round-toggle").pack(side=LEFT)
        self.create_info_icon(f_skip,
                              "地勤分配后不进行图标验证和颜色验证，直接开始；\n风险中等，可能导致飞机延误；\n仅推荐高峰期且有人在场时打开。").pack(side=LEFT, padx=5)

        ttkb.Separator(body).pack(fill=X, pady=10)
        ttkb.Label(body, text="防检测设置", font=("bold")).pack(anchor="w")

        f_rnd = ttkb.Frame(body);
        f_rnd.pack(fill=X, pady=5)
        # 【颜色统一】随机任务选择 -> 绿色开关
        ttkb.Checkbutton(f_rnd, text="随机任务选择", variable=self.var_random_task,
                         bootstyle="success-round-toggle").pack(side=LEFT)
        self.create_info_icon(f_rnd,
                              "开启后，脚本将在列表前3个任务中随机选择（80%概率），或从下方任务中随机选择（20%概率），以模拟真实操作。").pack(
            side=LEFT, padx=5)

        f_s = ttkb.Frame(body);
        f_s.pack(fill=X, pady=5)
        ttkb.Label(f_s, text="地勤分配—拖动随机耗时(ms):").pack(side=LEFT)
        e_min = ttkb.Entry(f_s, width=5);
        e_min.pack(side=LEFT, padx=5)
        e_min.insert(0, str(self.config.get("slide_min", 250)))
        ttkb.Label(f_s, text="-").pack(side=LEFT)
        e_max = ttkb.Entry(f_s, width=5);
        e_max.pack(side=LEFT, padx=5)
        e_max.insert(0, str(self.config.get("slide_max", 500)))
        self.create_info_icon(f_s,
                              "控制地勤分配界面中滑块操作的持续时间。\n建议范围 200-500ms。\n若最低值小于200，可能出现地勤分配时滑动不到位的情况").pack(
            side=LEFT, padx=5)

        f_t = ttkb.Frame(body);
        f_t.pack(fill=X, pady=5)
        ttkb.Label(f_t, text="随机思考时间:").pack(side=LEFT)
        c_th = ttkb.Combobox(f_t, values=("关闭", "短(0.1-0.4)", "中(0.3-1.0)", "长(0.8-2.0)"), state="readonly")
        cur = self.config.get("thinking_mode", 0)
        if 0 <= cur <= 3:
            c_th.current(cur)
        else:
            c_th.current(0)
        c_th.pack(side=LEFT, padx=5)
        self.create_info_icon(f_t,
                              "在点击操作前增加随机的“发呆”时间，\n模拟人类思考过程，大幅降低检测风险。\n追求极限速度可选择“关闭”。").pack(
            side=LEFT, padx=5)

        ttkb.Separator(body).pack(fill=X, pady=20)

        def save():
            old_cfg = dict(self.config)
            ap = e_adb.get().strip()
            if ap and os.path.exists(ap):
                self.config["adb_path"] = ap
                set_custom_adb_path(ap)
            else:
                if "adb_path" in self.config: del self.config["adb_path"]
                set_custom_adb_path(None)
            mp = e_mumu.get().strip()
            if mp and os.path.isdir(mp):
                self.config["mumu_path"] = mp
            else:
                if "mumu_path" in self.config: del self.config["mumu_path"]
            ctrl = ctrl_method.get().strip().lower()
            valid_ctrl = ("adb", "minitouch", "uiautomator2")
            self.config["control_method"] = ctrl if ctrl in valid_ctrl else "minitouch"
            sshot = scshot_method.get().strip().lower()
            self.config["screenshot_method"] = sshot if sshot in ("adb", "nemu_ipc", "uiautomator2", "droidcast_raw") else "nemu_ipc"
            try:
                vm = int(e_min.get());
                vx = int(e_max.get())
                if vm < 100: vm = 100
                if vx > 2000: vx = 2000
                if vx < vm: vx = vm
                self.config["slide_min"] = vm;
                self.config["slide_max"] = vx
            except:
                messagebox.showerror("错误", "输入整数", parent=win)
                return
            self.config["thinking_mode"] = c_th.current()
            self.config["speed_mode"] = self.var_speed_mode.get()
            self.config["skip_staff"] = self.var_skip_staff.get()
            self.config["cancel_stand_filter"] = self.var_cancel_stand_filter.get()
            self.config["random_task_order"] = self.var_random_task.get()

            changed = []
            if old_cfg.get("adb_path") != self.config.get("adb_path"):
                v = self.config.get("adb_path") or "自动"
                changed.append(("ADB 路径", str(v)))
            if old_cfg.get("mumu_path") != self.config.get("mumu_path"):
                v = self.config.get("mumu_path") or "自动"
                changed.append(("MuMu 安装路径", str(v)))
            if old_cfg.get("control_method") != self.config.get("control_method"):
                changed.append(("触控方式", self.config.get("control_method", "minitouch")))
            if old_cfg.get("screenshot_method") != self.config.get("screenshot_method"):
                changed.append(("截图方式", self.config.get("screenshot_method", "nemu_ipc")))
            if old_cfg.get("speed_mode") != self.config.get("speed_mode"):
                changed.append(("跳过二次校验", "开" if self.config.get("speed_mode") else "关"))
            if old_cfg.get("skip_staff") != self.config.get("skip_staff"):
                changed.append(("跳过地勤分配验证", "开" if self.config.get("skip_staff") else "关"))
            if old_cfg.get("cancel_stand_filter") != self.config.get("cancel_stand_filter"):
                changed.append(("塔台关闭取消停机位筛选", "开" if self.config.get("cancel_stand_filter") else "关"))

            anti_changed = (
                old_cfg.get("slide_min") != self.config.get("slide_min") or
                old_cfg.get("slide_max") != self.config.get("slide_max") or
                old_cfg.get("thinking_mode") != self.config.get("thinking_mode") or
                old_cfg.get("random_task_order") != self.config.get("random_task_order")
            )

            for name, val in changed:
                print(f">>> [高级设置] {name} 已更新: {val}")
            if anti_changed:
                print(f">>> [高级设置] 防检测: 随机任务={self.var_random_task.get()}, 滑块={vm}-{vx}ms, 思考时间={c_th.get()}")

            self.save_config()
            self.sync_all_configs_to_bot(from_advanced_save=True)
            win.destroy()

        ttkb.Button(body, text="保存设置", bootstyle="success", width=20, command=save).pack()
        win.after(50, lambda: self._center_toplevel_on_parent(win))

    def refresh_devices(self):
        print(">>> 正在扫描设备...")
        self.btn_scan.configure(text="扫描中...", state="disabled")
        for btn in [self.btn_main_start, self.btn_mini_start]:
            btn.configure(state="disabled")
        self.update()
        try:
            devs = AdbController.scan_devices(debug=True)
        except Exception as e:
            print(f">>> [扫描异常] {e}")
            devs = []
        finally:
            self.btn_scan.configure(text="智能扫描", state="normal")
            if not (getattr(self, 'bot', None) and self.bot.running):
                for btn in [self.btn_main_start, self.btn_mini_start]:
                    btn.configure(state="normal", text="▶ 启动脚本")
        self.combo_devices['values'] = devs
        if devs:
            self.combo_devices.current(0)
            print(f">>> 扫描完成: 发现 {len(devs)} 台设备")
        else:
            print(">>> 扫描完成: 未发现设备")

    def _try_use_mumu_adb_for_device(self, device_serial):
        """MuMu 设备（可读画面但无法点击时）自动切换到 MuMu 自带 adb"""
        if self.config.get("adb_path"):
            return
        if "127.0.0.1:" not in device_serial:
            return
        try:
            port = int(device_serial.split(":")[-1])
        except (ValueError, IndexError):
            return
        if port not in _MUMU_PORTS:
            return
        mumu_adb = AdbController._find_mumu_adb()
        if mumu_adb and os.path.isfile(mumu_adb):
            set_custom_adb_path(mumu_adb)
            print(f">>> [MuMu] 检测到 MuMu 设备，已切换至模拟器自带 ADB 以支持点击操作")

    def start_bot(self):
        device = self.combo_devices.get()
        if not device: messagebox.showwarning("提示", "请先选择设备"); return
        if self.bot and self.bot.running: return
        self.save_config()
        self._try_use_mumu_adb_for_device(device)
        for btn in [self.btn_main_start, self.btn_mini_start]:
            btn.configure(state="disabled", text="运行中...")
        for btn in [self.btn_main_stop, self.btn_mini_stop]:
            btn.configure(state="normal")
        self.combo_devices.configure(state="disabled")
        from main_adb import WoaBot
        self.bot = WoaBot(log_callback=self.log_to_queue, config_callback=self.on_bot_config_update)
        self.bot.set_device(device)
        self.sync_all_configs_to_bot()
        self.bot.start()

    def stop_bot(self):
        bot = self.bot
        if bot:
            bot.running = False
            bot.stop()
        self.bot = None
        for btn in [self.btn_main_start, self.btn_mini_start]:
            btn.configure(state="normal", text="▶ 启动脚本")
        for btn in [self.btn_main_stop, self.btn_mini_stop]:
            btn.configure(state="disabled")
        self.combo_devices.configure(state="readonly")
        print(">>> 脚本已停止")

    def on_confirm_tower_delay(self):
        self.sync_all_configs_to_bot()
        val_str = self.var_delay_count.get()
        if val_str == "0":
            print(f">>> [配置] 自动延时塔台: 已关闭")
        else:
            print(f">>> [配置] 自动延时塔台: 已更新为 {val_str} 次")

    def sync_all_configs_to_bot(self, from_advanced_save=False):
        no_log = from_advanced_save
        try:
            cnt = int(self.var_delay_count.get())
            if cnt < 0:
                cnt = 0
            elif cnt > 144:
                cnt = 144
        except ValueError:
            cnt = 0
        self.var_delay_count.set(str(cnt))
        self.config["auto_delay_count"] = cnt
        self.save_config()
        if self.bot:
            self.bot.set_bonus_staff_feature(self.var_bonus_staff.get())
            self.bot.set_vehicle_buy_feature(self.var_vehicle_buy.get())
            self.bot.set_speed_mode(self.var_speed_mode.get())
            self.bot.set_skip_staff_verify(self.var_skip_staff.get())
            self.bot.set_delay_bribe(self.var_delay_bribe.get())
            self.bot.set_auto_delay(cnt)
            self.bot.set_random_task_mode(self.var_random_task.get(), log_change=not no_log)
            self.bot.set_slide_duration_range(
                self.config.get("slide_min", 250), self.config.get("slide_max", 500), log_change=not no_log)
            self.bot.set_thinking_time_mode(self.config.get("thinking_mode", 0), log_change=not no_log)
            self.bot.set_cancel_stand_filter_when_tower_off(self.var_cancel_stand_filter.get())
            self.bot.set_control_method(self.config.get("control_method", "minitouch"))
            self.bot.set_screenshot_method(self.config.get("screenshot_method", "nemu_ipc"))
            self.bot.set_mumu_path(self.config.get("mumu_path", ""))

    def on_bot_config_update(self, key, value):
        if key == "auto_delay_count":
            self.var_delay_count.set(str(value))
        elif key == "vehicle_buy":
            self.var_vehicle_buy.set(bool(value))
        elif key == "mumu_path":
            self.config["mumu_path"] = value
            self.save_config()

    def log_to_queue(self, msg):
        self.log_queue.put(msg)

    def process_log_queue(self):
        # 【新增】UI防卡死保护：每次只处理最多50条日志
        # 即使后台疯狂报错，UI也不会因为要插入几千条日志而未响应
        count = 0
        while not self.log_queue.empty() and count < 50:
            self.log_queue.get()
            count += 1
        self.after(self.queue_check_interval, self.process_log_queue)


if __name__ == "__main__":
    try:
        app = Application()
        app.mainloop()
    except Exception:
        # 捕获 mainloop 中的异常并手动调用异常处理钩子
        if sys.excepthook:
            sys.excepthook(*sys.exc_info())
        else:
            traceback.print_exc()
            sys.exit(1)