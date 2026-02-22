import atexit
import ctypes
import socket
import subprocess
import time
import os
import sys
import re
import random
import threading

# 用于进程退出时清理残留（含非正常关闭）
_adb_instances = []


def kill_adb_server():
    """终止 adb server，释放对 adb_tools 目录的占用，避免无法删除打包文件夹"""
    adb_path = CURRENT_ADB_PATH if CURRENT_ADB_PATH and os.path.isfile(CURRENT_ADB_PATH) else None
    if not adb_path or adb_path == "adb":
        adb_path = get_bundled_resource_path(os.path.join("adb_tools", "adb.exe"))
    if not os.path.isfile(adb_path):
        return
    try:
        subprocess.run(
            [adb_path, "kill-server"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=3,
            creationflags=0x08000000,
        )
    except Exception:
        pass


def close_all_and_kill_server():
    """先关闭所有 AdbController（含 shell、minitouch、DroidCast 子进程），再终止 adb server，避免进程残留"""
    for c in list(_adb_instances):
        try:
            c.close()
        except Exception:
            pass
    kill_adb_server()


def _atexit_cleanup():
    for c in list(_adb_instances):
        try:
            c.close()
        except Exception:
            pass
    kill_adb_server()


atexit.register(_atexit_cleanup)

try:
    from emulator_discovery import (
        discover_all_serials_and_ports,
        get_mumu_adb_paths as _discover_mumu_adb,
    )
except ImportError:
    discover_all_serials_and_ports = None
    _discover_mumu_adb = None


def get_bundled_resource_path(relative_path):
    """获取资源路径，兼容 PyInstaller 与 Nuitka"""
    if getattr(sys, 'frozen', False):
        if hasattr(sys, '_MEIPASS'):
            # PyInstaller: 资源在临时解压目录
            base = sys._MEIPASS
        else:
            # Nuitka: 资源在 exe 同目录（--include-data-dir 输出位置）
            base = os.path.dirname(sys.executable)
        return os.path.join(base, relative_path)
    else:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


def _get_u2_jar_candidate_dirs():
    """返回 u2.jar 可能存在的目录（兼容开发环境与打包后，参考 ALAS uiautomator2cache）"""
    candidates = []
    if getattr(sys, "frozen", False):
        if hasattr(sys, "_MEIPASS"):
            candidates.append(sys._MEIPASS)
        candidates.append(os.path.dirname(sys.executable))
    candidates.append(os.path.dirname(os.path.abspath(__file__)))
    candidates.append(os.getcwd())
    try:
        import uiautomator2cache  # type: ignore[import-untyped]
        candidates.insert(0, os.path.join(os.path.dirname(uiautomator2cache.__file__), "cache"))
        candidates.insert(0, os.path.dirname(uiautomator2cache.__file__))
    except ImportError:
        pass
    return candidates


def _apply_u2_resource_patch():
    """打补丁：在 core 模块的 with_package_resource 处注入我们的查找逻辑（core 从 utils import 后持有引用，故需 patch core）"""
    try:
        import contextlib
        import pathlib
        u2_core = __import__("uiautomator2.core", fromlist=["with_package_resource"])
        if getattr(u2_core, "_woa_u2_patched", False):
            return
        _orig = getattr(u2_core, "with_package_resource", None)
        if _orig is None:
            return
        candidate_dirs = _get_u2_jar_candidate_dirs()

        @contextlib.contextmanager
        def _patched(filename):
            for base in candidate_dirs:
                p = pathlib.Path(base) / filename.replace("/", os.sep)
                if p.is_file():
                    yield p
                    return
            yield from _orig(filename)

        u2_core.with_package_resource = _patched
        u2_core._woa_u2_patched = True
    except Exception:
        pass


# 打包环境下提前打补丁，确保 u2 解压即用，无需用户额外配置
if getattr(sys, "frozen", False):
    try:
        _apply_u2_resource_patch()
    except Exception:
        pass


def find_adb_executable():
    """自动搜索 adb.exe"""
    internal_adb = get_bundled_resource_path(os.path.join("adb_tools", "adb.exe"))
    if os.path.exists(internal_adb):
        return internal_adb
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
        external_adb = os.path.join(exe_dir, "adb_tools", "adb.exe")
        if os.path.exists(external_adb):
            return external_adb
    return "adb"


DEFAULT_ADB_PATH = find_adb_executable()
CURRENT_ADB_PATH = DEFAULT_ADB_PATH


def set_custom_adb_path(path):
    global CURRENT_ADB_PATH
    if path and os.path.exists(path):
        if CURRENT_ADB_PATH != path:
            CURRENT_ADB_PATH = path
            print(f">>> [系统] ADB 路径已切换为: {CURRENT_ADB_PATH}")
    else:
        if CURRENT_ADB_PATH != DEFAULT_ADB_PATH:
            CURRENT_ADB_PATH = DEFAULT_ADB_PATH
            print(f">>> [系统] ADB 路径重置为默认")


class AdbController:
    """支持多种触控方案（参考 AzurLaneAutoScript）: adb, minitouch, uiautomator2"""
    VALID_CONTROL_METHODS = ("adb", "minitouch", "uiautomator2")

    def __init__(self, target_device=None, use_minitouch=False, screenshot_method="adb", control_method=None):
        self.device_serial = target_device
        if control_method is not None:
            m = (control_method or "adb").lower()
            self.control_method = m if m in self.VALID_CONTROL_METHODS else "adb"
        else:
            self.control_method = "minitouch" if use_minitouch else "adb"
        self.use_minitouch = (self.control_method == "minitouch")
        self.screenshot_method = (screenshot_method or "adb").lower()
        if self.screenshot_method not in ("adb", "nemu_ipc", "uiautomator2", "droidcast_raw"):
            self.screenshot_method = "adb"
        self.think_min = 0.0
        self.think_max = 0.0
        _adb_instances.append(self)

        # 【核心升级】持久化 Shell 进程
        self.shell_process = None
        self.shell_lock = threading.Lock()
        self._closed = False

        # minitouch（本脚本分辨率 1600x900，参考 ALAS 支持旋转）
        self._minitouch_client = None
        self._minitouch_port = 0
        self._minitouch_ready = False
        self._minitouch_lock = threading.Lock()
        self._minitouch_proc = None
        self._minitouch_max_x = 1600  # 从 ^ 行解析
        self._minitouch_max_y = 900
        self._minitouch_orientation = 0  # 0=正常 1=右 2=上 3=左（参考 ALAS）
        self._minitouch_screen_w = 1600
        self._minitouch_screen_h = 900

        # nemu_ipc 截图（MuMu12 专用）
        self._nemu_ipc_lib = None
        self._nemu_ipc_connect_id = 0
        self._nemu_ipc_folder = None
        self._nemu_ipc_instance_id = None
        self._nemu_ipc_logged = None  # "ok"/"fail" 用于避免重复日志

        # uiautomator2（可选，需 pip install uiautomator2）
        self._u2_device = None
        self._u2_fallback_logged = False
        self._u2_screenshot_fallback_logged = False

        # DroidCast_raw（参考 ALAS，需 assets/DroidCast_raw.apk）
        self._droidcast_port = 0
        self._droidcast_session = None
        self._droidcast_proc = None
        self._droidcast_fallback_logged = False

        # 初始化连接
        if self.device_serial:
            self._start_persistent_shell()

    def set_thinking_strategy(self, min_s, max_s):
        self.think_min = float(min_s)
        self.think_max = float(max_s)

    def set_control_method(self, method):
        m = (method or "adb").lower()
        if m in self.VALID_CONTROL_METHODS:
            self.control_method = m
            self.use_minitouch = (m == "minitouch")

    def _u2_ensure_assets(self, u2_module):
        """确保 u2.jar 可被 uiautomator2 找到，并打补丁（兼容打包后）"""
        _apply_u2_resource_patch()
        try:
            import shutil
            pkg_dir = os.path.dirname(os.path.abspath(u2_module.__file__))
            jar_in_pkg = os.path.join(pkg_dir, "assets", "u2.jar")
            if not os.path.isfile(jar_in_pkg):
                return
            for base in _get_u2_jar_candidate_dirs():
                assets_dir = os.path.join(base, "assets")
                jar_dst = os.path.join(assets_dir, "u2.jar")
                if os.path.isfile(jar_dst):
                    return
                try:
                    os.makedirs(assets_dir, exist_ok=True)
                    shutil.copy2(jar_in_pkg, jar_dst)
                    print(">>> [uiautomator2] 已复制 u2.jar 到 assets/")
                    return
                except (OSError, IOError):
                    continue
        except Exception:
            pass

    def set_screenshot_method(self, method):
        m = (method or "adb").lower()
        if m in ("adb", "nemu_ipc", "uiautomator2", "droidcast_raw"):
            self.screenshot_method = m

    def _u2_init(self):
        """初始化 uiautomator2 连接（参考 ALAS 对 u2.jar 路径的兼容处理）"""
        return self._u2_init_impl(for_control=True)

    def _u2_init_screenshot(self):
        """初始化 uiautomator2 连接（仅用于截图，不依赖 control_method）"""
        return self._u2_init_impl(for_control=False)

    def _u2_init_impl(self, for_control=True):
        """初始化 u2 设备：for_control=True 时需 control_method==u2；for_control=False 时需 screenshot_method==u2"""
        if self._u2_device is not None:
            return True
        if self._closed or not self.device_serial:
            return False
        if for_control and self.control_method != "uiautomator2":
            return False
        if not for_control and self.screenshot_method != "uiautomator2":
            return False
        try:
            import uiautomator2 as u2  # type: ignore[import-untyped]
            self._u2_ensure_assets(u2)  # 含补丁：core.with_package_resource 优先从 assets/ 查找
            self._u2_device = u2.connect(self.device_serial)
            msg = "触控将使用 atx-agent" if for_control else "截图将使用 atx-agent"
            print(f">>> [uiautomator2] 已启用，{msg}")
            return True
        except ImportError:
            flag = self._u2_fallback_logged if for_control else self._u2_screenshot_fallback_logged
            if not flag:
                if for_control:
                    self._u2_fallback_logged = True
                else:
                    self._u2_screenshot_fallback_logged = True
                print(">>> [uiautomator2] 未安装，请运行: pip install uiautomator2 ，回退到 ADB")
            return False
        except Exception as e:
            flag = self._u2_fallback_logged if for_control else self._u2_screenshot_fallback_logged
            if not flag:
                if for_control:
                    self._u2_fallback_logged = True
                else:
                    self._u2_screenshot_fallback_logged = True
                print(f">>> [uiautomator2] 连接失败: {e}，回退到 ADB")
            return False

    def _do_think(self):
        if self.think_max > 0:
            mu = (self.think_min + self.think_max) / 2
            sigma = (self.think_max - self.think_min) / 4
            wait = random.gauss(mu, sigma)
            wait = max(self.think_min, min(self.think_max, wait))
            time.sleep(wait)

    def _start_persistent_shell(self):
        """启动持久化 ADB Shell 会话"""
        if not self.device_serial or self._closed:
            return

        try:
            with self.shell_lock:
                if self.shell_process and self.shell_process.poll() is None:
                    return  # 依然存活

                cmd = [CURRENT_ADB_PATH, "-s", self.device_serial, "shell"]
                # 建立一个长连接管道，stdin 用于写入命令
                self.shell_process = subprocess.Popen(
                    cmd,
                    stdin=subprocess.PIPE,
                    stdout=subprocess.DEVNULL,  # 忽略输出以防阻塞
                    stderr=subprocess.DEVNULL,
                    creationflags=0x08000000  # 隐藏窗口
                )
                print(f">>> [底层] ADB Shell 长连接已建立: {self.device_serial}")
        except Exception as e:
            print(f"❌ [底层] 建立 Shell 失败: {e}")

    def _adb_forward(self, local, remote):
        """adb forward local remote"""
        r = self.run_cmd(["forward", local, remote], timeout=5)
        return r and r.returncode == 0

    def _minitouch_init(self):
        """初始化 minitouch：仅 x86 设备、推送二进制、启动服务、建立转发与 socket 连接"""
        with self._minitouch_lock:
            if self._minitouch_ready and self._minitouch_client:
                return True
            if self._closed or not self.device_serial or not self.use_minitouch:
                return False
        tout = 15 if "127.0.0.1" not in (self.device_serial or "") else 8
        abi_r = self.run_cmd(["shell", "getprop", "ro.product.cpu.abi"], timeout=tout)
        abi = ""
        if abi_r and abi_r.stdout:
            abi = (abi_r.stdout.decode("utf-8", errors="ignore") or "").strip().lower().replace("\r", "")
        # ABI -> minitouch 目录映射（须先匹配更具体的，避免 x86 误匹配 x86_64）
        _abi_map = [
            ("x86_64", ("minitouch-69336c8f34",)),
            ("arm64-v8a", ("minitouch-f7a806902f",)),
            ("x86", ("minitouch-1e3ccbf4fa", "minitouch_x86")),
            ("armeabi-v7a", ("minitouch-4575799dba",)),
        ]
        def _resolve_minitouch(name):
            # minitouch-xxx 目录下是 minitouch 文件；minitouch_x86 是根目录下的文件
            p = get_bundled_resource_path(os.path.join("adb_tools", name, "minitouch"))
            if os.path.isfile(p):
                return p
            p = get_bundled_resource_path(os.path.join("adb_tools", name))
            return p if os.path.isfile(p) else None

        local_path = None
        for key, candidates in _abi_map:
            if key in abi:
                for name in candidates:
                    local_path = _resolve_minitouch(name)
                    if local_path:
                        break
                break
        if not local_path:
            local_path = get_bundled_resource_path(os.path.join("adb_tools", "minitouch_x86"))
        if not os.path.isfile(local_path):
            print(f"⚠️ [minitouch] 未找到兼容 ABI={abi!r} 的 minitouch，回退到 ADB")
            self.use_minitouch = False
            return False
        _bin_name = os.path.basename(os.path.dirname(local_path)) if os.path.basename(local_path) == "minitouch" else os.path.basename(local_path)
        print(f">>> [minitouch] 设备 ABI={abi!r}，使用: {_bin_name}")
        r1 = self.run_cmd(["push", local_path, "/data/local/tmp/minitouch"], timeout=tout)
        if not r1 or r1.returncode != 0:
            err = (r1.stderr.decode("utf-8", errors="ignore") if r1 and r1.stderr else "未知")[:200]
            print(f"⚠️ [minitouch] 推送失败: {err}，回退到 ADB")
            self.use_minitouch = False
            return False
        self.run_cmd(["shell", "chmod", "755", "/data/local/tmp/minitouch"], timeout=tout)
        is_net = "127.0.0.1" not in (self.device_serial or "")
        port = 17392

        # 参考 ALAS：用持久子进程保持 minitouch 存活，避免 shell 退出后进程被收走
        self.run_cmd(["shell", "pkill", "-f", "minitouch"], timeout=tout)
        time.sleep(0.5)
        adb_path = get_bundled_resource_path(os.path.join("adb_tools", "adb.exe"))
        if not os.path.isfile(adb_path):
            adb_path = "adb"
        try:
            self._minitouch_proc = subprocess.Popen(
                [adb_path, "-s", self.device_serial, "shell", "/data/local/tmp/minitouch"],
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=0x08000000 if sys.platform == "win32" else 0,
            )
        except Exception as e:
            print(f"⚠️ [minitouch] 启动子进程失败: {e}，回退到 ADB")
            self.use_minitouch = False
            return False
        time.sleep(2.0 if is_net else 1.0)

        # ALAS 风格：forward + 连接重试循环（minitouch 可能稍慢就绪）
        self.run_cmd(["forward", "--remove", f"tcp:{port}"], timeout=5)
        if not self._adb_forward(f"tcp:{port}", "localabstract:minitouch"):
            if self._minitouch_proc:
                self._minitouch_proc.terminate()
                self._minitouch_proc = None
            print("⚠️ [minitouch] 端口转发失败，回退到 ADB")
            self.use_minitouch = False
            return False

        retry_count = 3
        client = None
        err_msg = "无输出"
        for attempt in range(retry_count):
            s = None
            try:
                print(f">>> [minitouch] 连接 127.0.0.1:{port} (尝试 {attempt + 1}/{retry_count})...")
                s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                s.settimeout(3)
                s.connect(("127.0.0.1", port))
                s.settimeout(2)
                out = s.makefile(mode="rb")
                first = out.readline().decode("utf-8", errors="replace").strip().replace("\r", "")
                if first and first[:1] in ("v", "^", "$"):
                    max_x, max_y = 1600, 900
                    line2 = out.readline().decode("utf-8", errors="replace").strip().replace("\r", "")
                    if line2.startswith("^"):
                        parts = line2.split()
                        if len(parts) >= 5:
                            max_x, max_y = int(parts[2]), int(parts[3])
                    out.readline()  # $ pid
                    self._minitouch_max_x, self._minitouch_max_y = max_x, max_y
                    self._minitouch_orientation = self._get_minitouch_orientation()
                    client = s
                    break
                s.close()
                err_msg = f"首行无效: {first[:60]!r}" if first else "无输出"
            except (socket.timeout, socket.error, OSError) as e:
                err_msg = str(e)
                if s:
                    try:
                        s.close()
                    except Exception:
                        pass
            if attempt < retry_count - 1:
                time.sleep(1.0)

        if client is None:
            if self._minitouch_proc:
                self._minitouch_proc.terminate()
                self._minitouch_proc = None
            print(f">>> [minitouch] handshake 失败: {err_msg}")
            raise RuntimeError("minitouch handshake failed")
        try:
            with self._minitouch_lock:
                self._minitouch_client = client
                self._minitouch_port = 17392
                self._minitouch_ready = True
            _o = self._minitouch_orientation
            _ominfo = f", orientation={_o}" if _o else ""
            print(f">>> [minitouch] 已启用 (1600x900{_ominfo})，触控将使用 socket 加速")
            return True
        except Exception as e:
            print(f"⚠️ [minitouch] 连接失败: {e}，回退到 ADB")
            self.use_minitouch = False
            return False

    def _get_minitouch_orientation(self):
        """获取设备方向 0/1/2/3（参考 ALAS dumpsys display）"""
        try:
            r = self.run_cmd(["shell", "dumpsys", "display"], timeout=8)
            if r and r.stdout:
                out = r.stdout.decode("utf-8", errors="ignore")
                # DisplayViewport{...orientation=N,...} 或 mCurrentOrientation=N
                m = re.search(r"orientation=(\d+)", out)
                if not m:
                    m = re.search(r"mCurrentOrientation=(\d+)", out)
                if m:
                    o = int(m.group(1))
                    if o in (0, 1, 2, 3):
                        return o
        except Exception:
            pass
        return 0

    def _minitouch_scale(self, x, y):
        """将屏幕坐标映射到 minitouch 设备坐标，含旋转（参考 ALAS convert）"""
        sw, sh = self._minitouch_screen_w, self._minitouch_screen_h
        mx, my = self._minitouch_max_x, self._minitouch_max_y
        o = self._minitouch_orientation
        if o == 0:
            pass
        elif o == 1:
            x, y = sh - y, x
            mx, my = my, mx
        elif o == 2:
            x, y = sw - x, sh - y
        elif o == 3:
            x, y = y, sw - x
            mx, my = my, mx
        return int(x * mx / sw), int(y * my / sh)

    def _adb_click_fallback(self, xi, yi):
        """ADB input tap 兜底"""
        if not self._write_shell_cmd(f"input tap {xi} {yi}"):
            self.run_cmd(["shell", "input", "tap", str(xi), str(yi)])

    def _minitouch_send(self, text):
        if not self._minitouch_client:
            return False
        try:
            self._minitouch_client.sendall(text.encode("utf-8"))
            time.sleep(0.012)
            return True
        except Exception:
            self._minitouch_ready = False
            self._minitouch_client = None
            return False

    def close(self):
        """关闭 Shell 长连接并终止子进程，避免进程残留导致文件夹无法删除"""
        try:
            _adb_instances.remove(self)
        except ValueError:
            pass
        self._closed = True
        with self._minitouch_lock:
            if self._minitouch_client:
                try:
                    self._minitouch_client.close()
                except Exception:
                    pass
                self._minitouch_client = None
                self._minitouch_ready = False
        if self._minitouch_proc:
            try:
                self._minitouch_proc.terminate()
                self._minitouch_proc.wait(timeout=2.0)
            except Exception:
                try:
                    self._minitouch_proc.kill()
                except Exception:
                    pass
            self._minitouch_proc = None
        self._u2_device = None
        if getattr(self, "_droidcast_proc", None):
            try:
                self._droidcast_proc.terminate()
                self._droidcast_proc.wait(timeout=2.0)
            except Exception:
                try:
                    self._droidcast_proc.kill()
                except Exception:
                    pass
            self._droidcast_proc = None
        self._droidcast_session = None
        self._droidcast_port = 0
        self.run_cmd(["forward", "--remove", "tcp:53516"], timeout=3)
        if self._nemu_ipc_connect_id and self._nemu_ipc_lib:
            try:
                self._nemu_ipc_lib.nemu_disconnect(self._nemu_ipc_connect_id)
            except Exception:
                pass
            self._nemu_ipc_connect_id = 0
        with self.shell_lock:
            if not self.shell_process:
                return
            try:
                if self.shell_process.poll() is None:
                    try:
                        if self.shell_process.stdin and not self.shell_process.stdin.closed:
                            self.shell_process.stdin.close()
                    except (OSError, BrokenPipeError):
                        pass
                    self.shell_process.terminate()
                    try:
                        self.shell_process.wait(timeout=2.0)
                    except subprocess.TimeoutExpired:
                        self.shell_process.kill()
                        self.shell_process.wait(timeout=1.0)
            except Exception:
                pass
            finally:
                self.shell_process = None

    def _write_shell_cmd(self, cmd_str):
        """通过长连接快速写入命令"""
        if not self.shell_process or self.shell_process.poll() is not None:
            self._start_persistent_shell()

        if self.shell_process:
            try:
                # 必须加 \n 模拟回车
                cmd_bytes = (cmd_str + "\n").encode('utf-8')
                self.shell_process.stdin.write(cmd_bytes)
                self.shell_process.stdin.flush()
                return True
            except (BrokenPipeError, OSError):
                # 管道断裂，尝试重启一次
                print("⚠️ [底层] Shell 管道断裂，尝试重连...")
                self._start_persistent_shell()
                try:
                    self.shell_process.stdin.write((cmd_str + "\n").encode('utf-8'))
                    self.shell_process.stdin.flush()
                    return True
                except:
                    return False
        return False

    def run_cmd(self, args, timeout=15):
        """普通一次性命令 (用于非交互式指令，如 pull/push/connect)"""
        cmd_list = [CURRENT_ADB_PATH or "adb"]
        if self.device_serial:
            cmd_list.extend(["-s", self.device_serial])

        if isinstance(args, list):
            cmd_list.extend([str(a) for a in args])
        elif isinstance(args, str):
            cmd_list.extend(args.split())

        try:
            result = subprocess.run(
                cmd_list,
                shell=False,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=timeout,
                creationflags=0x08000000
            )
            return result
        except subprocess.TimeoutExpired:
            print(f"❌ [ADB] 命令超时: {cmd_list[:4]}...")
            return None
        except Exception as e:
            print(f"❌ [ADB错误] {e}")
            return None

    @staticmethod
    def _find_mumu_adb():
        """尝试从常见路径及 ALAS 风格发现找到 MuMu 自带的 adb"""
        if _discover_mumu_adb:
            for p in _discover_mumu_adb():
                if p and os.path.isfile(p):
                    return p
        candidates = [
            r"E:\APP\MuMuPlayer\nx_main\adb.exe",
            os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), "Netease", "MuMu Player 12", "nx_main", "adb.exe"),
            os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), "Netease", "MuMu", "nx_main", "adb.exe"),
            os.path.join("D:", "Program Files", "Netease", "MuMu", "nx_main", "adb.exe"),
            os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), "Netease", "MuMu Player 6", "MuMu", "adb.exe"),
            os.path.join(os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)"), "Netease", "MuMu Player 12", "nx_main", "adb.exe"),
        ]
        for p in candidates:
            if p and os.path.isfile(p):
                return p
        for drive in ["C", "D", "E", "F"]:
            for base in [rf"{drive}:\Program Files\Netease", rf"{drive}:\APP", rf"{drive}:\Program Files (x86)\Netease"]:
                if not os.path.isdir(base):
                    continue
                try:
                    for name in os.listdir(base):
                        if "MuMu" in name or "MuMuPlayer" in name:
                            for sub in ["nx_main", "MuMu"]:
                                p = os.path.join(base, name, sub, "adb.exe")
                                if os.path.isfile(p):
                                    return p
                except Exception:
                    pass
        return None

    @staticmethod
    def _get_mumu_ports_from_vms():
        """从 MuMu vms 及注册表获取 ADB 端口（ALAS 风格 + 原有逻辑）"""
        if discover_all_serials_and_ports:
            _, ports = discover_all_serials_and_ports()
            if ports:
                return ports[:12]
        ports = []
        for base in [r"E:\APP\MuMuPlayer", os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), "Netease")]:
            vms_dir = os.path.join(base, "vms") if os.path.isdir(base) else None
            if not vms_dir or not os.path.isdir(vms_dir):
                continue
            try:
                for name in os.listdir(vms_dir):
                    if "MuMuPlayer" in name or "myandrovm" in name:
                        folder = os.path.join(vms_dir, name)
                        if not os.path.isdir(folder):
                            continue
                        for fname in os.listdir(folder):
                            if fname.endswith((".ini", ".cfg", ".vmx", ".nvram", ".nemu")) or "config" in fname.lower():
                                path = os.path.join(folder, fname)
                                if os.path.isfile(path):
                                    try:
                                        with open(path, "r", encoding="utf-8", errors="ignore") as f:
                                            content = f.read()
                                        for m in re.finditer(r"adb[_-]?port[\s\":=]*(\d+)", content, re.I):
                                            p = int(m.group(1))
                                            if 5000 < p < 70000 and p not in ports:
                                                ports.append(p)
                                        for m in re.finditer(r"hostport=\"(\d+)\"", content, re.I):
                                            p = int(m.group(1))
                                            if 5000 < p < 70000 and p not in ports:
                                                ports.append(p)
                                        for m in re.finditer(r"(\d{4,5})", content):
                                            p = int(m.group(1))
                                            if p in (16384, 16385, 16416, 16448, 5555, 5557, 5559) and p not in ports:
                                                ports.append(p)
                                    except Exception:
                                        pass
            except Exception:
                pass
        return ports[:12]

    @staticmethod
    def scan_devices(debug=False):
        """扫描所有可用模拟器（MuMu/雷电/BlueStacks 等），列出供用户选择"""
        creationflags = 0x08000000
        # 优先使用自带 adb_tools，以便同时连接 MuMu、雷电等不同模拟器
        bundled = get_bundled_resource_path(os.path.join("adb_tools", "adb.exe"))
        adb_path = CURRENT_ADB_PATH or "adb"
        if os.path.isfile(bundled):
            scan_adb = bundled
            if debug:
                print(f">>> [扫描调试] 使用自带 ADB 以识别多款模拟器: {scan_adb}")
        elif adb_path != "adb" and os.path.isfile(adb_path):
            scan_adb = adb_path
        else:
            mumu_adb = AdbController._find_mumu_adb()
            scan_adb = mumu_adb if (mumu_adb and os.path.isfile(mumu_adb)) else "adb"
            if debug and scan_adb != "adb":
                print(f">>> [扫描调试] 使用 MuMu ADB: {scan_adb}")
        if scan_adb != "adb" and not os.path.isfile(scan_adb):
            if debug:
                print(f">>> [扫描调试] ADB 路径无效: {scan_adb}")
            return []

        # 汇总所有可能端口（MuMu/雷电/BlueStacks 等）
        if discover_all_serials_and_ports:
            _, all_ports = discover_all_serials_and_ports()
            all_ports = list(dict.fromkeys(all_ports))
        else:
            vms_ports = AdbController._get_mumu_ports_from_vms()
            all_ports = list(dict.fromkeys(
                vms_ports + [5555, 5557, 5559, 5561, 16384, 16385, 16416, 16448, 7555,
                             5554, 5556, 62001, 62025, 16480, 16512, 16544, 16576]
            ))
        if debug:
            print(f">>> [扫描调试] 待连接端口数: {len(all_ports)}")

        def _devices_list():
            try:
                r = subprocess.run([scan_adb, "devices"], stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                   timeout=8, creationflags=creationflags)
                out = (r.stdout or b"").decode('utf-8', errors='ignore')
                err = (r.stderr or b"").decode('utf-8', errors='ignore')
                if debug and err.strip():
                    print(f">>> [扫描调试] adb devices stderr: {err.strip()[:300]}")
                devices = []
                for line in out.strip().split('\n'):
                    line = line.strip()
                    if not line or "List of devices" in line:
                        continue
                    parts = line.split()
                    if len(parts) >= 2 and parts[1] == 'device':
                        devices.append(parts[0])
                return devices
            except Exception as e:
                if debug:
                    print(f">>> [扫描调试] adb devices 异常: {e}")
                return []

        try:
            subprocess.run([scan_adb, "kill-server"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                           timeout=3, creationflags=creationflags)
        except Exception:
            pass
        time.sleep(0.25)
        try:
            subprocess.run([scan_adb, "start-server"], stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                           timeout=5, creationflags=creationflags)
        except Exception as e:
            if debug:
                print(f">>> [扫描调试] start-server 异常: {e}")
        time.sleep(0.2)

        # MuMuManager 连接 MuMu 实例 0、1、2（并发启动，提速）
        mumu_adb = AdbController._find_mumu_adb()
        if mumu_adb and os.path.isfile(mumu_adb):
            mumu_base = os.path.dirname(mumu_adb)
            for manager_name in ["MuMuManager.exe", "MuMuManager"]:
                manager_path = os.path.join(mumu_base, manager_name)
                if not os.path.isfile(manager_path):
                    manager_path = os.path.join(os.path.dirname(mumu_base), "shell", manager_name)
                if os.path.isfile(manager_path):
                    cwd = os.path.dirname(manager_path)
                    for instance_id in [0, 1, 2]:
                        try:
                            subprocess.Popen([manager_path, "adb", "-v", str(instance_id), "connect"],
                                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                             creationflags=creationflags, cwd=cwd)
                        except Exception:
                            pass
                    if debug:
                        print(f">>> [扫描调试] 已启动 MuMuManager 连接实例 0/1/2")
                    break

        # 用当前扫描用 adb 对全部端口执行 connect
        for port in all_ports:
            try:
                subprocess.Popen([scan_adb, "connect", f"127.0.0.1:{port}"],
                                 stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                 creationflags=creationflags)
            except Exception:
                pass
        time.sleep(1.5)
        devices = _devices_list()
        # 去重并保持顺序（127.0.0.1:xxx 与 emulator-xxxx 可能同时存在，保留用户可读的）
        seen = set()
        unique = []
        for d in devices:
            if d not in seen:
                seen.add(d)
                unique.append(d)
        if debug:
            print(f">>> [扫描调试] 最终设备列表: {unique if unique else '(无)'}")
            if not unique:
                print(">>> [扫描调试] 可尝试在高级设置中指定 ADB 路径")
        return unique

    def connect(self):
        if self.device_serial and ":" in self.device_serial:
            # 针对网络设备的重连逻辑
            self.run_cmd(["disconnect", self.device_serial])
            time.sleep(0.2)
            self.run_cmd(["connect", self.device_serial])

    def _nemu_ipc_find_folder_and_id(self):
        """从 serial 解析 MuMu12 安装路径和实例 ID（参考 ALAS NemuIpcImpl.serial_to_id）"""
        if not self.device_serial or ":" not in self.device_serial:
            return None, None
        try:
            port = int(self.device_serial.split(":")[1])
        except (ValueError, IndexError):
            return None, None
        index, offset = divmod(port - 16384 + 16, 32)
        offset -= 16
        if not (0 <= index < 32 and offset in (-2, -1, 0, 1, 2)):
            return None, None
        try:
            from emulator_discovery import get_mumu_serials_from_vms
        except ImportError:
            return None, None
        for serial, name, emu_dir in get_mumu_serials_from_vms():
            if serial != self.device_serial:
                continue
            if not emu_dir or not os.path.isdir(emu_dir):
                continue
            cand_folders = [emu_dir]
            for sub in ("MuMu Player 12", "MuMuPlayer-12.0", "MuMuPlayer12"):
                p = os.path.join(emu_dir, sub)
                if os.path.isdir(p):
                    cand_folders.append(p)
            for folder in cand_folders:
                for rel in ("shell/sdk/external_renderer_ipc.dll", "nx_device/12.0/shell/sdk/external_renderer_ipc.dll"):
                    fp = os.path.join(folder, rel)
                    if os.path.isfile(fp):
                        return folder, index
        return None, None

    def _nemu_ipc_capture_stderr(self, func, *args):
        """捕获 DLL 输出的 stderr（参考 ALAS CaptureNemuIpc）"""
        stderr_b = b''
        r = w = None
        try:
            fd_err = sys.stderr.fileno()
        except (ValueError, OSError):
            return stderr_b
        try:
            r, w = os.pipe()
            old_stderr = os.dup(fd_err)
            os.dup2(w, fd_err)
            sys.stderr.flush()
            try:
                func(*args)
            finally:
                os.dup2(old_stderr, fd_err)
                os.close(old_stderr)
                if w is not None:
                    os.close(w)
                    w = None
            chunks = []
            while r is not None:
                try:
                    chunk = os.read(r, 4096)
                    if not chunk:
                        break
                    chunks.append(chunk)
                except (OSError, BlockingIOError):
                    break
            stderr_b = b''.join(chunks)
        except Exception:
            pass
        if r is not None:
            try:
                os.close(r)
            except OSError:
                pass
        if w is not None:
            try:
                os.close(w)
            except OSError:
                pass
        return stderr_b

    def _nemu_ipc_debug_save(self, img_bgr, arr_raw, width, height, pixel_fmt, do_flip):
        """调试：保存 nemu_ipc 截图及对比用 ADB 截图，便于排查格式/方向问题"""
        import cv2
        import datetime
        debug_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "nemu_ipc_debug")
        try:
            os.makedirs(debug_dir, exist_ok=True)
        except OSError:
            return
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        cnt = getattr(self, "_nemu_ipc_debug_count", 0)
        if cnt >= 5:
            return
        self._nemu_ipc_debug_count = cnt + 1
        try:
            p_nemu = os.path.join(debug_dir, f"nemu_{ts}_fmt{pixel_fmt}_flip{int(do_flip)}.png")
            cv2.imwrite(p_nemu, img_bgr)
            bgr_mean = img_bgr.mean(axis=(0, 1))
            print(f">>> [nemu_ipc DEBUG] 已保存: {p_nemu}  shape={img_bgr.shape} BGR_mean={bgr_mean.round(1).tolist()}")
            adb_img = self.get_screenshot(force_method="adb")
            if adb_img is not None and adb_img.size > 0:
                p_adb = os.path.join(debug_dir, f"adb_{ts}.png")
                cv2.imwrite(p_adb, adb_img)
                adb_mean = adb_img.mean(axis=(0, 1))
                print(f">>> [nemu_ipc DEBUG] 已保存: {p_adb}  BGR_mean={adb_mean.round(1).tolist()}")
        except Exception as e:
            print(f">>> [nemu_ipc DEBUG] 保存失败: {e}")

    def _get_screenshot_nemu_ipc(self):
        """MuMu12 nemu_ipc 截图（参考 ALAS nemu_ipc.py）"""
        import cv2
        import numpy as np
        try:
            if self._nemu_ipc_lib is None:
                folder, instance_id = self._nemu_ipc_find_folder_and_id()
                if folder is None or instance_id is None:
                    if self._nemu_ipc_logged != "fail":
                        self._nemu_ipc_logged = "fail"
                        print(">>> [nemu_ipc] 未找到 MuMu12 或端口非 16xxx，回退到 ADB 截图")
                    return None
                for rel in ("shell/sdk/external_renderer_ipc.dll", "nx_device/12.0/shell/sdk/external_renderer_ipc.dll"):
                    dll_path = os.path.join(folder, rel)
                    if os.path.isfile(dll_path):
                        self._nemu_ipc_lib = ctypes.CDLL(dll_path)
                        self._nemu_ipc_lib.nemu_connect.argtypes = [ctypes.c_wchar_p, ctypes.c_int]
                        self._nemu_ipc_lib.nemu_connect.restype = ctypes.c_int
                        self._nemu_ipc_folder = folder
                        self._nemu_ipc_nx_device = "nx_device" in rel
                        self._nemu_ipc_instance_id = instance_id
                        print(f">>> [nemu_ipc] 已加载 DLL，folder={folder!r}，instance_id={instance_id}")
                        break
                if self._nemu_ipc_lib is None:
                    if self._nemu_ipc_logged != "fail":
                        self._nemu_ipc_logged = "fail"
                        print(">>> [nemu_ipc] 未找到 external_renderer_ipc.dll，回退到 ADB 截图")
                    return None
            if self._nemu_ipc_connect_id == 0:
                folders_to_try = [os.path.abspath(self._nemu_ipc_folder)]
                if getattr(self, "_nemu_ipc_nx_device", False):
                    alt = os.path.join(self._nemu_ipc_folder, "nx_device", "12.0")
                    if os.path.isdir(alt):
                        folders_to_try.append(os.path.abspath(alt))
                connect_id = 0
                last_stderr = b""
                for folder_path in folders_to_try:
                    connect_id = self._nemu_ipc_lib.nemu_connect(folder_path, self._nemu_ipc_instance_id)
                    if connect_id > 0:
                        break
                    last_stderr = self._nemu_ipc_capture_stderr(
                        lambda f=folder_path: self._nemu_ipc_lib.nemu_connect(f, self._nemu_ipc_instance_id)
                    )
                if connect_id == 0:
                    err_msg = last_stderr.decode("utf-8", errors="replace").strip() if last_stderr else ""
                    if self._nemu_ipc_logged != "fail":
                        self._nemu_ipc_logged = "fail"
                        reason = ">>> [nemu_ipc] nemu_connect 失败，回退到 ADB 截图。"
                        reason += " MuMu 5.x 可能与 nemu_ipc 不兼容，当前已使用 ADB 截图。"
                        reason += f" DLL stderr: {err_msg if err_msg else '(无)'}"
                        print(reason)
                    return None
                self._nemu_ipc_connect_id = connect_id
                if self._nemu_ipc_logged != "ok":
                    self._nemu_ipc_logged = "ok"
                    print(">>> [nemu_ipc] 已启用，仅 MuMu 可用，速度极快")
            w_ptr = ctypes.pointer(ctypes.c_int(0))
            h_ptr = ctypes.pointer(ctypes.c_int(0))
            ret = self._nemu_ipc_lib.nemu_capture_display(
                self._nemu_ipc_connect_id, 0, 0, w_ptr, h_ptr, None
            )
            if ret > 0:
                return None
            width, height = w_ptr.contents.value, h_ptr.contents.value
            if width <= 0 or height <= 0:
                return None
            length = width * height * 4
            buf = (ctypes.c_ubyte * length)()
            ret = self._nemu_ipc_lib.nemu_capture_display(
                self._nemu_ipc_connect_id, 0, length, w_ptr, h_ptr, ctypes.byref(buf)
            )
            if ret > 0:
                return None
            arr = np.ctypeslib.as_array(buf).reshape((height, width, 4)).copy()
            # ALAS 文档注明 nemu_ipc 返回 RGBA、图像倒置；MuMu 5.x 若仍不匹配可设 NEMU_IPC_PIXEL_FORMAT=bgra 或 NEMU_IPC_FLIP=0
            pixel_fmt = os.environ.get("NEMU_IPC_PIXEL_FORMAT", "rgba").lower()
            if pixel_fmt == "bgra":
                img = cv2.cvtColor(arr, cv2.COLOR_BGRA2BGR)
            else:
                img = cv2.cvtColor(arr, cv2.COLOR_RGBA2BGR)
            do_flip = os.environ.get("NEMU_IPC_FLIP", "1") != "0"
            if do_flip:
                img = cv2.flip(img, 0)
            if os.environ.get("NEMU_IPC_DEBUG") == "1":
                self._nemu_ipc_debug_save(img, arr, width, height, pixel_fmt, do_flip)
            return img
        except Exception as e:
            if self._nemu_ipc_logged != "fail":
                self._nemu_ipc_logged = "fail"
                print(f">>> [nemu_ipc] 异常: {e}，回退到 ADB 截图")
            return None

    def _get_droidcast_raw_apk_path(self):
        """获取 DroidCast_raw.apk 路径（参考 ALAS，需从 Torther/DroidCastS 下载）"""
        for base in _get_u2_jar_candidate_dirs():
            p = os.path.join(base, "assets", "DroidCast_raw.apk")
            if os.path.isfile(p):
                return p
        return None

    def _droidcast_init(self):
        """初始化 DroidCast_raw：推送 APK、启动服务、建立端口转发（参考 ALAS droidcast.py）"""
        if self._droidcast_session is not None and self._droidcast_port > 0:
            return True
        if self._closed or not self.device_serial or self.screenshot_method != "droidcast_raw":
            return False
        apk_path = self._get_droidcast_raw_apk_path()
        if not apk_path:
            if not self._droidcast_fallback_logged:
                self._droidcast_fallback_logged = True
                print(">>> [DroidCast_raw] 未找到 assets/DroidCast_raw.apk，请从 https://github.com/Torther/DroidCastS/releases 下载并放入 assets/")
            return False
        try:
            self._droidcast_stop()
            r = self.run_cmd(["push", apk_path, "/data/local/tmp/DroidCast_raw.apk"], timeout=15)
            if not r or r.returncode != 0:
                if not self._droidcast_fallback_logged:
                    self._droidcast_fallback_logged = True
                    print(">>> [DroidCast_raw] 推送 APK 失败，回退到 ADB")
                return False
            adb_path = CURRENT_ADB_PATH if CURRENT_ADB_PATH and os.path.isfile(CURRENT_ADB_PATH) else None
            if not adb_path:
                adb_path = get_bundled_resource_path(os.path.join("adb_tools", "adb.exe"))
            adb_path = adb_path if adb_path and os.path.isfile(adb_path) else "adb"
            try:
                self._droidcast_proc = subprocess.Popen(
                    [adb_path, "-s", self.device_serial, "shell",
                     "CLASSPATH=/data/local/tmp/DroidCast_raw.apk app_process / ink.mol.droidcast_raw.Main"],
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=0x08000000 if sys.platform == "win32" else 0,
                )
            except Exception as e:
                if not self._droidcast_fallback_logged:
                    self._droidcast_fallback_logged = True
                    print(f">>> [DroidCast_raw] 启动失败: {e}，回退到 ADB")
                return False
            time.sleep(2.0)
            self.run_cmd(["forward", "--remove", "tcp:53516"], timeout=5)
            if not self._adb_forward("tcp:53516", "tcp:53516"):
                if getattr(self, "_droidcast_proc", None):
                    try:
                        self._droidcast_proc.terminate()
                    except Exception:
                        pass
                    self._droidcast_proc = None
                if not self._droidcast_fallback_logged:
                    self._droidcast_fallback_logged = True
                    print(">>> [DroidCast_raw] 端口转发失败，回退到 ADB")
                return False
            self._droidcast_port = 53516
            self._droidcast_session = True  # 使用 urlopen 无需 session 对象
            print(">>> [DroidCast_raw] 已启用")
            return True
        except Exception as e:
            if not self._droidcast_fallback_logged:
                self._droidcast_fallback_logged = True
                print(f">>> [DroidCast_raw] 初始化异常: {e}，回退到 ADB")
            return False

    def _droidcast_stop(self):
        """停止 DroidCast 进程（参考 ALAS）"""
        try:
            self.run_cmd(["shell", "pkill", "-9", "-f", "ink.mol.droidcast_raw"], timeout=5)
        except Exception:
            pass
        try:
            self.run_cmd(["shell", "pkill", "-9", "-f", "DroidCast"], timeout=5)
        except Exception:
            pass
        self.run_cmd(["forward", "--remove", "tcp:53516"], timeout=3)

    def _get_screenshot_droidcast_raw(self):
        """DroidCast 截图（PNG模式）：使用 /preview 接口避免 RGB565 格式问题，请求 1600x900"""
        import cv2
        import numpy as np
        if not self._droidcast_init():
            return None
        
        # 尝试请求 PNG 预览图，指定分辨率
        w_req, h_req = 1600, 900
        try:
            from urllib.request import urlopen
            # DroidCast_raw 也支持 /preview 接口返回 PNG
            url = f"http://127.0.0.1:53516/preview?width={w_req}&height={h_req}"
            resp = urlopen(url, timeout=5)
            image = resp.read()
        except Exception:
            return None
            
        if not image:
            return None
            
        try:
            # 解码 PNG
            arr = np.frombuffer(image, np.uint8)
            img = cv2.imdecode(arr, cv2.IMREAD_COLOR)
            if img is None:
                return None
                
            # 检查分辨率并修正
            h, w = img.shape[:2]
            
            # 如果是 900x1600 (竖屏)，顺时针旋转 90° 变 1600x900
            if w == 900 and h == 1600:
                img = cv2.rotate(img, cv2.ROTATE_90_CLOCKWISE)
                
            # 如果是 1600x900，直接返回
            elif w == 1600 and h == 900:
                pass
            else:
                # 其他分辨率，返回 None 回退 ADB
                return None
                
            return img
        except Exception:
            return None

    def _get_screenshot_u2(self):
        """uiautomator2 截图（atx-agent，失败时返回 None 由上层回退 ADB）"""
        import cv2
        import numpy as np
        if not self._u2_init_screenshot() or not self._u2_device:
            return None
        try:
            # d.screenshot() 返回 PIL.Image；部分版本支持 format='opencv' 直接返回 BGR
            pil_img = self._u2_device.screenshot()
            if pil_img is None:
                return None
            arr = np.asarray(pil_img)
            if arr.ndim != 3 or arr.size == 0:
                return None
            if arr.shape[2] == 3:
                img = cv2.cvtColor(arr, cv2.COLOR_RGB2BGR)
            elif arr.shape[2] == 4:
                img = cv2.cvtColor(arr, cv2.COLOR_RGBA2BGR)
            else:
                return None
            return img
        except Exception:
            return None

    def get_screenshot(self, force_method=None):
        """force_method: 可选 'adb'，强制使用 ADB 截图（用于 nemu_ipc 检测异常时回退）"""
        import cv2
        import numpy as np
        import tempfile
        if not self.device_serial:
            return None
        use_nemu = (self.screenshot_method == "nemu_ipc" and force_method != "adb")
        use_u2 = (self.screenshot_method == "uiautomator2" and force_method != "adb")
        use_droidcast = (self.screenshot_method == "droidcast_raw" and force_method != "adb")

        if use_droidcast:
            img = self._get_screenshot_droidcast_raw()
            if img is not None:
                return img
            if not getattr(self, '_droidcast_raw_adb_fallback_logged', False):
                self._droidcast_raw_adb_fallback_logged = True
                print(">>> [DroidCast_raw] 截图失败，已回退到 ADB 截图")

        if use_u2:
            img = self._get_screenshot_u2()
            if img is not None:
                return img
            if not getattr(self, '_u2_screenshot_adb_fallback_logged', False):
                self._u2_screenshot_adb_fallback_logged = True
                print(">>> [uiautomator2] 截图失败，已回退到 ADB 截图")

        if use_nemu:
            if sys.platform != "win32":
                if self._nemu_ipc_logged != "fail":
                    self._nemu_ipc_logged = "fail"
                    print(">>> [nemu_ipc] 仅支持 Windows，已回退到 ADB 截图")
            else:
                img = self._get_screenshot_nemu_ipc()
                if img is not None:
                    return img
                if getattr(self, '_nemu_ipc_adb_fallback_logged', False) is not True:
                    self._nemu_ipc_adb_fallback_logged = True
                    print(">>> [nemu_ipc] 截图失败，已回退到 ADB 截图")

        def _try_exec_out():
            adb = CURRENT_ADB_PATH if CURRENT_ADB_PATH else "adb"
            cmd = [adb, "-s", self.device_serial, "exec-out", "screencap", "-p"]
            try:
                result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=8,
                                        creationflags=0x08000000, cwd=os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else None)
                if result.stdout and len(result.stdout) > 100:
                    image_array = np.frombuffer(result.stdout, np.uint8)
                    img = cv2.imdecode(image_array, cv2.IMREAD_COLOR)
                    if img is not None:
                        return img
            except Exception:
                pass
            return None

        def _try_pull():
            tmp_path = os.path.join(tempfile.gettempdir(), f"woa_sc_{os.getpid()}.png")
            try:
                r1 = self.run_cmd(["shell", "screencap", "-p", "/sdcard/woa_sc.png"], timeout=8)
                if r1 is None:
                    return None
                r2 = self.run_cmd(["pull", "/sdcard/woa_sc.png", tmp_path], timeout=8)
                if r2 is None:
                    return None
                if os.path.exists(tmp_path) and os.path.getsize(tmp_path) > 100:
                    img = cv2.imread(tmp_path)
                    try:
                        os.remove(tmp_path)
                    except Exception:
                        pass
                    if img is not None:
                        return img
            except Exception:
                pass
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass
            return None

        img = _try_exec_out()
        if img is not None:
            return img
        img = _try_pull()
        if img is not None:
            return img
        # 重试一次（部分模拟器偶发失败）
        time.sleep(0.2)
        img = _try_exec_out() or _try_pull()
        return img

    def get_pixel_color(self, x, y):
        screen = self.get_screenshot()
        if screen is None: return None
        try:
            b, g, r = screen[y, x]
            return int(b), int(g), int(r)
        except:
            return None

    def click(self, x, y, random_offset=5):
        self._do_think()
        if random_offset > 0:
            x += random.uniform(-random_offset, random_offset)
            y += random.uniform(-random_offset, random_offset)
        xi, yi = int(round(x)), int(round(y))

        # uiautomator2
        if self.control_method == "uiautomator2":
            if self._u2_init() and self._u2_device:
                try:
                    self._u2_device.click(xi, yi)
                    return
                except Exception:
                    if not self._u2_fallback_logged:
                        self._u2_fallback_logged = True
                        print(">>> [uiautomator2] 触控失败，回退到 ADB")
            self._adb_click_fallback(xi, yi)
            return

        if self.use_minitouch:
            if not self._minitouch_ready:
                self._minitouch_init()
            if self._minitouch_ready and self._minitouch_client:
                tx, ty = self._minitouch_scale(xi, yi)
                s = f"d 0 {tx} {ty} 50\nc\nu 0\nc\n"
                if self._minitouch_send(s):
                    return
                self._minitouch_ready = False
                if not getattr(self, '_minitouch_runtime_fallback_logged', False):
                    self._minitouch_runtime_fallback_logged = True
                    print(">>> [minitouch] 触控发送失败，已回退到 ADB")

        self._adb_click_fallback(xi, yi)

    def double_click(self, x, y, random_offset=20):
        self.click(x, y, random_offset)
        time.sleep(0.08)  # 双击间隔
        self.click(x, y, random_offset)

    def swipe(self, x1, y1, x2, y2, duration_ms=1000):
        self._do_think()
        x1i, y1i = int(round(x1)), int(round(y1))
        x2i, y2i = int(round(x2)), int(round(y2))

        if self.control_method == "uiautomator2":
            if self._u2_init() and self._u2_device:
                try:
                    self._u2_device.swipe(x1i, y1i, x2i, y2i, duration=duration_ms / 1000.0)
                    return
                except Exception:
                    pass
            self.run_cmd(["shell", "input", "swipe", str(x1i), str(y1i), str(x2i), str(y2i), str(duration_ms)])
            return

        if self.use_minitouch:
            if not self._minitouch_ready:
                self._minitouch_init()
            if self._minitouch_ready and self._minitouch_client:
                n = max(5, duration_ms // 20)
                pts = []
                for i in range(n + 1):
                    t = i / n
                    px = int(x1i + (x2i - x1i) * t)
                    py = int(y1i + (y2i - y1i) * t)
                    pts.append(self._minitouch_scale(px, py))
                sb = []
                for i, (tx, ty) in enumerate(pts):
                    if i == 0:
                        sb.append(f"d 0 {tx} {ty} 50\nc\n")
                    else:
                        sb.append(f"m 0 {tx} {ty} 50\nc\n")
                    sb.append(f"w {duration_ms // n}\n")
                sb.append("u 0\nc\nw 10\n")
                if self._minitouch_send("".join(sb)):
                    return
                self._minitouch_ready = False

        cmd = f"input swipe {x1i} {y1i} {x2i} {y2i} {duration_ms}"
        if not self._write_shell_cmd(cmd):
            self.run_cmd([
                "shell", "input", "swipe",
                str(x1i), str(y1i), str(x2i), str(y2i), str(duration_ms)
            ])

    # 图像识别相关方法保持不变，因为它们不涉及 ADB 通信
    def _read_image_safe(self, path):
        import cv2
        import numpy as np
        if not os.path.exists(path): return None
        try:
            img_array = np.fromfile(path, dtype=np.uint8)
            return cv2.imdecode(img_array, cv2.IMREAD_COLOR)
        except:
            return None

    def locate_image(self, template_path, confidence=0.8, screen_image=None):
        import cv2
        if screen_image is not None:
            screen = screen_image
        else:
            screen = self.get_screenshot()
        if screen is None: return None
        template = self._read_image_safe(template_path)
        if template is None: return None
        result = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        if max_val >= confidence:
            h, w = template.shape[:2]
            return (max_loc[0] + w // 2, max_loc[1] + h // 2, (max_loc[0], max_loc[1], w, h))
        return None

    def locate_all_images(self, template_path, confidence=0.8, screen_image=None):
        import cv2
        import numpy as np
        if screen_image is not None:
            screen = screen_image
        else:
            screen = self.get_screenshot()
        if screen is None: return []
        template = self._read_image_safe(template_path)
        if template is None: return []
        result = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
        h, w = template.shape[:2]
        loc = np.where(result >= confidence)
        found_items = []
        for pt in zip(*loc[::-1]):
            score = result[pt[1], pt[0]]
            is_duplicate = False
            for item in found_items:
                if abs(pt[0] - item['box'][0]) < 10 and abs(pt[1] - item['box'][1]) < 10:
                    if score > item['score']:
                        item['score'] = score
                        item['box'] = (pt[0], pt[1], w, h)
                        item['top'] = pt[1]
                        item['center'] = (pt[0] + w // 2, pt[1] + h // 2)
                    is_duplicate = True
                    break
            if not is_duplicate:
                found_items.append(
                    {'box': (pt[0], pt[1], w, h), 'top': pt[1], 'center': (pt[0] + w // 2, pt[1] + h // 2),
                     'score': score})
        return found_items
