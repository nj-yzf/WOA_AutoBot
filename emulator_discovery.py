# -*- coding: utf-8 -*-
"""
模拟器发现模块 - 参考 ALAS (AzurLaneAutoScript) 实现
从 vbox/nemu 配置文件、注册表等获取模拟器实例与 ADB 端口
"""

import os
import re
import sys
from typing import List, Optional, Tuple

# Windows 专用
if sys.platform == "win32":
    import winreg
    import codecs
else:
    winreg = None
    codecs = None


def _iter_folder(folder, is_dir=False, ext=None):
    """安全遍历目录"""
    try:
        files = os.listdir(folder)
    except (FileNotFoundError, PermissionError):
        return
    for f in files:
        sub = os.path.join(folder, f)
        if is_dir:
            if os.path.isdir(sub):
                yield sub.replace("\\", "/")
        elif ext:
            if os.path.isfile(sub) and sub.lower().endswith(ext):
                yield sub.replace("\\", "/")
        else:
            yield sub.replace("\\", "/")


def vbox_file_to_serial(filepath: str) -> str:
    """
    从 .nemu / .vbox 中解析 ADB 端口 (参考 ALAS vbox_file_to_serial)
    Returns: '127.0.0.1:PORT' 或 ''
    """
    # <Forwarding name="port2" proto="1" hostip="127.0.0.1" hostport="62026" guestport="5555"/>
    regex = re.compile(r'<[^>]*hostport="(\d+)"[^>]*guestport="5555"', re.I | re.S)
    try:
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
        m = regex.search(content)
        if m:
            return f"127.0.0.1:{m.group(1)}"
    except Exception:
        pass
    return ""


def _mum12_id_from_name(name: str) -> Optional[int]:
    """MuMu12 实例名 -> ID，如 MuMuPlayer-12.0-3 -> 3"""
    m = re.search(r"MuMuPlayer(?:Global)?-12\.0-(\d+)", name, re.I)
    if m:
        return int(m.group(1))
    m = re.search(r"YXArkNights-12\.0-(\d+)", name, re.I)
    if m:
        return int(m.group(1))
    return None


def get_mumu_serials_from_vms() -> List[Tuple[str, str, str]]:
    """
    从 MuMu vms 目录获取实例序列号
    Returns: [(serial, name, emu_dir), ...]
    """
    result = []
    bases = [
        r"E:\APP\MuMuPlayer",
        os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), "Netease"),
        os.path.join(os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)"), "Netease"),
    ]
    for base in bases:
        if not base or not os.path.isdir(base):
            continue
        # MuMu12: vms/MuMuPlayer-12.0-0
        vms = os.path.join(base, "vms")
        if not os.path.isdir(vms):
            continue
        try:
            for name in os.listdir(vms):
                if "MuMuPlayer" not in name and "myandrovm" not in name:
                    continue
                folder = os.path.join(vms, name)
                if not os.path.isdir(folder):
                    continue
                for fpath in _iter_folder(folder, ext=".nemu"):
                    serial = vbox_file_to_serial(fpath)
                    if serial:
                        emu_dir = os.path.dirname(os.path.dirname(os.path.dirname(fpath)))
                        result.append((serial, name, emu_dir))
                        break
                else:
                    # MuMu12 无 .nemu 端口记录时使用公式 16384 + 32 * id（参考 ALAS）
                    mid = _mum12_id_from_name(name)
                    if mid is not None:
                        port = 16384 + 32 * mid
                        emu_dir = base
                        result.append((f"127.0.0.1:{port}", name, emu_dir))
        except Exception:
            pass
    # MuMu6 单实例（端口 7555）
    seen_7555 = any(s[0] == "127.0.0.1:7555" for s in result)
    if not seen_7555:
        for base in bases:
            if not base:
                continue
            for cand in ["MuMu Player 6", "MuMu Player 12", "MuMuPlayer", "MuMuPlayer6"]:
                p = os.path.join(base, cand)
                if os.path.isdir(p):
                    result.append(("127.0.0.1:7555", "MuMu6", p))
                    break
    return result


def get_mumu_adb_paths() -> List[str]:
    """获取 MuMu 自带的 adb 路径"""
    candidates = [
        r"E:\APP\MuMuPlayer\nx_main\adb.exe",
        os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), "Netease", "MuMu Player 12", "nx_main", "adb.exe"),
        os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), "Netease", "MuMu", "nx_main", "adb.exe"),
        os.path.join("D:", "Program Files", "Netease", "MuMu", "nx_main", "adb.exe"),
        os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), "Netease", "MuMu Player 6", "MuMu", "adb.exe"),
        os.path.join(os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)"), "Netease", "MuMu Player 12", "nx_main", "adb.exe"),
    ]
    found = []
    for p in candidates:
        if p and os.path.isfile(p):
            found.append(os.path.normpath(p))
    for drive in ["C", "D", "E", "F"]:
        for base in [rf"{drive}:\Program Files\Netease", rf"{drive}:\APP\MuMuPlayer", rf"{drive}:\Program Files (x86)\Netease"]:
            if not os.path.isdir(base):
                continue
            try:
                for name in os.listdir(base):
                    if "MuMu" not in name:
                        continue
                    for sub in ["nx_main", "MuMu", "emulator\\nemu"]:
                        p = os.path.join(base, name, sub, "adb.exe")
                        if os.path.isfile(p):
                            found.append(os.path.normpath(p))
                    # MuMu9 vmonitor
                    vmon = os.path.join(base, name, "emulator", "nemu9", "vmonitor", "bin", "adb_server.exe")
                    if os.path.isfile(vmon):
                        found.append(os.path.normpath(vmon))
            except Exception:
                pass
    return list(dict.fromkeys(found))


def get_serials_from_registry() -> List[str]:
    """从注册表发现模拟器实例的 ADB 序列号 (参考 ALAS iter_uninstall_registry)"""
    if not winreg:
        return []
    serials = []
    paths = [
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        r"Software\Microsoft\Windows\CurrentVersion\Uninstall",
    ]
    names = [
        "Nox", "Nox64", "BlueStacks", "BlueStacks_nxt", "BlueStacks_cn", "BlueStacks_nxt_cn",
        "LDPlayer", "LDPlayer4", "LDPlayer9", "leidian", "leidian4", "leidian9",
        "Nemu", "Nemu9", "MuMuPlayer", "MuMuPlayer-12.0", "MuMu Player 12.0", "MEmu",
    ]
    for reg_path in paths:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as reg:
                i = 0
                while True:
                    try:
                        sub = winreg.EnumKey(reg, i)
                        i += 1
                    except OSError:
                        break
                    if sub not in names:
                        continue
                    try:
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, f"{reg_path}\\{sub}") as sk:
                            inst_dir, _ = winreg.QueryValueEx(sk, "InstallLocation")
                    except Exception:
                        continue
                    if not inst_dir or not os.path.isdir(inst_dir):
                        continue
                    inst_dir = inst_dir.rstrip("\\/")
                    # LDPlayer: 端口 5555 + instance*2，并补充常见多开端口
                    if "leidian" in sub.lower() or "ldplayer" in sub.lower():
                        vms = os.path.join(inst_dir, "vms")
                        if os.path.isdir(vms):
                            for d in _iter_folder(vms, is_dir=True):
                                bn = os.path.basename(d)
                                m = re.match(r"leidian(\d+)", bn, re.I)
                                if m:
                                    port = int(m.group(1)) * 2 + 5555
                                    serials.append(f"127.0.0.1:{port}")
                        for port in [5555, 5557, 5559, 5561, 5563, 5565]:
                            serials.append(f"127.0.0.1:{port}")
                    # BlueStacks5: 从 bluestacks.conf 读 (ALAS 格式)
                    if "bluestacks_nxt" in sub.lower():
                        try:
                            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, f"SOFTWARE\\{sub}") as sk:
                                ud, _ = winreg.QueryValueEx(sk, "UserDefinedDir")
                            conf = os.path.join(ud, "bluestacks.conf")
                            if os.path.isfile(conf):
                                with open(conf, encoding="utf-8", errors="ignore") as f:
                                    content = f.read()
                                for m in re.finditer(r'bst\.instance\.\w+\.status\.adb_port="(\d+)"', content):
                                    serials.append(f"127.0.0.1:{m.group(1)}")
                                if not re.search(r'adb_port="(\d+)"', content):
                                    serials.append("127.0.0.1:5555")
                        except Exception:
                            pass
        except Exception:
            pass
    return list(dict.fromkeys(serials))


def get_emulator_serial_pair(serial: str) -> Tuple[Optional[str], Optional[str]]:
    """
    127.0.0.1:5555 <-> emulator-5554 互换 (参考 ALAS get_serial_pair)
    Returns: (port_serial, emulator_serial)
    """
    if serial.startswith("127.0.0.1:"):
        try:
            port = int(serial[10:])
            if 5555 <= port <= 5555 + 32:
                return f"127.0.0.1:{port}", f"emulator-{port - 1}"
        except (ValueError, IndexError):
            pass
    if serial.startswith("emulator-"):
        try:
            port = int(serial[9:])
            if 5554 <= port <= 5554 + 32:
                return f"127.0.0.1:{port + 1}", f"emulator-{port}"
        except (ValueError, IndexError):
            pass
    return None, None


def discover_all_serials_and_ports() -> Tuple[List[str], List[int]]:
    """
    汇总所有发现的序列号与端口，供 scan_devices 使用
    Returns: (serials, ports)
    """
    serials = []
    ports_set = set()
    # MuMu vms（含 MuMu12 公式 16384+32*id）
    for serial, _, _ in get_mumu_serials_from_vms():
        if serial not in serials:
            serials.append(serial)
        if ":" in serial:
            try:
                ports_set.add(int(serial.split(":")[1]))
            except ValueError:
                pass
    # 注册表
    for s in get_serials_from_registry():
        if s not in serials:
            serials.append(s)
        if ":" in s:
            try:
                ports_set.add(int(s.split(":")[1]))
            except ValueError:
                pass
    # emulator-5554 等形式
    for s in list(serials):
        port_s, emu_s = get_emulator_serial_pair(s)
        if emu_s and emu_s not in serials:
            serials.append(emu_s)
    # 常用端口补充（MuMu/雷电/通用）
    for p in [
        5555, 5557, 5559, 5561, 5563, 5565, 5554, 5556, 7555,
        16384, 16385, 16416, 16448, 16480, 16512, 16544, 16576,
        62001, 62025, 59865, 21503, 6555,
    ]:
        ports_set.add(p)
    ports = list(ports_set)
    return serials, ports
