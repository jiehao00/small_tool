#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
企业微信自动发送文件工具 (GUI版)
功能：
  1. 遍历文件夹，按 "1_姓名_2020" 格式提取姓名
  2. 自动打开企业微信 → 搜索群名称 → 进群
  3. 在群成员列表中找到目标姓名 → 打开私聊 → 发送文件
  4. 发送成功后归档文件
依赖：pyautogui, pygetwindow, pyperclip, opencv-python（可选，用于图像识别定位）
"""


import os
import sys
import json
import time
import ctypes
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path

import pyautogui
import pygetwindow as gw
import pyperclip

# ========== 路径处理 ==========
if getattr(sys, 'frozen', False):
    APP_PATH = Path(sys.executable).parent
else:
    APP_PATH = Path(__file__).parent

CONFIG_FILE = APP_PATH / "config.json"
IMAGES_DIR = APP_PATH / "images"

# 检查 OpenCV 是否可用（pyautogui 图像识别依赖）
try:
    import cv2  # noqa: F401
    HAS_OPENCV = True
except ImportError:
    HAS_OPENCV = False


class UserInterruptedError(Exception):
    """用户通过移动鼠标手动中断了自动化操作"""
    pass

# ========== 默认配置 ==========
DEFAULT_CONFIG = {
    # ---- 文件夹设置 (GUI界面录入，这里为默认值) ----
    "watch_folder": r"D:\send_files",

    # ---- 企业微信 ----
    "wechat_title_keyword": "企业微信",

    # ---- 延时设置（秒）----
    "delay_click": 0.3,
    "delay_search_result": 1.2,
    "delay_chat_load": 1.0,
    "delay_member_list": 0.8,
    "delay_between_files": 2.0,

    # ---- 图像识别（推荐：不用校准坐标，自动定位按钮）----
    "image_confidence": 0.85,            # 图像匹配相似度阈值 (0~1)，越高越严格
    "member_first_y_offset": 25,         # 成员搜索框 → 第一个结果的垂直偏移(px)，图像识别时使用

    # ---- 屏幕坐标（回退方案：图像识别不可用时使用）----
    "search_box_x": 280,              # 企业微信主窗口搜索框
    "search_box_y": 50,
    "first_result_x": 280,            # 搜索结果第一项
    "first_result_y": 140,
    "member_panel_more_btn_x": 960,   # 成员面板右上角"..."按钮
    "member_panel_more_btn_y": 60,
    "member_search_entry_x": 960,     # "..."菜单中的「搜索群成员」按钮
    "member_search_entry_y": 120,
    "member_search_x": 820,           # 成员搜索框（点搜索群成员后出现）
    "member_search_y": 100,
    "member_first_x": 820,            # 成员搜索结果第一项
    "member_first_y": 145,
    "chat_input_x": 400,              # 聊天输入框区域（用于粘贴文件）
    "chat_input_y": 580,

    # ---- 校准 ----
    "calibrate_countdown": 3,        # 坐标校准倒计时秒数

    # ---- 文件 ----
    "supported_extensions": [
        ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
        ".txt", ".csv", ".zip", ".rar", ".7z",
        ".jpg", ".jpeg", ".png", ".gif", ".bmp",
    ],
}


def load_config():
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                user = json.load(f)
            cfg = DEFAULT_CONFIG.copy()
            cfg.update(user)
            return cfg
        except:
            pass
    return DEFAULT_CONFIG.copy()


def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)


# ========== 文件名解析 ==========

def extract_name(filename):
    """
    从文件名提取姓名：
      规则：按下划线拆分，取第二个字段作为姓名。
      示例：
        "序号_姓名_其他信息.pdf" → 姓名
        "test_张三.pdf"          → 张三
        "1_欧阳娜娜_2020.pdf"    → 欧阳娜娜
    """
    name_only, _ = os.path.splitext(filename)
    parts = name_only.split("_")
    if len(parts) >= 2:
        name = parts[1].strip()
        if name:
            return name
    return None


# ========== 企业微信自动化核心 ==========

# 企业微信常见安装路径
_WXWORK_PATHS = [
    os.path.join(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"), r"WXWork\WXWork.exe"),
    os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), r"WXWork\WXWork.exe"),
    os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Programs\WXWork\WXWork.exe"),
    os.path.join(os.environ.get("APPDATA", ""), r"WXWork\WXWork.exe"),
]


def _find_wxwork_exe():
    """查找企业微信可执行文件路径，找不到返回 None"""
    for p in _WXWORK_PATHS:
        if os.path.isfile(p):
            return p
    return None


def _is_wxwork_window(w):
    """
    判断窗口是否属于企业微信：
    1. ★ 优先用进程名 wxwork.exe（覆盖所有窗口类型：主窗口、分离聊天窗口）
    2. 回退：标题包含"企业微信"
    3. 窗口可见且非最小化、有实际尺寸
    4. 排除资源管理器、浏览器等误匹配窗口
    """
    try:
        title = w.title
    except Exception:
        return False

    # ★ 空标题窗口直接拒绝（辅助窗口、托盘窗口等，不是主窗口也不是聊天窗口）
    if not title or not title.strip():
        return False

    # 排除明显不相关窗口（含本工具自身的窗口）
    bad_keywords = [
        "文件资源管理器", "File Explorer",
        "Google Chrome", "Microsoft Edge", "Firefox",
        "记事本", "Notepad", "控制面板", "Control Panel",
        "文件夹", "Folder",
        "自动发送",      # ★ 排除本工具自己的窗口
        "坐标校准",      # ★ 排除本工具的辅助窗口
        "图片预览",      # ★ 排除本工具的辅助窗口
    ]
    for kw in bad_keywords:
        if kw in title:
            return False

    # 额外检查：窗口可见、有实际尺寸、属于 wxwork.exe 进程
    try:
        hwnd = int(w._hWnd)
        user32 = ctypes.windll.user32

        # IsWindowVisible → 窗口是否可见
        if not user32.IsWindowVisible(hwnd):
            return False

        # GetWindowRect → 窗口尺寸是否 > 0
        class RECT(ctypes.Structure):
            _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                        ("right", ctypes.c_long), ("bottom", ctypes.c_long)]
        rect = RECT()
        if user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            width = rect.right - rect.left
            height = rect.bottom - rect.top
            if width <= 0 or height <= 0:
                return False

        # ★ 标题含"企业微信" → 快速通道（主窗口，无需进程检测）
        if "企业微信" in title:
            return True

        # ★ 进程名校验：接受 wxwork.exe 的所有窗口（覆盖分离聊天窗口）
        pid = ctypes.c_ulong()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        proc_handle = ctypes.windll.kernel32.OpenProcess(
            0x0400 | 0x0010, False, pid.value)  # PROCESS_QUERY_INFORMATION | PROCESS_VM_READ
        if proc_handle:
            exe_buf = ctypes.create_unicode_buffer(260)
            exe_size = ctypes.c_ulong(260)
            if ctypes.windll.psapi.GetModuleBaseNameW(proc_handle, None, exe_buf, exe_size):
                process_name = exe_buf.value.lower()
                ctypes.windll.kernel32.CloseHandle(proc_handle)
                return process_name == "wxwork.exe"
            ctypes.windll.kernel32.CloseHandle(proc_handle)
        # 进程检测失败 → 拒绝（没有标题匹配也没有进程匹配，不安全）
        return False
    except Exception:
        # 窗口信息获取失败 → 拒绝
        return False


# ── 前台窗口 & 鼠标中断工具 ──────────────────────────────────────────────

def _get_foreground_info():
    """
    返回前台窗口信息: (hwnd, title, process_name)
    ★ 进程名检测比标题检测更可靠：企业微信分离聊天窗口标题不含"企业微信"，
       但进程名一定是 wxwork.exe。
    """
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return (0, '', '')
    title_buf = ctypes.create_unicode_buffer(256)
    user32.GetWindowTextW(hwnd, title_buf, 256)
    pid = ctypes.c_ulong()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    proc_name = ''
    try:
        handle = kernel32.OpenProcess(0x0400 | 0x0010, False, pid.value)
        if handle:
            exe_buf = ctypes.create_unicode_buffer(260)
            exe_size = ctypes.c_ulong(260)
            if ctypes.windll.psapi.GetModuleBaseNameW(handle, None, exe_buf, exe_size):
                proc_name = exe_buf.value.lower()
            kernel32.CloseHandle(handle)
    except Exception:
        pass
    return (hwnd, title_buf.value, proc_name)


def _is_wxwork_foreground():
    """
    判断当前前台窗口是不是企业微信（包括主窗口和分离聊天窗口）。
    ★ 优先用进程名 wxwork.exe，回退用标题"企业微信"。
    """
    _hwnd, title, proc_name = _get_foreground_info()
    if not _hwnd:
        return False
    # 进程名：最可靠，覆盖所有企业微信窗口类型
    if proc_name == 'wxwork.exe':
        return True
    # 回退：标题检查（排除本工具自身窗口）
    bad_kw = ("自动发送", "坐标校准", "图片预览")
    if "企业微信" in title and not any(kw in title for kw in bad_kw):
        return True
    return False


def _is_window_maximized(hwnd):
    """检查窗口是否处于最大化状态（通过 Win32 IsZoomed）"""
    if not hwnd:
        return False
    return bool(ctypes.windll.user32.IsZoomed(hwnd))


# ★ 全局缓存企业微信窗口 HWND，供 _reassert_wxwork 使用
_G_WXWORK_HWND = 0


def _reassert_wxwork():
    """
    在每次 pyautogui 操作后立即调用，把前台转回企业微信。
    原理：pyautogui 的 SendInput 会让本进程获得"前台权限"，
    紧接着调用 SetForegroundWindow(企业微信) 就能把权限转交，
    防止工具自身窗口或其他程序（如 IDE）抢夺焦点。
    """
    global _G_WXWORK_HWND
    if not _G_WXWORK_HWND:
        return
    user32 = ctypes.windll.user32
    if user32.GetForegroundWindow() != _G_WXWORK_HWND:
        user32.SetForegroundWindow(_G_WXWORK_HWND)
        user32.BringWindowToTop(_G_WXWORK_HWND)


def _ensure_wxwork_foreground(cfg, log_callback, max_retries=3):
    """
    确保前台窗口是企业微信。不是则尝试激活，失败则返回 False。
    ★ 先检查，已是企业微信则不动。
    ★ 激活策略改用直接窗口 API 操作，避免调用 activate_wechat
       的完整流程（其中 SendInput 会再次给工具进程前台权限，
       导致工具窗口抢回焦点 → 无限循环）。
    """
    global _G_WXWORK_HWND

    if _is_wxwork_foreground():
        # ★ 缓存前台 hwnd 到全局
        _G_WXWORK_HWND = ctypes.windll.user32.GetForegroundWindow()
        return True

    user32 = ctypes.windll.user32
    # 找企业微信窗口的 HWND（通过 gw 库）
    windows = gw.getWindowsWithTitle(cfg["wechat_title_keyword"])
    wxwork_hwnds = []
    for w in windows:
        if _is_wxwork_window(w):
            wxwork_hwnds.append(int(w._hWnd))
    if not wxwork_hwnds:
        # 回退：遍历所有窗口
        for w in gw.getAllWindows():
            if _is_wxwork_window(w):
                wxwork_hwnds.append(int(w._hWnd))
    if not wxwork_hwnds:
        log_callback(f"  ⚠ 找不到企业微信窗口")
        return False
    wxwork_hwnd = wxwork_hwnds[0]

    # ★ 缓存到全局，供 _reassert_wxwork 在每次 pyautogui 操作后使用
    _G_WXWORK_HWND = wxwork_hwnd

    for i in range(max_retries):
        fg = user32.GetForegroundWindow()
        if fg == wxwork_hwnd:
            return True
        log_callback(f"  ⚠ 前台窗口丢失({_get_foreground_info()[1] or '未知'})，恢复中...")
        # 直接 SetForegroundWindow + SwitchToThisWindow，不发送任何模拟输入
        user32.ShowWindow(wxwork_hwnd, 5)  # SW_SHOW
        user32.SetWindowPos(wxwork_hwnd, 0, 0, 0, 0, 0, 0x0003)
        # SwitchToThisWindow — 比 SetForegroundWindow 更激进
        try:
            ctypes.windll.user32.SwitchToThisWindow(wxwork_hwnd, True)
        except Exception:
            pass
        user32.SetForegroundWindow(wxwork_hwnd)
        user32.BringWindowToTop(wxwork_hwnd)
        time.sleep(0.2 + 0.15 * i)
        if user32.GetForegroundWindow() == wxwork_hwnd:
            return True
    return False


_def_tool_hwnd_cache = None  # 缓存工具窗口 HWND，避免反复查找


def _get_tool_window_hwnd():
    """获取本工具窗口的 HWND（带缓存）"""
    global _def_tool_hwnd_cache
    if _def_tool_hwnd_cache:
        return _def_tool_hwnd_cache
    user32 = ctypes.windll.user32
    # 枚举顶层窗口，匹配标题
    TOOL_TITLE = "企业微信自动发送文件工具"

    def enum_cb(hwnd, _lparam):
        buf = ctypes.create_unicode_buffer(512)
        user32.GetWindowTextW(hwnd, buf, 512)
        if TOOL_TITLE in buf.value and user32.IsWindowVisible(hwnd):
            nonlocal found
            found = hwnd
            return False  # 停止枚举
        return True

    found = 0
    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_ulong, ctypes.c_ulong)
    try:
        ctypes.windll.user32.EnumWindows(WNDENUMPROC(enum_cb), 0)
    except Exception:
        pass
    _def_tool_hwnd_cache = found or 0
    return _def_tool_hwnd_cache


def _minimize_tool_window(log_callback=None):
    """最小化工具自身窗口，防止抢企业微信焦点"""
    hwnd = _get_tool_window_hwnd()
    if hwnd:
        ctypes.windll.user32.ShowWindow(hwnd, 6)  # SW_MINIMIZE
        if log_callback:
            log_callback(f"  🔽 工具窗口已最小化")


def _restore_tool_window():
    """恢复工具窗口"""
    hwnd = _get_tool_window_hwnd()
    if hwnd:
        ctypes.windll.user32.ShowWindow(hwnd, 9)  # SW_RESTORE
        ctypes.windll.user32.SetForegroundWindow(hwnd)


def _get_mouse_pos():
    """获取当前鼠标屏幕坐标 (x, y)"""
    class POINT(ctypes.Structure):
        _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
    pt = POINT()
    ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
    return (pt.x, pt.y)


def _check_mouse_moved(before_pos, threshold=80):
    """
    检查鼠标是否被用户手动移动。
    阈值 80px 足以容忍受 pyautogui 操作带来的轻微偏移，
    但能拦截用户手动大幅移动鼠标的中断意图。
    返回: True = 被移动了(中断), False = 正常
    """
    now = _get_mouse_pos()
    dx = now[0] - before_pos[0]
    dy = now[1] - before_pos[1]
    return (dx * dx + dy * dy) > (threshold * threshold)


def _mouse_safe_click(x, y, log_msg=""):
    """
    点击后检测用户是否手动移动鼠标。
    ★ 点击操作本身会把鼠标移到 (x,y)，所以比较"点击后位置"与
       "点击目标 (x,y)"，而不是与"点击前位置"比较。
    ★ 点击后立即 _reassert_wxwork，把前台权限转回企业微信。
    """
    pyautogui.click(x, y)
    _reassert_wxwork()
    now = _get_mouse_pos()
    dx = now[0] - x
    dy = now[1] - y
    if (dx * dx + dy * dy) > (80 * 80):
        raise UserInterruptedError(f"用户移动鼠标，操作已中止: {log_msg}")


def _mouse_safe_press(key, log_msg=""):
    """键盘操作。操作后立即转回企业微信前台。"""
    pyautogui.press(key)
    _reassert_wxwork()


def _mouse_safe_hotkey(*keys, log_msg=""):
    """组合键。操作后立即转回企业微信前台。"""
    pyautogui.hotkey(*keys)
    _reassert_wxwork()


def _mouse_safe_doubleclick(x, y, log_msg=""):
    """双击（比较点击后位置与目标）。操作后立即转回企业微信前台。"""
    pyautogui.doubleClick(x, y)
    _reassert_wxwork()
    now = _get_mouse_pos()
    dx = now[0] - x
    dy = now[1] - y
    if (dx * dx + dy * dy) > (80 * 80):
        raise UserInterruptedError(f"用户移动鼠标，操作已中止: {log_msg}")


def _launch_wxwork(log_callback):
    """尝试启动企业微信，返回 True/False"""
    exe = _find_wxwork_exe()
    if not exe:
        log_callback("  ⚠ 未找到企业微信安装路径，请手动启动")
        return False
    try:
        log_callback(f"  🚀 正在启动企业微信: {exe}")
        subprocess.Popen([exe], shell=False)
        return True
    except Exception as e:
        log_callback(f"  ⚠ 启动企业微信失败: {e}")
        return False


def _send_key_for_foreground():
    """
    通过 SendInput 模拟一次无害按键(Ctrl)，授予本进程前台激活权限。
    ★ 用 Ctrl 替代 Alt：Alt 会触发 Electron 应用（企业微信）的菜单栏，
       导致焦点被菜单栏抢走，激活后窗口立刻失去前台状态。
       Ctrl 单独按下不产生任何 UI 副作用。
    """
    user32 = ctypes.windll.user32
    VK_CONTROL = 0x11
    KEYEVENTF_KEYUP = 0x0002

    buf_size = 40
    inp = ctypes.create_string_buffer(buf_size)

    # 按下 Ctrl
    ctypes.memset(inp, 0, buf_size)
    ctypes.cast(inp, ctypes.POINTER(ctypes.c_ulong))[0] = 1      # type = INPUT_KEYBOARD
    ctypes.cast(ctypes.c_void_p(
        ctypes.addressof(inp) + 8),
        ctypes.POINTER(ctypes.c_ushort))[0] = VK_CONTROL          # wVk = VK_CONTROL
    user32.SendInput(1, inp, buf_size)

    # 释放 Ctrl
    ctypes.memset(inp, 0, buf_size)
    ctypes.cast(inp, ctypes.POINTER(ctypes.c_ulong))[0] = 1
    ctypes.cast(ctypes.c_void_p(
        ctypes.addressof(inp) + 8),
        ctypes.POINTER(ctypes.c_ushort))[0] = VK_CONTROL
    ctypes.cast(ctypes.c_void_p(
        ctypes.addressof(inp) + 12),
        ctypes.POINTER(ctypes.c_ulong))[0] = KEYEVENTF_KEYUP    # dwFlags = KEYUP
    user32.SendInput(1, inp, buf_size)


def _send_mouse_click(x, y):
    """
    通过 SendInput 模拟鼠标左键点击。
    用 ctypes.Structure 正确定义 INPUT + MOUSEINPUT，避免手动计算偏移量出错。
    """
    user32 = ctypes.windll.user32

    # --- 正确定义 Windows 结构体，让 ctypes 自动处理对齐和偏移 ---
    class MOUSEINPUT(ctypes.Structure):
        _fields_ = [
            ("dx",         ctypes.c_long),
            ("dy",         ctypes.c_long),
            ("mouseData",  ctypes.c_ulong),
            ("dwFlags",    ctypes.c_ulong),
            ("time",       ctypes.c_ulong),
            ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
        ]

    class KEYBDINPUT(ctypes.Structure):
        _fields_ = [
            ("wVk",         ctypes.c_ushort),
            ("wScan",       ctypes.c_ushort),
            ("dwFlags",     ctypes.c_ulong),
            ("time",        ctypes.c_ulong),
            ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
        ]

    class HARDWAREINPUT(ctypes.Structure):
        _fields_ = [
            ("uMsg",    ctypes.c_ulong),
            ("wParamL", ctypes.c_ushort),
            ("wParamH", ctypes.c_ushort),
        ]

    class _INPUT_UNION(ctypes.Union):
        _fields_ = [
            ("mi", MOUSEINPUT),
            ("ki", KEYBDINPUT),
            ("hi", HARDWAREINPUT),
        ]

    class INPUT(ctypes.Structure):
        _fields_ = [
            ("type", ctypes.c_ulong),
            ("u",    _INPUT_UNION),
        ]

    # --- 移动鼠标 + 构造两个 INPUT ---
    user32.SetCursorPos(x, y)
    time.sleep(0.02)

    down           = INPUT()
    down.type      = 0  # INPUT_MOUSE
    down.u.mi.dwFlags = 0x0002  # MOUSEEVENTF_LEFTDOWN

    up             = INPUT()
    up.type        = 0  # INPUT_MOUSE
    up.u.mi.dwFlags = 0x0004  # MOUSEEVENTF_LEFTUP

    inputs = (INPUT * 2)(down, up)
    user32.SendInput(2, ctypes.byref(inputs), ctypes.sizeof(INPUT))


def _force_activate_window(w):
    """
    Win10/11 强制激活窗口 — 多策略组合，任一成功即返回 True。
    策略0: 最小化工具窗口（防抢焦点）
    策略1: 已前台 → 直接返回
    策略2: AllowSetForegroundWindow + SetForegroundWindow（优先尝试）
    策略3: SwitchToThisWindow（激进模式）
    策略4: SendInput 模拟输入 → 获取前台权限 → SetForegroundWindow
    策略5: AttachThreadInput 绑定输入线程 → SetForegroundWindow
    策略6: SendInput 模拟鼠标点击标题栏 → 激活
    """
    hwnd = int(w._hWnd)
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32

    # --- 快速路径：已经是前台窗口 ---
    if user32.GetForegroundWindow() == hwnd:
        return True

    # --- 预处理：恢复/显示目标窗口 ---
    if user32.IsIconic(hwnd):
        user32.ShowWindow(hwnd, 9)  # SW_RESTORE
    user32.ShowWindow(hwnd, 5)      # SW_SHOW
    user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 0x0003)  # HWND_TOP, SWP_NOMOVE|SWP_NOSIZE

    # ====== 策略1: AllowSetForegroundWindow + SetForegroundWindow ======
    # 显式告知 Windows：允许后续的 SetForegroundWindow 调用成功
    try:
        ctypes.windll.user32.AllowSetForegroundWindow(-1)  # ASFW_ANY = -1
        user32.SetForegroundWindow(hwnd)
        user32.BringWindowToTop(hwnd)
        time.sleep(0.15)
        if user32.GetForegroundWindow() == hwnd:
            return True
    except Exception:
        pass

    # ====== 策略2: SwitchToThisWindow — 绕过前台锁限制 ======
    try:
        ctypes.windll.user32.SwitchToThisWindow(hwnd, True)
        time.sleep(0.15)
        if user32.GetForegroundWindow() == hwnd:
            return True
    except Exception:
        pass

    # ====== 策略3: SendInput 模拟 Ctrl 键 + SetForegroundWindow ======
    try:
        _send_key_for_foreground()
        user32.SetForegroundWindow(hwnd)
        user32.BringWindowToTop(hwnd)
        time.sleep(0.15)
        if user32.GetForegroundWindow() == hwnd:
            return True
    except Exception:
        pass

    # ====== 策略4: AttachThreadInput + SetForegroundWindow ======
    try:
        fg_hwnd = user32.GetForegroundWindow()
        if fg_hwnd and fg_hwnd != hwnd:
            our_tid = kernel32.GetCurrentThreadId()
            fg_tid = user32.GetWindowThreadProcessId(fg_hwnd, 0)
            if fg_tid and fg_tid != our_tid:
                user32.AttachThreadInput(our_tid, fg_tid, True)
                user32.SetForegroundWindow(hwnd)
                user32.BringWindowToTop(hwnd)
                user32.AttachThreadInput(our_tid, fg_tid, False)
                time.sleep(0.1)
                if user32.GetForegroundWindow() == hwnd:
                    return True
    except Exception:
        pass

    # ====== 策略5: SendInput 鼠标点击标题栏 → 模拟真实点击 ======
    try:
        if w.left > 0 and w.top > 0:
            _send_mouse_click(w.left + w.width // 2, w.top + 15)
            time.sleep(0.1)
            user32.SetForegroundWindow(hwnd)
            time.sleep(0.1)
            if user32.GetForegroundWindow() == hwnd:
                return True
    except Exception:
        pass

    # 最终尝试
    user32.SetForegroundWindow(hwnd)
    time.sleep(0.2)
    return user32.GetForegroundWindow() == hwnd


def _find_hidden_wxwork_hwnds():
    """
    查找隐藏的企业微信窗口（进程为 wxwork.exe 但 IsWindowVisible 为 False）。
    企业微信点"关闭"会隐藏窗口到托盘，HWND 仍存在但不可见。
    ★ 按窗口尺寸过滤（>200x200），不依赖标题（隐藏后标题可能为空）。
    返回: list[int] — 隐藏窗口的 HWND 列表
    """
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    found = []

    def _get_process_name(hwnd):
        """通过 HWND 获取进程名"""
        pid = ctypes.c_ulong()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        try:
            handle = kernel32.OpenProcess(0x0400 | 0x0010, False, pid.value)
            if handle:
                buf = ctypes.create_unicode_buffer(260)
                size = ctypes.c_ulong(260)
                if ctypes.windll.psapi.GetModuleBaseNameW(handle, None, buf, size):
                    kernel32.CloseHandle(handle)
                    return buf.value.lower()
                kernel32.CloseHandle(handle)
        except Exception:
            pass
        return ''

    def _enum_cb(hwnd, _lparam):
        # 跳过可见窗口（已由正常流程处理）
        if user32.IsWindowVisible(hwnd):
            return True
        # 检查进程名
        proc = _get_process_name(hwnd)
        if proc == 'wxwork.exe':
            # 按尺寸过滤：主窗口即使隐藏也有较大尺寸，辅助窗口（托盘/tooltip）很小
            class RECT(ctypes.Structure):
                _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                            ("right", ctypes.c_long), ("bottom", ctypes.c_long)]
            rect = RECT()
            if user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                w = rect.right - rect.left
                h = rect.bottom - rect.top
                if w > 200 and h > 200:  # 主窗口尺寸阈值
                    found.append(hwnd)
        return True

    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_ulong, ctypes.c_ulong)
    try:
        user32.EnumWindows(WNDENUMPROC(_enum_cb), 0)
    except Exception:
        pass
    return found


def activate_wechat(title_keyword, log_callback=None, timeout=15):
    """
    激活企业微信窗口，找不到时自动尝试启动
    用严格匹配防止误激活其他窗口（如资源管理器、浏览器标签等）
    """
    if log_callback is None:
        log_callback = print

    deadline = time.time() + timeout
    all_tried_hwnds = set()  # 避免对同一个 hwnd 重复激活

    # 第一次尝试：找已有窗口
    for attempt in range(6):  # 前6秒找已有窗口
        if time.time() > deadline:
            break
        # 先用精准标题查找
        windows = gw.getWindowsWithTitle(title_keyword)
        for w in windows:
            if _is_wxwork_window(w):
                hwnd = int(w._hWnd)
                if hwnd in all_tried_hwnds:
                    continue
                all_tried_hwnds.add(hwnd)
                try:
                    if _force_activate_window(w):
                        log_callback(f"  ✅ 已激活企业微信窗口: {w.title}")
                        return True
                    else:
                        log_callback(f"  ⚠ 窗口找到但激活未生效(系统限制): {w.title}")
                except Exception as e:
                    log_callback(f"  ⚠ 激活异常(title匹配): {e}")
        # 再遍历所有窗口
        for w in gw.getAllWindows():
            if _is_wxwork_window(w):
                hwnd = int(w._hWnd)
                if hwnd in all_tried_hwnds:
                    continue
                all_tried_hwnds.add(hwnd)
                try:
                    if _force_activate_window(w):
                        log_callback(f"  ✅ 已激活企业微信窗口: {w.title}")
                        return True
                except Exception as e:
                    continue
        time.sleep(1)

    # ★ 兜底：正常搜索没找到可见窗口 → 可能隐藏到托盘了
    hidden_hwnds = _find_hidden_wxwork_hwnds()
    if hidden_hwnds:
        log_callback(f"  🔍 发现 {len(hidden_hwnds)} 个隐藏的企业微信窗口(托盘模式)，尝试恢复...")
        user32 = ctypes.windll.user32
        for hwnd in hidden_hwnds:
            try:
                # 先恢复显示隐藏的窗口（用 SW_SHOW 保持原最大化状态，不用 SW_RESTORE）
                user32.ShowWindow(hwnd, 5)   # SW_SHOW
                time.sleep(0.5)
                # 恢复后重新通过正常流程查找并激活（此时窗口已可见）
                for w in gw.getAllWindows():
                    if int(w._hWnd) == hwnd and _is_wxwork_window(w):
                        if _force_activate_window(w):
                            log_callback(f"  ✅ 已从托盘恢复企业微信窗口: {w.title}")
                            return True
                # 如果 getAllWindows 没匹配到，直接用 HWND 激活
                title_buf = ctypes.create_unicode_buffer(256)
                user32.GetWindowTextW(hwnd, title_buf, 256)
                user32.SetForegroundWindow(hwnd)
                time.sleep(0.2)
                if user32.GetForegroundWindow() == hwnd:
                    log_callback(f"  ✅ 已从托盘恢复企业微信窗口: {title_buf.value}")
                    return True
            except Exception as e:
                log_callback(f"  ⚠ 恢复隐藏窗口失败: {e}")
                continue

    # 没找到 → 尝试自动启动
    log_callback("  ⚠ 未找到企业微信窗口，尝试自动启动...")
    if _launch_wxwork(log_callback):
        # 启动后等待窗口出现（额外等最多 15 秒）
        launch_deadline = time.time() + 15
        while time.time() < launch_deadline:
            windows = gw.getWindowsWithTitle(title_keyword)
            for w in windows:
                if _is_wxwork_window(w):
                    try:
                        if _force_activate_window(w):
                            time.sleep(0.2)
                            log_callback("  ✅ 企业微信已启动")
                            return True
                    except Exception:
                        pass
            time.sleep(1)

    log_callback("  ❌ 无法启动或找到企业微信窗口")
    return False


def click_and_type(cfg, x, y, text, clear_first=True):
    """点击坐标 → 清空 → 粘贴文本（带鼠标中断检测）"""
    _mouse_safe_click(x, y, log_msg=f"click_and_type click({x},{y})")
    time.sleep(cfg["delay_click"])
    if clear_first:
        _mouse_safe_hotkey("ctrl", "a", log_msg="click_and_type Ctrl+A")
        time.sleep(0.1)
        _mouse_safe_press("delete", log_msg="click_and_type Delete")
        time.sleep(0.1)
    pyperclip.copy(text)
    _mouse_safe_hotkey("ctrl", "v", log_msg="click_and_type Ctrl+V")
    time.sleep(0.2)


def search_and_enter_group(cfg, group_name, log_callback=None):
    """在企业微信中搜索群名并进入（纯坐标定位 + 鼠标中断检测）"""
    if log_callback is None:
        log_callback = print
    log_callback(f"  📌 搜索群聊 → 纯坐标: 搜索框({cfg['search_box_x']},{cfg['search_box_y']}) 结果项({cfg['first_result_x']},{cfg['first_result_y']})")
    # 点击搜索框 → 输入群名
    click_and_type(cfg, cfg["search_box_x"], cfg["search_box_y"], group_name)
    time.sleep(cfg["delay_search_result"])

    # 点击第一个搜索结果（应该是群聊）
    _mouse_safe_click(cfg["first_result_x"], cfg["first_result_y"], log_msg="进入群聊")
    time.sleep(cfg["delay_chat_load"])
    return True


def _locate_center(image_name, confidence=0.85, log_callback=None):
    """用图像识别在屏幕上找图片，返回中心坐标；找不到返回 None"""
    if not HAS_OPENCV:
        if log_callback:
            log_callback(f"  ⚠ 图像识别不可用：缺少 OpenCV (cv2) 模块")
        return None
    img_path = IMAGES_DIR / image_name
    if not img_path.exists():
        if log_callback:
            log_callback(f"  ⚠ 图像识别失败：图片文件不存在 → {img_path}")
        return None
    try:
        loc = pyautogui.locateOnScreen(str(img_path), confidence=confidence, grayscale=True)
        if loc is not None:
            return pyautogui.center(loc)
        else:
            if log_callback:
                log_callback(f"  ⚠ 图像识别未找到：{image_name} (confidence={confidence})")
                log_callback(f"     → 可能原因：屏幕分辨率/DPI不同，需在该电脑上重新截图替换 images/ 中的图片")
    except Exception as e:
        if log_callback:
            log_callback(f"  ⚠ 图像识别异常：{image_name} → {e}")
    return None


def find_member_in_group(cfg, person_name, log_callback=None):
    """
    在群成员面板中搜索目标成员 → 点击打开私聊。

    流程（优先图像识别，回退坐标）：
      ① 点击成员面板右上角 "..." 按钮    (图像识别 > 坐标)
      ② 点击弹出菜单中的「搜索群成员」   (图像识别 > 坐标)
      ③ 在搜索框输入姓名                (图像识别 > 坐标)
      ④ 双击第一个匹配结果              (图像偏移计算 > 坐标)

    图像识别图片位置: images/ 目录
    所有坐标均可通过 GUI 的「坐标校准」按钮设置（作为回退方案）。
    """
    if log_callback is None:
        log_callback = print

    confidence = cfg.get("image_confidence", 0.85)

    # === 步骤1: 点击成员面板右上角 "..." 按钮 ===
    pos = _locate_center("more_btn.png", confidence, log_callback)
    if pos:
        log_callback(f"  ① 图像识别 → more_btn.png → ({pos.x}, {pos.y})")
        _mouse_safe_click(pos.x, pos.y, log_msg="成员面板...按钮(图像)")
    else:
        log_callback(f"  ① 坐标回退 → ({cfg['member_panel_more_btn_x']}, {cfg['member_panel_more_btn_y']})")
        pyautogui.moveTo(cfg["member_first_x"], cfg["member_first_y"])
        time.sleep(0.3)
        _mouse_safe_click(cfg["member_panel_more_btn_x"], cfg["member_panel_more_btn_y"], log_msg="成员面板...按钮(坐标)")
    time.sleep(cfg["delay_member_list"])

    # === 步骤2: 点击「搜索群成员」 ===
    pos = _locate_center("search_member_entry.png", confidence, log_callback)
    if pos:
        log_callback(f"  ② 图像识别 → search_member_entry.png → ({pos.x}, {pos.y})")
        _mouse_safe_click(pos.x, pos.y, log_msg="搜索群成员(图像)")
    else:
        log_callback(f"  ② 坐标回退 → ({cfg['member_search_entry_x']}, {cfg['member_search_entry_y']})")
        _mouse_safe_click(cfg["member_search_entry_x"], cfg["member_search_entry_y"], log_msg="搜索群成员(坐标)")
    time.sleep(cfg["delay_click"])

    # === 步骤3: 点击成员搜索框 → 输入姓名 ===
    #  优先图像识别定位搜索框（面板大小变化时位置稳定）
    #  记录是否命中图像，步骤4需要据此计算第一个结果位置
    search_box_pos = _locate_center("member_search_box.png", confidence, log_callback)
    if search_box_pos:
        log_callback(f"  ③ 图像识别 → member_search_box.png → ({search_box_pos.x}, {search_box_pos.y})")
        _mouse_safe_click(search_box_pos.x, search_box_pos.y, log_msg="成员搜索框(图像)")
        time.sleep(cfg["delay_click"])
        _mouse_safe_hotkey("ctrl", "a", log_msg="成员搜索框 Ctrl+A")
        time.sleep(0.1)
        _mouse_safe_press("delete", log_msg="成员搜索框 Delete")
        time.sleep(0.1)
        pyperclip.copy(person_name)
        _mouse_safe_hotkey("ctrl", "v", log_msg="成员搜索框 Ctrl+V")
        time.sleep(0.2)
    else:
        log_callback(f"  ③ 坐标回退 → ({cfg['member_search_x']}, {cfg['member_search_y']})")
        click_and_type(cfg, cfg["member_search_x"], cfg["member_search_y"], person_name)
    time.sleep(cfg["delay_search_result"])

    # === 步骤4: 双击第一个匹配成员 ===
    #  如果步骤3用了图像定位，用搜索框位置+偏移计算第一个结果位置
    if search_box_pos:
        offset_y = cfg.get("member_first_y_offset", 55)
        log_callback(f"  ④ 图像偏移计算 → 搜索框({search_box_pos.x}, {search_box_pos.y}) + 偏移({offset_y}px)")
        _mouse_safe_doubleclick(search_box_pos.x, search_box_pos.y + offset_y, log_msg="双击成员(图像)")
    else:
        log_callback(f"  ④ 坐标回退 → ({cfg['member_first_x']}, {cfg['member_first_y']})")
        _mouse_safe_doubleclick(cfg["member_first_x"], cfg["member_first_y"], log_msg="双击成员(坐标)")
    time.sleep(cfg["delay_chat_load"])
    return True


def _copy_file_via_powershell(file_path):
    """策略2: PowerShell 复制文件到剪贴板（含路径转义 / 备用方案）"""
    # 处理路径中的特殊字符：内部双引号需要转义
    escaped = file_path.replace('"', '`"')
    ps_cmd = f'Get-Item -LiteralPath "{escaped}" | Set-Clipboard'
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_cmd],
            capture_output=True, text=True, timeout=10,
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
        )
        return result.returncode == 0
    except Exception:
        return False


def _copy_file_via_cfhdrop(file_path):
    """
    策略1（主方案）: 直接通过 Win32 API 写入 CF_HDROP 格式到剪贴板。
    零外部依赖，不依赖 PowerShell，兼容企业域环境组策略限制。
    CF_HDROP 是资源管理器复制文件的标准剪贴板格式，
    企业微信的 Ctrl+V 粘贴附件能力直接依赖此格式。
    """
    from ctypes import wintypes

    # ---- 常量 ----
    CF_HDROP = 15
    GMEM_MOVEABLE = 0x0002
    GMEM_ZEROINIT = 0x0040

    # ---- DROPFILES 结构（定义文件拖放元数据）----
    class DROPFILES(ctypes.Structure):
        _fields_ = [
            ("pFiles", wintypes.DWORD),   # 文件路径列表偏移（字节），通常 = sizeof(DROPFILES)
            ("pt",     wintypes.POINT),   # 拖放释放点（0,0 即可）
            ("fNC",    wintypes.BOOL),    # 是否使用非客户区坐标
            ("fWide",  wintypes.BOOL),    # TRUE = Unicode 路径（UTF-16LE）
        ]

    # ---- 准备数据 ----
    # fWide=True 表示 UTF-16LE 编码，CF_HDROP 标准要求 double-null 结尾
    file_wide = (file_path + "\0\0").encode("utf-16-le")  # 末尾 \0\0 → 编码后 4 个零字节
    dropfiles = DROPFILES()
    dropfiles.pFiles = ctypes.sizeof(DROPFILES)  # 文件路径紧跟在结构体之后
    dropfiles.fWide = 1

    header = bytes(dropfiles)
    total_size = len(header) + len(file_wide)

    # ---- 分配全局内存 ----
    # kernel32 是 ctypes.windll.kernel32 的别名
    k32 = ctypes.windll.kernel32
    u32 = ctypes.windll.user32

    hGlobal = k32.GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, total_size)
    if not hGlobal:
        return False

    try:
        ptr = k32.GlobalLock(hGlobal)
        if not ptr:
            return False
        try:
            # 写入 DROPFILES 头 + 文件路径列表
            buf = (ctypes.c_char * total_size).from_address(ptr)
            buf[:len(header)] = header
            buf[len(header):] = file_wide
        finally:
            k32.GlobalUnlock(hGlobal)

        # ---- 写入剪贴板 ----
        if not u32.OpenClipboard(0):
            return False
        try:
            u32.EmptyClipboard()
            u32.SetClipboardData(CF_HDROP, hGlobal)
            # 所有权已转移给剪贴板，不再由我们释放
            hGlobal = 0
        finally:
            u32.CloseClipboard()
    finally:
        if hGlobal:
            k32.GlobalFree(hGlobal)

    return True


def copy_file_to_clipboard(file_path):
    """
    将文件复制到剪贴板（模拟资源管理器中 Ctrl+C 文件的效果）。
    双策略回退：
      1. ctypes CF_HDROP 原生写入（主方案，零外部依赖）
      2. PowerShell Set-Clipboard（备用方案）
    """
    file_path = os.path.abspath(file_path)

    # 策略1（主方案）: ctypes 原生 CF_HDROP
    if _copy_file_via_cfhdrop(file_path):
        return True

    # 策略2（备用）: PowerShell
    return _copy_file_via_powershell(file_path)


def send_file_to_chat(cfg, file_path, message_text=None):
    """在已打开的私聊窗口中，粘贴文件 → 粘贴附加文字 → 发送（带鼠标中断检测）"""
    file_path = os.path.abspath(file_path)
    if not os.path.exists(file_path):
        return False

    # 1. 将文件复制到剪贴板
    if not copy_file_to_clipboard(file_path):
        return False
    time.sleep(0.4)

    # 2. 点击聊天输入区域，确保焦点在输入框
    _mouse_safe_click(cfg["chat_input_x"], cfg["chat_input_y"], log_msg="点击聊天输入框")
    time.sleep(0.3)

    # 3. Ctrl+V 粘贴文件到输入框（作为附件）
    _mouse_safe_hotkey("ctrl", "v", log_msg="粘贴文件 Ctrl+V")
    time.sleep(0.6)

    # 4. 如果有附加文字，复制 → 粘贴
    if message_text:
        pyperclip.copy(message_text)
        time.sleep(0.2)
        _mouse_safe_hotkey("ctrl", "v", log_msg="粘贴附加文字 Ctrl+V")
        time.sleep(0.4)

    # 5. 回车发送（文件和文字一起发送）
    _mouse_safe_press("enter", log_msg="回车发送")
    time.sleep(cfg["delay_between_files"])
    return True


def process_one_file(cfg, group_name, file_path, log_callback, message_text=None):
    """
    处理单个文件的完整流程（带前台校验 + 鼠标中断检测）：
      提取姓名 → 激活/确认企业微信 → 进群 → 找成员 → 发文件(+附加文字)
    ★ 每一步操作前都会校验前台窗口是否仍为企业微信，丢失时自动恢复。
    ★ 每一步操作后检测用户是否移动了鼠标，移动超过 80px 即视为中断。
    """
    filename = os.path.basename(file_path)
    log_callback(f"\n{'─'*45}\n📄 处理: {filename}")

    # 1. 提取姓名
    name = extract_name(filename)
    if not name:
        log_callback(f"  ⚠ 无法提取姓名，跳过")
        return "skip"
    log_callback(f"  👤 提取姓名: {name}")

    try:
        # 2. 激活/确认企业微信（★ 进程名 + 标题双重校验，覆盖分离聊天窗口）
        _, title, proc_name = _get_foreground_info()
        if _is_wxwork_foreground():
            log_callback(f"  ✅ 当前已在前台: {title} (进程: {proc_name})")
        else:
            log_callback(f"  🔍 激活企业微信... 当前前台: {title or '无窗口'}")
            if not activate_wechat(cfg["wechat_title_keyword"], log_callback=log_callback):
                log_callback(f"  ❌ 无法激活企业微信窗口")
                return "failed"
            # 二次确认
            if not _ensure_wxwork_foreground(cfg, log_callback, max_retries=2):
                _hw, _title, _pn = _get_foreground_info()
                log_callback(f"  ❌ 前台窗口不是企业微信: {_title or '无窗口'}，已中止")
                return "failed"

        # ── 3. 搜索并进入群聊 ──（★ 移除了"回到消息列表"Escape 步骤）
        # ★ 校验企业微信窗口是否最大化，未最大化则中断
        if not _is_window_maximized(_G_WXWORK_HWND):
            log_callback(f"  ❌ 企业微信窗口未最大化，请最大化后重试")
            return "failed"

        log_callback(f"  💬 搜索群: {group_name}")
        if not _ensure_wxwork_foreground(cfg, log_callback):
            return "failed"
        search_and_enter_group(cfg, group_name, log_callback=log_callback)

        # ── 5. 在群成员中查找并打开私聊 ──
        log_callback(f"  🔎 在群成员中搜索: {name}")
        if not _ensure_wxwork_foreground(cfg, log_callback):
            return "failed"
        find_member_in_group(cfg, name, log_callback=log_callback)

        # ── 6. 发送文件 + 附加文字 ──
        if message_text:
            log_callback(f"  📤 发送文件 + 附加文字")
        else:
            log_callback(f"  📤 发送文件")
        if not _ensure_wxwork_foreground(cfg, log_callback):
            return "failed"
        ok = send_file_to_chat(cfg, file_path, message_text=message_text)
        if not ok:
            log_callback(f"  ❌ 发送失败")
            return "failed"

        log_callback(f"  ✅ 发送成功")

        return "success"

    except UserInterruptedError as e:
        log_callback(f"  🛑 用户中断: {e}")
        return "interrupted"
    except Exception as e:
        log_callback(f"  ❌ 处理异常: {e}")
        return "failed"


# ========== GUI 界面 ==========

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("企业微信自动发送文件工具")
        self.root.geometry("860x620")
        self.root.resizable(True, True)
        self.root.minsize(760, 520)

        # 加载配置
        self.cfg = load_config()

        # 状态变量
        self.running = False
        self.worker_thread = None
        self.stop_flag = threading.Event()

        # 设置样式
        self._setup_style()
        # 构建界面
        self._build_ui()
        # 加载上次使用的值
        self._restore_values()

        # 窗口协议
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # 界面配色（参考财务工具风格：深蓝侧边栏 + 蓝色主按钮）
    SIDEBAR_BG = "#2c3e50"
    SIDEBAR_BTN_BG = "#34495e"
    SIDEBAR_BTN_ACTIVE = "#2b9bd4"
    SIDEBAR_TEXT = "#ecf0f1"
    SIDEBAR_TEXT_ACTIVE = "#ffffff"
    HEADER_BG = "#2c3e50"
    ACTION_BLUE = "#2b9bd4"
    ACTION_BLUE_ACTIVE = "#1d8bc3"
    CONTENT_BG = "#f0f0f0"
    CARD_BG = "#ffffff"

    def _setup_style(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TLabel", font=("微软雅黑", 10), background=self.CONTENT_BG)
        style.configure("TButton", font=("微软雅黑", 10), padding=4)
        style.configure("TEntry", font=("微软雅黑", 10))
        style.configure("TLabelframe", font=("微软雅黑", 10, "bold"), background=self.CONTENT_BG)
        style.configure("TLabelframe.Label", background=self.CONTENT_BG)
        style.configure("Send.TButton", font=("微软雅黑", 11, "bold"),
                        background=self.ACTION_BLUE, foreground="white")
        style.map("Send.TButton",
                  background=[("active", self.ACTION_BLUE_ACTIVE), ("disabled", "#bdc3c7")],
                  foreground=[("disabled", "#7f8c8d")])

    # ==================== 页面切换 ====================

    def _switch_page(self, page_name):
        """切换右侧内容区域显示的页面"""
        for page in (self.page_func, self.page_help, self.page_about):
            page.pack_forget()
        if page_name == "func":
            self.page_func.pack(fill="both", expand=True)
            self._set_active_sidebar("nav_func")
        elif page_name == "help":
            self.page_help.pack(fill="both", expand=True)
            self._set_active_sidebar("nav_help")
        elif page_name == "about":
            self.page_about.pack(fill="both", expand=True)
            self._set_active_sidebar("nav_about")

    def _set_active_sidebar(self, active_key):
        """设置侧边栏按钮的激活样式"""
        for key, btn in self.sidebar_btns.items():
            if key == active_key:
                btn.configure(bg=self.SIDEBAR_BTN_ACTIVE, fg=self.SIDEBAR_TEXT_ACTIVE)
            else:
                btn.configure(bg=self.SIDEBAR_BTN_BG, fg=self.SIDEBAR_TEXT)

    # ==================== 构建主界面 ====================

    def _build_ui(self):
        # ===== 顶部标题栏 =====
        header = tk.Frame(self.root, bg=self.HEADER_BG, height=56)
        header.pack(side="top", fill="x")
        header.pack_propagate(False)

        tk.Label(header, text="📨  企业微信自动发送文件工具",
                 bg=self.HEADER_BG, fg="white",
                 font=("微软雅黑", 14, "bold")).pack(side="left", padx=(16, 10), pady=10)

        # ===== 主体容器 =====
        outer = tk.Frame(self.root, bg=self.CONTENT_BG)
        outer.pack(side="top", fill="both", expand=True)

        # ===== 左侧导航栏 =====
        sidebar = tk.Frame(outer, bg=self.SIDEBAR_BG, width=150)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        tk.Frame(sidebar, bg="#3d566e", height=1).pack(fill="x", padx=15, pady=(10, 8))

        self.sidebar_btns = {}
        nav_items = [
            ("nav_func", "  ⚙️  功能"),
            ("nav_help", "  📖  使用说明"),
            ("nav_about", "  ℹ️  关于"),
        ]

        for key, text in nav_items:
            btn = tk.Button(sidebar, text=text,
                            bg=self.SIDEBAR_BTN_BG, fg=self.SIDEBAR_TEXT,
                            font=("微软雅黑", 10), bd=0, relief="flat", cursor="hand2",
                            anchor="w", padx=18, pady=11,
                            activebackground=self.SIDEBAR_BTN_ACTIVE,
                            activeforeground=self.SIDEBAR_TEXT_ACTIVE)
            btn.pack(fill="x", padx=10, pady=3)
            self.sidebar_btns[key] = btn

        self.sidebar_btns["nav_func"].configure(command=lambda: self._switch_page("func"))
        self.sidebar_btns["nav_help"].configure(command=lambda: self._switch_page("help"))
        self.sidebar_btns["nav_about"].configure(command=lambda: self._switch_page("about"))

        tk.Label(sidebar, text="", bg=self.SIDEBAR_BG, font=("", 1)).pack(fill="both", expand=True)

        # ===== 右侧内容区域 =====
        self.content_area = tk.Frame(outer, bg=self.CONTENT_BG, padx=2, pady=2)
        self.content_area.pack(side="left", fill="both", expand=True)

        self.page_func = tk.Frame(self.content_area, bg=self.CONTENT_BG)
        self.page_help = tk.Frame(self.content_area, bg=self.CONTENT_BG)
        self.page_about = tk.Frame(self.content_area, bg=self.CONTENT_BG)

        self._build_func_page()
        self._build_help_page()
        self._build_about_page()

        self.page_func.pack(fill="both", expand=True)
        self._set_active_sidebar("nav_func")

    # ==================== 功能页面 ====================

    def _build_func_page(self):
        page = self.page_func
        pad = {"padx": 14, "pady": 8}

        # ===== 基本设置 =====
        settings_frame = ttk.LabelFrame(page, text="基本设置", padding=12)
        settings_frame.pack(fill="x", **pad)

        # 文件夹路径
        ttk.Label(settings_frame, text="文件夹路径").grid(
            row=0, column=0, sticky="w", pady=(2, 4))
        self.folder_var = tk.StringVar()
        self.folder_entry = ttk.Entry(settings_frame, textvariable=self.folder_var)
        self.folder_entry.grid(row=1, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(settings_frame, text="浏览...", command=self._browse_folder).grid(
            row=1, column=1, sticky="e")

        # 群名称
        ttk.Label(settings_frame, text="群名称").grid(
            row=2, column=0, sticky="w", pady=(10, 4))
        self.group_var = tk.StringVar()
        self.group_entry = ttk.Entry(settings_frame, textvariable=self.group_var)
        self.group_entry.grid(row=3, column=0, sticky="ew", padx=(0, 6))
        tk.Frame(settings_frame, bg=self.CONTENT_BG).grid(row=3, column=1)

        # 附加文字
        ttk.Label(settings_frame, text="附加文字（选填）").grid(
            row=4, column=0, sticky="w", pady=(10, 4))
        self.message_var = tk.StringVar()
        self.message_entry = ttk.Entry(settings_frame, textvariable=self.message_var)
        self.message_entry.grid(row=5, column=0, columnspan=2, sticky="ew")

        # 校准倒计时
        ttk.Label(settings_frame, text="校准倒计时（秒）").grid(
            row=6, column=0, sticky="w", pady=(10, 4))
        countdown_sub = tk.Frame(settings_frame)
        countdown_sub.grid(row=7, column=0, columnspan=2, sticky="w")
        self.countdown_var = tk.StringVar(value="3")
        self.countdown_entry = ttk.Entry(countdown_sub, textvariable=self.countdown_var,
                                         width=6)
        self.countdown_entry.pack(side="left")
        ttk.Label(countdown_sub,
                  text="  点击校准按钮后等待的秒数，建议3-5秒",
                  foreground="#7f8c8d", font=("微软雅黑", 9)).pack(side="left")

        settings_frame.columnconfigure(0, weight=1)

        # ===== 操作按钮区 =====
        action_frame = tk.Frame(page, bg=self.CONTENT_BG)
        action_frame.pack(fill="x", padx=14, pady=(6, 4))

        self.send_btn = tk.Button(action_frame, text="▶  开始发送",
                                  bg=self.ACTION_BLUE, fg="white",
                                  font=("微软雅黑", 12, "bold"),
                                  bd=0, relief="flat", cursor="hand2",
                                  activebackground=self.ACTION_BLUE_ACTIVE,
                                  activeforeground="white",
                                  command=self._start_send)
        self.send_btn.pack(side="left", padx=(0, 12), ipadx=30, ipady=8)

        self.stop_btn = tk.Button(action_frame, text="⏹ 停止",
                                  bg="#e74c3c", fg="white",
                                  font=("微软雅黑", 10), bd=0, relief="flat",
                                  cursor="hand2", state="disabled",
                                  activebackground="#c0392b",
                                  activeforeground="white",
                                  command=self._stop_send)
        self.stop_btn.pack(side="left", padx=(0, 8), ipadx=14, ipady=5)

        self.calibrate_btn = tk.Button(action_frame, text="🎯 坐标校准",
                                       bg="#3498db", fg="white",
                                       font=("微软雅黑", 10), bd=0, relief="flat",
                                       cursor="hand2",
                                       activebackground="#2980b9",
                                       activeforeground="white",
                                       command=self._open_calibrate)
        self.calibrate_btn.pack(side="left", ipadx=14, ipady=5)

        self.progress_var = tk.StringVar(value="就绪")
        ttk.Label(action_frame, textvariable=self.progress_var,
                  foreground="#7f8c8d").pack(side="right", padx=(10, 0))

        # ===== 运行日志 =====
        log_frame = ttk.LabelFrame(page, text="运行日志", padding=6)
        log_frame.pack(fill="both", expand=True, padx=14, pady=(8, 10))

        self.log_area = scrolledtext.ScrolledText(
            log_frame, font=("Consolas", 10), wrap=tk.WORD,
            bg=self.CARD_BG, fg="#2c3e50", insertbackground="#2c3e50",
            relief="solid", borderwidth=1
        )
        self.log_area.pack(fill="both", expand=True)
        self.log_area.configure(state="disabled")

    # ==================== 使用说明页面 ====================

    def _build_help_page(self):
        page = self.page_help
        pad = {"padx": 16, "pady": 6}

        tk.Label(page, text="📖 使用说明",
                 bg=self.CONTENT_BG, fg="#2c3e50",
                 font=("微软雅黑", 16, "bold")).pack(anchor="w", **pad)

        help_text_frame = ttk.LabelFrame(page, text="操作步骤", padding=8)
        help_text_frame.pack(fill="both", expand=True, padx=16, pady=4)

        help_text = tk.Text(
            help_text_frame, font=("微软雅黑", 10), wrap=tk.WORD, bd=1,
            relief="solid", bg=self.CARD_BG, fg="#2c3e50")
        help_text.pack(fill="both", expand=True)

        content = """【第一步：准备工作】
  1. 将需要发送的文件放入一个文件夹中
  2. 文件命名格式：按下划线拆分，第二个字段为姓名
     格式：任意_姓名_其他信息.扩展名
     示例：1_张三_01.txt、test_李四.pdf、序号_欧阳娜娜_2020.docx
  3. 程序会自动取文件名第二个下划线字段作为收件人姓名

【第二步：基本设置】
  1. 点击「浏览...」选择文件所在的文件夹
  2. 在「群名称」输入框中填写目标群的名称
  3. 如需附加文字，在「附加文字」框中输入（选填）

【第三步：坐标校准（首次使用必做）】
  1. 将企业微信窗口最大化，确保界面元素完整显示
  2. 点击「🎯 坐标校准」按钮
  3. 按照提示，依次点击各按钮并移动鼠标
  4. 倒计时后自动记录坐标（秒数可在基本设置中调整）
  5. 所有位置记录完成后点击「保存到配置」

   提示：程序会自动使用 images/ 目录下的截图做图像识别定位。
   需要 3 张截图：more_btn.png / search_member_entry.png / member_search_box.png
   只要有这 3 张图，③④⑤⑥ 步骤就无需手动校准。详见 images/截图说明.txt

【第四步：发送文件】
  1. 将企业微信窗口最大化，确认已登录
  2. 点击「▶ 开始发送」
  3. 确认弹窗中的文件清单
  4. 程序自动执行发送流程
  5. 期间请勿移动鼠标

【注意事项】
  • 文件夹不支持选择桌面
  • 坐标校准和发送文件时，企业微信需保持最大化
  • 企业微信窗口被关到托盘时会自动恢复
  • 操作期间请勿移动鼠标
  • 支持的文件类型：.pdf/.doc/.docx/.xls/.xlsx
    .ppt/.pptx/.txt/.csv/.png/.jpg/.jpeg"""
        help_text.insert("1.0", content)
        help_text.configure(state="disabled")

        tip_frame = tk.Frame(page, bg=self.CONTENT_BG)
        tip_frame.pack(fill="x", padx=16, pady=4)
        tk.Label(tip_frame,
                 text="💡 提示：如遇到问题，请检查企业微信是否已登录、坐标校准是否正确。",
                 bg=self.CONTENT_BG, fg="#7f8c8d", font=("微软雅黑", 9)).pack(anchor="w")

    # ==================== 关于页面 ====================

    def _build_about_page(self):
        page = self.page_about
        pad = {"padx": 16, "pady": 8}

        tk.Label(page, text="ℹ️ 关于",
                 bg=self.CONTENT_BG, fg="#2c3e50",
                 font=("微软雅黑", 16, "bold")).pack(anchor="w", **pad)

        info_frame = tk.Frame(page, bg=self.CARD_BG, bd=1, relief="solid")
        info_frame.pack(fill="x", padx=16, pady=(4, 8))

        about_info = [
            ("软件名称", "企业微信自动发送文件工具"),
            ("版本号",   "v3.0"),
            ("发布日期", "2026-07-14"),
            ("运行环境", "Windows 10/11 + 企业微信客户端"),
            ("开发语言", "Python 3.8 + Tkinter"),
            ("核心依赖", "pyautogui, pygetwindow, pyperclip, opencv-python"),
        ]

        for i, (label, value) in enumerate(about_info):
            bg_color = self.CARD_BG
            tk.Label(info_frame, text=label + "：", bg=bg_color, fg="#7f8c8d",
                     font=("微软雅黑", 10, "bold"), anchor="e").grid(
                row=i, column=0, sticky="e", padx=(20, 2), pady=3)
            tk.Label(info_frame, text=value, bg=bg_color, fg="#2c3e50",
                     font=("微软雅黑", 10), anchor="w").grid(
                row=i, column=1, sticky="w", padx=(2, 20), pady=3)

        info_frame.columnconfigure(1, weight=1)

        # 使用提示
        stats_frame = ttk.LabelFrame(page, text="使用提示", padding=8)
        stats_frame.pack(fill="x", padx=16, pady=4)

        tips = [
            "• 首次使用请先在「坐标校准」中设置各关键位置的坐标",
            "• 程序已内置按钮截图，自动使用图像识别定位，无需手动配置",
            "• 操作期间请勿移动鼠标，否则会自动中断",
            "• 如有问题或建议，请联系开发者",
        ]
        for tip in tips:
            tk.Label(stats_frame, text=tip, bg=self.CONTENT_BG, fg="#555",
                     font=("微软雅黑", 9)).pack(anchor="w", pady=1)

        tk.Label(page, text="© 2026 企业微信自动发送工具  |  仅供内部使用",
                 bg=self.CONTENT_BG, fg="#95a5a6", font=("微软雅黑", 8)).pack(side="bottom", pady=10)

    # ---------- 日志输出 ----------

    def log(self, msg):
        """线程安全地输出日志到GUI"""
        def _write():
            self.log_area.configure(state="normal")
            self.log_area.insert(tk.END, msg + "\n")
            self.log_area.see(tk.END)
            self.log_area.configure(state="disabled")
        self.root.after(0, _write)

    # ---------- 配置存取 ----------

    def _restore_values(self):
        """从 config 恢复上次填入的值"""
        wf = self.cfg.get("watch_folder", "")
        if wf:
            self.folder_var.set(wf)
        gn = self.cfg.get("last_group_name", "")
        if gn:
            self.group_var.set(gn)
        mt = self.cfg.get("last_message_text", "")
        if mt:
            self.message_var.set(mt)
        ct = self.cfg.get("calibrate_countdown", 3)
        self.countdown_var.set(str(ct))

    def _save_values(self):
        """保存当前 GUI 中的值到 config"""
        self.cfg["watch_folder"] = self.folder_var.get().strip()
        self.cfg["last_group_name"] = self.group_var.get().strip()
        self.cfg["last_message_text"] = self.message_var.get().strip()
        try:
            self.cfg["calibrate_countdown"] = int(self.countdown_var.get().strip())
        except ValueError:
            self.cfg["calibrate_countdown"] = 3
        save_config(self.cfg)

    # ---------- 浏览文件夹 ----------

    def _is_desktop(self, path):
        """判断路径是否为桌面"""
        path = os.path.normpath(os.path.abspath(path)).lower()
        # 常见桌面路径
        userprofile = os.environ.get("USERPROFILE", "").lower()
        desktops = []
        if userprofile:
            desktops.append(os.path.normpath(os.path.join(userprofile, "Desktop")))
            desktops.append(os.path.normpath(os.path.join(userprofile, "桌面")))
        # OneDrive 桌面
        onedrive = os.environ.get("OneDrive", "")
        if onedrive:
            desktops.append(os.path.normpath(os.path.join(onedrive.lower(), "Desktop")))
            desktops.append(os.path.normpath(os.path.join(onedrive.lower(), "桌面")))
        return any(path == d for d in desktops)

    def _browse_folder(self):
        path = filedialog.askdirectory(title="选择待发送文件所在的文件夹")
        if path:
            if self._is_desktop(path):
                messagebox.showwarning(
                    "警告",
                    "❌ 不允许选择桌面作为源文件夹！\n\n"
                    "请将需要发送的文件放到一个专门的文件夹中，\n"
                    "再选择那个文件夹进行操作。"
                )
                return
            self.folder_var.set(path)

    # ---------- 发送主逻辑 ----------

    def _start_send(self):
        folder = self.folder_var.get().strip()
        group_name = self.group_var.get().strip()

        if not folder:
            messagebox.showwarning("提示", "请先选择文件夹路径")
            return
        if not os.path.isdir(folder):
            messagebox.showerror("错误", f"文件夹不存在:\n{folder}")
            return
        if not group_name:
            messagebox.showwarning("提示", "请输入企业微信群名称")
            return

        # 禁止桌面
        if self._is_desktop(folder):
            messagebox.showwarning(
                "警告",
                "❌ 不允许选择桌面作为源文件夹！\n\n"
                "请将需要发送的文件放到一个专门的文件夹中，\n"
                "再选择那个文件夹进行操作。"
            )
            return

        # 扫描文件
        extensions = [e.lower() for e in self.cfg["supported_extensions"]]
        files = []
        for f in sorted(os.listdir(folder)):
            full = os.path.join(folder, f)
            if os.path.isfile(full) and os.path.splitext(f)[1].lower() in extensions:
                files.append(full)

        if not files:
            messagebox.showinfo("提示", f"文件夹中没有可处理的文件\n{folder}")
            return

        # 确认
        names = []
        for fp in files:
            n = extract_name(os.path.basename(fp))
            names.append(n if n else "???")
        preview = "\n".join(
            f"  {os.path.basename(fp)}  →  {nm}"
            for fp, nm in zip(files, names)
        )
        ok = messagebox.askokcancel(
            "确认发送",
            f"群名称: {group_name}\n"
            f"文件数: {len(files)}\n\n"
            f"文件清单:\n{preview}\n\n"
            f"确认开始发送？\n(操作期间请勿移动鼠标)"
        )
        if not ok:
            return

        # 保存设置
        self._save_values()

        # 更新UI状态
        self.running = True
        self.stop_flag.clear()
        self.send_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.calibrate_btn.configure(state="disabled")

        # 后台线程执行
        self.worker_thread = threading.Thread(
            target=self._do_send, args=(folder, group_name, files), daemon=True
        )
        self.worker_thread.start()

    def _do_send(self, folder, group_name, files):
        total = len(files)
        success = 0
        fail = 0
        skip = 0
        message_text = self.message_var.get().strip() or None

        self.log(f"\n{'='*50}")
        self.log(f"开始发送 | 群: {group_name} | 文件数: {total}")
        if message_text:
            self.log(f"附加文字: {message_text}")
        self.log(f"{'='*50}")

        for i, fp in enumerate(files, 1):
            if self.stop_flag.is_set():
                self.log(f"\n⏹ 用户停止")
                break

            self.progress_var.set(f"进度: {i}/{total}")
            result = process_one_file(self.cfg, group_name, fp, self.log, message_text=message_text)
            if result == "success":
                success += 1
            elif result == "failed":
                fail += 1
            elif result == "interrupted":
                self.log(f"\n⏹ 用户移动鼠标中断，停止发送")
                break
            else:
                skip += 1

        self.log(f"\n{'='*50}")
        self.log(f"发送完成 | 成功: {success} | 失败: {fail} | 跳过: {skip}")
        self.log(f"{'='*50}")

        # 恢复UI
        self.root.after(0, self._finish_send)

    def _finish_send(self):
        self.running = False
        self.send_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.calibrate_btn.configure(state="normal")
        self.progress_var.set("就绪")

    def _stop_send(self):
        if messagebox.askyesno("确认", "确定要停止发送吗？"):
            self.stop_flag.set()
            self.log("\n⏹ 正在停止...")

    # ---------- 坐标校准窗口 ----------

    def _open_calibrate(self):
        win = tk.Toplevel(self.root)
        win.title("坐标校准")
        win.geometry("780x720")
        win.resizable(True, True)
        win.minsize(640, 640)

        # 读取倒计时设置：优先从输入框读取，实时生效
        try:
            countdown_sec = int(self.countdown_var.get().strip())
        except ValueError:
            countdown_sec = self.cfg.get("calibrate_countdown", 3)
        if countdown_sec < 1:
            countdown_sec = 3

        ttk.Label(
            win,
            text=f"🎯 使用方法：点击按钮 → 倒计时{countdown_sec}秒 → 把鼠标移到目标位置 → 等待自动记录",
            justify="center", font=("微软雅黑", 11, "bold"), foreground="#d4380d"
        ).pack(padx=16, pady=(12, 4))

        info_text = (
            "💡 推荐：截取按钮图片放到 images/ 目录，程序自动图像识别定位\n"
            "    （不依赖绝对坐标，不受窗口大小影响）\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "  截图准备（窗口最大化状态下截取，详见 images/截图说明.txt）：\n"
            "    ① more_btn.png          → 群成员面板右上角「...」按钮\n"
            "    ② search_member_entry.png → 菜单中的「搜索群成员」文字\n"
            "    ③ member_search_box.png  → 搜索群成员后出现的空搜索框\n"
            "  有这 3 张截图后，下方③④⑤⑥ 步骤无需手动校准！\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
            "需要校准的位置（图像识别不可用时作为回退坐标）：\n"
            "  【主窗口】\n"
            "    ① 顶部搜索框      ② 搜索结果第一项\n"
            "  【群聊 → 右侧成员面板】\n"
            "    ③ 「...」按钮     ④ 「搜索群成员」菜单项\n"
            "    ⑤ 成员搜索框      ⑥ 搜索结果第一项\n"
            "  【发送】\n"
            "    ⑦ 聊天窗口输入框区域"
        )
        ttk.Label(win, text=info_text, justify="left", font=("微软雅黑", 10)).pack(
            padx=16, pady=(4, 8))

        # 实时坐标显示
        coord_var = tk.StringVar(value="X: ---   Y: ---")
        ttk.Label(win, textvariable=coord_var,
                  font=("Consolas", 14, "bold"), foreground="#07C160").pack(pady=4)

        # 倒计时标签
        countdown_var = tk.StringVar(value="")
        ttk.Label(win, textvariable=countdown_var,
                  font=("Consolas", 16, "bold"), foreground="#ff4d4f").pack(pady=2)

        def update_coord():
            if win.winfo_exists():
                x, y = pyautogui.position()
                coord_var.set(f"X: {x:5d}   Y: {y:5d}")
                win.after(80, update_coord)

        update_coord()

        # 校准按钮区域
        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill="x", padx=16, pady=8)

        positions = [
            ("① 搜索框", "search_box"),
            ("② 搜索结果第一项", "first_result"),
            ("③ 成员面板「...」", "member_panel_more_btn"),
            ("④ 「搜索群成员」", "member_search_entry"),
            ("⑤ 成员搜索框", "member_search"),
            ("⑥ 成员列表第一项", "member_first"),
            ("⑦ 输入框区域", "chat_input"),
        ]

        self.calib_vars = {}
        row, col = 0, 0

        def start_countdown(key, var, btn, label):
            """启动延时录音：禁止所有按钮 → 倒计时 → 录音 → 恢复"""
            # 禁用所有按钮，防止重复点击
            for child in btn_frame.winfo_children():
                child.configure(state="disabled")

            seconds = [countdown_sec]  # 用列表以便闭包内修改

            def tick():
                if seconds[0] > 0:
                    countdown_var.set(f"⏳ 将在 {seconds[0]} 秒后记录「{label}」的坐标...")
                    seconds[0] -= 1
                    win.after(1000, tick)
                else:
                    countdown_var.set("📸 已记录！")
                    # 此刻鼠标已在目标位置，直接记录
                    x, y = pyautogui.position()
                    var.set(f"X:{x} Y:{y}")
                    self.log(f"[校准] {key} → X:{x} Y:{y}")
                    # 恢复所有按钮
                    for child in btn_frame.winfo_children():
                        child.configure(state="normal")
                    # 清除倒计时
                    win.after(800, lambda: countdown_var.set(""))

            countdown_var.set(f"⏳ 将在 {countdown_sec} 秒后记录「{label}」的坐标...")
            seconds[0] = countdown_sec - 1  # 第一次 tick 立即显示 N，然后 N-1, ..., 0
            win.after(1000, tick)

        for label, key in positions:
            var = tk.StringVar(value="未设置")
            self.calib_vars[key] = var
            btn = ttk.Button(
                btn_frame, text=label,
            )
            btn.configure(
                command=lambda k=key, v=var, b=btn, lb=label: start_countdown(k, v, b, lb)
            )
            btn.grid(row=row, column=col * 2, sticky="w", padx=4, pady=3)
            ttk.Label(btn_frame, textvariable=var,
                      font=("Consolas", 9), foreground="gray").grid(
                row=row, column=col * 2 + 1, padx=4, pady=3)
            row += 1
            if row >= 3:
                row = 0
                col += 1

        # 底部按钮
        bottom = ttk.Frame(win)
        bottom.pack(pady=10)

        ttk.Button(bottom, text="💾 保存到配置",
                   command=lambda: self._save_calibration(win)).pack(
            side="left", padx=8)

        ttk.Button(bottom, text="关闭", command=win.destroy).pack(side="left", padx=8)

        # 加载已有坐标
        for key in self.calib_vars:
            x = self.cfg.get(f"{key}_x", 0)
            y = self.cfg.get(f"{key}_y", 0)
            if x and y:
                self.calib_vars[key].set(f"X:{x} Y:{y}")

    def _save_calibration(self, win):
        for key, var in self.calib_vars.items():
            txt = var.get()
            if txt.startswith("X:"):
                parts = txt.replace("X:", "").replace("Y:", "").split()
                if len(parts) >= 2:
                    try:
                        self.cfg[f"{key}_x"] = int(parts[0])
                        self.cfg[f"{key}_y"] = int(parts[1])
                    except ValueError:
                        pass
        save_config(self.cfg)
        self.log("[校准] 坐标已保存到 config.json")
        messagebox.showinfo("完成", "坐标已保存到配置文件")
        win.destroy()

    # ---------- 关闭 ----------

    def _on_close(self):
        if self.running:
            if not messagebox.askyesno("确认", "正在发送中，确定要退出吗？"):
                return
            self.stop_flag.set()
        self._save_values()
        self.root.destroy()


# ========== 入口 ==========

def main():
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.08

    # 检查配置文件
    if not CONFIG_FILE.exists():
        save_config(DEFAULT_CONFIG)

    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()