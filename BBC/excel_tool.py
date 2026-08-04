#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BBC百宝箱 — 主入口
只负责：窗口框架、左侧导航栏、页面切换、公共日志
各功能页面独立存放在 pages/ 目录下
"""

import tkinter as tk
from tkinter import ttk
from datetime import datetime

from pages import (
    create_performance_page,
    create_salary_slip_page,
    create_help_page,
    create_about_page,
)
from pages.logo_data import LOGO_NAV_B64, b64_to_photoimage


class ExcelToolApp:
    """主应用框架"""

    # ========== 配色常量 ==========
    COLOR_NAV_BG          = "#2c3e50"
    COLOR_NAV_HOVER       = "#34495e"
    COLOR_NAV_ACTIVE      = "#3498db"
    COLOR_CONTENT_BG      = "#f5f5f5"
    COLOR_BTN_PRIMARY     = "#3498db"
    COLOR_BTN_PRIMARY_HOVER = "#2980b9"

    # ========== 导航配置 ==========
    # (显示文本, page_id, 页面构建函数, 是否懒加载)
    NAV_ITEMS = [
        ("业绩生成", "performance", create_performance_page, False),
        ("工资条生成", "salary_slip", create_salary_slip_page, False),
        ("使用说明", "help",      create_help_page,     True),
        ("关于",     "about",     create_about_page,    True),
    ]

    def __init__(self, root):
        self.root = root
        self.root.title("BBC百宝箱")
        self.root.geometry("900x800")
        self.root.minsize(850, 750)
        self.root.configure(bg=self.COLOR_CONTENT_BG)

        self.nav_buttons = {}        # page_id → Label
        self.page_builders = {}      # page_id → build_func
        self.pages = {}              # page_id → Frame（懒加载）
        self.current_page = None
        self._log_callback = None    # 可选的公共日志回调

        self._init_ttk_styles()
        self._build_ui()
        self._switch_page("performance")

    # ===== 初始化 ttk 主题样式 =====
    def _init_ttk_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")  # clam 主题更易于自定义
        except tk.TclError:
            pass

        # ---- 按钮样式 ----
        # 主按钮（蓝色，大号）
        style.configure("Primary.TButton",
            font=("Microsoft YaHei", 11, "bold"),
            padding=(30, 8),
            borderwidth=0,
            relief="flat",
            foreground="white",
            background="#3498db",
            focusthickness=0)
        style.map("Primary.TButton",
            background=[("active", "#2980b9"), ("disabled", "#aab7c4")],
            foreground=[("disabled", "#e0e0e0")])

        # 次要按钮（白色，带边框，用于"浏览...")
        style.configure("Secondary.TButton",
            font=("Microsoft YaHei", 9),
            padding=(12, 4),
            borderwidth=1,
            relief="solid",
            foreground="#2c3e50",
            background="#ecf0f1",
            focusthickness=0,
            bordercolor="#d5dce0")
        style.map("Secondary.TButton",
            background=[("active", "#d5dbdb"), ("disabled", "#f0f3f4")],
            bordercolor=[("active", "#bdc3c7"), ("disabled", "#d5dce0")])

        # 启动按钮（绿色）
        style.configure("Action.TButton",
            font=("Microsoft YaHei", 11, "bold"),
            padding=(30, 8),
            borderwidth=0,
            relief="flat",
            foreground="white",
            background="#27ae60",
            focusthickness=0)
        style.map("Action.TButton",
            background=[("active", "#219a52"), ("disabled", "#aab7c4")],
            foreground=[("disabled", "#e0e0e0")])

        # ---- 复选框样式 ----
        style.configure("Card.TCheckbutton",
            font=("Microsoft YaHei", 10),
            foreground="#2c3e50",
            background="#f5f5f5",
            focusthickness=0)
        style.map("Card.TCheckbutton",
            background=[("active", "#f5f5f5"), ("selected", "#f5f5f5")])

        # ---- 输入框样式 ----
        style.configure("Normal.TEntry",
            font=("Microsoft YaHei", 10),
            padding=(8, 6),
            relief="solid",
            borderwidth=1,
            fieldbackground="white",
            foreground="#2c3e50",
            lightcolor="#d5dce0",
            darkcolor="#d5dce0",
            bordercolor="#d5dce0")
        style.configure("Readonly.TEntry",
            font=("Microsoft YaHei", 10),
            padding=(8, 6),
            relief="solid",
            borderwidth=1,
            fieldbackground="#f5f5f5",
            foreground="#555555",
            lightcolor="#d5dce0",
            darkcolor="#d5dce0",
            bordercolor="#d5dce0")
        style.configure("CondEntry.TEntry",
            font=("Microsoft YaHei", 10),
            padding=(6, 4),
            relief="solid",
            borderwidth=1,
            fieldbackground="white",
            foreground="#2c3e50",
            lightcolor="#d5dce0",
            darkcolor="#d5dce0",
            bordercolor="#d5dce0")

        # ---- 条件筛选下拉框样式 ----
        # 弹出列表（Listbox）的样式：圆角不可达，但通过 option_add 改善字体/选中色
        self.root.option_add("*TCombobox*Listbox.font", ("Microsoft YaHei", 10))
        self.root.option_add("*TCombobox*Listbox.background", "white")
        self.root.option_add("*TCombobox*Listbox.foreground", "#2c3e50")
        self.root.option_add("*TCombobox*Listbox.selectBackground", "#3498db")
        self.root.option_add("*TCombobox*Listbox.selectForeground", "white")
        self.root.option_add("*TCombobox*Listbox.relief", "solid")
        self.root.option_add("*TCombobox*Listbox.borderWidth", 1)
        self.root.option_add("*TCombobox*Listbox.highlightThickness", 0)

        style.configure("CondCol.TCombobox",
            font=("Microsoft YaHei", 10),
            padding=(10, 6),
            borderwidth=1,
            relief="solid",
            bordercolor="#d5dce0",
            lightcolor="#d5dce0",
            darkcolor="#d5dce0",
            arrowcolor="#7f8c8d",
            arrowsize=14,
            fieldbackground="white",
            background="white",
            foreground="#2c3e50")
        style.map("CondCol.TCombobox",
            fieldbackground=[("readonly", "white"), ("disabled", "#f0f0f0"), ("focus", "white"), ("active", "white")],
            background=[("readonly", "white"), ("disabled", "#f0f0f0"), ("focus", "white"), ("active", "white")],
            foreground=[("readonly", "#2c3e50"), ("disabled", "#aaaaaa")],
            bordercolor=[("focus", "#3498db"), ("active", "#3498db")])
        style.configure("CondOp.TCombobox",
            font=("Microsoft YaHei", 10),
            padding=(10, 6),
            borderwidth=1,
            relief="solid",
            bordercolor="#d5dce0",
            lightcolor="#d5dce0",
            darkcolor="#d5dce0",
            arrowcolor="#7f8c8d",
            arrowsize=14,
            fieldbackground="white",
            background="white",
            foreground="#2c3e50")
        style.map("CondOp.TCombobox",
            fieldbackground=[("readonly", "white"), ("disabled", "#f0f0f0"), ("focus", "white"), ("active", "white")],
            background=[("readonly", "white"), ("disabled", "#f0f0f0"), ("focus", "white"), ("active", "white")],
            foreground=[("readonly", "#2c3e50"), ("disabled", "#aaaaaa")],
            bordercolor=[("focus", "#3498db"), ("active", "#3498db")])
        style.configure("CondConn.TCombobox",
            font=("Microsoft YaHei", 9, "bold"),
            padding=(8, 6),
            borderwidth=1,
            relief="solid",
            bordercolor="#d5dce0",
            lightcolor="#d5dce0",
            darkcolor="#d5dce0",
            arrowcolor="#3498db",
            arrowsize=14,
            fieldbackground="#f8f9fa",
            background="#f8f9fa",
            foreground="#3498db")
        style.map("CondConn.TCombobox",
            fieldbackground=[("readonly", "#f8f9fa"), ("focus", "white"), ("active", "white")],
            background=[("readonly", "#f8f9fa"), ("focus", "white"), ("active", "white")],
            foreground=[("readonly", "#3498db")],
            bordercolor=[("focus", "#3498db"), ("active", "#3498db")])

        # ---- LabelFrame 卡片样式 ----
        # 边框几乎透明，只用留白和标题做分区
        style.configure("Card.TLabelframe",
            background=self.COLOR_CONTENT_BG,
            bordercolor="#f0f0f0",
            relief="solid",
            borderwidth=1,
            labelmargins=(10, 2))
        style.configure("Card.TLabelframe.Label",
            font=("Microsoft YaHei", 11, "bold"),
            foreground="#2c3e50",
            background=self.COLOR_CONTENT_BG)

        # ---- 滚动条样式（扁平化） ----
        style.configure("Vertical.TScrollbar",
            background="#e8ecf0",
            troughcolor="#f5f7fa",
            bordercolor="#f5f7fa",
            arrowcolor="#7f8c8d",
            lightcolor="#f5f7fa",
            darkcolor="#f5f7fa",
            borderwidth=0,
            arrowsize=14)
        style.map("Vertical.TScrollbar",
            background=[("active", "#d5dce0"), ("pressed", "#bdc3c7")],
            arrowcolor=[("active", "#2c3e50")])
        style.configure("Horizontal.TScrollbar",
            background="#e8ecf0",
            troughcolor="#f5f7fa",
            bordercolor="#f5f7fa",
            arrowcolor="#7f8c8d",
            lightcolor="#f5f7fa",
            darkcolor="#f5f7fa",
            borderwidth=0,
            arrowsize=14)
        style.map("Horizontal.TScrollbar",
            background=[("active", "#d5dce0"), ("pressed", "#bdc3c7")],
            arrowcolor=[("active", "#2c3e50")])

        # 将 style 引用保存，供页面模块获取
        self.ttk_style = style

    # ===== 构建主框架 =====
    def _build_ui(self):
        # 左侧导航栏
        self.nav_frame = tk.Frame(self.root, width=160, bg=self.COLOR_NAV_BG)
        self.nav_frame.pack(side=tk.LEFT, fill=tk.Y)
        self.nav_frame.pack_propagate(False)

        # 侧边栏小 Logo
        nav_logo_img = b64_to_photoimage(LOGO_NAV_B64)
        nav_logo_label = tk.Label(self.nav_frame, image=nav_logo_img,
                                  bg=self.COLOR_NAV_BG)
        nav_logo_label.image = nav_logo_img  # 保持引用
        nav_logo_label.pack(pady=(18, 12))

        for text, pid, builder, _lazy in self.NAV_ITEMS:
            self.page_builders[pid] = builder
            btn = tk.Label(self.nav_frame, text=text,
                font=("Microsoft YaHei", 11),
                bg=self.COLOR_NAV_BG, fg="white",
                padx=20, pady=12, anchor=tk.W, cursor="hand2")
            btn.pack(fill=tk.X)
            btn.bind("<Button-1>", lambda e, p=pid: self._switch_page(p))
            btn.bind("<Enter>",
                lambda e, b=btn: b.config(bg=self.COLOR_NAV_HOVER))
            btn.bind("<Leave>",
                lambda e, b=btn, p=pid: b.config(
                    bg=self.COLOR_NAV_ACTIVE if self.current_page == p else self.COLOR_NAV_BG))
            self.nav_buttons[pid] = btn

        # 右侧内容容器
        self.content_container = tk.Frame(self.root, bg=self.COLOR_CONTENT_BG)
        self.content_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=15, pady=15)

    # ===== 页面切换 =====
    def _switch_page(self, page_id):
        # 导航按钮高亮
        for pid, btn in self.nav_buttons.items():
            btn.config(bg=self.COLOR_NAV_ACTIVE if pid == page_id else self.COLOR_NAV_BG)
        self.current_page = page_id

        # 隐藏所有页面
        for p in self.pages.values():
            p.pack_forget()

        # 懒加载
        if page_id not in self.pages:
            builder = self.page_builders[page_id]
            self.pages[page_id] = builder(self.content_container, self)

        self.pages[page_id].pack(fill=tk.BOTH, expand=True)

    # ===== 公共日志 =====
    def set_log_callback(self, callback):
        """设置外部日志回调（如公共日志框），供页面模块调用"""
        self._log_callback = callback

    def log(self, message):
        """页面模块可调用此方法输出公共日志"""
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {message}"
        print(line)
        if self._log_callback:
            self._log_callback(line + "\n")


# ===== 主入口 =====
def main():
    root = tk.Tk()
    ExcelToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
