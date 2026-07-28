# -*- coding: utf-8 -*-
"""关于页面"""

import tkinter as tk


VERSION = "1.0.0"
RELEASE_DATE = "2026-07-24"


def create_page(parent, app):
    bg = app.COLOR_CONTENT_BG
    frame = tk.Frame(parent, bg=bg)

    # ---- 顶部标题区 ----
    header = tk.Frame(frame, bg=bg)
    header.pack(fill=tk.X, pady=(40, 0))

    tk.Label(
        header,
        text="Excel 通用小工具",
        font=("Microsoft YaHei", 22, "bold"),
        bg=bg,
        fg="#2c3e50"
    ).pack()

    tk.Label(
        header,
        text=f"版本: {VERSION}",
        font=("Microsoft YaHei", 11),
        bg=bg,
        fg="#7f8c8d"
    ).pack(pady=(6, 2))

    tk.Label(
        header,
        text=f"上线时间: {RELEASE_DATE}",
        font=("Microsoft YaHei", 10),
        bg=bg,
        fg="#95a5a6"
    ).pack(pady=(0, 24))

    # ---- 简介卡片 ----
    card = tk.Frame(frame, bg="white", padx=24, pady=20)
    card.pack(fill=tk.X, padx=60, pady=(0, 20))

    tk.Label(
        card,
        text="一款轻量化的 Excel 辅助工具，专注于日常办公中常见的表格处理需求。",
        font=("Microsoft YaHei", 11),
        bg="white",
        fg="#34495e",
        wraplength=600,
        justify=tk.CENTER
    ).pack()

    # ---- 功能特性 ----
    tk.Label(
        frame,
        text="主要功能",
        font=("Microsoft YaHei", 13, "bold"),
        bg=bg,
        fg="#2c3e50"
    ).pack(anchor=tk.W, padx=60, pady=(10, 12))

    features = [
        ("填充数据", "按匹配列自动从另一个表格填充缺失内容"),
        ("数据对比", "比较两个表格，生成交集、仅A有、仅B有三部分结果"),
        ("数据合并", "按主键去重并对数值列进行聚合统计"),
        ("数据拆分", "将一列中多个值拆分为多行"),
        ("多文件追加", "纵向拼接多个结构相似或不同的 Excel 文件"),
        ("按列值拆分", "按指定列的唯一值拆分为多个独立文件"),
    ]

    features_frame = tk.Frame(frame, bg=bg)
    features_frame.pack(fill=tk.X, padx=60, pady=(0, 20))

    for idx, (title, desc) in enumerate(features):
        row = tk.Frame(features_frame, bg=bg)
        row.pack(fill=tk.X, pady=(0, 8))

        # 蓝色圆点
        dot = tk.Label(
            row,
            text="●",
            font=("Microsoft YaHei", 8),
            bg=bg,
            fg="#3498db",
            width=2,
            anchor=tk.W
        )
        dot.pack(side=tk.LEFT)

        tk.Label(
            row,
            text=title,
            font=("Microsoft YaHei", 10, "bold"),
            bg=bg,
            fg="#2c3e50",
            width=12,
            anchor=tk.W
        ).pack(side=tk.LEFT)

        tk.Label(
            row,
            text=desc,
            font=("Microsoft YaHei", 10),
            bg=bg,
            fg="#555555",
            anchor=tk.W
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

    # ---- 技术与环境 ----
    tk.Label(
        frame,
        text="技术栈",
        font=("Microsoft YaHei", 13, "bold"),
        bg=bg,
        fg="#2c3e50"
    ).pack(anchor=tk.W, padx=60, pady=(10, 12))

    tech_frame = tk.Frame(frame, bg=bg)
    tech_frame.pack(fill=tk.X, padx=60, pady=(0, 20))

    techs = [
        ("Python", "核心开发语言"),
        ("tkinter", "图形用户界面"),
        ("openpyxl", "Excel 读写处理"),
    ]

    for name, desc in techs:
        row = tk.Frame(tech_frame, bg=bg)
        row.pack(fill=tk.X, pady=(0, 6))
        tk.Label(
            row,
            text=f"{name}：",
            font=("Microsoft YaHei", 10, "bold"),
            bg=bg,
            fg="#2c3e50",
            width=12,
            anchor=tk.W
        ).pack(side=tk.LEFT)
        tk.Label(
            row,
            text=desc,
            font=("Microsoft YaHei", 10),
            bg=bg,
            fg="#555555",
            anchor=tk.W
        ).pack(side=tk.LEFT)

    # ---- 底部版权 ----
    tk.Label(
        frame,
        text="© 2025 Excel 通用小工具\n使用过程中如有问题，请查看「使用说明」页面",
        font=("Microsoft YaHei", 9),
        bg=bg,
        fg="#95a5a6",
        justify=tk.CENTER
    ).pack(side=tk.BOTTOM, pady=(0, 30))

    return frame
