# -*- coding: utf-8 -*-
"""关于页面"""

import tkinter as tk
from .logo_data import LOGO_ABOUT_B64, b64_to_photoimage


VERSION = "1.0.0"


def create_page(parent, app):
    bg = app.COLOR_CONTENT_BG
    frame = tk.Frame(parent, bg=bg)

    # ---- Logo ----
    logo_img = b64_to_photoimage(LOGO_ABOUT_B64)
    logo_label = tk.Label(frame, image=logo_img, bg=bg)
    logo_label.image = logo_img  # 保持引用防止被 GC
    logo_label.pack(pady=(20, 0))

    # ---- 名称 + 版本 ----
    header = tk.Frame(frame, bg=bg)
    header.pack(fill=tk.X, pady=(8, 0))

    tk.Label(
        header,
        text="BBC百宝箱",
        font=("Microsoft YaHei", 22, "bold"),
        bg=bg,
        fg="#2c3e50"
    ).pack()

    tk.Label(
        header,
        text=f"版本 {VERSION}",
        font=("Microsoft YaHei", 11),
        bg=bg,
        fg="#7f8c8d"
    ).pack(pady=(6, 16))

    # ---- 简介卡片 ----
    card = tk.Frame(frame, bg="white", padx=28, pady=18)
    card.pack(fill=tk.X, padx=80, pady=(0, 16))

    tk.Label(
        card,
        text="一款实用、高效的 Excel 数据填充工具，通过键值匹配自动补齐缺失字段。",
        font=("Microsoft YaHei", 11),
        bg="white",
        fg="#34495e",
        wraplength=600,
        justify=tk.CENTER
    ).pack()

    # ---- 主要功能 ----
    tk.Label(
        frame,
        text="主要功能",
        font=("Microsoft YaHei", 14, "bold"),
        bg=bg,
        fg="#2c3e50"
    ).pack(anchor=tk.W, padx=80, pady=(8, 12))

    features = [
        ("业绩生成", "导入原始数据，自动生成 BBC 业绩报表"),
        ("工资条生成", "按人员匹配汇总与明细，批量拆分工资条"),
    ]

    features_frame = tk.Frame(frame, bg=bg)
    features_frame.pack(fill=tk.X, padx=80, pady=(0, 8))

    # 单列布局
    cols = [features]
    for col_features in cols:
        col = tk.Frame(features_frame, bg=bg)
        col.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 16))

        for title, desc in col_features:
            row = tk.Frame(col, bg=bg)
            row.pack(fill=tk.X, pady=(0, 10))

            dot = tk.Label(
                row, text="●",
                font=("Microsoft YaHei", 8),
                bg=bg, fg="#3498db", width=2, anchor=tk.W
            )
            dot.pack(side=tk.LEFT)

            tk.Label(
                row, text=title,
                font=("Microsoft YaHei", 11, "bold"),
                bg=bg, fg="#2c3e50", anchor=tk.W
            ).pack(side=tk.LEFT, padx=(0, 8))

            tk.Label(
                row, text=desc,
                font=("Microsoft YaHei", 10),
                bg=bg, fg="#7f8c8d", anchor=tk.W
            ).pack(side=tk.LEFT)

    # ---- 底部版权 ----
    tk.Label(
        frame,
        text="© 2025 财务Excel工具箱\n使用过程中如有问题，请查看「使用说明」页面",
        font=("Microsoft YaHei", 9),
        bg=bg,
        fg="#95a5a6",
        justify=tk.CENTER
    ).pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 16))

    return frame
