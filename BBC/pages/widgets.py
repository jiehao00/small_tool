# -*- coding: utf-8 -*-
"""
共享 UI 组件库
提供圆角复选框等自定义控件，各页面统一引用使用。
"""

import tkinter as tk


class RoundedCheckbox(tk.Canvas):
    """圆角复选框
    用法：
        var = tk.BooleanVar(value=True)
        cb = RoundedCheckbox(parent, variable=var, bg=bg)
        cb.pack(side=tk.LEFT)
        tk.Label(parent, text="选项文字", ...).pack(side=tk.LEFT)
    """

    CHECK_SIZE = 18
    RADIUS = 5

    def __init__(self, parent, variable=None, **kw):
        self._var = variable or tk.BooleanVar(value=False)
        self._bg = kw.pop("bg", "#f5f5f5")

        super().__init__(parent, width=self.CHECK_SIZE + 4, height=self.CHECK_SIZE + 4,
                         bg=self._bg, highlightthickness=0, bd=0,
                         cursor="hand2", **kw)
        self.bind("<Button-1>", self._toggle)
        # 监听 BooleanVar 外部变化（例如全选/全不选按钮）
        self._trace_id = self._var.trace_add("write", self._on_var_change)
        self._draw()

    @property
    def var(self):
        return self._var

    def _draw(self):
        self.delete("all")
        c = self.CHECK_SIZE // 2 + 1
        r = self.RADIUS

        if self._var.get():
            # 选中状态：蓝色填充
            self._round_rect(c - self.CHECK_SIZE // 2, c - self.CHECK_SIZE // 2,
                             c + self.CHECK_SIZE // 2, c + self.CHECK_SIZE // 2,
                             r, fill="#3498db", outline="#3498db")
            # 白色对勾
            self.create_line(c - 4, c, c - 1, c + 3, c + 4, c - 4,
                             fill="white", width=2, capstyle=tk.ROUND,
                             joinstyle=tk.ROUND)
        else:
            # 未选中状态：空心圆角框
            self._round_rect(c - self.CHECK_SIZE // 2, c - self.CHECK_SIZE // 2,
                             c + self.CHECK_SIZE // 2, c + self.CHECK_SIZE // 2,
                             r, fill="white", outline="#bdc3c7")

    def _round_rect(self, x1, y1, x2, y2, r, **kw):
        """画圆角矩形（用 smooth polygon 模拟）"""
        points = (x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
                  x2, y2 - r, x2, y2, x2 - r, y2,
                  x1 + r, y2, x1, y2, x1, y2 - r,
                  x1, y1 + r, x1, y1)
        self.create_polygon(points, smooth=True, **kw)

    def _on_var_change(self, *args):
        """BooleanVar 外部变化时自动重绘"""
        self._draw()

    def _toggle(self, event=None):
        self._var.set(not self._var.get())
        # 注意：set() 会触发 _on_var_change 自动调用 _draw()，无需手动再调

    def pack(self, **kw):
        super().pack(**kw)
        return self

    def grid(self, **kw):
        super().grid(**kw)
        return self
