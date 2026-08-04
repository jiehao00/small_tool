# -*- coding: utf-8 -*-
"""
页面模块注册中心

每个页面模块对外暴露一个 create_page(parent, app) 函数：
  - parent : 父级 Frame 容器
  - app    : ExcelToolApp 实例，提供 log()、配色常量等
  返回构建好的页面 Frame。

在 ExcelToolApp.NAV_ITEMS 中引用时直接传入 create_page 函数即可。
"""

from .performance_gen import create_page as create_performance_page
from .salary_slip import create_page as create_salary_slip_page
from .help_page import create_page as create_help_page
from .about import create_page as create_about_page

__all__ = [
    "create_performance_page",
    "create_salary_slip_page",
    "create_help_page",
    "create_about_page",
]
