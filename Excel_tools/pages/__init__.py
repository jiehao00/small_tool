# -*- coding: utf-8 -*-
"""
页面模块注册中心

每个页面模块对外暴露一个 create_page(parent, app) 函数：
  - parent : 父级 Frame 容器
  - app    : ExcelToolApp 实例，提供 log()、配色常量等
  返回构建好的页面 Frame。

在 ExcelToolApp.NAV_ITEMS 中引用时直接传入 create_page 函数即可。
"""

from .merge_data import create_page as create_merge_page
from .data_compare import create_page as create_compare_page
from .data_merge import create_page as create_data_merge_page
from .data_split import create_page as create_split_page
from .data_append import create_page as create_append_page
from .column_split import create_page as create_column_split_page
from .help_page import create_page as create_help_page
from .about import create_page as create_about_page

__all__ = [
    "create_merge_page",
    "create_compare_page",
    "create_data_merge_page",
    "create_split_page",
    "create_append_page",
    "create_column_split_page",
    "create_help_page",
    "create_about_page",
]
