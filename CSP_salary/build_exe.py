#!/usr/bin/env python3
"""
一键打包脚本：将 工资条批量生成工具 打包为单个 .exe 文件
用法: python build_exe.py
输出: dist/工资条生成工具.exe

注意事项：
  - 生成的 EXE 不含 Excel/Word 源文件，请将它们放在 EXE 同级目录下
  - 对方电脑需要安装 Word / WPS / LibreOffice 之一（用于 DOCX→PDF 转换）
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

PROJECT_DIR = Path(__file__).parent
MAIN_SCRIPT = PROJECT_DIR / "generate_salary_slips.py"
DIST_DIR = PROJECT_DIR / "dist"
EXE_NAME = "工资条生成工具"

# 需要显式导入的隐藏模块
HIDDEN_IMPORTS = [
    "pandas._libs.tslibs",
    "pandas._libs.join",
    "pandas._libs.ops",
    "pandas._libs.parsers",
    "openpyxl",
    "openpyxl.cell._writer",
    "openpyxl.utils",
    "openpyxl.styles",
    "lxml.etree",
    "lxml._elementpath",
    "fitz",
    "fitz.fitz",
    "PIL",
    "PIL._imaging",
    "PIL._imagingft",
    "win32com",
    "win32com.client",
    "pythoncom",
    "tkinter",
    "tkinter.filedialog",
    "tkinter.messagebox",
    "xml",
    "xml.etree",
    "json",
    "zipfile",
    "re",
    "subprocess",
]


def build():
    print("=" * 50)
    print("  工资条批量生成工具 - EXE 打包")
    print("=" * 50)
    print()

    # 清理旧构建
    for d in [DIST_DIR, PROJECT_DIR / "build"]:
        if d.exists():
            shutil.rmtree(d, ignore_errors=True)

    # 构建命令
    hidden_args = []
    for imp in HIDDEN_IMPORTS:
        hidden_args.extend(["--hidden-import", imp])

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",                          # 单文件 EXE
        "--noconsole",                        # 不显示控制台
        "--clean",
        "--name", EXE_NAME,

        # 收集关键包的二进制和数据文件
        "--collect-binaries", "fitz",
        "--collect-submodules", "fitz",
        "--collect-binaries", "lxml",
        "--collect-submodules", "lxml",
        "--collect-binaries", "win32com",
        "--collect-submodules", "win32com",
        "--collect-binaries", "openpyxl",
        "--collect-submodules", "openpyxl",
        "--collect-binaries", "PIL",
        "--collect-submodules", "PIL",

        "--noconfirm",

        *hidden_args,

        str(MAIN_SCRIPT),
    ]

    print(f"[INFO] 开始 PyInstaller 打包（约 1-2 分钟）...")
    print(f"[INFO] Python: {sys.executable}")
    print(f"[INFO] 入口: {MAIN_SCRIPT}")
    print()

    result = subprocess.run(cmd, cwd=str(PROJECT_DIR))

    if result.returncode != 0:
        print("\n[FAIL] 打包失败，请检查上方错误信息。")
        return 1

    exe_path = DIST_DIR / f"{EXE_NAME}.exe"
    if exe_path.exists():
        size_mb = exe_path.stat().st_size / (1024 * 1024)
        print(f"\n{'=' * 50}")
        print(f"  [OK] 打包成功！")
        print(f"  文件: {exe_path}")
        print(f"  大小: {size_mb:.1f} MB")
        print(f"{'=' * 50}")
        print()
        print("发给他人使用方法：")
        print("  1. 准备以下文件放在一起：")
        print("     - 工资条生成工具.exe   （主程序）")
        print("     - 工资表.docx          （Word 模板）")
        print("     - 工资条1.xlsx          （Excel 数据）")
        print("     （文件名可自定义，运行时自行选择）")
        print("  2. 双击 EXE 运行 → 打开选择界面")
        print("  3. 点击「选择文件」分别选取 Excel 和 Word 文件")
        print("  4. 点击「开始生成」，生成图片在 output/ 目录下")
        print()
        print("  要求：电脑需安装 Word / WPS / LibreOffice 之一")
        return 0
    else:
        print("[FAIL] EXE 文件未生成！")
        return 1


if __name__ == "__main__":
    sys.exit(build())
