#!/usr/bin/env python3
"""
工资条批量生成工具
功能：读取 Excel 员工数据 → 复制 Word 模板 → 替换 <<字段名>> 占位符 → 导出 PNG 图片 / PDF 文件
输出：{姓名}_{工号}.png 或 {姓名}_{工号}.pdf
"""

import os
import re
import sys
import math
import shutil
import tempfile
import zipfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from pathlib import Path
from typing import Optional

# 保证 Windows 控制台也能正确输出中文（PyInstaller 打包后无需处理）
if not getattr(sys, 'frozen', False):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    try:
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass


def _get_app_dir() -> Path:
    """
    返回应用根目录：
    - 正常 Python 运行：脚本所在目录
    - PyInstaller 打包后：EXE 所在目录（数据文件如 Excel/Word 模板放在 EXE 旁边）
    """
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包后的环境
        return Path(sys.executable).parent
    return Path(__file__).parent


# 注意：pandas / lxml 等重库已改为延迟导入，放在实际使用的函数中，
# 确保主界面能快速弹出，用户体验更好。

# ===================== 配置区 =====================

EXCEL_PATH = "工资条1.xlsx"          # Excel 数据源
TEMPLATE_PATH = "工资表.docx"         # Word 模板（含 <<字段>> 或邮件合并域）
OUTPUT_DIR = "output"                 # 输出目录
IMG_DPI = 300                         # 图片 DPI（高清）
CROP_MODE = "page"                    # "page" = 整页（推荐） / "table" = 自动截取表格区域（实验性）

DELETE_TEMP = True                    # 是否删除临时文件

# ==================================================


NS = {
    'w':  'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r':  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'w14':'http://schemas.microsoft.com/office/word/2010/wordml',
}


# ---------- 步骤 1: 读取 Excel ----------

def read_excel(path: str):
    import pandas as pd
    df = pd.read_excel(path)
    # 清理列名空格
    df.columns = [str(c).strip() for c in df.columns]
    # 工号转为字符串（避免科学计数法）
    if '工号' in df.columns:
        df['工号'] = df['工号'].apply(
            lambda x: str(int(x)) if pd.notna(x) else ''
        )
    return df


# ---------- 步骤 2: 替换占位符 (XML 级别) ----------

def _replace_placeholders_in_xml(xml_str: str, row: dict) -> str:
    """
    替换 document.xml 中的占位符：
      1. Word 邮件合并域 MERGEFIELD  →  直接替换为实际数值（删除整个域结构）
      2. <<字段名>> 或 «字段名» 纯文本占位符  →  替换为实际数值
    """
    import pandas as pd
    from lxml import etree

    W = NS['w']

    # 步骤 1: 将 MERGEFIELD 域结构整体替换为纯文本
    tree = etree.fromstring(xml_str.encode('utf-8'))
    for paragraph in tree.iter(f'{{{W}}}p'):
        children = list(paragraph)
        new_children = []
        i = 0
        while i < len(children):
            child = children[i]
            if child.tag != f'{{{W}}}r':
                new_children.append(child)
                i += 1
                continue

            begin = child.find(f'{{{W}}}fldChar')
            if begin is not None and begin.get(f'{{{W}}}fldCharType') == 'begin':
                field_name = None
                j = i
                while j < len(children):
                    c = children[j]
                    if c.tag == f'{{{W}}}r':
                        instr = c.find(f'{{{W}}}instrText')
                        if instr is not None and instr.text:
                            m = re.search(r'MERGEFIELD\s+([\S\u4e00-\u9fff（）()]+)', instr.text)
                            if m:
                                field_name = m.group(1).strip()
                        end = c.find(f'{{{W}}}fldChar')
                        if end is not None and end.get(f'{{{W}}}fldCharType') == 'end':
                            break
                    j += 1

                if field_name and j < len(children):
                    val = row.get(field_name)
                    replace_text = _fmt(val) if val is not None and not (isinstance(val, float) and pd.isna(val)) else ''
                    new_run = etree.Element(f'{{{W}}}r')
                    t = etree.SubElement(new_run, f'{{{W}}}t')
                    t.text = replace_text
                    new_children.append(new_run)
                    i = j + 1
                    continue
                else:
                    new_children.append(child)
                    i += 1
                    continue

            new_children.append(child)
            i += 1

        paragraph.clear()
        for c in new_children:
            paragraph.append(c)

    xml_str = etree.tostring(tree, encoding='unicode')

    # 步骤 2: 替换剩余的 <<字段名>> / «字段名» 纯文本占位符
    def repl_text(m):
        key = m.group(1).strip()
        val = row.get(key)
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return ''
        return _fmt(val)

    xml_str = re.sub(r'(?:&lt;&lt;|<<|«)(.+?)(?:&gt;&gt;|>>|»)', repl_text, xml_str)
    return xml_str







def _fmt(val) -> str:
    """格式化数值"""
    if isinstance(val, float):
        if math.isnan(val):
            return ''
        if val == int(val):
            return str(int(val))
        return f'{val:.2f}'
    return str(val)


def fill_template(template_path: str, row: dict, output_path: str):
    """用一行数据填充模板，生成临时 docx（纯 XML 替换，无需 Office）"""
    with zipfile.ZipFile(template_path, 'r') as zin:
        file_list = zin.namelist()
        doc_xml = zin.read('word/document.xml').decode('utf-8')

    doc_xml = _replace_placeholders_in_xml(doc_xml, row)

    header_footer_files = {
        'word/header1.xml': 'header',
        'word/header2.xml': 'header',
        'word/footer1.xml': 'footer',
        'word/footer2.xml': 'footer',
    }

    with zipfile.ZipFile(template_path, 'r') as zin:
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in file_list:
                if item == 'word/document.xml':
                    zout.writestr(item, doc_xml.encode('utf-8'))
                elif item in header_footer_files:
                    raw = zin.read(item).decode('utf-8')
                    raw = _replace_placeholders_in_xml(raw, row)
                    zout.writestr(item, raw.encode('utf-8'))
                else:
                    zout.writestr(item, zin.read(item))


# ---------- 步骤 3: DOCX → PDF ----------


def _find_office_tool():
    """
    自动查找可用的 Office 工具，返回 (名称, exe_path, 转换命令构建函数)
    检测顺序: Word COM → WPS COM → LibreOffice
                   Word 更快更稳定，优先使用
    """
    import subprocess
    import glob

    # 1) 优先 Word COM（速度最快）
    try:
        import win32com.client  # noqa: F401
        import pythoncom  # noqa: F401
        return ('word_com', None, None)
    except ImportError:
        pass

    # 2) 查找 WPS
    #    先检查注册表，再检查常见路径，最后全局搜索
    wps_path = _find_wps_via_registry()
    if wps_path:
        return ('wps', wps_path, None)

    wps_paths = [
        os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Kingsoft', 'WPS Office', 'office6', 'wps.exe'),
        os.path.join(os.environ.get('PROGRAMFILES', ''), 'Kingsoft', 'WPS Office', 'office6', 'wps.exe'),
        os.path.join(os.environ.get('PROGRAMFILES(X86)', ''), 'Kingsoft', 'WPS Office', 'office6', 'wps.exe'),
        r'C:\Program Files\Kingsoft\WPS Office\office6\wps.exe',
        r'C:\Program Files (x86)\Kingsoft\WPS Office\office6\wps.exe',
    ]
    for p in wps_paths:
        if os.path.exists(p):
            return ('wps', p, None)

    # 尝试通过 PATH 找到 wps.exe
    try:
        result = subprocess.run(['where', 'wps.exe'], capture_output=True, timeout=5, text=True)
        if result.returncode == 0:
            lines = result.stdout.strip().split('\n')
            if lines:
                return ('wps', os.path.normpath(lines[0].strip()), None)
    except Exception:
        pass

    # 全局搜索 (用户目录、Program Files)
    try:
        search_roots = [
            os.environ.get('USERPROFILE', ''),
            os.environ.get('LOCALAPPDATA', ''),
            r'C:\Program Files',
            r'C:\Program Files (x86)',
            'D:/',
            'E:/',
        ]
        for root in search_roots:
            if not root or not os.path.isdir(root):
                continue
            for ext_dir in ['Kingsoft', 'WPS*']:
                pattern = os.path.join(root, ext_dir, '**', 'wps.exe')
                matches = glob.glob(pattern, recursive=True)
                if matches:
                    return ('wps', os.path.normpath(matches[0]), None)
    except Exception:
        pass

    # 3) 查找 LibreOffice
    lo_paths = [
        os.path.join(os.environ.get('PROGRAMFILES', ''), 'LibreOffice', 'program', 'soffice.exe'),
        os.path.join(os.environ.get('PROGRAMFILES(X86)', ''), 'LibreOffice', 'program', 'soffice.exe'),
        r'C:\Program Files\LibreOffice\program\soffice.exe',
        r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
    ]
    for p in lo_paths:
        if os.path.exists(p):
            return ('libreoffice', p, None)

    # 4) 尝试 PATH 中的 soffice
    try:
        result = subprocess.run(['soffice', '--version'], capture_output=True, timeout=5)
        if result.returncode == 0:
            return ('libreoffice', 'soffice', None)
    except Exception:
        pass

    return (None, None, None)


def _find_wps_via_registry():
    """通过 Windows 注册表查找 WPS 安装路径"""
    import subprocess
    reg_queries = [
        r'HKLM\SOFTWARE\WOW6432Node\Kingsoft\WPS Office',
        r'HKLM\SOFTWARE\Kingsoft\WPS Office',
        r'HKCU\SOFTWARE\Kingsoft\WPS Office',
    ]
    for key in reg_queries:
        try:
            result = subprocess.run(
                ['reg', 'query', key, '/v', 'InstallRoot'],
                capture_output=True, timeout=5, text=True
            )
            if result.returncode == 0:
                for line in result.stdout.split('\n'):
                    if 'InstallRoot' in line:
                        parts = line.strip().split()
                        # 最后一段是路径
                        root = parts[-1].strip()
                        wps = os.path.join(root, 'office6', 'wps.exe')
                        if os.path.exists(wps):
                            return os.path.normpath(wps)
        except Exception:
            pass
    return None


class DocxToPdfConverter:
    """DOCX→PDF 转换器（复用 Office 实例，避免重复启动）"""

    def __init__(self):
        self._com_initialized = False
        self._word = None
        self._tool_name, self._tool_path, _ = _find_office_tool()

        # WPS 和 Word 都使用 COM 接口
        if self._tool_name in ('wps', 'word_com'):
            import pythoncom
            pythoncom.CoInitialize()
            self._com_initialized = True

    def convert(self, docx_path: str, pdf_path: str):
        """转换单个文件"""
        abs_docx = os.path.abspath(docx_path)
        abs_pdf = os.path.abspath(pdf_path)

        if self._tool_name in ('wps', 'word_com'):
            self._convert_word_com(abs_docx, abs_pdf)
        elif self._tool_name == 'libreoffice' and self._tool_path:
            docx_to_pdf_libreoffice(self._tool_path, abs_docx, abs_pdf)
        else:
            raise RuntimeError(
                '未找到可用的 Office 工具！\n请安装 WPS / Word / LibreOffice 之一'
            )

    def _convert_word_com(self, abs_docx: str, abs_pdf: str):
        """复用同一个 Office 实例（WPS 或 Word）"""
        from win32com import client as wc

        if self._word is None:
            if self._tool_name == 'wps':
                prog_id = 'KWPS.Application'
            else:
                prog_id = 'Word.Application'
            self._word = wc.DispatchEx(prog_id)
            self._word.Visible = False
            self._word.DisplayAlerts = 0

        doc = self._word.Documents.Open(abs_docx)
        try:
            doc.ExportAsFixedFormat(
                abs_pdf,
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0,
                CreateBookmarks=0,
            )
        finally:
            doc.Close(SaveChanges=False)

    def close(self):
        """关闭 Office 实例，释放 COM 资源"""
        if self._word is not None:
            try:
                self._word.Quit()
            except Exception:
                pass
            self._word = None
        if self._com_initialized:
            import pythoncom
            pythoncom.CoUninitialize()
            self._com_initialized = False


def docx_to_pdf_libreoffice(lo_exe: str, docx_path: str, pdf_path: str):
    """LibreOffice 命令行转换"""
    import subprocess
    out_dir = os.path.dirname(os.path.abspath(pdf_path))
    subprocess.run([
        lo_exe,
        '--headless',
        '--convert-to', 'pdf',
        '--outdir', out_dir,
        os.path.abspath(docx_path),
    ], check=True, timeout=60)


def docx_to_pdf(docx_path: str, pdf_path: str):
    """自动选择最佳 DOCX→PDF 转换方式（单文件便捷调用）"""
    converter = DocxToPdfConverter()
    try:
        converter.convert(docx_path, pdf_path)
        tool_label = {'wps': 'WPS', 'word_com': 'Word', 'libreoffice': 'LibreOffice'}
        print(f'  [OK] PDF 已生成 ({tool_label.get(converter._tool_name, converter._tool_name)})')
    finally:
        converter.close()


# ---------- 步骤 4: PDF → 高清 PNG ----------

def pdf_to_png(pdf_path: str, png_path: str, crop_mode: str = 'table'):
    """PyMuPDF 将 PDF 每页转为高清 PNG，可选截取表格区域"""
    import fitz  # PyMuPDF

    doc = fitz.open(pdf_path)
    page_count = len(doc)

    if page_count == 0:
        raise ValueError('PDF 无页面')

    # 渲染高清图片 (300 DPI)
    zoom = IMG_DPI / 72.0
    mat = fitz.Matrix(zoom, zoom)

    if page_count == 1:
        page = doc[0]
        pix = page.get_pixmap(matrix=mat)

        if crop_mode == 'table':
            cropped = _crop_to_table(page, pix, mat)
            if cropped is not None:
                cropped.save(png_path)
            else:
                pix.save(png_path)
        else:
            pix.save(png_path)
    else:
        # 多页：垂直拼接
        from PIL import Image
        images = []
        total_height = 0
        max_width = 0
        for page in doc:
            pix = page.get_pixmap(matrix=mat)
            if crop_mode == 'table':
                cropped = _crop_to_table(page, pix, mat)
                if cropped is not None:
                    img = cropped
                else:
                    img = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
            else:
                img = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)

            images.append(img)

            total_height += img.height
            max_width = max(max_width, img.width)

        combined = Image.new('RGB', (max_width, total_height), 'white')
        y_offset = 0
        for img in images:
            combined.paste(img, (0, y_offset))
            y_offset += img.height
        combined.save(png_path)

    doc.close()



def _crop_to_table(page, pix, mat):
    """
    自动检测页面内容边界并裁切（优先表格，其次文本包围盒）
    返回 PIL.Image 或 None（裁切失败）
    """
    from PIL import Image

    bbox = None
    try:
        tables = page.find_tables()
        if tables and tables.tables:
            x0 = min(t.bbox[0] for t in tables.tables)
            y0 = min(t.bbox[1] for t in tables.tables)
            x1 = max(t.bbox[2] for t in tables.tables)
            y1 = max(t.bbox[3] for t in tables.tables)
            bbox = (x0, y0, x1, y1)
    except Exception:
        pass

    # 回退：通过所有文本块计算包围盒
    if bbox is None:
        try:
            blocks = page.get_text('blocks')
            if blocks:
                x0 = min(b[0] for b in blocks)
                y0 = min(b[1] for b in blocks)
                x1 = max(b[2] for b in blocks)
                y1 = max(b[3] for b in blocks)
                bbox = (x0, y0, x1, y1)
        except Exception:
            pass

    if bbox is None:
        return None

    x0 = int(bbox[0] * mat[0])
    y0 = int(bbox[1] * mat[1])
    x1 = int(bbox[2] * mat[0])
    y1 = int(bbox[3] * mat[1])
    pad = 20
    x0 = max(0, x0 - pad)
    y0 = max(0, y0 - pad)
    x1 = min(pix.width, x1 + pad)
    y1 = min(pix.height, y1 + pad)

    img = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
    return img.crop((x0, y0, x1, y1))






# ---------- 步骤 5: 主选择界面 ----------

class MainSelector:
    """主界面：选择文件 → 点击开始 → 实时日志 → 处理完成弹窗"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title('工资条批量生成工具')
        self.root.resizable(False, False)

        self.excel_path: Optional[Path] = None
        self.word_path: Optional[Path] = None
        self.output_format = tk.StringVar(value='png')  # png / pdf / both
        self.output_dir: Path = _get_app_dir() / OUTPUT_DIR  # 默认输出目录
        self.processing = False   # 防止重复点击

        self._build_ui()

        # 居中
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f'+{(sw - w) // 2}+{(sh - h) // 2}')
        self.root.focus_force()
        self.root.protocol('WM_DELETE_WINDOW', self._on_close)

    # ------------------------------------------------------------------
    # UI 构建
    # ------------------------------------------------------------------

    def _build_ui(self):
        """构建主界面布局"""

        # ---- 标题 ----
        title_frame = tk.Frame(self.root, bg='#2c3e50', height=48)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)
        tk.Label(title_frame, text='工资条批量生成工具', font=('Microsoft YaHei', 13, 'bold'),
                 fg='white', bg='#2c3e50').pack(expand=True)

        # ---- 配置区域（固定上半部分）----
        config_frame = tk.Frame(self.root)
        config_frame.pack(fill=tk.X, padx=16, pady=(16, 0))

        # Excel 文件选择
        section1 = tk.LabelFrame(config_frame, text=' Excel 数据文件 ', padx=10, pady=10)
        section1.pack(fill=tk.X)
        row1 = tk.Frame(section1)
        row1.pack(fill=tk.X)
        tk.Label(row1, text='文件路径:', width=8, anchor=tk.E).pack(side=tk.LEFT)
        self.excel_var = tk.StringVar(value='（未选择）')
        excel_entry = tk.Entry(row1, textvariable=self.excel_var, state='readonly', width=50)
        excel_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 8))
        tk.Button(row1, text='选择文件', width=10,
                  command=self._choose_excel).pack(side=tk.LEFT)

        # Word 文件选择
        section2 = tk.LabelFrame(config_frame, text=' Word 模板文件 ', padx=10, pady=10)
        section2.pack(fill=tk.X, pady=(10, 0))
        row2 = tk.Frame(section2)
        row2.pack(fill=tk.X)
        tk.Label(row2, text='文件路径:', width=8, anchor=tk.E).pack(side=tk.LEFT)
        self.word_var = tk.StringVar(value='（未选择）')
        word_entry = tk.Entry(row2, textvariable=self.word_var, state='readonly', width=50)
        word_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 8))
        tk.Button(row2, text='选择文件', width=10,
                  command=self._choose_word).pack(side=tk.LEFT)

        # 输出格式
        fmt_section = tk.LabelFrame(config_frame, text=' 输出格式 ', padx=10, pady=10)
        fmt_section.pack(fill=tk.X, pady=(10, 0))
        fmt_row = tk.Frame(fmt_section)
        fmt_row.pack(fill=tk.X)
        tk.Label(fmt_row, text='格式:', width=8, anchor=tk.E).pack(side=tk.LEFT)
        tk.Radiobutton(fmt_row, text='PNG 图片', variable=self.output_format,
                       value='png').pack(side=tk.LEFT, padx=(8, 4))
        tk.Radiobutton(fmt_row, text='PDF 文件', variable=self.output_format,
                       value='pdf').pack(side=tk.LEFT, padx=4)
        tk.Radiobutton(fmt_row, text='两者都生成', variable=self.output_format,
                       value='both').pack(side=tk.LEFT, padx=4)

        # 输出目录
        out_section = tk.LabelFrame(config_frame, text=' 输出目录 ', padx=10, pady=10)
        out_section.pack(fill=tk.X, pady=(10, 0))
        out_row = tk.Frame(out_section)
        out_row.pack(fill=tk.X)
        tk.Label(out_row, text='保存至:', width=8, anchor=tk.E).pack(side=tk.LEFT)
        self.out_dir_var = tk.StringVar(value=str(self.output_dir))
        out_entry = tk.Entry(out_row, textvariable=self.out_dir_var, state='readonly', width=50)
        out_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 8))
        tk.Button(out_row, text='选择目录', width=10,
                  command=self._choose_output_dir).pack(side=tk.LEFT)

        # 开始按钮
        btn_frame = tk.Frame(config_frame)
        btn_frame.pack(pady=(16, 4))
        self.start_btn = tk.Button(btn_frame, text='开始生成', width=20, height=2,
                                    bg='#27ae60', fg='white', activebackground='#219a52',
                                    activeforeground='white', bd=0,
                                    command=self._on_start)
        self.start_btn.pack()

        # ---- 日志区域（占据剩余空间，可滚动）----
        log_frame = tk.LabelFrame(self.root, text=' 处理日志 ', padx=4, pady=4)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=(8, 12))

        self.log_text = scrolledtext.ScrolledText(
            log_frame, wrap=tk.WORD, state='disabled',
            font=('Consolas', 9), bg='#1e1e1e', fg='#d4d4d4',
            insertbackground='white', relief=tk.FLAT,
            height=12,
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 配置 text tag 用于不同颜色
        self.log_text.tag_configure('info', foreground='#569cd6')
        self.log_text.tag_configure('ok', foreground='#6a9955')
        self.log_text.tag_configure('fail', foreground='#f44747')
        self.log_text.tag_configure('header', foreground='#4ec9b0')
        self.log_text.tag_configure('summary', foreground='#dcdcaa')

        # ---- 提示信息（底部）----
        tip_label = tk.Label(self.root, text='提示：请确保 Excel 列标题与 Word 模板中的占位符名称一致',
                             fg='#888', font=('Microsoft YaHei', 8))
        tip_label.pack(pady=(0, 8))

    # ------------------------------------------------------------------
    # 文件选择
    # ------------------------------------------------------------------

    def _choose_excel(self):
        path = filedialog.askopenfilename(
            parent=self.root,
            title='选择 Excel 数据文件',
            filetypes=[('Excel 文件', '*.xlsx *.xls'), ('所有文件', '*.*')],
        )
        if path:
            self.excel_path = Path(path)
            self.excel_var.set(path)

    def _choose_word(self):
        path = filedialog.askopenfilename(
            parent=self.root,
            title='选择 Word 模板文件',
            filetypes=[('Word 文档', '*.docx'), ('所有文件', '*.*')],
        )
        if path:
            self.word_path = Path(path)
            self.word_var.set(path)

    def _choose_output_dir(self):
        """选择输出根目录"""
        path = filedialog.askdirectory(
            parent=self.root,
            title='选择输出目录',
            initialdir=str(self.output_dir),
        )
        if path:
            self.output_dir = Path(path)
            self.out_dir_var.set(path)

    # ------------------------------------------------------------------
    # 日志输出
    # ------------------------------------------------------------------

    def _log(self, msg: str, tag: str = 'info'):
        """线程安全地追加日志到文本框（支持在非 GUI 线程中调用）"""
        def _append():
            self.log_text.configure(state='normal')
            self.log_text.insert(tk.END, msg + '\n', tag)
            self.log_text.see(tk.END)            # 自动滚动到底部
            self.log_text.configure(state='disabled')
        self.root.after(0, _append)

    def _clear_log(self):
        """清空日志区域"""
        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', tk.END)
        self.log_text.configure(state='disabled')

    # ------------------------------------------------------------------
    # 开始处理
    # ------------------------------------------------------------------

    def _on_start(self):
        """校验输入 → 启动后台处理线程"""
        if self.processing:
            return   # 正在处理中，忽略重复点击

        # 校验
        if self.excel_path is None:
            messagebox.showwarning('提示', '请先选择 Excel 数据文件', parent=self.root)
            return
        if not self.excel_path.exists():
            messagebox.showerror('错误', f'Excel 文件不存在:\n{self.excel_path}', parent=self.root)
            return
        if self.excel_path.suffix.lower() not in ('.xlsx', '.xls'):
            messagebox.showerror('错误', '请选择 .xlsx 或 .xls 格式的 Excel 文件', parent=self.root)
            return
        if self.word_path is None:
            messagebox.showwarning('提示', '请先选择 Word 模板文件', parent=self.root)
            return
        if not self.word_path.exists():
            messagebox.showerror('错误', f'Word 文件不存在:\n{self.word_path}', parent=self.root)
            return
        if self.word_path.suffix.lower() != '.docx':
            messagebox.showerror('错误', '请选择 .docx 格式的 Word 模板文件', parent=self.root)
            return

        # 设置处理状态
        self.processing = True
        self.start_btn.configure(text='处理中...', bg='#888', state='disabled')
        self._clear_log()

        self._log('=' * 50, 'header')
        self._log('  工资条批量生成工具', 'header')
        self._log('=' * 50, 'header')
        self._log(f'  Excel 数据: {self.excel_path}')
        self._log(f'  Word 模板:  {self.word_path}')
        fmt_label = {'png': 'PNG 图片', 'pdf': 'PDF 文件', 'both': 'PNG + PDF'}
        self._log(f'  输出格式:    {fmt_label.get(self.output_format.get(), self.output_format.get())}')
        self._log('-' * 50, 'header')
        self._log('')

        # 在后台线程中执行处理，不阻塞 UI
        excel = self.excel_path
        word = self.word_path
        fmt = self.output_format.get()
        out_dir = self.output_dir
        threading.Thread(
            target=self._do_process,
            args=(excel, word, fmt, out_dir),
            daemon=True,
        ).start()

    # ------------------------------------------------------------------
    # 后台处理逻辑
    # ------------------------------------------------------------------

    def _do_process(self, excel_path: Path, template_path: Path, output_format: str, output_base: Path):
        """在后台线程中执行工资条生成"""
        converter = None
        temp_dir = None
        success = 0
        failed = 0

        try:
            # --- 输出目录（使用用户选择）---
            base_output = output_base
            base_output.mkdir(parents=True, exist_ok=True)

            png_dir = base_output / 'png' if output_format in ('png', 'both') else None
            pdf_dir = base_output / 'pdf' if output_format in ('pdf', 'both') else None

            if png_dir:
                png_dir.mkdir(exist_ok=True)
            if pdf_dir:
                pdf_dir.mkdir(exist_ok=True)

            temp_dir = Path(tempfile.mkdtemp(prefix='salary_'))

            # --- 读取数据 ---
            self._log(f'[INFO] 读取 Excel: {excel_path}', 'info')
            df = read_excel(str(excel_path))
            self._log(f'  共 {len(df)} 条记录')
            self._log(f'  列标题: {", ".join(df.columns.tolist())}')
            self._log('')

            # --- 创建转换器 ---
            converter = DocxToPdfConverter()
            tool_label = {'word_com': 'Word', 'wps': 'WPS', 'libreoffice': 'LibreOffice'}
            self._log(f'[INFO] 使用转换工具: {tool_label.get(converter._tool_name, converter._tool_name)}', 'info')
            self._log('')

            # --- 逐行处理 ---
            for idx, row in df.iterrows():
                row_dict = row.to_dict()
                name = str(row_dict.get('姓名', f'row{idx}')).strip()
                emp_id = str(row_dict.get('工号', '')).strip()
                base_name = f'{name}_{emp_id}' if emp_id else name
                safe_name = f'{idx+1}_{base_name}'
                safe_name = re.sub(r'[\\/:*?"<>|]', '_', safe_name)

                self._log(f'[{idx+1}/{len(df)}] 处理: {safe_name}')

                try:
                    # 填充模板
                    filled_docx = str(temp_dir / f'{safe_name}.docx')
                    fill_template(str(template_path), row_dict, filled_docx)

                    # DOCX → PDF
                    pdf_temp = str(temp_dir / f'{safe_name}.pdf')
                    converter.convert(filled_docx, pdf_temp)

                    # 按格式输出
                    output_files = []
                    if pdf_dir:
                        pdf_out = str(pdf_dir / f'{safe_name}.pdf')
                        shutil.copy2(pdf_temp, pdf_out)
                        output_files.append(pdf_out)
                    if png_dir:
                        png_out = str(png_dir / f'{safe_name}.png')
                        pdf_to_png(pdf_temp, png_out, CROP_MODE)
                        output_files.append(png_out)

                    self._log(f'  [OK] → {", ".join(output_files)}', 'ok')
                    success += 1

                except Exception as e:
                    self._log(f'  [FAIL] {e}', 'fail')
                    failed += 1

            # --- 关闭转换器 ---
            if converter:
                converter.close()

            # --- 清理 ---
            if temp_dir and DELETE_TEMP:
                shutil.rmtree(temp_dir, ignore_errors=True)

            # --- 结果汇总 ---
            self._log('')
            self._log('===== 完成 =====', 'header')
            self._log(f'成功: {success} 条  |  失败: {failed} 条', 'summary')
            if png_dir:
                self._log(f'PNG 目录: {png_dir.absolute()}')
            if pdf_dir:
                self._log(f'PDF 目录: {pdf_dir.absolute()}')
            if not png_dir and not pdf_dir:
                self._log(f'输出目录: {base_output.absolute()}')

            # 弹窗显示结果（在主线程中）
            self.root.after(0, lambda: self._show_result(success, failed, base_output, png_dir, pdf_dir))

        except Exception as e:
            # 异常时清理并弹窗
            if converter:
                try:
                    converter.close()
                except Exception:
                    pass
            if temp_dir and DELETE_TEMP:
                shutil.rmtree(temp_dir, ignore_errors=True)
            self._log(f'[ERROR] {e}', 'fail')
            self.root.after(0, lambda: self._show_error(str(e)))

        finally:
            # 恢复按钮状态（在主线程中）
            self.root.after(0, self._reset_ui)

    # ------------------------------------------------------------------
    # 结果弹窗 & UI 恢复
    # ------------------------------------------------------------------

    def _show_result(self, success, failed, base_output, png_dir, pdf_dir):
        """处理完成弹窗"""
        lines = ['处理完成！', '']
        lines.append(f'成功: {success} 条')
        lines.append(f'失败: {failed} 条')
        if png_dir:
            lines.append(f'PNG 目录: {png_dir.absolute()}')
        if pdf_dir:
            lines.append(f'PDF 目录: {pdf_dir.absolute()}')
        if not png_dir and not pdf_dir:
            lines.append(f'输出目录: {base_output.absolute()}')

        if failed > 0:
            messagebox.showwarning('工资条生成结果', '\n'.join(lines), parent=self.root)
        else:
            messagebox.showinfo('工资条生成结果', '\n'.join(lines), parent=self.root)

    def _show_error(self, err_msg: str):
        """错误弹窗"""
        messagebox.showerror('错误', f'程序运行出错:\n{err_msg}', parent=self.root)

    def _reset_ui(self):
        """恢复按钮和状态"""
        self.processing = False
        self.start_btn.configure(text='开始生成', bg='#27ae60', state='normal')

    # ------------------------------------------------------------------
    # 启动 / 退出
    # ------------------------------------------------------------------

    def _on_close(self):
        """窗口关闭"""
        if self.processing:
            if not messagebox.askyesno('确认', '正在处理中，确定要退出吗？', parent=self.root):
                return
        self.root.destroy()

    def run(self):
        """运行主界面，阻塞直到窗口关闭"""
        self.root.mainloop()


# ---------- 步骤 6: 主入口 ----------

def main():
    """启动主界面，阻塞直到用户关闭窗口"""
    selector = MainSelector()
    selector.run()


if __name__ == '__main__':
    main()
