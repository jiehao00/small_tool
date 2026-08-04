# -*- coding: utf-8 -*-
"""使用说明页面（可折叠）"""

import tkinter as tk


def create_page(parent, app):
    bg = app.COLOR_CONTENT_BG
    frame = tk.Frame(parent, bg=bg)

    tk.Label(frame, text="★ 使用说明", font=("Microsoft YaHei", 16, "bold"),
             bg=bg, fg="#2c3e50").pack(anchor=tk.W, pady=(0, 12))

    tk.Frame(frame, height=1, bg="#e8ecf0").pack(fill=tk.X, pady=(0, 14))

    # ---- Canvas + Scrollbar 滚动区域 ----
    canvas = tk.Canvas(frame, bg=bg, highlightthickness=0, relief="flat")
    scrollbar = tk.Scrollbar(frame, orient=tk.VERTICAL, command=canvas.yview)
    scrollable = tk.Frame(canvas, bg=bg)

    canvas.create_window((0, 0), window=scrollable, anchor=tk.NW, tags="inner")
    canvas.configure(yscrollcommand=scrollbar.set)

    _reflow_scheduled = {"pending": False}

    def _reflow():
        """统一刷新：同步内容高度 + 按需显示滚动条"""
        canvas.update_idletasks()
        width = canvas.winfo_width()
        if width <= 1:
            # Canvas 尚未渲染，稍后重试
            return
        req_height = scrollable.winfo_reqheight()
        canvas.itemconfig("inner", width=width, height=req_height)
        canvas.configure(scrollregion=(0, 0, width, req_height))
        view_height = canvas.winfo_height()
        if req_height <= view_height:
            # 内容未超出，隐藏滚动条并回到顶部
            if scrollbar.winfo_ismapped():
                scrollbar.pack_forget()
            canvas.yview_moveto(0)
        else:
            if not scrollbar.winfo_ismapped():
                scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _schedule_reflow():
        """合并多次刷新请求，避免重复计算"""
        if _reflow_scheduled["pending"]:
            return
        _reflow_scheduled["pending"] = True
        def _do():
            _reflow_scheduled["pending"] = False
            _reflow()
        frame.after_idle(_do)

    # Canvas 尺寸变化时（页面显示、窗口拉伸）重新评估
    def _on_canvas_configure(event):
        _schedule_reflow()
    canvas.bind("<Configure>", _on_canvas_configure)

    # 鼠标滚轮支持
    def _on_mousewheel(event):
        # 仅在内容超出可视区域时才滚动
        if scrollable.winfo_reqheight() > canvas.winfo_height():
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    def _bind_wheel(e):
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
    def _unbind_wheel(e):
        canvas.unbind_all("<MouseWheel>")
    canvas.bind("<Enter>", _bind_wheel)
    canvas.bind("<Leave>", _unbind_wheel)

    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # ---- 折叠面板数据 ----
    sections = [
        ("业绩生成", [
            ("适用场景", "导入 BBC 原始数据 Excel，按营业区→营业部→业务员三级层级汇总，生成格式化的业绩报表。"),
            ("报表结构",
             "生成的报表包含 12 列：\n"
             "  A 营业区名称\n"
             "  B 月度目标\n"
             "  C 营业区昨日累计业绩\n"
             "  D 营业区本月累计业绩\n"
             "  E 月度达成率\n"
             "  F 营业部名称\n"
             "  G 营业部昨日出单业绩\n"
             "  H 营业部本月出单业绩\n"
             "  I 业务员编码\n"
             "  J 姓名\n"
             "  K 昨日出单业绩\n"
             "  L 本月累计业绩\n"
             "· 月度目标从基础数据第二个Sheet读取\n"
             "· 月度达成率 = 营业区本月累计业绩 ÷ 月度目标 × 100%\n"
             "其中营业区和营业部级别的汇总数据会自动合并单元格并着色，层级清晰。"),
            ("操作步骤",
             "1. 点击「浏览」选择待处理的原始数据 Excel 文件\n"
             "2. 文件加载后，右侧会显示文件基本信息（行数、列数、工作表名等）\n"
             "3. 选择「基础数据」文件（必传）：\n"
             "   Sheet1：补入清单中有但原始数据中无业绩的人员（显示零业绩）\n"
             "   Sheet2：提供各营业区的「月度目标」数据（表头：营业区名称、月度目标）\n"
             "4. （可选）启用「排除条件」：选择某列，当该列的值等于指定值时跳过该行\n"
             '   例如：排除"缴费状态"="退保"的行，可从下拉框中选择列名和值\n'
             '5. 自定义「报表标题」（默认"BBC业绩报表"）\n'
             "6. 设置「输出目录」和「文件名前缀」\n"
             "   输出文件名会自动追加日期，格式：前缀_YYYY-MM-DD.xlsx\n"
             "7. 在「输出格式」中勾选需要的格式（Excel / PNG / PDF，支持多选）\n"
             "   勾选 PNG 或 PDF 时会自动转换（需 Office 软件支持）\n"
             "8. 点击「⚡ 开始处理」"),
            ("关于昨日数据", "昨日的业绩和件数根据当日数据自动计算：\n• 昨日标保：所有记录标保的合计\n• 昨日件数：按保单号去重计数，同一保单号只计 1 件\n营业部级别的昨日数据基于该部下所有业务员的昨日数据汇总。"),
            ("注意事项",
             "• Excel 文件的第一行必须是表头（列名），程序会自动跳过空行定位表头\n"
             "• 目前支持 .xlsx 格式（.xls 文件请先用 Excel 另存为 .xlsx）\n"
             "• 原始文件不会被修改，结果另存为新文件\n"
             "• 导出 PDF/PNG 需要电脑安装 Microsoft Excel 或 WPS Office\n"
             "  （程序会自动检测并优先使用 Excel，其次 WPS）\n"
             "• 如人员清单中某人员在原始数据中不存在，将自动生成零业绩行"),
        ]),
        ("工资条生成", [
            ("适用场景", "根据汇总表和明细表，按人员拆分生成单独的工资条文件，每人一个独立 Excel。"),
            ("操作步骤",
             "1. 选择汇总表文件（包含每个人员的基本信息和匹配键值）\n"
             "2. 选择明细表文件（包含每个人员的详细薪酬数据）\n"
             "3. 设置「汇总对比列」（汇总表中用于匹配的列名，如：工号）\n"
             "4. 设置「明细对比列」（明细表中用于匹配的列名，如：业务员代码）\n"
             "5. （可选）设置「明细剔除列」和「剔除列值」\n"
             "   当明细中某列的值为剔除值时，该行将被跳过（如剔除实发佣金=0的行）\n"
             "6. 选择输出目录\n"
             "7. 点击「⚡ 开始生成」"),
            ("输出说明",
             "• 每人生成一个 .xlsx 文件，命名格式：序号_姓名_工号.xlsx\n"
             "• 每个文件包含：汇总表头 → 汇总数据行 → 各明细 Sheet 区块\n"
             "• 不同明细 Sheet 用不同底色区分，便于阅读\n"
             "• 长数字（如身份证号、银行卡号）自动转为文本，避免科学计数法"),
            ("注意事项",
             "• 两个文件的第一行都必须是表头（列名）\n"
             "• 对比列的列名必须与文件中的表头完全一致（区分大小写）\n"
             "• 汇总表和明细表都支持包含多个 Sheet，程序会遍历所有 Sheet 进行匹配\n"
             "• 支持 .xls 和 .xlsx 两种格式混用"),
        ]),
    ]

    # ---- 渲染折叠面板 ----
    COLOR_HEADER_BG = "#f0f3f5"
    COLOR_HEADER_FG = "#2c3e50"
    COLOR_BODY_BG = "#fafbfc"
    COLOR_ACCENT = "#3498db"
    COLOR_BORDER = "#dce0e4"
    COLOR_DESC = "#555"

    toggle_state = {}  # section_index -> BooleanVar

    for idx, (title, items) in enumerate(sections):
        visible = tk.BooleanVar(value=(idx == 0))  # 默认展开第一个
        toggle_state[idx] = visible

        # --- 标题栏 ---
        header = tk.Frame(scrollable, bg=COLOR_HEADER_BG, cursor="hand2",
                          padx=12, pady=10)
        header.pack(fill=tk.X, padx=0, pady=(8 if idx > 0 else 0, 0))

        # 左侧蓝色竖条
        bar = tk.Frame(header, bg=COLOR_ACCENT, width=4)
        bar.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))

        # 折叠图标
        icon = tk.Label(header, text="▼", font=("Microsoft YaHei", 10, "bold"),
                        bg=COLOR_HEADER_BG, fg=COLOR_ACCENT, width=2)
        icon.pack(side=tk.LEFT)

        # 标题文字
        tk.Label(header, text=title, font=("Microsoft YaHei", 12, "bold"),
                 bg=COLOR_HEADER_BG, fg=COLOR_HEADER_FG).pack(side=tk.LEFT)

        # 分隔线
        sep = tk.Frame(scrollable, height=1, bg=COLOR_BORDER)
        sep.pack(fill=tk.X, padx=0)

        # 内容体：默认第一个展开，其余不打包（避免影响初始高度）
        body = tk.Frame(scrollable, bg=COLOR_BODY_BG, padx=16, pady=12)
        if idx == 0:
            body.pack(fill=tk.X, padx=0, after=sep)
            icon.config(text="▼")
        else:
            icon.config(text="▶")

        for label, desc in items:
            row = tk.Frame(body, bg=COLOR_BODY_BG)
            row.pack(fill=tk.X, pady=(0, 8))
            tk.Label(row, text=label + "：", font=("Microsoft YaHei", 10, "bold"),
                     bg=COLOR_BODY_BG, fg="#2c3e50",
                     width=8, anchor=tk.NW, justify=tk.LEFT).pack(side=tk.LEFT)
            tk.Label(row, text=desc, font=("Microsoft YaHei", 10),
                     bg=COLOR_BODY_BG, fg=COLOR_DESC,
                     anchor=tk.W, justify=tk.LEFT, wraplength=580).pack(
                         side=tk.LEFT, fill=tk.X, expand=True)

        # 底部留白
        tk.Frame(body, bg=COLOR_BODY_BG, height=2).pack(fill=tk.X)

        # --- 折叠/展开逻辑 ---
        def make_toggle(v, ic, bd, sp):
            def toggle(e=None):
                if v.get():
                    bd.pack_forget()
                    ic.config(text="▶")
                    v.set(False)
                else:
                    bd.pack(fill=tk.X, padx=0, after=sp)
                    ic.config(text="▼")
                    v.set(True)
                # 同步高度并按需显示滚动条
                _reflow()
            return toggle

        toggle_fn = make_toggle(visible, icon, body, sep)
        # 点击标题栏任意位置均可触发展开/折叠
        header.bind("<Button-1>", toggle_fn)
        for child in header.winfo_children():
            child.bind("<Button-1>", toggle_fn)

        # hover 颜色变化
        def make_hover(hd):
            def on_enter(e):
                hd.config(bg="#e4e9ed")
                for child in hd.winfo_children():
                    try:
                        child.config(bg="#e4e9ed")
                    except tk.TclError:
                        pass
            def on_leave(e):
                hd.config(bg=COLOR_HEADER_BG)
                for child in hd.winfo_children():
                    try:
                        child.config(bg=COLOR_HEADER_BG)
                    except tk.TclError:
                        pass
            hd.bind("<Enter>", on_enter)
            hd.bind("<Leave>", on_leave)
        make_hover(header)

    # 底部留白
    tk.Frame(scrollable, bg=bg, height=20).pack()

    # 初始布局完成后，调度一次刷新（等 Canvas 真正渲染后执行）
    _schedule_reflow()

    return frame
