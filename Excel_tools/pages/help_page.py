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
        ("功能一：填充数据（Excel 内容填充）", [
            ("适用场景", "A 表需要补充数据，B 表是数据源，根据共同列匹配后自动填充。"),
            ("操作步骤",
             "1. 选择 A 文件（需要被填充的表）— 列名自动加载\n"
             "2. 选择 B 文件（数据来源表）— 列名自动加载\n"
             "3. 点击匹配列旁的 📋 按钮，在弹出的列名列表中勾选两表的匹配列\n"
             "   （如 A[编号] ↔ B[ID]），Ctrl+点击多选，按选择顺序配对\n"
             "4. 同样点击填充列旁的 📋 按钮，勾选 A 表哪列要填、B 表哪列是数据来源\n"
             "5. 设置输出文件路径，点击「开始转换」"),
            ("匹配规则",
             "• 当 A 文件的匹配列值与 B 文件的匹配列值一致时，认为两行匹配\n"
             "• 只有当 A 文件的填充列为空时，才会用 B 文件的内容覆盖\n"
             "• 支持多列组合匹配/填充，Ctrl+点击按顺序选中即可，无需手动输入逗号\n"
             "  例：在两列勾选框中分别选中「编号,姓名」，则两列都相等才算匹配成功"),
        ]),
        ("功能二：数据对比（交集差集对比）", [
            ("适用场景", "A 表与 B 表进行差异分析，找出独有/共有数据。"),
            ("操作步骤",
             "1. 选择 A 文件和 B 文件，列名自动加载，右侧显示列数\n"
             "2. 点击 📋 按钮，在弹出的列名列表中 Ctrl+点击勾选关联列\n"
             "   如 A📋 → 勾选 [编号, 姓名]   B📋 → 勾选 [ID, Name]\n"
             "   选中的列名自动填入文本框，按选择顺序一一配对\n"
             "3. 核对两侧列数一致后，设置输出路径，点击「开始对比」"),
            ("输出结果（一个 Excel 包含三个 Sheet）",
             "• 「合并结果」：两表的交集，A 表列 + B 表独有列\n"
             "• 「仅A有」：A 表中有但 B 表中没有的行\n"
             "• 「仅B有」：B 表中有但 A 表中没有的行"),
        ]),
        ("功能三：财务对账（汇总表 vs 明细累加对比）", [
            ("适用场景", "有汇总表和明细表，需要核对汇总金额是否等于明细累加值。\n"
             "如：财务总账金额 vs 发票明细加总、业务汇总 vs 流水明细等。"),
            ("操作步骤",
             "1. 选择汇总表（A 文件）和明细表（B 文件），列名自动加载\n"
             "2. 点击主键列旁的 📋 按钮，分别选择两表的匹配列\n"
             "   如 汇总[发票号] ↔ 明细[发票号]，Ctrl+点击可多选复合主键\n"
             "3. 点击对比列旁的 📋 按钮，分别选择两表要核对的数值列\n"
             "   如 汇总[金额, 税额] ↔ 明细[金额, 税额]，两侧列数须一致\n"
             "4. 设置输出文件路径，点击「开始对账」"),
            ("处理逻辑",
             "• 明细表先按主键分组，对比列自动求和（SUM）\n"
             "• 汇总表逐行与明细累加值比对，不一致则整行标黄色高亮\n"
             "• 汇总有但明细无的记录记录在日志中\n"
             "• 明细有但汇总无的记录也会在日志中提示"),
        ]),
        ("功能四：发票验重 & 连号检测", [
            ("适用场景", "扫描发票号码列，检测重复报销或连号发票风险。\n"
             "审计、财务审核时快速定位异常发票。"),
            ("操作步骤",
             "1. 浏览选择包含发票数据的 Excel 文件，列名自动加载\n"
             "2. 在列表框中 Ctrl+点击 选择比对依据列\n"
             "3. 勾选需要的检测项（可同时勾选）\n"
             "4. 设置输出路径，点击「开始检测」"),
            ("输出结果（一个 Excel 包含三个 Sheet）",
             "• 「检测结果」：全部数据保持原序，末尾追加验重结果、连号结果等标记列\n"
             "• 「重复明细」：仅重复行，按号码分组集中展示\n"
             "• 「连号明细」：仅连号行，按连号组集中，组内号码升序排列，相邻组交替背景色便于区分"),
            ("重要说明",
             "• 重复/连号检测都基于所选列的组合值，逻辑统一\n"
             "• 验重：多列组合值完全相同时判定为重复或连号\n"
             "• 连号检测提取组合值中最长连续数字段进行比较\n"
             "• 建议：只勾选发票号码列做连号检测，加选商品列做精准验重\n"
             "• 连号阈值默认 ≥2 张，可自行调整（如设为 ≥3 张，则两连号不标出）"),
        ]),
        ("功能五：数据合并（单表按主键去重 + 聚合统计）", [
            ("适用场景", "一个 Excel 表中存在相同主键的多行数据，需要按主键合并并统计数值列。"),
            ("操作步骤",
             "1. 点击「浏览」选择要处理的 Excel 文件（仅支持 .xlsx）\n"
             "2. 在主键列列表中勾选合并依据的列（支持 Ctrl/Shift 多选作为复合主键）\n"
             "3. 勾选需要的聚合方式（求和 / 计数 / 平均值 / 最大值 / 最小值，默认求和）\n"
             "4. 设置输出文件保存路径\n"
             "5. 点击「开始合并」"),
            ("说明",
             "• 相同主键的行，数值列按勾选方式汇总（如\"金额_求和\"\"数量_计数\"）\n"
             "• 非数值列取对应主键组中第一行的值\n"
             "• 如果没有重复主键，将原样输出"),
        ]),
        ("功能六：数据拆分（拆分列→多行）", [
            ("适用场景", "将某列中按分隔符存放的多个值，拆分成多行独立数据。"),
            ("操作步骤",
             "1. 点击「浏览...」选择要处理的 Excel 文件\n"
             "2. 在左侧表格中勾选需要拆分的列（可多选）\n"
             "3. 选择分隔符（逗号、分号、空格等），也可输入自定义分隔符\n"
             "4. 设置输出文件路径，点击「开始拆分」"),
            ("拆分规则",
             "• 选中列的单元格内容按分隔符切分后，每个值生成一行新数据\n"
             "• 其余列的内容原样复制到每一行\n"
             "• 例：单元格 \"张三,李四,王五\"，按逗号拆分后变成三行"),
        ]),
        ("功能七：多文件追加合并（纵向拼接）", [
            ("适用场景", "多个结构相似或不同的 Excel 文件，按行纵向拼成一个总表。"),
            ("操作步骤",
             "1. 点击「添加文件」选择一个或多个 Excel 文件\n"
             "2. 在列表中拖拽调整文件顺序（按从上到下的顺序拼接）\n"
             "3. 选择列对齐方式：取并集（保留所有列）/ 取交集（仅共有列）\n"
             "4. 勾选「每文件首行为列名」则按列名对齐，否则按位置拼接\n"
             "5. 设置输出文件路径，点击「开始追加合并」"),
            ("输出说明",
             "• 末尾自动添加「_来源文件」列，记录每行数据来源于哪个文件\n"
             "• 取并集时，某文件没有的列自动留空\n"
             "• 取交集时，仅保留所有文件共有的列"),
        ]),
        ("功能八：按列值拆分文件（一个组合 → 一个文件）", [
            ("适用场景", "根据一列或多列的组合取值，将 Excel 数据拆分为多个独立文件。\n"
             "如：按「省份」拆分，生成 北京.xlsx、上海.xlsx…\n"
             "或按「省份+年份」组合，生成 北京_2023.xlsx、北京_2024.xlsx…"),
            ("操作步骤",
             "1. 点击「浏览」选择一个 Excel 文件 — 列名自动加载到列表\n"
             "2. 在列名列表中选中拆分依据列（可按住 Ctrl 多选，如同时选：省份、年份）\n"
             "3. 选择输出目录\n"
             "4. 勾选输出格式：Excel / PNG / PDF，可多选（点击「全选 / 全不选」快捷切换）\n"
             "5. 设置文件名前缀，点击「开始拆分」"),
            ("输出说明",
             "• 根据所选格式，在输出目录下分别创建 excel / png / pdf 子目录\n"
             "• 每个唯一组合值生成一个文件，命名格式：前缀 + 组合值(下划线拼接) + 对应扩展名\n"
             "• 空值统一归入「(空值)」\n"
             "• 非法文件名（含 \\/:*?\"<>|）会自动替换为下划线"),
        ]),
        ("功能九：条件筛选提取（按条件过滤数据）", [
            ("适用场景", "按自定义条件从 Excel 中筛选符合要求的行，输出指定列的新文件。\n"
             "如：提取金额>1000 的记录、筛选某省份的数据、过滤未填写的行等。"),
            ("操作步骤",
             "1. 浏览选择待筛选的 Excel 文件，列名自动加载\n"
             "2. 在「输出列」区域勾选要输出的列（默认全选，可全选/全不选快捷切换）\n"
             "3. 点击「＋ 添加条件」，逐个设置筛选规则：\n"
             "   - 选择列名（下拉列表自动列出所有列）\n"
             "   - 选择运算符：=、≠、>、<、>=、<=、包含、不包含、为空、不为空\n"
             "   - 输入比较值（为空/不为空时无需输入值）\n"
             "   - 每个条件左侧可选择 AND/OR，点击 ✕ 可删除条件\n"
             "4. 不添加任何筛选条件时，输出全部行（相当于仅选择输出列）\n"
             "5. 设置输出文件路径，点击「开始筛选」"),
            ("AND/OR 分组规则",
             "• 条件默认用 AND 连接（全部条件同时满足才提取）\n"
             "• 遇到 OR 则开启一组新条件：组内 AND（交集），组间 OR（并集）\n"
             "• 例：[A>10 AND B包含'北京' OR C='是' AND D='否']\n"
             "  含义：(A>10 且 B包含'北京') 或 (C='是' 且 D='否') 的行都提取"),
            ("运算符说明",
             "• = ≠：精确匹配，区分大小写\n"
             "• > < >= <=：先尝试日期比较（支持 2024-01-01 等格式），再尝试数值比较，最后字符串比较\n"
             "• 包含 / 不包含：文本模糊匹配，即值中是否含有输入的内容\n"
             "• 为空：单元格为空白的行；不为空：单元格有内容的行"),
            ("处理逻辑",
             "• 结果另存为新文件，原始文件不受影响\n"
             "• 空条件时输出所有行、指定列，适合快速提取部分列\n"
             "• 日期列自动标准化为 YYYY-MM-DD 进行比较，支持多种日期格式"),
        ]),
        ("通用注意事项", [
            ("文件格式",
             "• Excel 文件的第一行必须是表头（列名），各功能通过列名识别数据列\n"
             "• 目前支持 .xlsx 格式（.xls 文件请先用 Excel 另存为 .xlsx ）\n"
             "• 列名支持中文/英文，无需特殊处理"),
            ("数据安全",
             "• 所有功能均「读取式」处理，原始文件不会被修改或覆盖\n"
             "• 结果默认输出为新文件，保存前请确认路径无误"),
            ("数据规范",
             "• 日期列建议统一使用 YYYY-MM-DD 或 YYYY/MM/DD 格式\n"
             "• 金额列为纯数字时匹配精度最高，避免带单位（如\"元\"）\n"
             "• 大量空白单元格为正常现象，不影响各功能正常运行"),
            ("性能提示",
             "• 处理上万行数据时请稍候，日志框会显示实时进度\n"
             "• 对账/比对功能建议先确保两表行数大致相当，效率更高"),
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
