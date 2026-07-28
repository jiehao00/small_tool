const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, LevelFormat, PageNumber, PageBreak
} = require("docx");

// ── 辅助：创建普通段落 ──
function p(text, opts = {}) {
  const runs = [];
  if (typeof text === "string") {
    runs.push(new TextRun({ text, bold: opts.bold, font: "Microsoft YaHei", size: 21, color: opts.color }));
  } else if (Array.isArray(text)) {
    text.forEach(t => {
      if (typeof t === "string") runs.push(new TextRun({ text: t, font: "Microsoft YaHei", size: 21 }));
      else runs.push(new TextRun({ ...t, font: "Microsoft YaHei", size: t.size || 21 }));
    });
  }
  return new Paragraph({ spacing: { before: opts.before || 60, after: opts.after || 60 }, children: runs, ...(opts.alignment ? { alignment: opts.alignment } : {}) });
}

// ── 辅助：标题 ──
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 120 },
    children: [new TextRun({ text, font: "Microsoft YaHei", size: 32, bold: true, color: "2c3e50" })]
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 80 },
    children: [new TextRun({ text, font: "Microsoft YaHei", size: 26, bold: true, color: "2c3e50" })]
  });
}

// ── 分隔线 ──
function divider() {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "CCCCCC", space: 1 } },
    children: []
  });
}

// ── 文档内容 ──
const children = [
  // ──── 封面标题 ────
  new Paragraph({ spacing: { before: 600 }, children: [] }),
  p("工资条批量生成工具", { bold: true, alignment: AlignmentType.CENTER, before: 0, after: 200 }),
  p("使用说明", { alignment: AlignmentType.CENTER, before: 0, after: 400 }),
  divider(),

  // ── 功能简介 ──
  h2("功能简介"),
  p("本工具可以批量将 Excel 中的员工工资数据，自动填充到 Word 模板中，并导出为每位员工的独立文件。支持三种输出格式："),
  p([{ text: "●   PNG 图片", bold: true }, " — 高清 300 DPI，适合微信/邮件发送"], { before: 40 }),
  p([{ text: "●   PDF 文件", bold: true }, " — 电子文档，适合存档打印"], { before: 40 }),
  p([{ text: "●   PNG + PDF", bold: true }, " — 同时生成两种格式"], { before: 40 }),
  p("文件名格式：{序号}_{姓名}_{工号}.扩展名，示例：1_张三_001.png", { before: 80 }),
  divider(),

  // ── 使用前提 ──
  h2("使用前提"),
  p("本工具为 Windows GUI 程序（无命令行窗口），运行时需要以下条件之一："),
  p([{ text: "●   Microsoft Word", bold: true }, "（推荐，转换速度最快）"], { before: 40 }),
  p([{ text: "●   WPS Office", bold: true }, "（免费，自动作为备选）"], { before: 40 }),
  p([{ text: "●   LibreOffice", bold: true }, "（免费开源）"], { before: 40 }),
  p("检测顺序：Word → WPS → LibreOffice，只要装了任一即可。", { before: 80 }),
  divider(),

  // ── 文件准备 ──
  h2("文件准备"),
  p("工具为图形界面程序，无需手动放置文件在同一目录。运行后弹出窗口，通过按钮选择以下文件："),
  p("1.   Excel 数据文件（.xlsx 或 .xls） — 员工工资表", { before: 40 }),
  p("2.   Word 模板文件（.docx） — 须包含占位符", { before: 40 }),
  divider(),

  // ── 使用步骤 ──
  h2("使用步骤"),

  h1("第一步：准备 Excel 数据文件"),
  p([{ text: "●   " }, "第一行必须是列标题（表头），例如："]),
  p("姓名  |  工号  |  应发工资  |  实发工资  |  FYC  |  ...", { before: 40 }),
  p([{ text: "●   " }, "后续每行为一名员工的数据，每列对应一个字段。"]),
  p([{ text: "●   " }, "工号列建议设为文本格式，避免科学计数法显示。"]),
  p([{ text: "●   " }, "姓名和工号不是必须的，但建议保留\"姓名\"列，便于生成有意义的文件名。"]),

  h1("第二步：制作 Word 模板"),
  p("在 Word 文档中需要展示数据的位置，插入占位符，支持两种格式："),
  p([{ text: "●   方式一（推荐）：" }, "使用 Word 邮件合并域", { bold: true }], { before: 60 }),
  p("    在 Word 中点击「插入」→「文档部件」→「域」→ 选择 MergeField，域名与 Excel 列标题完全一致。", { before: 20 }),
  p([{ text: "●   方式二：" }, "手动输入纯文本占位符", { bold: true }], { before: 60 }),
  p("    格式为 <<字段名>>，例如：<<姓名>>、<<应发工资>>。字段名必须与 Excel 列标题完全一致（区分中英文、空格、标点）。", { before: 20 }),
  p("模板中可以只放需要的字段，不一定要包含 Excel 所有列。", { before: 80 }),

  h1("第三步：运行程序"),
  p("双击\"工资条生成工具.exe\"启动，界面从上到下依次为："),
  p([{ text: "① " }, "Excel 文件选择     ", { bold: true }, "→ 点击「选择文件」选取员工工资表"]),
  p([{ text: "② " }, "Word 模板选择      ", { bold: true }, "→ 点击「选择文件」选取模板文件"]),
  p([{ text: "③ " }, "输出格式选择       ", { bold: true }, "→ 勾选 PNG / PDF / PNG+PDF"]),
  p([{ text: "④ " }, "输出目录选择       ", { bold: true }, "→ 点击「选择目录」自定义保存位置（默认为 output）"]),
  p([{ text: "⑤ " }, "开始生成按钮       ", { bold: true }, "→ 点击后启动处理"]),
  p([{ text: "⑥ " }, "实时日志区域       ", { bold: true }, "→ 深色日志窗口，彩色标注处理状态"]),
  p(""),
  p("点击「开始生成」后，日志区实时显示处理进度："),
  p([{ text: "●   蓝色文字 [INFO]" }, "：读取数据、使用的转换工具"], { before: 40, color: "569cd6" }),
  p([{ text: "●   绿色文字 [OK]" }, "：每条记录生成成功"], { before: 40, color: "6a9955" }),
  p([{ text: "●   红色文字 [FAIL]" }, "：某条记录处理失败及原因"], { before: 40, color: "f44747" }),
  p([{ text: "●   青色/黄色" }, "：标题与汇总信息"], { before: 40, color: "4ec9b0" }),
  p(""),
  p([{ text: "●   " }, "处理期间按钮变灰禁用，防止重复点击，UI 不会卡死。"]),
  p([{ text: "●   " }, "处理完毕后自动弹窗显示结果（成功/失败数量、输出路径）。"]),
  p([{ text: "●   " }, "关闭弹窗后按钮恢复，可继续处理下一批数据。"]),

  h1("第四步：查看结果"),
  p("输出目录结构（以默认 output 为例）："),
  p("output\\", { before: 40 }),
  p("├── png\\          ← PNG 图片", { before: 20 }),
  p("│   ├── 1_张三_001.png", { before: 20 }),
  p("│   └── 2_李四_002.png", { before: 20 }),
  p("└── pdf\\          ← PDF 文件", { before: 20 }),
  p("    ├── 1_张三_001.pdf", { before: 20 }),
  p("    └── 2_李四_002.pdf", { before: 20 }),
  p(""),
  p([{ text: "●   " }, "PNG 分辨率为 300 DPI，清晰度适合打印和发送。"]),
  p([{ text: "●   " }, "PDF 为 Office 原生导出，排版与 Word 一致。"]),

  new Paragraph({ children: [new PageBreak()] }),

  // ── 常见问题 ──
  h2("常见问题"),
  divider(),

  p([{ text: "Q1：运行时报\"未找到可用的 Office 工具\"怎么办？", bold: true }], { before: 120 }),
  p("A：说明电脑上没有安装 Word、WPS 或 LibreOffice 中的任何一个。请安装 WPS Office（免费）或 LibreOffice（免费）即可。"),

  p([{ text: "Q2：生成的图片是空白的，或者缺少数据？", bold: true }], { before: 120 }),
  p("A：请检查 Word 模板中的占位符名称是否与 Excel 列标题完全一致（包括中英文、空格、括号等，必须严格匹配）。"),

  p([{ text: "Q3：图片太宽了，能不能只截取表格区域？", bold: true }], { before: 120 }),
  p("A：可以修改配置（需要源码版本），将 CROP_MODE 改为 \"table\"。EXE 版本默认使用 \"page\"（整页）模式。"),

  p([{ text: "Q4：生成速度太慢怎么办？", bold: true }], { before: 120 }),
  p("A：本工具优先使用 Word COM（速度最快），WPS 作为备选。如果电脑装了 Word，会自动使用 Word 转换。处理人数较多时（如 50+ 人），稍等片刻即可，程序会复用 Office 实例加快速度。"),

  p([{ text: "Q5：支持多少个字段？", bold: true }], { before: 120 }),
  p("A：没有上限，Excel 有多少列就支持多少字段。"),

  p([{ text: "Q6：Excel 中没有\"工号\"列可以吗？", bold: true }], { before: 120 }),
  p("A：可以。文件名格式为 {序号}_{姓名}.扩展名。如果\"姓名\"列也没有，则会显示 row0、row1 等占位名称。"),

  p([{ text: "Q7：正在处理时能关闭窗口吗？", bold: true }], { before: 120 }),
  p("A：点击关闭按钮会弹出确认提示，确认后退出。建议等待处理完成后再关闭。"),

  p([{ text: "Q8：想在多个不同位置输出，需要每次都选目录吗？", bold: true }], { before: 120 }),
  p("A：是的。每次处理前都可以重新选择输出目录，默认记住上次的选择。"),

  divider(),

  // ── 注意事项 ──
  h2("注意事项"),
  p([{ text: "●   " }, "无命令行窗口，所有操作和反馈均在图形界面中完成。"]),
  p([{ text: "●   " }, "文件名中的特殊字符（\\ / : * ? \" < > |）会自动替换为下划线。"]),
  p([{ text: "●   " }, "数字类型字段会自动格式化（整数不显示小数点，小数保留两位）。"]),
  p([{ text: "●   " }, "空单元格在生成时显示为空白。"]),
  p([{ text: "●   " }, "Word 模板建议设置为单页，方便生成单张图片。"]),
  p([{ text: "●   " }, "支持 .xlsx 和 .xls 格式的 Excel 文件。"]),

  divider(),

  // ── 页脚信息 ──
  p(""),
  p("技术支持：如遇问题，可联系程序提供者。", { before: 200 }),
  p("更新日期：2026年7月16日", { before: 40, color: "888888" }),
];

// ── 生成文档 ──
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Microsoft YaHei", size: 21 } }
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Microsoft YaHei", color: "2c3e50" },
        paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Microsoft YaHei", color: "1a5276" },
        paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 1,
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "3498db", space: 4 } } }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },  // A4
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 4 } },
          children: [new TextRun({ text: "工资条批量生成工具 · 使用说明", font: "Microsoft YaHei", size: 18, color: "999999", italics: true })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 4 } },
          children: [
            new TextRun({ text: "第 ", font: "Microsoft YaHei", size: 18, color: "999999" }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Microsoft YaHei", size: 18, color: "999999" }),
            new TextRun({ text: " 页", font: "Microsoft YaHei", size: 18, color: "999999" }),
          ]
        })]
      })
    },
    children,
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("e:/worksun202109-/workboddy_code/CSP_salary/工资条批量生成工具-使用说明.docx", buffer);
  console.log("[OK] 文档已生成: 工资条批量生成工具-使用说明.docx");
});
