# 工资条批量生成工具

根据 Excel 员工数据，自动填充 Word 模板中的 `<<字段名>>` 或 Word 邮件合并域 `«字段名»`，并将填充后的文档导出为高清 PNG 图片。

## 功能

- 读取 Excel 中所有员工行数据
- 复制 Word 模板，逐行替换占位符为实际数值
- 将 Word 页面转为高清图片（默认 300 DPI）
- 批量输出：`{姓名}_{工号}.png`
- 支持纯文本占位符 `<<字段名>>` 和邮件合并域 `«字段名»`
- **0 外部依赖**：模板填充纯 Python 实现，PDF 转换自动检测已安装的 Office

## 文件说明

| 文件 | 说明 |
|------|------|
| `generate_salary_slips.py` | 主程序 |
| `requirements.txt` | Python 依赖 |
| `工资条1.xlsx` | 员工数据源示例 |
| `工资表.docx` | Word 模板示例 |
| `output/` | 输出目录 |

## 安装依赖

```bash
pip install -r requirements.txt
```

## 运行

```bash
python generate_salary_slips.py
```

## PDF 转换引擎（自动检测）

脚本启动时自动按优先级查找可用引擎：

| 优先级 | 引擎 | 说明 |
|--------|------|------|
| 1 | Microsoft Word | 需安装 Office（效果最佳） |
| 2 | WPS Office | 国产免费，自动查找安装路径 |
| 3 | LibreOffice | 免费开源，跨平台 |

只要装了其中任意一个，就能正常使用。

## 打包给他人使用

### 方案 A：便携版 LibreOffice（推荐）

1. 下载 [LibreOffice Portable](https://www.libreoffice.org/download/portable-versions/) 免安装版
2. 解压到项目目录下的 `LibreOffice` 文件夹
3. 使用 PyInstaller 打包：

```bash
pip install pyinstaller
pyinstaller --onefile --add-data "LibreOffice;LibreOffice" generate_salary_slips.py
```

用户拿到 `.exe` 后解压即用，无需安装任何软件。

### 方案 B：要求用户装 WPS

大多数人电脑上已经有 WPS，脚本已自动支持检测。用户只需：

1. 安装 Python 和依赖
2. 运行脚本

---

## 配置

编辑 `generate_salary_slips.py` 顶部的配置区：

```python
EXCEL_PATH = "工资条1.xlsx"          # Excel 数据源
TEMPLATE_PATH = "工资表.docx"        # Word 模板
OUTPUT_DIR = "output"                # 输出目录
IMG_DPI = 300                        # 图片 DPI
CROP_MODE = "page"                   # "page" = 整页 / "table" = 自动截取表格区域
DELETE_TEMP = True                   # 是否删除临时文件
```

## 模板规则

1. Excel 第一行必须为列标题（如 `姓名`、`工号`、`FYC` 等）
2. Word 模板中需要出现与列标题对应的占位符：
   - `<<字段名>>` 纯文本占位符
   - `«字段名»` Word 邮件合并域占位符
3. 工具会自动按行填充并导出为图片
