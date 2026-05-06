# 整体解决方案撰写 Skill

## Skill 名称
`solution-writer` - 整体解决方案撰写专家

## 触发条件
当用户需要：
- 编写投标解决方案文档
- 将需求文档转化为解决方案
- 生成带格式的技术方案文档
- 创建解决方案架构图
- 扩写已有投标文档(增加字数、补充图表)

## 核心能力

### 1. 需求解析能力
- 读取并解析甲方业务需求文档(.docx/.txt/.md)
- 解析评分细则/文档结构要求
- 建立"需求→方案"映射关系
- 覆盖性校验（确保 100% 覆盖原始需求）

### 2. 方案撰写能力
- 严格遵循评分细则章节结构
- 使用售前投标语言（正式、严谨、非口语化）
- 成段描述为主，分点说明为辅
- 每个模块包含：功能说明、实现方式、应用价值
- 支持扩写模式：在原文基础上扩展字数、补充技术细节

### 3. 文档生成能力
- 生成带格式的 Word 文档（.docx）
- **标题使用 Heading 样式**，确保 Word 导航检索可用
- 字体格式：H1 黑体18号、H2 仿宋15号加粗、H3 仿宋14号加粗、H4 仿宋13号加粗、正文仿宋12号
- 行距设置：固定行距 24 磅（EXACTLY 模式）
- 首行缩进：正文首行缩进 24 磅（2 字符）
- 支持图片+居中图题、表格（含表头加粗）
- 支持合并多个章节文档

### 4. 架构图生成能力
- 使用 PIL 生成系统架构图 PNG
- 展示多层架构体系（如六层架构）
- 支持安全体系、数据流等可视化
- 自动插入到 Word 文档对应章节

### 5. 质量校验能力
- 字数统计与达标验证
- 标题层级一致性检查
- Word 导航可用性验证
- 需求覆盖率检查

## 执行流程

```
┌─────────────────────────────────────────────────────────┐
│                   整体解决方案撰写流程                    │
└─────────────────────────────────────────────────────────┘

1. 输入确认
   ├── 原始需求文档路径
   ├── 文档结构/评分细则路径
   ├── 输出文档名称
   └── 目标字数（如 30000）

2. 需求解析（内部执行）
   ├── 功能需求拆解
   ├── 业务流程梳理
   ├── 数据需求分析
   ├── 管理/监管需求
   └── 展示与可视化需求

3. 结构解析（内部执行）
   ├── 章节结构确认
   ├── 每章撰写目标与字数分配
   ├── 内容边界确定
   └── 深度要求确认

4. 方案撰写（单脚本一次性生成）
   ├── 逐章节输出内容
   ├── 100% 映射原始需求
   ├── 使用投标语言
   └── 结构化拆解模块

5. 架构图生成（如需要）
   ├── 识别需要架构图的章节
   ├── 生成架构 PNG 图片
   └── 插入到 Word 文档

6. 质量校验
   ├── 字数统计与达标验证
   ├── 标题层级一致性检查
   ├── 需求覆盖性检查
   └── 格式规范性检查

7. 文档输出
   └── 保存最终文档
```

## 输出规范

### 内容撰写规范

**每个模块必须包含：**
1. 功能说明（做什么）
2. 实现方式（怎么做，适度技术描述）
3. 应用价值（解决什么问题，带来什么收益）

**表达要求：**
- ✅ 有总有分（段落说明 + 分点补充）
- ✅ 逻辑递进清晰
- ✅ 专业表达（符合政府/军用项目文档风格）
- ❌ 避免纯罗列式短句
- ❌ 避免空洞表述
- ❌ 避免模糊描述

## 技术实现

### Python 核心代码结构

```python
import os, sys
sys.stdout.reconfigure(encoding="utf-8")  # Windows 中文输出

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from PIL import Image, ImageDraw, ImageFont

# ============================================================
# 辅助函数
# ============================================================

def set_run_font(run, font_name="FangSong", size=12, bold=False, color=None):
    """设置 run 的字体（同时设置中英文）"""
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    if color:
        run.font.color.rgb = RGBColor(*color)


def set_paragraph_format(p, line_spacing=24, space_before=0, space_after=0, indent_first=True):
    """设置段落格式"""
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY  # 固定行距
    pf.line_spacing = Pt(line_spacing)
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if indent_first:
        pf.first_line_indent = Pt(24)  # 首行缩进 2 字符


def add_heading(doc, text, level=1):
    """添加标题（使用 Heading 样式，支持 Word 导航）"""
    p = doc.add_paragraph(style=f"Heading {level}")
    run = p.add_run(text)

    if level == 1:
        set_run_font(run, "SimHei", 18, bold=True, color=(0, 0, 0))
        p.paragraph_format.space_before = Pt(20)
        p.paragraph_format.space_after = Pt(12)
        p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    elif level == 2:
        set_run_font(run, "FangSong", 15, bold=True, color=(0, 0, 0))
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(8)
    elif level == 3:
        set_run_font(run, "FangSong", 14, bold=True, color=(0, 0, 0))
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(6)
    elif level == 4:
        set_run_font(run, "FangSong", 13, bold=True, color=(0, 0, 0))
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)

    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    return p


def add_body(doc, text, indent=True):
    """添加正文段落（仿宋 12 号，固定行距 24 磅，首行缩进）"""
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, "FangSong", 12)
    set_paragraph_format(p, line_spacing=24, indent_first=indent)
    return p


def add_image_with_caption(doc, image_path, caption, width_cm=15.5):
    """插入居中图片 + 居中图题"""
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run()
    run.add_picture(image_path, width=Cm(width_cm))

    cap = doc.add_paragraph()
    cap.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    cap_run = cap.add_run(caption)
    set_run_font(cap_run, "FangSong", 11, bold=True)
    cap.paragraph_format.space_before = Pt(2)
    cap.paragraph_format.space_after = Pt(8)
    cap.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE


def add_table(doc, headers, rows, col_widths=None):
    """添加表格（表头加粗居中）"""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = ""
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = p.add_run(h)
        set_run_font(run, "FangSong", 11, bold=True)

    for r, row in enumerate(rows):
        cells = table.rows[r + 1].cells
        for c, val in enumerate(row):
            cells[c].text = ""
            p = cells[c].paragraphs[0]
            run = p.add_run(str(val))
            set_run_font(run, "FangSong", 11)

    if col_widths:
        for row in table.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Cm(w)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    return table


def init_document():
    """初始化文档（设置页面边距和默认样式）"""
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.8)

    style = doc.styles["Normal"]
    style.font.name = "FangSong"
    style.font.size = Pt(12)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "FangSong")
    return doc


def verify_document(doc, target_chars=30000):
    """质量校验：字数统计、标题层级检查"""
    total = sum(len(p.text) for p in doc.paragraphs)
    headings = [p for p in doc.paragraphs if p.style.name.startswith("Heading")]

    print(f"[字数统计] 总字符数: {total}")
    print(f"[标题统计] 标题数量: {len(headings)}")
    for h in headings:
        print(f"  {h.style.name}: {h.text[:50]}")

    if total >= target_chars:
        print(f"[达标] 字符数 ≥ {target_chars}")
    else:
        print(f"[未达标] 还差 {target_chars - total} 字符")

    return total >= target_chars
```

### 完整示例代码

```python
# 初始化文档
doc = init_document()

# 添加标题（使用 Heading 样式，Word 导航可用）
add_heading(doc, '1. 技术方案(技术要求响应)', level=1)
add_heading(doc, '1.1 需求分析', level=2)
add_heading(doc, '一、项目背景', level=3)
add_heading(doc, '(一)数据仿真模型需求', level=4)

# 添加正文（首行缩进，固定行距）
add_body(doc, '信息指挥调度平台建设是落实单位信息化建设的重要举措...')

# 添加架构图（居中图片 + 图题）
add_image_with_caption(doc, '业务架构图.png', '图 1-1 平台业务架构图', 16)

# 添加表格
add_table(doc,
    headers=["序号", "技术参数", "招标要求", "投标响应", "满足情况"],
    rows=[["1", "平台架构", "C/S+B/S", "C/S+B/S", "满足"]],
    col_widths=[1.5, 6, 6, 6, 3])

# 保存文档
doc.save('解决方案.docx')

# 质量校验
verify_document(doc, target_chars=30000)
```

## 使用示例

### 方式一：直接使用 Skill（推荐）

```
/solution-writer 请根据以下文档生成整体解决方案：
- 原始需求：C:/项目/原始需求.md
- 文档结构：C:/项目/文档撰写需求.md
- 输出名称：某某系统解决方案.docx
- 目标字数：30000
```

### 方式二：扩写已有文档

```
/solution-writer 请扩写以下投标文档至3万字：
- 原始文档：C:/项目/投标文件_软件部分.docx
- 输出名称：投标文件_软件部分_扩写版.docx
- 保持原有章节结构和标题格式
- 文图结合，插入架构图
```

### 方式三：使用 Python 工具

```python
from solution_writer import SolutionDocGenerator, ArchitectureDiagramGenerator

# 1. 创建文档
gen = SolutionDocGenerator(output_path='C:/项目/')

# 2. 添加章节
gen.add_heading('1. 技术方案(技术要求响应)', level=1)
gen.add_body('本平台采用...')

# 3. 生成并插入架构图
arch = ArchitectureDiagramGenerator()
arch.generate_technical_architecture()
gen.add_image_with_caption('总体技术架构图.png', '图 1-1 总体技术架构图')

# 4. 保存并校验
gen.save('解决方案.docx')
gen.verify(target_chars=30000)
```

### Skill 输出
```
✅ 已生成完整解决方案文档

📁 文件位置：
   C:/.../解决方案.docx

📊 文档统计：
   - 总字符数：32,807
   - 标题数量：68 个（Heading 1-4）
   - 架构图：8 张 PNG 已插入

✅ 质量校验：
   - 字数达标：≥30,000
   - 标题层级：一致，Word 导航可用
   - 需求覆盖率：100%
   - 格式规范性：投标文档标准
```

## 可扩展场景

此 Skill 可扩展应用于：
- 政府/军用信息化项目投标
- 自然资源行业解决方案
- 智慧城市项目方案
- 企业数字化转型方案
- 其他需要结构化方案文档的场景

## 注意事项

1. **需求覆盖**：必须 100% 覆盖原始需求，不允许遗漏
2. **结构遵循**：严格按照评分细则章节结构，不增删改
3. **语言风格**：售前投标语言，正式严谨
4. **格式规范**：符合政府项目文档格式要求
5. **架构图示**：复杂技术方案需配架构图说明
6. **单脚本生成**：避免分段生成再合并（会丢失格式），应一次性生成全部章节

## 字体字号规范（重要）

| 元素 | 字体 | 字号 | 加粗 | 行距 | 样式 |
|------|------|------|------|------|------|
| 一级标题 | 黑体 | 18 号 | 是 | SINGLE | Heading 1 |
| 二级标题 | 仿宋 | 15 号 | 是 | SINGLE | Heading 2 |
| 三级标题 | 仿宋 | 14 号 | 是 | SINGLE | Heading 3 |
| 四级标题 | 仿宋 | 13 号 | 是 | SINGLE | Heading 4 |
| 正文 | 仿宋 | 12 号 | 否 | 固定 24 磅 | Normal |
| 图题 | 仿宋 | 11 号 | 是 | SINGLE | Normal |

**标题层级说明**：
- H1（level=1）：如"1. 技术方案" → 黑体 18 号加粗
- H2（level=2）：如"1.1 需求分析" → 仿宋 15 号加粗
- H3（level=3）：如"一、项目背景" → 仿宋 14 号加粗
- H4（level=4）：如"(一)数据仿真模型需求" → 仿宋 13 号加粗

**关键：必须使用 `doc.add_paragraph(style="Heading N")` 而非 `add_paragraph()`，否则 Word 导航不可用**

## 常见陷阱与最佳实践

### 1. 标题必须使用 Heading 样式

**错误做法**：
```python
p = doc.add_paragraph()
run = p.add_run('1.1 需求分析')
run.font.bold = True  # 只是视觉上加粗，Word 导航不可用
```

**正确做法**：
```python
p = doc.add_paragraph(style="Heading 2")
run = p.add_run('1.1 需求分析')  # Word 导航可用
```

### 2. 字体设置必须同时设置英文和中文

**错误做法**：
```python
run.font.name = 'FangSong'  # 只设置了英文字体
```

**正确做法**：
```python
run.font.name = 'FangSong'
run._element.rPr.rFonts.set(qn('w:eastAsia'), 'FangSong')  # 同时设置中文字体
```

### 3. 固定行距必须用 EXACTLY 模式

**错误做法**：
```python
p.paragraph_format.line_spacing = Pt(24)  # 默认是 AT_LEAST，不是固定
```

**正确做法**：
```python
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
p.paragraph_format.line_spacing = Pt(24)
```

### 4. 避免分段生成再合并

**错误做法**：
```python
# Part1 生成 1.1-1.3，Part2 生成 1.4-4.4，然后合并
# 合并时段落格式、字体设置会丢失
```

**正确做法**：
```python
# 单脚本一次性生成全部章节
doc = init_document()
build_chapter_1(doc)
build_chapter_2(doc)
build_chapter_3(doc)
build_chapter_4(doc)
doc.save(OUTPUT)
```

### 5. 避免无意义空行

**错误做法**：
```python
add_body(doc, '')  # 会产生空行段落
add_heading(doc, '标题')
add_body(doc, '')
```

**正确做法**：
```python
# 通过标题的 space_before/space_after 控制间距
add_heading(doc, '标题')
add_body(doc, '正文内容直接跟随')
```

### 6. 文件保存前关闭 Word

生成的 docx 文件如果被 Word 打开会导致无法覆盖保存。解决方式：
- 保存前确保 Word 已关闭
- 或使用临时文件名保存后再替换

### 7. Windows 环境中文输出

**必须设置**：
```python
import sys
sys.stdout.reconfigure(encoding="utf-8")
```
否则 `print()` 输出中文会报编码错误。

### 8. 首行缩进

正文段落需要首行缩进 2 字符（约 24 磅），标题和图题不需要缩进。

---

*此 Skill 基于葫芦岛信息指挥调度平台项目实战经验优化，适用于政府/军用行业投标解决方案撰写*
