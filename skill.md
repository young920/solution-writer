# solution-writer - 整体解决方案撰写 Skill

## 触发条件

当用户需要：
- 编写投标解决方案文档
- 将需求文档转化为解决方案
- 生成带格式的技术方案文档
- 创建解决方案架构图

## 核心能力

### 1. 需求解析能力
- 读取并解析甲方业务需求文档
- 解析评分细则/文档结构要求
- 建立"需求→方案"映射关系
- 覆盖性校验（确保 100% 覆盖原始需求）

### 2. 方案撰写能力
- 严格遵循评分细则章节结构
- 使用售前投标语言（正式、严谨、非口语化）
- 成段描述为主，分点说明为辅
- 每个模块包含：功能说明、实现方式、应用价值

### 3. 文档生成能力
- 生成带格式的 Word 文档（.docx）
- 字体格式：正文仿宋 12 号，二级标题仿宋 15 号加粗，三级标题仿宋 14 号加粗
- 行距设置：固定行距 20 磅
- 支持合并多个章节文档

### 4. 架构图生成能力
- 使用 PIL 生成系统架构图 PNG
- 展示多层架构体系（如六层架构）
- 支持安全体系、数据流等可视化
- 自动插入到 Word 文档对应章节

## 执行流程

```
1. 输入确认 → 2. 需求解析 → 3. 结构解析 → 4. 方案撰写 → 5. 架构图生成 → 6. 文档合并 → 7. 质量校验
```

## 输出规范

### 文档结构示例
```
3.2.1 总体技术方案
  └── 总体架构、业务架构、数据架构、安全架构 + 架构图

3.2.2 用地用矿用海用林用草协同审批模块
  └── 数据库建设、表单流程管理、审批规则及查询统计...
```

### 内容撰写规范

**每个模块必须包含：**
1. 功能说明（做什么）
2. 实现方式（怎么做，适度技术描述）
3. 应用价值（解决什么问题，带来什么收益）

**表达要求：**
- ✅ 有总有分（段落说明 + 分点补充）
- ✅ 逻辑递进清晰
- ✅ 专业表达（符合政府项目文档风格）
- ❌ 避免纯罗列式短句
- ❌ 避免空洞表述
- ❌ 避免模糊描述

## 技术实现

### Python 核心代码结构

```python
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from PIL import Image, ImageDraw, ImageFont

# 1. 文档初始化
def init_document():
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'FangSong'  # 仿宋
    style.font.size = Pt(12)       # 12 号
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'FangSong')
    return doc

# 2. 标题添加
def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    run = p.add_run(text)
    if level == 1:
        # 章标题：仿宋 15 号加粗
        run.font.name = 'FangSong'
        run.font.size = Pt(15)
        run.font.bold = True
    elif level == 2:
        # 节标题：仿宋 14 号加粗
        run.font.name = 'FangSong'
        run.font.size = Pt(14)
        run.font.bold = True
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'FangSong')
    return p

# 3. 正文添加（仿宋 12 号，固定行距 20 磅）
def add_body(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = 'FangSong'
    run.font.size = Pt(12)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'FangSong')
    p.paragraph_format.line_spacing = Pt(20)  # 固定行距 20 磅
    return p

# 4. 架构图生成
def generate_architecture_diagram():
    # 使用 PIL 绘制多层架构图
    pass

# 5. 文档合并
def merge_documents():
    # 合并多个章节文档
    pass
```

## 字体字号规范（重要）

| 元素 | 字体 | 字号 | 加粗 | 行距 |
|------|------|------|------|------|
| 正文 | 仿宋 | 12 号 | 否 | 固定 20 磅 |
| 二级标题 | 仿宋 | 15 号 | 是 | - |
| 三级标题 | 仿宋 | 14 号 | 是 | - |

## 常见陷阱与最佳实践

### 1. 避免无意义空行
通过标题的 `space_before/space_after` 控制间距，不要添加空行段落。

### 2. 字体设置必须同时设置英文和中文
```python
run.font.name = 'FangSong'
run._element.rPr.rFonts.set(qn('w:eastAsia'), 'FangSong')
```

### 3. 文件保存前关闭 Word
生成的 docx 文件如果被 Word 打开会导致无法覆盖保存。

## 注意事项

1. **需求覆盖**：必须 100% 覆盖原始需求，不允许遗漏
2. **结构遵循**：严格按照评分细则章节结构，不增删改
3. **语言风格**：售前投标语言，正式严谨
4. **格式规范**：符合政府项目文档格式要求
5. **架构图示**：复杂技术方案需配架构图说明

## 可扩展场景

- 政府信息化项目投标
- 自然资源行业解决方案
- 智慧城市项目方案
- 企业数字化转型方案
