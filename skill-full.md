# 整体解决方案撰写 Skill

## Skill 名称
`solution-writer` - 整体解决方案撰写专家

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
┌─────────────────────────────────────────────────────────┐
│                   整体解决方案撰写流程                    │
└─────────────────────────────────────────────────────────┘

1. 输入确认
   ├── 原始需求文档路径
   ├── 文档结构/评分细则路径
   └── 输出文档名称

2. 需求解析（内部执行）
   ├── 功能需求拆解
   ├── 业务流程梳理
   ├── 数据需求分析
   ├── 管理/监管需求
   └── 展示与可视化需求

3. 结构解析（内部执行）
   ├── 章节结构确认
   ├── 每章撰写目标
   ├── 内容边界确定
   └── 深度要求确认

4. 方案撰写
   ├── 逐章节输出内容
   ├── 100% 映射原始需求
   ├── 使用投标语言
   └── 结构化拆解模块

5. 架构图生成（如需要）
   ├── 识别需要架构图的章节
   ├── 生成架构 PNG 图片
   └── 插入到 Word 文档

6. 文档合并
   ├── 合并所有章节
   └── 生成完整版文档

7. 质量校验
   ├── 需求覆盖性检查
   ├── 结构完整性检查
   └── 格式规范性检查
```

## 输出规范

### 文档结构示例（自然资源项目）
```
3.2.1 总体技术方案
  └── 总体架构、业务架构、数据架构、安全架构 + 架构图

3.2.2 用地用矿用海用林用草协同审批模块
  └── 数据库建设、表单流程管理、审批规则及查询统计、
      数据交换文件、许可及证书管理、三级矿业权全生命周期管理

3.2.3 用地数字化审核模块
  └── 地类审核、权属审核、规划审核、补充耕地审核、
      矿产压覆审核、土地利用计划指标管理

3.2.4 矿业权出让管理模块
  └── 出让区块来源管理、出让区块综合评价管理、
      省级出让矿业权核实业务管理、省级矿业权出让项目管理

3.2.5 矿业权项目档案管理模块
  └── 数据库设计、表单设计、新建/修改/删除功能、
      附件管理、数据调用、数据录入接口

3.2.6 项目找资源智能选址模块
  └── 限制性要求自动查询、量化比对预选址调整、
      合规性审查报告生成

3.2.7 资源找项目模块
  └── 存量资源展示、条件筛选匹配、数据动态更新

3.2.8 项目全生命周期监管模块
  └── 要素保障清单、全链条数据监管、可视化核查
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

# 2. 标题添加（注意：level=1 是章标题 15 号，level=2 是节标题 14 号）
def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    run = p.add_run(text)
    if level == 1:
        # 章标题：仿宋 15 号加粗
        run.font.name = 'FangSong'
        run.font.size = Pt(15)
        run.font.bold = True
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
    elif level == 2:
        # 节标题：仿宋 14 号加粗
        run.font.name = 'FangSong'
        run.font.size = Pt(14)
        run.font.bold = True
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)
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

### 完整示例代码

```python
# 初始化文档
doc = init_document()

# 添加章标题
add_heading(doc, '3.2.1 总体技术方案', level=1)

# 添加节标题
add_heading(doc, '一、总体架构设计', level=2)

# 添加正文（直接跟随，不要加空行）
add_body(doc, '本系统总体架构设计遵循...')
add_body(doc, '基础设施层：基于省政务云资源...')
add_body(doc, '数据资源层：构建统一的...')

# 添加架构图
add_image(doc, '总体技术架构图.png', width_cm=18)
add_body(doc, '图 3-1 总体技术架构图')

# 保存文档
doc.save('3.2.1 总体技术方案.docx')
```

## 使用示例

### 方式一：直接使用 Skill（推荐）

```
/solution-writer 请根据以下文档生成整体解决方案：
- 原始需求：C:/项目/原始需求.md
- 文档结构：C:/项目/文档撰写需求.md
- 输出名称：某某系统解决方案.docx
```

### 方式二：使用 Python 工具

```python
from solution_writer import SolutionDocGenerator, ArchitectureDiagramGenerator, DocumentMerger

# 1. 创建文档
gen = SolutionDocGenerator(output_path='C:/项目/')

# 2. 添加章节
gen.add_heading('3.2.1 总体技术方案', level=1)
gen.add_body('本系统总体架构设计遵循...')

# 3. 生成并插入架构图
arch = ArchitectureDiagramGenerator()
arch.generate_technical_architecture()
gen.add_image('总体技术架构图.png')

# 4. 保存
gen.save('解决方案.docx')
```

### 方式三：命令行调用

```bash
# 生成完整解决方案
python solution_writer.py --requirements 原始需求.md --structure 文档撰写需求.md --output 解决方案.docx

# 仅生成架构图
python solution_writer.py --diagram technical --output 架构图.png

# 合并多个文档
python solution_writer.py --merge 3.2.1.docx 3.2.2.docx 3.2.3.docx --output 完整版.docx
```

### Skill 输出
```
✅ 已生成完整解决方案文档

📁 文件位置：
   C:/.../用地用矿用海用林用草协同审批系统解决方案_完整版.docx

📊 文档统计：
   - 总段落数：306 段
   - 标题数量：171 个
   - 文件大小：0.06 MB
   - 章节数量：8 章（3.2.1-3.2.8）

📐 架构图：
   - 总体技术架构图.png 已插入 3.2.1 章节
   - 展示六层架构体系 + 安全体系

✅ 质量校验：
   - 需求覆盖率：100%
   - 结构完整性：符合评分细则
   - 格式规范性：投标文档标准
```

## 可扩展场景

此 Skill 可扩展应用于：
- 政府信息化项目投标
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

## 字体字号规范（重要）

| 元素 | 字体 | 字号 | 加粗 | 行距 |
|------|------|------|------|------|
| 正文 | 仿宋 | 12 号 | 否 | 固定 20 磅 |
| 二级标题 | 仿宋 | 15 号 | 是 | - |
| 三级标题 | 仿宋 | 14 号 | 是 | - |

**标题层级说明**：
- 章标题（level=1）：如"3.2.1 总体技术方案" → 仿宋 15 号加粗
- 节标题（level=2）：如"一、总体架构设计" → 仿宋 14 号加粗
- 无一级标题（16 号）定义，不要使用

## 常见陷阱与最佳实践

### 1. 避免无意义空行

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

### 3. 文件保存前关闭 Word

生成的 docx 文件如果被 Word 打开会导致无法覆盖保存。解决方式：
- 保存前确保 Word 已关闭
- 或使用临时文件名保存后再替换

### 4. 标题层级不要超过三级

投标文档一般只需要：
- 章标题（3.2.1 格式）→ level=1 → 15 号加粗
- 节标题（一、二、三）→ level=2 → 14 号加粗
- 如有更细分的小标题 → 可用加粗正文代替

---

*此 Skill 基于实际项目经验总结，适用于自然资源/政府行业投标解决方案撰写*
