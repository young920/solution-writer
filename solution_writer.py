# -*- coding: utf-8 -*-
"""
整体解决方案撰写工具
用于生成符合投标要求的解决方案文档

更新记录（基于葫芦岛项目实战）：
- 修复中文字体设置（需同时设置 eastAsia）
- 固定行距改为 EXACTLY 模式
- 增加首行缩进支持
- 增加质量校验函数
- 单脚本一次性生成全部章节
"""

import os
import sys
sys.stdout.reconfigure(encoding="utf-8")  # Windows 中文输出

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from copy import deepcopy
from PIL import Image, ImageDraw, ImageFont


# ==================== 核心辅助函数 ====================

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


# ==================== 文档生成工具类 ====================

class SolutionDocGenerator:
    """解决方案文档生成器"""

    def __init__(self, output_path=None):
        self.base_path = output_path or os.getcwd()
        self.doc = self._init_document()

    def _init_document(self):
        """初始化 Word 文档"""
        doc = Document()
        # 设置页面边距
        section = doc.sections[0]
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.8)
        section.right_margin = Cm(2.8)

        # 设置默认样式
        style = doc.styles['Normal']
        style.font.name = 'FangSong'
        style.font.size = Pt(12)
        style.element.rPr.rFonts.set(qn("w:eastAsia"), "FangSong")
        return doc

    def add_heading(self, text, level=1):
        """添加标题（使用 Heading 样式，支持 Word 导航）

        Args:
            text: 标题文本
            level: 1=一级标题(H1), 2=二级标题(H2), 3=三级标题(H3), 4=四级标题(H4)
        """
        p = self.doc.add_paragraph(style=f"Heading {level}")
        run = p.add_run(text)

        if level == 1:
            # 一级标题：黑体 18 号加粗
            set_run_font(run, "SimHei", 18, bold=True, color=(0, 0, 0))
            p.paragraph_format.space_before = Pt(20)
            p.paragraph_format.space_after = Pt(12)
            p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        elif level == 2:
            # 二级标题：仿宋 15 号加粗
            set_run_font(run, "FangSong", 15, bold=True, color=(0, 0, 0))
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(8)
        elif level == 3:
            # 三级标题：仿宋 14 号加粗
            set_run_font(run, "FangSong", 14, bold=True, color=(0, 0, 0))
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(6)
        elif level == 4:
            # 四级标题：仿宋 13 号加粗
            set_run_font(run, "FangSong", 13, bold=True, color=(0, 0, 0))
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(4)

        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        return p

    def add_body(self, text, indent=True):
        """添加正文段落（仿宋 12 号，固定行距 24 磅，首行缩进）"""
        p = self.doc.add_paragraph()
        run = p.add_run(text)
        set_run_font(run, "FangSong", 12)
        set_paragraph_format(p, line_spacing=24, indent_first=indent)
        return p

    def add_image_with_caption(self, image_path, caption, width_cm=15.5):
        """插入居中图片 + 居中图题"""
        if not os.path.exists(image_path):
            print(f'图片不存在：{image_path}')
            return False

        # 图片
        p = self.doc.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run()
        run.add_picture(image_path, width=Cm(width_cm))

        # 图题
        cap = self.doc.add_paragraph()
        cap.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cap_run = cap.add_run(caption)
        set_run_font(cap_run, "FangSong", 11, bold=True)
        cap.paragraph_format.space_before = Pt(2)
        cap.paragraph_format.space_after = Pt(8)
        cap.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        return True

    def add_image(self, image_path, width_cm=15.5):
        """插入图片（无图题，简单版）"""
        if not os.path.exists(image_path):
            print(f'图片不存在：{image_path}')
            return False

        p = self.doc.add_paragraph()
        run = p.add_run()
        run.add_picture(image_path, width=Cm(width_cm))
        return True

    def add_table(self, headers, rows, col_widths=None):
        """添加表格（表头加粗居中）"""
        table = self.doc.add_table(rows=1 + len(rows), cols=len(headers))
        table.style = "Table Grid"
        table.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # 表头
        hdr_cells = table.rows[0].cells
        for i, h in enumerate(headers):
            hdr_cells[i].text = ""
            p = hdr_cells[i].paragraphs[0]
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run = p.add_run(h)
            set_run_font(run, "FangSong", 11, bold=True)

        # 数据行
        for r, row in enumerate(rows):
            cells = table.rows[r + 1].cells
            for c, val in enumerate(row):
                cells[c].text = ""
                p = cells[c].paragraphs[0]
                run = p.add_run(str(val))
                set_run_font(run, "FangSong", 11)

        # 列宽
        if col_widths:
            for row in table.rows:
                for i, w in enumerate(col_widths):
                    row.cells[i].width = Cm(w)

        self.doc.add_paragraph().paragraph_format.space_after = Pt(4)
        return table

    def save(self, filename):
        """保存文档"""
        path = os.path.join(self.base_path, filename)
        self.doc.save(path)
        print(f'文档已保存：{path}')
        return path

    def verify(self, target_chars=30000):
        """质量校验"""
        total = sum(len(p.text) for p in self.doc.paragraphs)
        headings = [p for p in self.doc.paragraphs if p.style.name.startswith("Heading")]

        print(f"\n========== 质量校验 ==========")
        print(f"[字数统计] 总字符数: {total}")
        print(f"[标题统计] 标题数量: {len(headings)}")

        if total >= target_chars:
            print(f"[达标] 字符数 ≥ {target_chars} ✓")
            return True
        else:
            print(f"[未达标] 还差 {target_chars - total} 字符 ✗")
            return False


# ==================== 架构图生成工具类 ====================

class ArchitectureDiagramGenerator:
    """架构图生成器"""

    def __init__(self, output_path=None):
        self.base_path = output_path or os.getcwd()
        self.font = self._load_font()

    def _load_font(self):
        """加载中文字体"""
        font_paths = [
            'C:/Windows/Fonts/simsun.ttc',
            'C:/Windows/Fonts/simsun.ttc',
            'C:/Windows/Fonts/msyh.ttc',
            'C:/Windows/Fonts/simhei.ttf'
        ]
        for fp in font_paths:
            if os.path.exists(fp):
                try:
                    return ImageFont.truetype(fp, 16)
                except:
                    continue
        return ImageFont.load_default()

    def generate_technical_architecture(self, filename='总体技术架构图.png'):
        """生成总体技术架构图（六层架构 + 安全体系）"""
        width, height = 1400, 900
        img = Image.new('RGB', (width, height), color='white')
        draw = ImageDraw.Draw(img)

        title_font = ImageFont.truetype('C:/Windows/Fonts/simsun.ttc', 24)
        layer_font = ImageFont.truetype('C:/Windows/Fonts/simsun.ttc', 14)
        desc_font = ImageFont.truetype('C:/Windows/Fonts/simsun.ttc', 12)

        # Title
        draw.text((width//2, 20), '总体技术架构图', fill='black', anchor='mm', font=title_font)

        # Colors
        colors = {
            'user': (227, 242, 253),
            'biz': (255, 243, 224),
            'support': (232, 245, 233),
            'platform': (252, 228, 236),
            'data': (243, 229, 245),
            'infra': (224, 247, 250),
            'security': (255, 235, 238)
        }

        def draw_layer(y, title, items, color, box_width=180, box_height=50):
            draw.text((20, y), title, fill='#666666', font=layer_font)
            x = 20
            box_y = y + 20
            for item in items:
                draw.rectangle([x, box_y, x + box_width, box_y + box_height],
                             fill=color, outline='#333', width=2)
                # 文本居中
                if len(item) > 8:
                    mid = len(item) // 2
                    draw.text((x + box_width//2, box_y + box_height//2 - 8), item[:mid],
                             fill='black', anchor='mm', font=desc_font)
                    draw.text((x + box_width//2, box_y + box_height//2 + 8), item[mid:],
                             fill='black', anchor='mm', font=desc_font)
                else:
                    draw.text((x + box_width//2, box_y + box_height//2), item,
                             fill='black', anchor='mm', font=desc_font)
                x += box_width + 10
            return y + box_height + 40

        # Draw layers
        y = 60
        y = draw_layer(y, '【用户交互层】', ['C/S 客户端', 'B/S 客户端', '大屏展示系统'], colors['user'])
        y = draw_layer(y, '【业务应用层】', ['单位概览', '安防管理', '网络资源', '设备运维', '业务建设', '服务保障'], colors['biz'])
        y = draw_layer(y, '【应用支撑层】', ['GIS 引擎', '3D 渲染引擎', '数据总线', '告警服务', '权限管理'], colors['support'])
        y = draw_layer(y, '【数据资源层】', ['空间信息数据库', '三维模型数据库', '业务数据库', '实时数据库'], colors['data'])
        y = draw_layer(y, '【基础设施层】', ['服务器集群', '存储系统', '网络设备', '安全设备'], colors['infra'])

        # Security
        draw.text((20, y), '【安全体系】（贯穿各层，符合信息安全等级保护要求）', fill='#666666', font=layer_font)
        y += 28
        security_items = ['身份认证与访问控制', '数据加密存储', '边界防护', '安全审计', '国密算法/TLS加密', '冗余容错']
        x = 20
        for item in security_items:
            draw.rectangle([x, y, x + 200, y + 35], fill=colors['security'], outline='#c62828', width=2)
            draw.text((x + 100, y + 17), item, fill='black', anchor='mm', font=desc_font)
            x += 210

        # Save
        path = os.path.join(self.base_path, filename)
        img.save(path)
        print(f'架构图已保存：{path}')
        return path

    def generate_business_architecture(self, filename='业务架构图.png'):
        """生成业务架构图（四层架构）"""
        width, height = 1200, 700
        img = Image.new('RGB', (width, height), color='white')
        draw = ImageDraw.Draw(img)

        title_font = ImageFont.truetype('C:/Windows/Fonts/simsun.ttc', 24)
        layer_font = ImageFont.truetype('C:/Windows/Fonts/simsun.ttc', 14)
        desc_font = ImageFont.truetype('C:/Windows/Fonts/simsun.ttc', 12)

        draw.text((width//2, 20), '业务架构图', fill='black', anchor='mm', font=title_font)

        colors = {
            'perception': (225, 245, 254),
            'aggregation': (255, 249, 196),
            'application': (232, 245, 233),
            'presentation': (252, 228, 236)
        }

        def draw_layer(y, title, items, color):
            draw.text((20, y), title, fill='#666666', font=layer_font)
            x = 20
            box_y = y + 22
            for item in items:
                draw.rectangle([x, box_y, x + 165, box_y + 45], fill=color, outline='#666', width=1)
                draw.text((x + 82, box_y + 22), item, fill='black', anchor='mm', font=desc_font)
                x += 175
            return y + 80

        y = 70
        y = draw_layer(y, '【数据感知层】', ['GIS 地理信息采集', '3D 模型构建', '视频监控接入', '传感器采集', '报表录入'], colors['perception'])
        y = draw_layer(y, '【数据汇聚层】', ['空间信息数据库', '三维模型数据库', '业务数据库', '实时监测数据库'], colors['aggregation'])
        y = draw_layer(y, '【业务应用层】', ['数据仿真模型', '数据采集服务', '数据可视化服务', '后台管理服务'], colors['application'])
        y = draw_layer(y, '【服务展现层】', ['3D 可视化', '图表统计', '实时监控', '告警推送'], colors['presentation'])

        path = os.path.join(self.base_path, filename)
        img.save(path)
        print(f'业务架构图已保存：{path}')
        return path

    def generate_deployment_architecture(self, filename='部署架构图.png'):
        """生成部署架构图"""
        width, height = 1200, 800
        img = Image.new('RGB', (width, height), color='white')
        draw = ImageDraw.Draw(img)

        title_font = ImageFont.truetype('C:/Windows/Fonts/simsun.ttc', 24)
        layer_font = ImageFont.truetype('C:/Windows/Fonts/simsun.ttc', 14)
        desc_font = ImageFont.truetype('C:/Windows/Fonts/simsun.ttc', 12)

        draw.text((width//2, 20), '部署架构图', fill='black', anchor='mm', font=title_font)

        # 绘制云/机房区域
        def draw_box(x, y, w, h, text, color, border='#333'):
            draw.rectangle([x, y, x + w, y + h], fill=color, outline=border, width=2)
            draw.text((x + w//2, y + h//2), text, fill='black', anchor='mm', font=desc_font)

        # 内部网络区域
        draw_box(300, 80, 600, 180, '单位内部机房', (224, 247, 250))

        # 服务器
        draw_box(320, 100, 150, 50, '应用服务器\n(64核ARM)', (255, 243, 224))
        draw_box(480, 100, 150, 50, '数据库服务器\n(MySQL主备)', (255, 243, 224))
        draw_box(640, 100, 150, 50, '存储服务器\n(8TB×4)', (255, 243, 224))

        # 网络设备
        draw_box(320, 160, 150, 40, '核心交换机', (232, 245, 233))
        draw_box(480, 160, 150, 40, '防火墙', (232, 245, 233))
        draw_box(640, 160, 150, 40, 'IDS', (232, 245, 233))

        # 客户端区域
        draw_box(80, 350, 140, 50, '指挥大厅大屏', (227, 242, 253))
        draw_box(240, 350, 140, 50, '业务工作站', (227, 242, 253))
        draw_box(400, 350, 140, 50, '移动终端', (227, 242, 253))
        draw_box(560, 350, 140, 50, 'Web浏览器', (227, 242, 253))

        # 连接线示意
        draw.line([(170, 400), (170, 280), (320, 280)], fill='#666', width=2)
        draw.line([(330, 280), (450, 280), (450, 400)], fill='#666', width=2)

        path = os.path.join(self.base_path, filename)
        img.save(path)
        print(f'部署架构图已保存：{path}')
        return path

    def generate_security_architecture(self, filename='安全防护体系图.png'):
        """生成安全防护体系图"""
        width, height = 1200, 700
        img = Image.new('RGB', (width, height), color='white')
        draw = ImageDraw.Draw(img)

        title_font = ImageFont.truetype('C:/Windows/Fonts/simsun.ttc', 24)
        layer_font = ImageFont.truetype('C:/Windows/Fonts/simsun.ttc', 14)
        desc_font = ImageFont.truetype('C:/Windows/Fonts/simsun.ttc', 12)

        draw.text((width//2, 20), '安全防护体系图', fill='black', anchor='mm', font=title_font)

        # 中心安全域
        center_x, center_y = 600, 350

        # 安全层次
        layers = [
            ('边界安全', [(80, 350, 150, 60, '防火墙', (227, 242, 253)),
                         (80, 420, 150, 60, 'IDS/IPS', (227, 242, 253)),
                         (80, 490, 150, 60, '网闸', (227, 242, 253))]),
            ('网络安全', [(280, 350, 150, 60, 'VLAN隔离', (255, 243, 224)),
                         (280, 420, 150, 60, '访问控制', (255, 243, 224)),
                         (280, 490, 150, 60, '流量监控', (255, 243, 224))]),
            ('主机安全', [(480, 350, 150, 60, '冗余电源', (232, 245, 233)),
                         (480, 420, 150, 60, '热插拔硬盘', (232, 245, 233)),
                         (480, 490, 150, 60, '集群部署', (232, 245, 233))]),
            ('应用安全', [(680, 350, 150, 60, '身份认证', (243, 229, 245)),
                         (680, 420, 150, 60, '权限控制', (243, 229, 245)),
                         (680, 490, 150, 60, '安全审计', (243, 229, 245))]),
            ('数据安全', [(880, 350, 150, 60, 'AES-256加密', (255, 235, 238)),
                         (880, 420, 150, 60, 'TLS传输', (255, 235, 238)),
                         (880, 490, 150, 60, '备份容灾', (255, 235, 238))])
        ]

        for title, boxes in layers:
            draw.text((80, 80), f'【{title}】', fill='#333', font=layer_font)
            for x, y, w, h, text, color in boxes:
                draw.rectangle([x, y, x + w, y + h], fill=color, outline='#666', width=1)
                draw.text((x + w//2, y + h//2), text, fill='black', anchor='mm', font=desc_font)

        path = os.path.join(self.base_path, filename)
        img.save(path)
        print(f'安全防护体系图已保存：{path}')
        return path


# ==================== 文档合并工具类 ====================

class DocumentMerger:
    """文档合并器"""

    def __init__(self, base_path=None):
        self.base_path = base_path or os.getcwd()

    def merge(self, file_list, output_filename='合并文档.docx'):
        """合并多个 Word 文档

        注意：不推荐使用合并方式，会丢失格式。建议单脚本一次性生成。
        """
        master_doc = Document()
        style = master_doc.styles['Normal']
        style.font.name = 'FangSong'
        style.font.size = Pt(12)
        style.element.rPr.rFonts.set(qn("w:eastAsia"), "FangSong")

        for filename in file_list:
            path = os.path.join(self.base_path, filename)
            if not os.path.exists(path):
                print(f'文件不存在：{path}')
                continue

            src = Document(path)
            for elem in src.element.body:
                master_doc.element.body.append(deepcopy(elem))
            print(f'已添加：{filename}')

        output_path = os.path.join(self.base_path, output_filename)
        master_doc.save(output_path)
        print(f'合并完成：{output_path}')
        return output_path


# ==================== 使用示例 ====================

if __name__ == '__main__':
    base_path = r'C:\Users\36193\Desktop\项目'

    # 1. 创建文档生成器
    gen = SolutionDocGenerator(base_path)

    # 2. 添加章节（使用 Heading 样式）
    gen.add_heading('1. 技术方案(技术要求响应)', level=1)
    gen.add_heading('1.1 需求分析', level=2)
    gen.add_heading('一、项目背景', level=3)
    gen.add_body('信息指挥调度平台建设是落实单位信息化建设的重要举措...')

    # 3. 生成架构图
    arch_gen = ArchitectureDiagramGenerator(base_path)
    arch_path = arch_gen.generate_technical_architecture()

    # 4. 插入架构图
    gen.add_image_with_caption(arch_path, '图 1-1 总体技术架构图', 16)

    # 5. 添加表格
    gen.add_table(
        headers=["序号", "技术参数", "招标要求", "投标响应", "满足情况"],
        rows=[["1", "平台架构", "C/S+B/S", "C/S+B/S", "满足"]],
        col_widths=[1.5, 6, 6, 6, 3]
    )

    # 6. 保存文档
    gen.save('解决方案示例.docx')

    # 7. 质量校验
    gen.verify(target_chars=30000)

    print('\n✅ 解决方案文档生成完成！')