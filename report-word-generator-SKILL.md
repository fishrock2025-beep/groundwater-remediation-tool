---
name: report-word-generator
title: 地下水调查报告Word文档生成
description: 根据官方报告编制格式规范，将txt文本文件转换为格式规范的简体中文Word文档(.docx)
category: productivity
---

# 地下水调查报告Word文档生成

根据**报告编制格式.docx**模板规范，将文本文件(.txt)转换为格式规范的Word文档(.docx)。

## 适用场景

- 将水文地质调查报告、地下水污染调查报告等txt文本文件转为Word文档
- 需要统一格式的各类环境调查报告

## 格式规范

### 1. 页面设置
| 项目 | 设置 |
|------|------|
| 纸张 | A4 |
| 页边距 | 上3cm、下3.5cm、左3.17cm、右3.17cm |
| 页眉 | 五号、楷体_GB2312，居中 |
| 页码 | 五号、Times New Roman，居中 |
| 章节 | 每个章节单独起页 |

### 2. 标题格式
| 级别 | 格式 |
|------|------|
| 报告标题 | 二号，黑体（英文Times New Roman），1.5倍行距，居中 |
| 封面单位/时间 | 小二号，黑体，1.5倍行距 |
| 一级标题（章） | 三号，黑体，居中/左对齐，1.5倍行距 |
| 二级标题（节） | 小三号，仿宋_GB2312，加粗，左对齐，1.5倍行距 |
| 三级标题（小节） | 小三号，楷体_GB2312，加粗，左对齐，1.5倍行距 |
| 四级标题 | 四号/小四号，仿宋_GB2312，加粗，左对齐，1.5倍行距 |
| 四级以下 | 按(1)、1)、①、A、a顺序，仿宋_GB2312，首行缩进2字符 |

### 3. 正文格式
| 项目 | 设置 |
|------|------|
| 字体 | 四号/小四号，仿宋_GB2312（英文Times New Roman） |
| 行距 | 1.5倍行距 |
| 缩进 | 首行缩进2字符 |
| 段前段后 | 均为0 |

### 4. 表格格式
| 项目 | 设置 |
|------|------|
| 表名 | 五号，黑体，居中，1.5倍行距，序号后空两格（如"表1-1  xxx"） |
| 表内文字 | 五号，仿宋_GB2312 |
| 行距 | 单倍行距 |
| 排列 | 网格式，居中 |
| 表序 | 按章编排（表1-1、表1-2...） |

### 5. 图件格式
| 项目 | 设置 |
|------|------|
| 图名 | 五号，黑体，居中，1.5倍行距，序号后空两格（如"图1-1  xxx"） |
| 格式 | 嵌入式，居中 |
| 颜色 | 自动彩色 |
| 图序 | 按章编排（图1-1、图1-2...） |

### 6. 字号对照表（中文字号 → 磅值）
| 中文字号 | 磅值(pt) | EMU值 |
|---------|---------|-------|
| 二号 | 22pt | 563880 |
| 小二 | 18pt | 457200 |
| 三号 | 16pt | 406400 |
| 小三 | 15pt | 381000 |
| 四号 | 14pt | 355600 |
| 小四 | 12pt | 304800 |
| 五号 | 10.5pt | 266700 |
| 小五 | 9pt | 228600 |

## 使用方法

### 方式一：直接调用生成函数（推荐）

```python
import sys
sys.path.insert(0, '/tmp/docx_venv/lib/python3.12/site-packages')
```

### 方式二：命令行一键生成

```bash
# 创建虚拟环境（首次）
python3 -m venv /tmp/docx_venv
/tmp/docx_venv/bin/pip install python-docx

# 生成Word文档
/tmp/docx_venv/bin/python3 /path/to/script.py
```

## Python生成脚本

以下是完整的Word文档生成脚本，保存为 `generate_report_docx.py` 使用：

```python
#!/usr/bin/env python3
"""
地下水调查报告 Word文档生成工具
根据报告编制格式.docx规范，将txt文本转换为格式规范的Word文档
"""

import sys
import os
import re

sys.path.insert(0, '/tmp/docx_venv/lib/python3.12/site-packages')

try:
    from docx import Document
    from docx.shared import Pt, Inches, Cm, RGBColor, Emu
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.enum.section import WD_ORIENT
    from docx.oxml.ns import qn, nsdecls
    from docx.oxml import parse_xml
    print("模块导入成功")
except ImportError as e:
    print(f"请先安装python-docx: pip install python-docx")
    print(f"错误: {e}")
    sys.exit(1)


def set_font(run, font_name_cn='仿宋_GB2312', font_name_en='Times New Roman', 
             size_pt=12, bold=False, color=None):
    """设置字体属性"""
    run.font.name = font_name_en
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    # 设置中文字体
    r = run._element
    rPr = r.find(qn('w:rPr'))
    if rPr is None:
        rPr = parse_xml(f'<w:rPr {nsdecls("w")}></w:rPr>')
        r.insert(0, rPr)
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")}></w:rFonts>')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_name_cn)
    if color:
        run.font.color.rgb = color


def set_paragraph_format(paragraph, line_spacing=1.5, first_line_indent=None,
                         alignment=None, space_before=0, space_after=0):
    """设置段落格式"""
    pf = paragraph.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = line_spacing
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if first_line_indent is not None:
        pf.first_line_indent = first_line_indent
    if alignment is not None:
        paragraph.alignment = alignment


def setup_page(section, margin_top_cm=3, margin_bottom_cm=3.5, 
               margin_left_cm=3.17, margin_right_cm=3.17):
    """设置页面"""
    section.top_margin = Cm(margin_top_cm)
    section.bottom_margin = Cm(margin_bottom_cm)
    section.left_margin = Cm(margin_left_cm)
    section.right_margin = Cm(margin_right_cm)
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)


def add_page_header(doc, text):
    """添加页眉"""
    header = doc.sections[0].header
    p = header.paragraphs[0]
    run = p.add_run(text)
    set_font(run, '楷体_GB2312', 'Times New Roman', 10.5)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def add_page_number(doc):
    """添加页码"""
    footer = doc.sections[0].footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    fldChar1 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>')
    run._element.append(fldChar1)
    run2 = p.add_run()
    instrText = parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>')
    run2._element.append(instrText)
    run3 = p.add_run()
    fldChar2 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
    run3._element.append(fldChar2)


def txt_to_docx(input_path, output_path=None):
    """
    将txt报告文件转换为格式规范的Word文档
    
    参数:
        input_path: 输入txt文件路径
        output_path: 输出docx文件路径（None则自动生成）
    
    返回:
        docx文件路径
    """
    if output_path is None:
        base = os.path.splitext(input_path)[0]
        output_path = base + '.docx'
    
    with open(input_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    if lines and 'warning: setlocale' in lines[0]:
        lines = lines[1:]
    
    doc = Document()
    setup_page(doc.sections[0])
    
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    style.element.find(qn('w:rPr'))
    rPr = style.element.find(qn('w:rPr'))
    if rPr is not None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:eastAsia="仿宋_GB2312"/>')
        rPr.append(rFonts)
    
    chapter_pat = re.compile(r'^第[一二三四五六七八九十百]+章\s+')
    section_pat = re.compile(r'^\d+\.\d+\s+')
    subsection_pat = re.compile(r'^\d+\.\d+\.\d+\s+')
    list1_pat = re.compile(r'^（[一二三四五六七八九十百]+）')
    list2_pat = re.compile(r'^（\d+）')
    list3_pat = re.compile(r'^\(\d+\)')
    
    for line in lines:
        line = line.strip()
        if not line:
            doc.add_paragraph()
            continue
        
        if '|' in line:
            parts = line.split('|', 1)
            if len(parts) > 1 and parts[0].strip().isdigit():
                line = parts[1].strip()
        
        if not line:
            continue
        
        if chapter_pat.match(line):
            heading = doc.add_heading('', 1)
            run = heading.add_run(line)
            set_font(run, '黑体', 'Times New Roman', 16, bold=True)
            set_paragraph_format(heading, 1.5, alignment=WD_ALIGN_PARAGRAPH.LEFT)
            doc.add_page_break()
            
        elif section_pat.match(line):
            heading = doc.add_heading('', 2)
            run = heading.add_run(line)
            set_font(run, '仿宋_GB2312', 'Times New Roman', 15, bold=True)
            set_paragraph_format(heading, 1.5, alignment=WD_ALIGN_PARAGRAPH.LEFT)
            
        elif subsection_pat.match(line):
            heading = doc.add_heading('', 3)
            run = heading.add_run(line)
            set_font(run, '楷体_GB2312', 'Times New Roman', 15, bold=True)
            set_paragraph_format(heading, 1.5, alignment=WD_ALIGN_PARAGRAPH.LEFT)
            
        elif list1_pat.match(line) or list2_pat.match(line) or list3_pat.match(line):
            p = doc.add_paragraph()
            run = p.add_run(line)
            set_font(run, '仿宋_GB2312', 'Times New Roman', 12, bold=False)
            set_paragraph_format(p, 1.5, first_line_indent=Emu(355600))
            
        elif line.startswith('- ') or line.startswith('• '):
            p = doc.add_paragraph()
            run = p.add_run(line[2:])
            set_font(run, '仿宋_GB2312', 'Times New Roman', 12)
            set_paragraph_format(p, 1.5, first_line_indent=Emu(355600))
            
        else:
            p = doc.add_paragraph()
            run = p.add_run(line)
            set_font(run, '仿宋_GB2312', 'Times New Roman', 12)
            set_paragraph_format(p, 1.5, first_line_indent=Emu(355600))
    
    doc.save(output_path)
    file_size = os.path.getsize(output_path)
    print(f"文档已生成: {output_path}")
    print(f"段落数: {len(doc.paragraphs)}")
    print(f"文件大小: {file_size/1024:.1f} KB")
    
    return output_path


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='将地下水调查报告txt转换为格式规范的Word文档')
    parser.add_argument('input', help='输入的txt文件路径')
    parser.add_argument('-o', '--output', help='输出的docx文件路径（默认自动生成）')
    args = parser.parse_args()
    
    txt_to_docx(args.input, args.output)
```

## 一键生成命令

```bash
# 1. 创建虚拟环境（首次需要）
python3 -m venv /tmp/docx_venv
/tmp/docx_venv/bin/pip install python-docx

# 2. 运行脚本
/tmp/docx_venv/bin/python3 generate_report_docx.py "输入文件.txt" -o "输出文件.docx"
```

## 常见问题

### 1. python-docx未安装
```bash
python3 -m venv /tmp/docx_venv
/tmp/docx_venv/bin/pip install python-docx
```

### 2. 中文字体无法显示
```bash
fc-list :lang=zh | head -5
sudo apt install fonts-wqy-zenhei fonts-wqy-microhei
```

### 3. 页眉页脚需根据实际项目调整

### 4. 表格和图片需单独处理

## 验证清单

- [ ] 纸张大小为A4
- [ ] 页边距为上3、下3.5、左3.17、右3.17（cm）
- [ ] 标题格式正确（字体、字号、加粗）
- [ ] 正文使用仿宋_GB2312
- [ ] 首行缩进2字符
- [ ] 行距为1.5倍
- [ ] 无乱码
- [ ] 页眉页脚正确

## 引用的模板文件

- `报告编制格式.docx` — 位于 `e:\code\baogao\` 目录下
