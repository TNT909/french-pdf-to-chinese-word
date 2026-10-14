#!/usr/bin/env python3
"""法语PDF翻译为中文Word - 完整流程脚本"""
import os
import sys
import json
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============ 配置区 ============
# 根据实际文档修改以下配置

DOC_CONFIG = {
    "title": "文档标题",
    "subtitle": "副标题",
    "org": "机构名称",
    "ref_no": "编号",
    "nif": "NIF编号",
}

# 页码映射（根据PDF实际页码填写）
PAGE_MAPPING = {
    # '一、招标公告': 3,
    # '二、': 12,
}

# 颜色定义
DARK_BLUE = RGBColor(0x1F, 0x3A, 0x5F)
MEDIUM_BLUE = RGBColor(0x2E, 0x5A, 0x88)
LIGHT_BLUE = RGBColor(0x3A, 0x6B, 0x9E)
LIGHT_GRAY = RGBColor(0x55, 0x55, 0x55)
BLACK = RGBColor(0x1A, 0x1A, 0x1A)
PAGE_GRAY = RGBColor(0x99, 0x99, 0x99)

# ============ 函数区 ============

def create_document():
    """创建基础文档并设置全局样式"""
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = '宋体'
    style.font.size = Pt(12)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    style.paragraph_format.line_spacing = 1.5
    return doc

def add_cover(doc, config):
    """添加封面页"""
    for _ in range(4):
        doc.add_paragraph('')
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(config['org'])
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.color.rgb = DARK_BLUE
    
    doc.add_paragraph('')
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(config['title'])
    run.font.size = Pt(22)
    run.font.bold = True
    run.font.color.rgb = BLACK
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"（{config['ref_no']}）")
    run.font.size = Pt(16)
    run.font.bold = True
    
    doc.add_paragraph('')
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('中文译本')
    run.font.size = Pt(14)
    
    for _ in range(4):
        doc.add_paragraph('')
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(config['nif'])
    run.font.size = Pt(12)
    
    doc.add_page_break()

def add_toc(doc, items):
    """添加目录页
    items: [(text, level), ...] level 1或2
    """
    doc.add_heading('目  录', level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph('')
    for text, level in items:
        p = doc.add_paragraph(text)
        if level == 1:
            for run in p.runs:
                run.font.bold = True
                run.font.size = Pt(12)
        else:
            p.paragraph_format.left_indent = Cm(1)
            for run in p.runs:
                run.font.size = Pt(11)
    doc.add_page_break()

def add_page_label(doc, page_num, style='article'):
    """添加页码标签"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    if style == 'chapter':
        r = p.add_run(f'—— 原文第{page_num}页 ——')
        r.font.size = Pt(9)
        r.font.color.rgb = PAGE_GRAY
        r.font.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        r = p.add_run(f'[原文页码 {page_num}]')
        r.font.size = Pt(7)
        r.font.color.rgb = RGBColor(0xBB, 0xBB, 0xBB)
    r.font.name = '宋体'
    r._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

def add_heading_with_page(doc, text, level, page_num=None):
    """添加标题，可选带页码标签"""
    if page_num:
        add_page_label(doc, page_num, 'chapter' if level <= 2 else 'article')
    doc.add_heading(text, level=level)

def add_table(doc, headers, rows_data, style='Table Grid'):
    """添加表格"""
    cols = len(headers)
    table = doc.add_table(rows=1, cols=cols)
    table.style = style
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, t in enumerate(headers):
        hdr[i].text = t
    for row_data in rows_data:
        row = table.add_row().cells
        for i, t in enumerate(row_data[:cols]):
            row[i].text = t
    return table

# ============ 主流程 ============

if __name__ == "__main__":
    print("法语PDF翻译为中文Word - 工具库")
    print("\n可用函数:")
    print("  create_document() - 创建基础文档")
    print("  add_cover(doc, config) - 添加封面")
    print("  add_toc(doc, items) - 添加目录")
    print("  add_page_label(doc, page_num, style) - 添加页码标签")
    print("  add_heading_with_page(doc, text, level, page_num) - 添加标题")
    print("  add_table(doc, headers, rows_data) - 添加表格")
    print("\n使用示例:")
    print("  doc = create_document()")
    print("  add_cover(doc, DOC_CONFIG)")
    print("  add_heading_with_page(doc, '一、招标公告', 1, 3)")
    print("  doc.save('output.docx')")
