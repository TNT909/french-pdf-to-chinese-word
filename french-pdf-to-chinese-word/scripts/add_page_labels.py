#!/usr/bin/env python3
"""页码标签插入脚本"""
import sys
import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

PAGE_GRAY = RGBColor(0x99, 0x99, 0x99)
PAGE_GRAY_LIGHT = RGBColor(0xBB, 0xBB, 0xBB)

def add_page_label(paragraph, page_num, style='article'):
    """在段落前插入页码标签"""
    label_p = paragraph.insert_paragraph_before()
    label_p.paragraph_format.space_before = Pt(1)
    label_p.paragraph_format.space_after = Pt(1)
    
    if style == 'chapter':
        r = label_p.add_run(f'—— 原文第{page_num}页 ——')
        r.font.size = Pt(9)
        r.font.color.rgb = PAGE_GRAY
        r.font.bold = True
        label_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        r = label_p.add_run(f'[原文页码 {page_num}]')
        r.font.size = Pt(7)
        r.font.color.rgb = PAGE_GRAY_LIGHT
    
    r.font.name = '宋体'
    r._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

def add_page_labels(docx_path, page_mapping):
    """
    根据页码映射插入标签
    page_mapping: dict，键为标题关键词，值为页码
    例如: {'一、招标公告': 3, '二、': 12, '附件B1：': 86}
    """
    doc = Document(docx_path)
    inserted = 0
    
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        
        for keyword, page_num in page_mapping.items():
            if text.startswith(keyword) or keyword in text[:20]:
                style = 'chapter' if p.style.name in ('Heading 1', 'Heading 2') else 'article'
                add_page_label(p, page_num, style)
                inserted += 1
                break
    
    doc.save(docx_path)
    print(f"已插入 {inserted} 个页码标签 → {docx_path}")
    return inserted

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法:")
        print("  python add_page_labels.py output.docx")
        print("\n然后在代码中定义 page_mapping 字典后调用 add_page_labels()")
        sys.exit(1)
    
    # 示例：根据实际文档修改
    page_mapping = {
        '一、招标公告': 3,
        '二、': 12,
        '三、': 55,
        '四、': 60,
        '五、': 86,
        '六、': 125,
        '七、': 38,
    }
    add_page_labels(sys.argv[1], page_mapping)
