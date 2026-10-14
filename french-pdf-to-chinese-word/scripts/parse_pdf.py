#!/usr/bin/env python3
"""PDF解析脚本：提取文本块（带页码标签）和渲染扫描件页面"""
import fitz
import json
import os
import sys

def extract_text_blocks(pdf_path, output_json):
    """文本层PDF：提取文本块并标记页码"""
    doc = fitz.open(pdf_path)
    blocks_with_page = []
    for page_num in range(doc.page_count):
        page = doc[page_num]
        blocks = page.get_text("blocks")
        for block in blocks:
            text = block[4].strip()
            if text:
                blocks_with_page.append({
                    "page": page_num + 1,
                    "text": text
                })
    doc.close()
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(blocks_with_page, f, ensure_ascii=False, indent=2)
    print(f"提取 {len(blocks_with_page)} 个文本块 → {output_json}")
    return blocks_with_page

def render_pages(pdf_path, output_dir, dpi=200):
    """扫描件PDF：渲染页面为图片"""
    os.makedirs(output_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=dpi)
        pix.save(os.path.join(output_dir, f"{base_name}_p{i+1:03d}.png"))
    doc.close()
    print(f"渲染 {doc.page_count} 页 → {output_dir}/")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("用法:")
        print("  文本层PDF: python parse_pdf.py text input.pdf output.json")
        print("  扫描件PDF: python parse_pdf.py render input.pdf output_dir/")
        sys.exit(1)
    
    mode = sys.argv[1]
    pdf_path = sys.argv[2]
    output = sys.argv[3]
    
    if mode == "text":
        extract_text_blocks(pdf_path, output)
    elif mode == "render":
        render_pages(pdf_path, output)
    else:
        print(f"未知模式: {mode}")
        sys.exit(1)
