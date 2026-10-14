#!/usr/bin/env python3
"""表格重复行修复脚本"""
import sys
from docx import Document

def fix_duplicate_rows(docx_path):
    doc = Document(docx_path)
    fixed = 0
    
    for tbl_idx, table in enumerate(doc.tables):
        rows_to_remove = []
        for i, row in enumerate(table.rows):
            if i == 0:
                continue
            cells = [c.text.strip() for c in row.cells]
            prev_cells = [c.text.strip() for c in table.rows[i-1].cells]
            
            # 检查当前行是否完全包含在前一行中（PDF换行导致的重复）
            if cells[0] and cells[0] in prev_cells[0]:
                rows_to_remove.append(i)
                print(f"  表格{tbl_idx+1}: 删除第{i+1}行 '{cells[0][:40]}'")
        
        for i in reversed(rows_to_remove):
            table._tbl.remove(table.rows[i]._tr)
            fixed += 1
    
    if fixed > 0:
        doc.save(docx_path)
        print(f"\n共修复 {fixed} 行重复数据 → {docx_path}")
    else:
        print("未发现重复行")
    
    return fixed

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python fix_duplicate_rows.py output.docx")
        sys.exit(1)
    fix_duplicate_rows(sys.argv[1])
