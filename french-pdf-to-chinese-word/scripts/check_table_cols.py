#!/usr/bin/env python3
"""表格列数检查脚本：防止python-docx表格列数不匹配"""
import ast
import sys

def check_table_cols(script_path, var_names=None, expected_cols=None):
    """
    检查Python脚本中表格数据的列数
    var_names: 要检查的变量名列表，如 ['items', 'rows_data']
    expected_cols: 期望的列数
    """
    with open(script_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    tree = ast.parse(content)
    found_issues = []
    
    for node in ast.walk(tree):
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name):
                    var_name = target.id
                    if var_names and var_name not in var_names:
                        continue
                    try:
                        data = ast.literal_eval(node.value)
                        if isinstance(data, list) and data:
                            for i, row in enumerate(data):
                                if isinstance(row, tuple) or isinstance(row, list):
                                    cols = len(row)
                                    if expected_cols and cols != expected_cols:
                                        found_issues.append(f"{var_name}[{i}]: {cols}列 (期望{expected_cols}列) - {row[0] if row else 'empty'}")
                                    elif not expected_cols:
                                        print(f"{var_name}[{i}]: {cols}列")
                    except:
                        pass
    
    if found_issues:
        print(f"发现 {len(found_issues)} 个列数问题:")
        for issue in found_issues:
            print(f"  ❌ {issue}")
    elif expected_cols:
        print(f"✅ 所有数据均为 {expected_cols} 列")
    
    return found_issues

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法:")
        print("  python check_table_cols.py script.py")
        print("  python check_table_cols.py script.py --vars items,rows_data --cols 4")
        sys.exit(1)
    
    script_path = sys.argv[1]
    var_names = None
    expected_cols = None
    
    if '--vars' in sys.argv:
        idx = sys.argv.index('--vars')
        var_names = sys.argv[idx+1].split(',')
    if '--cols' in sys.argv:
        idx = sys.argv.index('--cols')
        expected_cols = int(sys.argv[idx+1])
    
    check_table_cols(script_path, var_names, expected_cols)
