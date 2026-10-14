---
name: french-pdf-to-chinese-word
description: Translate French PDFs to Chinese Word.
---

# 法语PDF翻译为中文Word文档

将法语PDF翻译为中文Word文档，保留原格式、文档结构、导航目录，并添加原文页码标签方便对照查阅。

## 适用场景

- 法语招标文件、技术规格书、合同等长文档翻译
- 需要保留原文排版和结构的正式文件
- 需要对照原文页码快速查阅的场合
- 文本层PDF和扫描件PDF均可处理

## 核心要求

### 1. 格式保留
- 字体、字号、颜色、加粗、斜体、下划线
- 段落缩进、行距、对齐方式
- 表格结构（行列数、合并单元格、边框）
- 封面、目录、分隔页设计

### 2. 文档结构保留
- 章节层级（标题1/2/3）
- 编号体系（1.1.1、Article 1、Annexe A）
- 目录页（翻译目录文字）
- 附录、附件、清单

### 3. 导航目录
- 可点击的导航目录
- 书签/超链接结构
- 交叉引用保持对应

### 4. 翻译质量
- 法语→中文，工程/商务文档语域
- 技术术语保留原文括注（如 "soupape de sécurité (安全阀)"）
- 公司名称、人名、地名保留原文（SONELGAZ、DCQAV等）
- 标准编号（ISO、EN、API、IEC）不翻译
- 金额、日期、单位保持原格式

### 5. 页码标签
- 每个译文段落前标注原文页码：`[原文页码 X]`
- 章节级标签：`—— 原文第X页 ——`（9pt加粗深灰，居中）
- 条款级标签：`[原文页码 X]`（7pt浅灰，左对齐）
- 页码在PDF解析阶段提取，翻译时携带，Word重建时渲染

## 工作流程

### Step 1: PDF解析

**文本层PDF（如A.O招标文件）：**
```python
import fitz, json
doc = fitz.open('input.pdf')
blocks_with_page = []
for page_num in range(doc.page_count):
    page = doc[page_num]
    blocks = page.get_text("blocks")
    for block in blocks:
        text = block[4].strip()
        if text:
            blocks_with_page.append({"page": page_num + 1, "text": text})
doc.close()
with open('blocks_with_page.json', 'w', encoding='utf-8') as f:
    json.dump(blocks_with_page, f, ensure_ascii=False, indent=2)
```

**扫描件PDF（如技术规格书）：**
```python
import fitz, os
doc = fitz.open('input.pdf')
for i, page in enumerate(doc):
    pix = page.get_pixmap(dpi=200)
    pix.save(f'pages/page_{i+1:03d}.png')
doc.close()
```

### Step 2: 内容提取

**文本层PDF：** 直接使用Step 1提取的文本块

**扫描件PDF：** 用delegate_task派发子代理，每个子代理用vision_analyze读取3-4页图片并转录法语内容。注意429限流，每批不超过5个子代理。

### Step 3: 翻译

- 逐段翻译，保持上下文一致
- 页码标签作为元数据携带，不翻译
- 公司名、标准号保留原文

### Step 4: Word重建

使用python-docx重建文档：
- 封面页（标题+编号+机构）
- 目录页（美化：黑体大标题深蓝色+底部边框，一级条目黑体中蓝色+左侧竖线，二级条目宋体深灰色缩进）
- 正文（H1黑体18pt深蓝+底部边框，H2黑体15pt中蓝，H3黑体13pt浅蓝）
- 表格（Table Grid样式，居中对齐）
- 页码标签（灰色小字插入段落前）

### Step 5: 美化

**目录美化：**
- 标题：黑体26pt、深蓝色(#1F3A5F)、加粗、底部装饰线
- 一级条目：黑体14pt、中蓝色(#2E5A88)、加粗、左侧蓝色竖线
- 二级条目：宋体12pt、深灰色(#555555)、缩进2cm

**标题美化：**
- H1：黑体18pt、深蓝色、加粗、底部边框
- H2：黑体15pt、中蓝色、加粗
- H3：黑体13pt、浅蓝色、加粗
- 正文：1.5倍行距、段后6pt

### Step 6: 校验

**6.1 文档结构校验：**
- 对比原文与译文的章节层级（H1/H2/H3数量及顺序）
- 检查附件编号完整性（A1-A12、B1-B8等）
- 检查条款编号连续性
- 检查标准编号保留

**6.2 内容校验：**
- 法语残留检查（排除公司名称等保留原文的情况）
- 技术术语一致性
- 数字、金额、日期、单位一致性

**6.3 表格校验：**
- 对比原文与译文表格行数、列数
- 检查合并单元格位置
- 检查表头和数据一致性
- **重点：检查是否有因PDF换行导致的重复行**（如供货清单中同一项被拆成两行）

**6.4 页码标签校验：**
- 标签数量
- 页码覆盖范围
- 标签位置是否正确（在标题前）

### Step 7: 修复

根据校验结果修订问题，直到校验通过。

### Step 8: 交付

复制到用户指定目录，文件名为 `原文件名_中文.docx`。

## 常见问题与解决方案

### 1. PDF表格换行导致重复行
**问题：** PDF文本提取时，表格中同一供货项因换行被拆成多个文本块，导致Word表格多出行。
**解决：** 提取后人工检查表格最后几行，删除重复行。校验脚本应增加"表格行数与原文对比"步骤。

### 2. 429 API限流
**问题：** 并行派发大量vision_analyze子代理时触发限流。
**解决：** 每批不超过5个子代理，每个子代理处理3-4页。失败任务等待1分钟后重试。

### 3. python-docx表格列数不匹配
**问题：** 创建表格时cols参数与数据列数不一致，导致IndexError。
**解决：** 用ast.literal_eval检查数据行列数，确保与表格cols一致。

### 4. 公司名称被误判为法语残留
**问题：** 校验脚本检测法语单词时，公司名称（如SONELGAZ-Services Spa）被误报。
**解决：** 法语残留检查应排除已知公司名称和标准编号。

### 5. 页码标签覆盖不全
**问题：** 只在章节级插入页码标签，条款级和正文段落级缺失。
**解决：** 建立章节→页码和条款→页码的映射表，在Word重建时逐级插入。

## 关键脚本模板

### 表格重复行修复
```python
from docx import Document
doc = Document('output.docx')
for table in doc.tables:
    rows_to_remove = []
    for i, row in enumerate(table.rows):
        cells = [c.text for c in row.cells]
        if i > 0:
            prev_cells = [c.text for c in table.rows[i-1].cells]
            if cells[0] and cells[0] in prev_cells[0]:
                rows_to_remove.append(i)
    for i in reversed(rows_to_remove):
        table._tbl.remove(table.rows[i]._tr)
doc.save('output.docx')
```

### 页码标签插入
```python
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
doc = Document('output.docx')
PAGE_GRAY = RGBColor(0x99, 0x99, 0x99)
def add_page_label(paragraph, page_num):
    label_p = paragraph.insert_paragraph_before()
    r = label_p.add_run(f'[原文页码 {page_num}]')
    r.font.size = Pt(7)
    r.font.color.rgb = PAGE_GRAY
    r.font.name = '宋体'
    r._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
```

### 表格列数检查
```python
import ast
with open('gen_script.py', 'r', encoding='utf-8') as f:
    tree = ast.parse(f.read())
for node in ast.walk(tree):
    if isinstance(node, ast.Assign):
        for target in node.targets:
            if isinstance(target, ast.Name):
                try:
                    data = ast.literal_eval(node.value)
                    if isinstance(data, list) and data:
                        for i, row in enumerate(data):
                            if len(row) != expected_cols:
                                print(f'Row {i}: {len(row)} cols')
                except:
                    pass
```

## 文件命名规范

- 输出文件：`原文件名_中文.docx`
- 提取文本：`{文档名}_text.txt`
- 页码映射：`{文档名}_blocks_with_page.json`
- 页面图片：`pages/{文档名}_p{页码:03d}.png`

## 质量检查清单

- [ ] 封面页完整（标题、编号、机构）
- [ ] 目录页美化（黑体标题、蓝色条目）
- [ ] H1/H2/H3标题样式统一
- [ ] 表格无重复行
- [ ] 表格行列数与原文一致
- [ ] 附件编号完整
- [ ] 条款编号连续
- [ ] 公司名称保留原文
- [ ] 标准编号保留
- [ ] 页码标签覆盖关键章节和条款
- [ ] 无法语残留（排除公司名称）
- [ ] 1.5倍行距、段后6pt