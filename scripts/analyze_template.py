#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""分析 .docx 模板文件，提取格式信息作为 JSON 输出。

用法:
    python analyze_template.py <template.docx> [--output format.json]

输出 JSON 结构:
    {
      "page_setup": { "top_mm": 25, "bottom_mm": 20, ... },
      "styles": {
        "Title": { "font_name": "黑体", "font_size_pt": 16, "bold": true, ... },
        "Heading1": { ... },
        ...
      },
      "sections": [ ... ],
      "sample_paragraphs": [ ... ]
    }
"""

import json
import sys
import os
from pathlib import Path

# 在技能安装目录查找 python-docx，回退到系统路径
_script_dir = os.path.dirname(os.path.abspath(__file__))
_skill_dir = os.path.dirname(_script_dir)
_temp_paths = [
    os.path.join(_skill_dir, '..', '.temp'),
    os.path.join(_skill_dir, '.temp'),
]
for _p in _temp_paths:
    if os.path.isdir(_p):
        sys.path.insert(0, _p)

# 确保 python-docx 可用
try:
    from docx import Document
    from docx.shared import Pt, Cm, Inches, Emu
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    print("错误: 需要 python-docx 库。请运行: pip install python-docx")
    sys.exit(1)


def mm_to_emu(mm):
    """毫米转 EMU"""
    return int(mm * 360000)


def emu_to_cm(emu):
    """EMU 转厘米"""
    if emu is None:
        return None
    return round(emu / 360000, 2)


def emu_to_mm(emu):
    """EMU 转毫米"""
    if emu is None:
        return None
    return round(emu / 36000, 1)


def pt_to_half_pt(pt):
    """磅转半磅"""
    if pt is None:
        return None
    return pt * 2


def analyze_alignment(alignment):
    """分析对齐方式"""
    if alignment is None:
        return "left"
    mapping = {
        WD_ALIGN_PARAGRAPH.LEFT: "left",
        WD_ALIGN_PARAGRAPH.CENTER: "center",
        WD_ALIGN_PARAGRAPH.RIGHT: "right",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
    }
    return mapping.get(alignment, "left")


def analyze_font(run):
    """分析文字格式"""
    font = run.font
    result = {}

    if font.name:
        result["font_name"] = font.name
    if font.size:
        result["font_size_pt"] = round(font.size / 12700, 1)
    if font.bold is not None:
        result["bold"] = font.bold
    if font.italic is not None:
        result["italic"] = font.italic
    if font.underline is not None:
        result["underline"] = font.underline
    if font.color and font.color.rgb:
        result["color"] = str(font.color.rgb)

    return result


def analyze_paragraph_format(para):
    """分析段落格式"""
    pf = para.paragraph_format
    result = {}

    if pf.alignment is not None:
        result["alignment"] = analyze_alignment(pf.alignment)

    if pf.line_spacing is not None:
        result["line_spacing"] = round(pf.line_spacing, 1)

    if pf.space_before is not None:
        result["space_before_pt"] = round(pf.space_before / 12700, 1)

    if pf.space_after is not None:
        result["space_after_pt"] = round(pf.space_after / 12700, 1)

    if pf.first_line_indent is not None:
        result["first_line_indent_mm"] = emu_to_mm(pf.first_line_indent)

    return result


def analyze_page_setup(doc):
    """分析页面设置"""
    result = {}
    for i, section in enumerate(doc.sections):
        sec_info = {
            "page_width_mm": emu_to_mm(section.page_width),
            "page_height_mm": emu_to_mm(section.page_height),
            "top_margin_mm": emu_to_mm(section.top_margin),
            "bottom_margin_mm": emu_to_mm(section.bottom_margin),
            "left_margin_mm": emu_to_mm(section.left_margin),
            "right_margin_mm": emu_to_mm(section.right_margin),
            "gutter_mm": emu_to_mm(section.gutter),
        }
        result[f"section_{i}"] = sec_info
    return result


def analyze_styles(doc):
    """分析文档中使用的样式"""
    styles_info = {}

    # 收集所有段落中实际使用的格式模式
    format_patterns = {}
    for para in doc.paragraphs:
        style_name = para.style.name if para.style else "Normal"
        text = para.text.strip()

        if not text:
            continue

        # 分析文本格式
        fonts = []
        for run in para.runs:
            f = analyze_font(run)
            if f:
                fonts.append(f)

        # 分析段落格式
        para_fmt = analyze_paragraph_format(para)

        key = style_name
        if key not in format_patterns:
            format_patterns[key] = {
                "style_name": style_name,
                "fonts": fonts[:3] if fonts else [],
                "paragraph_format": para_fmt,
                "sample_text": text[:80],
                "count": 1,
            }
        else:
            format_patterns[key]["count"] += 1

    # 按出现频率排序
    styles_info = dict(
        sorted(
            format_patterns.items(),
            key=lambda x: x[1]["count"],
            reverse=True,
        )
    )

    return styles_info


def analyze_document_structure(doc):
    """分析文档结构（标题层级等）"""
    structure = []
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        style_name = para.style.name if para.style else "Normal"

        # 判断是否为标题类段落
        is_heading = any(
            keyword in style_name.lower()
            for keyword in ["heading", "title", "toc"]
        )

        entry = {
            "index": i,
            "style": style_name,
            "text_preview": text[:120],
            "is_heading": is_heading,
        }

        # 检测标题层级
        if "heading 1" in style_name.lower() or "heading1" in style_name.lower():
            entry["level"] = 1
        elif "heading 2" in style_name.lower() or "heading2" in style_name.lower():
            entry["level"] = 2
        elif "heading 3" in style_name.lower() or "heading3" in style_name.lower():
            entry["level"] = 3
        elif "title" in style_name.lower():
            entry["level"] = 0

        structure.append(entry)

    return structure[:100]  # 最多返回 100 条


def analyze_template(filepath):
    """分析整个模板文件"""
    doc = Document(filepath)

    result = {
        "file": str(filepath),
        "page_setup": analyze_page_setup(doc),
        "paragraph_count": len(doc.paragraphs),
        "table_count": len(doc.tables),
        "section_count": len(doc.sections),
        "styles": analyze_styles(doc),
        "structure": analyze_document_structure(doc),
    }

    # 分析表格格式
    tables_info = []
    for i, table in enumerate(doc.tables):
        t_info = {
            "index": i,
            "rows": len(table.rows),
            "cols": len(table.columns),
            "sample_cells": [],
        }
        # 采样前几行
        for row_idx, row in enumerate(table.rows[:3]):
            for col_idx, cell in enumerate(row.cells[:5]):
                if cell.text.strip():
                    t_info["sample_cells"].append(
                        {
                            "row": row_idx,
                            "col": col_idx,
                            "text": cell.text.strip()[:60],
                        }
                    )
        tables_info.append(t_info)

    result["tables"] = tables_info

    return result


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    filepath = sys.argv[1]
    if not os.path.exists(filepath):
        print(f"错误: 文件不存在: {filepath}")
        sys.exit(1)

    result = analyze_template(filepath)

    # 输出
    output_path = sys.argv[3] if len(sys.argv) > 3 and sys.argv[2] == "--output" else None

    if output_path:
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print(f"分析结果已保存到: {output_path}")
    else:
        print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
