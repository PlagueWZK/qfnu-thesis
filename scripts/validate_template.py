#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""校验标准模板 JSON 文件是否符合 schema 规范。

用法:
    python scripts/validate_template.py <template.json>

    # 校验模板库中所有模板:
    python scripts/validate_template.py --all

退出码: 0=通过, 1=校验失败
"""

import json
import re
import sys
import os
from pathlib import Path

REQUIRED_TOP_KEYS = {
    "id", "name", "description", "source",
    "page_setup", "styles", "page_order",
    "toc", "title_page", "heading_levels", "references",
}

VALID_SOURCE_VALUES = {"builtin", "imported", "normalized"}
VALID_ALIGNMENTS = {"left", "center", "right", "justify"}
VALID_SECTIONS = {"cover", "toc", "title_page", "abstract_zh", "abstract_en", "body", "references"}
VALID_PLACEHOLDER_TYPES = {"paragraph", "table_cell"}

STYLE_REQUIRED_KEYS = {"font_ascii", "font_east_asia", "font_size_pt"}


def check(condition, message):
    """断言，失败时抛出带消息的异常。"""
    if not condition:
        raise ValueError(message)


def validate_style(name, style, errors):
    """校验单个样式定义。"""
    try:
        check(isinstance(style, dict), f"styles.{name}: 必须是对象")
        for key in STYLE_REQUIRED_KEYS:
            check(key in style, f"styles.{name}: 缺少必需字段 '{key}'")
            if key == "font_size_pt":
                check(isinstance(style[key], (int, float)),
                      f"styles.{name}.font_size_pt: 必须是数字")
            else:
                check(isinstance(style[key], str),
                      f"styles.{name}.{key}: 必须是字符串")

        check(isinstance(style["font_size_pt"], (int, float)) and style["font_size_pt"] > 0,
              f"styles.{name}.font_size_pt: 必须为正数")

        if "alignment" in style:
            check(style["alignment"] in VALID_ALIGNMENTS,
                  f"styles.{name}.alignment: 无效值 '{style['alignment']}'，有效值: {VALID_ALIGNMENTS}")

        if "color" in style:
            check(isinstance(style["color"], str) and len(style["color"]) == 6,
                  f"styles.{name}.color: 必须是 6 位 hex 字符串")

    except ValueError as e:
        errors.append(str(e))


def validate_template(data, filepath="<unknown>"):
    """校验整个模板 JSON。返回错误列表。"""
    errors = []

    try:
        # 顶层必需字段
        for key in REQUIRED_TOP_KEYS:
            check(key in data, f"缺少顶层字段 '{key}'")

        check(data["source"] in VALID_SOURCE_VALUES,
              f"source: 无效值 '{data['source']}'，有效值: {VALID_SOURCE_VALUES}")

        check(re.match(r'^[a-z0-9][a-z0-9-]*$', data["id"]),
              f"id: 格式无效 '{data['id']}'，仅允许小写字母、数字、连字符")

        # page_setup
        ps = data["page_setup"]
        check("paper" in ps and "margins" in ps, "page_setup: 缺少 paper 或 margins")
        check("width_mm" in ps.get("paper", {}) and "height_mm" in ps.get("paper", {}),
              "page_setup.paper: 缺少 width_mm 或 height_mm")
        for m in ("top_mm", "bottom_mm", "left_mm", "right_mm"):
            check(m in ps.get("margins", {}), f"page_setup.margins: 缺少 {m}")

        # styles
        check(isinstance(data["styles"], dict) and len(data["styles"]) > 0,
              "styles: 必须是非空对象")
        for name, style in data["styles"].items():
            validate_style(name, style, errors)

        # page_order
        po = data["page_order"]
        check(isinstance(po, list) and len(po) > 0, "page_order: 必须是非空数组")
        seen = set()
        for i, entry in enumerate(po):
            check(isinstance(entry, dict), f"page_order[{i}]: 必须是对象")
            check("section" in entry, f"page_order[{i}]: 缺少 section")
            check(entry["section"] in VALID_SECTIONS,
                  f"page_order[{i}].section: 无效值 '{entry['section']}'")
            check(entry["section"] not in seen,
                  f"page_order[{i}].section: 重复的 section '{entry['section']}'")
            seen.add(entry["section"])
            if "page_break_after" in entry:
                check(isinstance(entry["page_break_after"], bool),
                      f"page_order[{i}].page_break_after: 必须是布尔值")

        # cover (可选)
        if "cover" in data and data["cover"] is not None:
            cover = data["cover"]
            if cover.get("enabled"):
                if cover.get("source_template"):
                    check(isinstance(cover["source_template"], str),
                          "cover.source_template: 必须是字符串")
                if cover.get("paragraphs_boundary"):
                    pb = cover["paragraphs_boundary"]
                    check("start" in pb and "end" in pb,
                          "cover.paragraphs_boundary: 缺少 start 或 end")
                if "placeholders" in cover:
                    for pk, pv in cover["placeholders"].items():
                        check("type" in pv, f"cover.placeholders.{pk}: 缺少 type")
                        check(pv["type"] in VALID_PLACEHOLDER_TYPES,
                              f"cover.placeholders.{pk}.type: 无效值 '{pv['type']}'")

        # toc
        toc = data["toc"]
        check("title_style" in toc, "toc: 缺少 title_style")
        check(toc["title_style"] in data["styles"], f"toc.title_style: 样式 '{toc['title_style']}' 未在 styles 中定义")
        check("levels" in toc, "toc: 缺少 levels")
        for i, lv in enumerate(toc.get("levels", [])):
            check("heading_level" in lv, f"toc.levels[{i}]: 缺少 heading_level")
            check("style" in lv, f"toc.levels[{i}]: 缺少 style")
            check(lv["style"] in data["styles"], f"toc.levels[{i}].style: 样式 '{lv['style']}' 未在 styles 中定义")

        # title_page
        tp = data["title_page"]
        check("title_cn_style" in tp, "title_page: 缺少 title_cn_style")
        check(tp["title_cn_style"] in data["styles"],
              f"title_page.title_cn_style: 样式 '{tp['title_cn_style']}' 未在 styles 中定义")

        # heading_levels
        hl = data["heading_levels"]
        check(isinstance(hl, list) and len(hl) > 0, "heading_levels: 必须是非空数组")
        for i, h in enumerate(hl):
            check("level" in h and "style" in h, f"heading_levels[{i}]: 缺少 level 或 style")
            check(h["style"] in data["styles"],
                  f"heading_levels[{i}].style: 样式 '{h['style']}' 未在 styles 中定义")

        # references
        ref = data["references"]
        check("title_style" in ref, "references: 缺少 title_style")
        check("body_style" in ref, "references: 缺少 body_style")
        check(ref["body_style"] in data["styles"],
              f"references.body_style: 样式 '{ref['body_style']}' 未在 styles 中定义")

    except ValueError as e:
        errors.append(str(e))
    except KeyError as e:
        errors.append(f"缺少字段: {e}")

    return errors


def main():
    import argparse

    parser = argparse.ArgumentParser(description="校验 qfnu-thesis 标准模板 JSON")
    parser.add_argument("file", nargs="?", help="模板 JSON 文件路径")
    parser.add_argument("--all", action="store_true", help="校验模板库中所有模板")
    args = parser.parse_args()

    if args.all:
        skill_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        builtin_dir = os.path.join(skill_dir, "assets", "templates")
        user_dir = os.path.join(skill_dir, "templates")

        all_ok = True
        for label, directory in [("内置", builtin_dir), ("用户", user_dir)]:
            if not os.path.isdir(directory):
                continue
            for fname in sorted(os.listdir(directory)):
                if not fname.endswith(".json") or fname == "index.json":
                    continue
                fpath = os.path.join(directory, fname)
                with open(fpath, "r", encoding="utf-8") as f:
                    data = json.load(f)
                errors = validate_template(data, fpath)
                if errors:
                    all_ok = False
                    print(f"❌ [{label}] {fname}:")
                    for e in errors:
                        print(f"   - {e}")
                else:
                    print(f"✅ [{label}] {fname}")
        sys.exit(0 if all_ok else 1)

    if not args.file:
        parser.print_help()
        sys.exit(1)

    with open(args.file, "r", encoding="utf-8") as f:
        data = json.load(f)

    errors = validate_template(data, args.file)
    if errors:
        print(f"❌ {args.file}:")
        for e in errors:
            print(f"   - {e}")
        sys.exit(1)
    else:
        print(f"✅ {args.file}: 校验通过 (id={data.get('id')}, styles={len(data.get('styles', {}))})")


if __name__ == "__main__":
    main()
