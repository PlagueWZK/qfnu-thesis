# 标准模板 JSON Schema

本文档定义 qfnu-thesis 技能的标准模板 JSON 格式。所有内置模板和用户归一化后的模板都必须符合此结构。

---

## 顶层结构

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `id` | string | ✅ | 唯一标识，仅含小写字母、数字、连字符。例：`science`、`course`、`my-math-template` |
| `name` | string | ✅ | 显示名称。例：「理科毕业论文」 |
| `description` | string | ✅ | 简短说明，1-2 句话描述模板用途 |
| `source` | string | ✅ | `builtin`（官方）/ `imported`（从 docx 导入）/ `normalized`（从描述归一化） |
| `page_setup` | object | ✅ | 页面设置 |
| `styles` | object | ✅ | 样式字典，key 为样式名，value 为样式定义 |
| `page_order` | array | ✅ | 页面顺序及换页规则 |
| `cover` | object | ❌ | 封面定义，无封面则省略或 `enabled: false` |
| `toc` | object | ✅ | 目录格式 |
| `title_page` | object | ✅ | 标题页格式 |
| `heading_levels` | array | ✅ | 正文标题层级定义 |
| `references` | object | ✅ | 参考文献格式 |
| `word_count_estimate` | object | ❌ | 字数估算参数，默认 `chars_per_page_single_spacing: 1500` |

---

## `page_setup` — 页面设置

```json
{
  "paper": { "width_mm": 210, "height_mm": 297 },
  "margins": { "top_mm": 25, "bottom_mm": 20, "left_mm": 25, "right_mm": 20 },
  "gutter_mm": 5,
  "default_line_spacing": 1.0
}
```

| 字段 | 类型 | 说明 |
|------|------|------|
| `paper` | object | 纸张尺寸，`width_mm` / `height_mm` |
| `margins` | object | 页边距，`top_mm` / `bottom_mm` / `left_mm` / `right_mm` |
| `gutter_mm` | number | 装订线宽度 |
| `default_line_spacing` | number | 全文默认行距倍数（1.0=单倍, 1.5=1.5倍, 2.0=双倍） |

---

## `styles` — 样式字典

样式名（key）为字符串，对应 Word 样式名称或语义角色。推荐的样式名约定：

| 样式名 | 语义 | 典型用途 |
|--------|------|---------|
| `BodyText` | 正文 | 普通段落 |
| `Heading1` | 一级标题 | 章标题 |
| `Heading2` | 二级标题 | 节标题 |
| `Heading3` | 三级标题 | 条标题 |
| `Title` | 论文题目（中文） | 标题页 |
| `Title_EN` | 论文题目（英文） | 标题页 |
| `Abstract` | 中文摘要正文 | 摘要页 |
| `Abstract_EN` | 英文摘要正文 | 摘要页 |
| `AbstractLabel` | 摘要标签 | 「摘要」二字 |
| `Keywords` | 中文关键词 | 关键词 |
| `Keywords_EN` | 英文关键词 | Key words |
| `TOC_Title` | 目录标题 | 「目  录」 |
| `TOC_Level1` | 目录一级条目 | 章条目 |
| `TOC_Level2` | 目录二级条目 | 节条目 |
| `Reference` | 参考文献正文 | 参考文献条目 |
| `FigureCaption` | 图题 | 图下方 |
| `TableCaption` | 表题 | 表上方 |
| `TitleInfo` | 标题页信息 | 姓名/导师/院系 |
| `HeaderFooter` | 页眉页脚 | 页码 |

每个样式定义的结构：

```json
{
  "font_ascii": "Times New Roman",
  "font_east_asia": "宋体",
  "font_size_pt": 12,
  "bold": false,
  "italic": false,
  "color": "000000",
  "alignment": "justify",
  "line_spacing": 1.0,
  "space_before_pt": 0,
  "space_after_pt": 0,
  "first_line_indent_chars": 2
}
```

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `font_ascii` | string | ✅ | 拉丁字符字体名（如 "Times New Roman"） |
| `font_east_asia` | string | ✅ | CJK 字符字体名（如 "宋体"、"黑体"） |
| `font_size_pt` | number | ✅ | 字号（磅） |
| `bold` | boolean | ❌ | 是否加粗，默认 false |
| `italic` | boolean | ❌ | 是否斜体，默认 false |
| `color` | string | ❌ | RGB 颜色，6 位 hex，默认 "000000" |
| `alignment` | string | ❌ | 对齐：`left` / `center` / `right` / `justify`，默认 "left" |
| `line_spacing` | number | ❌ | 行距倍数，继承 `page_setup.default_line_spacing` |
| `space_before_pt` | number | ❌ | 段前间距（磅），默认 0 |
| `space_after_pt` | number | ❌ | 段后间距（磅），默认 0 |
| `first_line_indent_chars` | number | ❌ | 首行缩进字符数，0=无缩进，2=两字符缩进 |

---

## `page_order` — 页面顺序与换页规则

```json
[
  { "section": "cover", "page_break_after": true },
  { "section": "toc", "page_break_after": true },
  { "section": "title_page", "page_break_after": false },
  { "section": "abstract_zh", "page_break_after": false },
  { "section": "abstract_en", "page_break_after": false },
  { "section": "body", "page_break_after": false },
  { "section": "references", "page_break_after": false }
]
```

| `section` 值 | 说明 |
|-------------|------|
| `cover` | 封面（学校名、Logo、信息表） |
| `toc` | 目录 |
| `title_page` | 论文题目 + 学生/导师信息 |
| `abstract_zh` | 中文摘要 + 关键词 |
| `abstract_en` | 英文摘要 + Keywords |
| `body` | 正文各章节 |
| `references` | 参考文献 |

`page_break_after: true` 表示该 section 结束后换页。

**换页符数量** = `page_break_after: true` 的条目数。例如模板 A 为 1 处（toc→title_page），模板 B 为 2 处（cover→toc, toc→title_page）。

---

## `cover` — 封面定义（可选）

```json
{
  "enabled": true,
  "source_template": "assets/qfnu-template-course.docx",
  "paragraphs_boundary": { "start": 0, "end": 16 },
  "placeholders": {
    "course_name": {
      "type": "paragraph",
      "style": "CoverCourse",
      "default": "<课程名称>"
    },
    "student_name": {
      "type": "table_cell",
      "table_index": 0,
      "label": "姓    名",
      "default": "学生姓名"
    },
    "department": {
      "type": "table_cell",
      "table_index": 0,
      "label": "院    系",
      "default": "xxxx学院"
    },
    "major": {
      "type": "table_cell",
      "table_index": 0,
      "label": "专    业",
      "default": "xxxx专业"
    },
    "advisor": {
      "type": "table_cell",
      "table_index": 0,
      "label": "指导教师",
      "default": "指导教师姓名"
    },
    "date": {
      "type": "paragraph",
      "style": "StudentInfo",
      "format": "年   月   日"
    }
  }
}
```

| 字段 | 类型 | 说明 |
|------|------|------|
| `enabled` | boolean | 是否有封面 |
| `source_template` | string\|null | 封面来源 .docx 路径（相对于技能根目录），从该文件复制封面。`null` 表示从零构建 |
| `paragraphs_boundary` | object\|null | 封面在源文件中的段落范围，`{start, end}`。`null` 表示需遍历确认 |
| `placeholders` | object | 占位符字典，key 为语义名，value 定义位置和默认值 |

**占位符类型**：

| `type` | 说明 |
|--------|------|
| `paragraph` | 段落级占位符，通过 `style` 定位段落并替换其文本 |
| `table_cell` | 表格单元格占位符，通过 `table_index` + `label` 定位标签相邻单元格 |

**占位符替换规则**：
- 多 run 段落（如 bold 标签 + 非 bold 空格）必须在替换时 `para.clear()` 后 `add_run()`
- 若用户未提供某占位符的值，使用 `default`

---

## `toc` — 目录格式

```json
{
  "title_style": "TOC_Title",
  "levels": [
    { "heading_level": 1, "style": "TOC_Level1", "indent_chars": 0 },
    { "heading_level": 2, "style": "TOC_Level2", "indent_chars": 2 }
  ],
  "tab_stop_cm": 14.5
}
```

| 字段 | 类型 | 说明 |
|------|------|------|
| `title_style` | string | 目录标题「目  录」使用的样式名 |
| `levels` | array | 目录层级定义 |
| `levels[].heading_level` | number | 对应正文标题层级（1/2/3） |
| `levels[].style` | string | 该级目录条目使用的样式名 |
| `levels[].indent_chars` | number | 缩进字符数（0=顶格） |
| `tab_stop_cm` | number | 右对齐制表位位置（cm），用于页码对齐+点线前导符 |

---

## `title_page` — 标题页格式

```json
{
  "title_cn_style": "Title",
  "title_en_style": "Title_EN",
  "info_styles": ["TitleInfo"],
  "info_fields": [
    { "key": "student", "label": "学生姓名", "style": "TitleInfo" },
    { "key": "advisor", "label": "指导教师", "style": "TitleInfo" }
  ]
}
```

---

## `heading_levels` — 正文标题层级

```json
[
  { "level": 1, "style": "Heading1", "numbering_pattern": "\\d+\\s+" },
  { "level": 2, "style": "Heading2", "numbering_pattern": "\\d+\\.\\d+\\s+" },
  { "level": 3, "style": "Heading3", "numbering_pattern": "\\d+\\.\\d+\\.\\d+\\s+" }
]
```

| 字段 | 类型 | 说明 |
|------|------|------|
| `level` | number | 标题层级（1/2/3） |
| `style` | string | 使用的样式名（必须在 `styles` 中定义） |
| `numbering_pattern` | string | 编号的正则表达式，用于目录扫描时识别标题 |

---

## `references` — 参考文献格式

```json
{
  "title_style": "Heading1",
  "body_style": "Reference",
  "title_text": "参考文献"
}
```

---

## `word_count_estimate` — 字数估算参数（可选）

```json
{
  "chars_per_page_single_spacing": 1500
}
```

用于目录页码估算公式：`每页字符数 = chars_per_page_single_spacing / line_spacing`

---

## 完整示例

参见 `assets/templates/science.json` 和 `assets/templates/course.json`。
