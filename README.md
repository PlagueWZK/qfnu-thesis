# qfnu-thesis

曲阜师范大学论文 .docx 生成器 —— Claude Code 技能。

根据曲阜师范大学论文模板，自动生成格式严格的学术论文 `.docx` 文件。支持理科毕业论文、课程论文报告和用户自定义模板三种模式。

## 安装

### 前提条件

- [Claude Code](https://claude.ai/code) 已安装
- [docx 技能](https://github.com/anthropics/claude-code) 已安装：
  ```bash
  claude plugins install docx
  ```

### 安装本技能

```bash
# 克隆到 Claude Code 全局技能目录
git clone https://github.com/PlagueWZK/qfnu-thesis.git ~/.claude/skills/qfnu-thesis
```

或者手动下载后，将整个 `qfnu-thesis/` 目录复制到 `~/.claude/skills/` 下。

## 支持的模板

| 模板 | 说明 | 结构 |
|------|------|------|
| **A — 理科毕业论文** | 曲阜师范大学本科毕业论文格式（理科），标准学术论文结构 | 目录 → 题目 → 中英文摘要 → 正文 → 参考文献 |
| **B — 课程论文** | 曲阜师范大学课程论文报告格式，带学校 Logo 封面和信息表 | 封面 → 目录 → 题目 → 中英文摘要 → 正文 → 参考文献 |
| **自定义模板** | 用户提供的 `.docx` 模板文件 | 使用 `analyze_template.py` 自动分析并适配 |

## 使用方式

在 Claude Code 对话中直接描述需求即可触发：

```
帮我生成一篇理科毕业论文，题目是《基于深度学习的网络入侵检测系统研究》，
计算机科学与技术专业，字数 8000-10000 字。
```

```
帮我生成一篇课程论文，题目是《量子计算在密码学中的应用》，
课程名称是量子信息导论，姓名王明，院系物理工程学院，
指导教师张伟教授，字数 5000-7000 字。
```

用户提到「生成论文」「毕业论文」「课程论文」「QFNU 论文」等关键词时，技能自动激活。

### 自定义模板

如果使用自己的 `.docx` 模板，技能会先用 `analyze_template.py` 分析模板的页面设置、字体层级、样式定义，然后严格按照分析结果生成论文——**不使用 QFNU 默认格式**。

支持情况：
- ✅ 格式完整的 `.docx`（含预定义样式）—— 完美支持
- ⚠️ 仅有直接格式的 `.docx`（手动改字体，无样式）—— 部分支持
- ❌ `.doc` 旧格式 —— 需先用 Word 另存为 `.docx`

## 格式保证

本技能处理以下格式细节，确保产出符合曲阜师范大学论文规范：

- **页面设置**：A4 纸，25mm/20mm 边距，5mm 装订线
- **字体层级**：黑体标题、仿宋体一级标题、宋体正文（模板 A）；课程论文封面使用华文行楷
- **目录生成**：手动生成带点线前导符和右对齐页码的立即可见目录
- **换页规则**：整个正文区域（题目→摘要→正文→参考文献）连续排版，不插入多余换页
- **中文字体修复**：解决 python-docx 默认生成 MS Mincho 日文字体的问题
- **封面复制**：模板 B 封面从模板直接复制，保证 Logo 和布局 1:1 还原

## 目录结构

```
qfnu-thesis/
├── SKILL.md                         # 技能定义（6 步工作流）
├── scripts/
│   ├── analyze_template.py          # 模板格式分析工具（输出 JSON）
│   └── update_fields.py             # LibreOffice 域更新（TOC/页码自动填充）
├── references/
│   ├── format-science.md            # 模板 A 格式规范
│   └── format-course.md             # 模板 B 格式规范
├── assets/
│   ├── qfnu-template-science.docx   # 模板 A：理科毕业论文
│   ├── qfnu-template-course.docx    # 模板 B：课程论文（含封面）
│   └── qfnu-logo.png                # 学校 Logo
└── evals/
    └── evals.json                   # 测试用例
```

## 依赖

| 依赖 | 用途 | 安装方式 |
|------|------|----------|
| `python-docx` | Python .docx 读写 | `pip install python-docx` |
| `lxml` | XML 处理（字体修复） | `pip install lxml` |
| `docx` (npm) | JavaScript .docx 创建 | `npm install -g docx` |
| LibreOffice（可选） | TOC 域自动更新 | 系统包管理器安装 |

> 注：`python-docx` 和 `lxml` 也可放置在技能目录下的 `.temp/` 中，避免全局安装。

## 关键设计

### 模板先于格式

技能启动时不预加载任何格式规则。先确定用户选择哪个模板，再加载对应的格式参考文件。这避免了不同模板的格式规则在 LLM 上下文中互相干扰。

### python-docx 中文字体修复

`python-docx` 的 `style.font.name` 只设置西文字体槽位（`w:ascii` / `w:hAnsi`），不设置中文字体槽位（`w:eastAsia`）。默认 `themeFontLang` 指向日语（`ja-JP`），导致中文字符回退为日文字体 MS Mincho。本技能在生成每个 `.docx` 后自动修复此问题。

### 手动目录生成

不依赖 Word 的 TOC 域代码（需要用户手动右键「更新域」）。改为扫描所有标题、估算页码、用制表位 + 点线前导符生成立即可见的目录。

## 许可

MIT
