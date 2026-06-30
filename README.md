# qfnu-thesis

曲阜师范大学论文 .docx 生成器 —— Claude Code 技能。

根据曲阜师范大学论文模板，自动生成格式严格的学术论文 `.docx` 文件。支持内置模板、用户自定义模板库，以及从口头描述或非标准 .docx 归一化创建新模板。

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

| 模板 | 说明 | 来源 |
|------|------|------|
| **理科毕业论文** | 曲阜师范大学本科毕业论文格式（理科），标准学术论文结构 | 内置 |
| **课程论文** | 曲阜师范大学课程论文报告格式，带学校 Logo 封面和信息表 | 内置 |
| **我的模板** | 用户自行归一化或导入的自定义模板，存储在 `templates/` | 用户 |
| **一次性模板** | 从口头描述或非标准 .docx 临时归一化，不持久化 | 临时 |

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

### 创建自定义模板

技能支持将非标准格式输入**归一化**为标准模板 JSON，可选持久化到本地模板库：

- **从 .docx 模板归一化**：`analyze_template.py --output-template` 自动提取样式定义，输出标准模板 JSON
- **从口头描述归一化**：LLM 按 schema 生成模板 JSON
- **持久化**：归一化后可保存到 `templates/<id>.json`，纳入模板库供后续复用
- **临时使用**：不持久化则写入 `.temp/`，本次使用后丢弃

### 内置模板

内置的两个模板以标准模板 JSON 形式存储在 `assets/templates/`：

- `science.json` — 理科毕业论文格式
- `course.json` — 课程论文格式

技能生成论文时直接读取这些 JSON 文件获取格式参数，不依赖 `.md` 参考文件。

## 格式保证

- **页面设置**：从模板 JSON 的 `page_setup` 读取
- **字体层级**：从模板 JSON 的 `styles` 读取每个角色的字体/字号/对齐
- **目录生成**：手动生成带点线前导符和右对齐页码的立即可见目录
- **换页规则**：由模板 JSON 的 `page_order` 决定
- **中文字体修复**：生成后自动修复 python-docx 的 east-asia 字体缺陷
- **封面复制**：封面从源 .docx 模板直接复制，保证 Logo 和布局 1:1 还原

## 目录结构

```
qfnu-thesis/
├── SKILL.md                         # 技能定义
├── scripts/
│   ├── analyze_template.py          # 模板分析工具（支持 --output-template 输出标准模板 JSON）
│   ├── validate_template.py         # 模板 JSON 校验工具
│   └── update_fields.py             # LibreOffice 域更新（TOC/页码）
├── references/
│   ├── template-schema.md           # 标准模板 JSON Schema 定义
│   ├── format-science.md            # 理科毕业论文格式（人类参考）
│   └── format-course.md             # 课程论文格式（人类参考）
├── assets/
│   ├── templates/                   # 内置官方模板 JSON（git 跟踪）
│   │   ├── index.json               # 官方模板索引
│   │   ├── science.json             # 模板：理科毕业论文
│   │   └── course.json              # 模板：课程论文
│   ├── qfnu-template-science.docx   # 模板 A 原始 .docx
│   ├── qfnu-template-course.docx    # 模板 B 原始 .docx
│   └── qfnu-logo.png                # 学校 Logo
├── templates/                       # 用户自定义模板（gitignore 排除）
└── evals/
    └── evals.json                   # 测试用例
```

## 依赖

| 依赖 | 用途 | 安装方式 |
|------|------|----------|
| `python-docx` | Python .docx 读写 | `pip install --target=.temp python-docx` |
| `lxml` | XML 处理（字体修复） | `pip install --target=.temp lxml` |
| `docx` (npm) | JavaScript .docx 创建 | `npm install -g docx` |
| LibreOffice（可选） | TOC 域自动更新 | 系统包管理器安装 |

> 注：`python-docx` 和 `lxml` 建议放置在技能目录下的 `.temp/` 中，避免全局安装。

## 关键设计

### 模板 JSON 是单一真相来源

所有格式参数（字体、字号、行距、页边距、页面顺序、标题层级等）均从标准模板 JSON 文件中读取。`references/format-*.md` 仅为人类参考文档，不直接驱动生成。

### 模板归一化

非标准输入（口头描述、无样式的 .docx）通过 `analyze_template.py` + LLM 推断归一化为标准模板 JSON。归一化结果可临时使用或持久化到本地模板库。

### 模板先于格式

技能启动时不预加载任何格式规则。先确定用户选择哪个模板，再加载对应的模板 JSON。避免了不同模板的格式规则在 LLM 上下文中互相干扰。

### python-docx 中文字体修复

`python-docx` 的 `style.font.name` 只设置西文字体槽位（`w:ascii` / `w:hAnsi`），不设置中文字体槽位（`w:eastAsia`）。默认 `themeFontLang` 指向日语（`ja-JP`），导致中文字符回退为日文字体 MS Mincho。本技能在生成每个 `.docx` 后自动修复此问题。

### 手动目录生成

不依赖 Word 的 TOC 域代码（需要用户手动右键「更新域」）。改为扫描所有标题、估算页码、用制表位 + 点线前导符生成立即可见的目录。

## 许可

MIT
