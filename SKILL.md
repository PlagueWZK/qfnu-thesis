---
name: qfnu-thesis
description: >
  根据模板生成严格格式的论文 .docx 文件。当用户提到「生成论文」「写毕业论文」
  「学位论文」「曲阜师范大学论文」「QFNU论文格式」「课程论文」「生成docx论文」
  或需要将内容排版成正式学术论文格式时，必须使用此技能。即使用户没有明确说
  「按模板」，只要涉及生成学术论文 .docx 文件就应该触发。支持内置模板、用户
  自定义模板库，以及从口头描述或非标准 .docx 归一化创建新模板。此技能依赖
  docx 技能处理底层 Word 文档操作。
---

# QFNU 论文生成器

根据用户选择的模板文件生成格式严格的学术论文 .docx 文件。

## 核心原则

- **模板先于格式**：启动时不加载任何格式规则——先确定用哪个模板，再加载对应
  的格式规范。这避免了不同模板的格式规则在 LLM 上下文中互相干扰产生幻觉。
- **模板 JSON 是单一真相来源**：所有格式参数（字体、字号、行距、页边距、页面顺序、
  标题层级等）均从标准模板 JSON 文件中读取。`references/format-*.md` 仅为人类参考
  文档，不直接驱动生成。
- **格式精确性**：字体、字号、行距、页边距必须与选定模板一致。学术论文格式是
  硬性要求，格式错误会导致论文被退回。
- **内容适应性**：根据用户提供的主题和专业方向自动调整论文章节结构。
- **依赖 docx 技能**：所有底层 .docx 操作通过 `docx` 技能完成。本技能负责高层
  逻辑：模板选择、需求收集、结构规划、格式指导、质量检查。

## 工作流程

### 第零步：环境检查（必须最先执行）

检查当前环境是否具备以下依赖：

#### 0.1 docx 技能检查

查看可用技能列表中是否存在 `docx`：

- **docx 技能可用** → 继续
- **docx 技能不可用** → **停止**，引导用户安装：

  > ⚠️ 此技能依赖 **docx 技能** 来处理 Word 文档底层操作。未检测到 docx 技能。
  > 请先安装：在 Claude Code 中运行 `/plugin install docx`，安装后重新触发本技能。

#### 0.2 Python 依赖检查

本技能的模板分析和字体修复脚本依赖 `python-docx` 和 `lxml`。运行以下命令检查：

```bash
python -c "import docx; import lxml; print('OK')" 2>&1
```

- **输出 `OK`** → 依赖就绪，继续
- **报 `ModuleNotFoundError`** → 先安装依赖再继续：

  > ⚠️ 缺少 Python 依赖 `python-docx` 和/或 `lxml`。
  > 请运行以下命令安装后重新触发本技能：
  > ```
  > pip install --target=.temp python-docx lxml
  > ```

项目 `.temp/` 目录已配置在 `analyze_template.py` 的搜索路径中，安装到该目录可避免全局污染。

#### 0.3 模板库初始化

确保用户模板目录存在：

```bash
mkdir -p templates/
```

读取 `assets/templates/index.json` 获取内置模板列表；扫描 `templates/` 目录获取用户自定义模板列表（`.json` 文件，排除 `index.json`）。合并时对同名 ID，用户模板优先。

### 第一步：确定模板（先于一切格式加载）

**在加载任何格式规则之前**，先和用户确认使用哪个模板。不要在此时读取参考文件。

从第零步初始化的模板库中列出可用模板，分两组展示：

#### 官方模板

从 `assets/templates/index.json` 读取并列出。

#### 我的模板

扫描 `templates/*.json` 列出用户自定义模板。

向用户展示时用表格：
```
📁 官方模板:
  1. science — 理科毕业论文
  2. course  — 课程论文

📁 我的模板 (templates/):
  3. my-math-template — 数学学院毕业论文格式
  (无用户模板则显示「暂无自定义模板」)

🆕 创建新模板 — 从口头描述或 .docx 文件归一化为标准模板
```

用户选择后，**加载对应模板 JSON** 作为格式基准：

- 选内置模板 → `json.load(open("assets/templates/<id>.json"))`
- 选用户模板 → `json.load(open("templates/<id>.json"))`
- 选「创建新模板」→ 进入 **第一步附：模板归一化** 流程（见下方）

加载后将模板 JSON 的核心字段（`id`, `name`, `styles`, `page_order`, `cover` 等）记入上下文，后续所有格式决策以此为准。

#### 第一步附：模板归一化（创建新模板）

当用户选择「创建新模板」或用户提供的模板不是标准模板 JSON 时，进入归一化流程。

**支持的输入类型与处理方式**：

| 输入类型 | 处理方式 |
|----------|---------|
| 标准 .docx（样式齐全） | `analyze_template.py --output-template` 自动提取 → 展示预览 → 确认 |
| 非标准 .docx（仅直接格式） | `analyze_template.py` 提取 + LLM 推断样式映射 → 展示预览 → 确认 |
| 口头描述（"正文楷体四号, 标题黑体三号…"） | LLM 按 schema 生成模板 JSON → 展示预览 → 确认 |
| 混合（标准模板 + 口头覆盖部分样式） | 加载标准模板 JSON，用户指定的样式项覆盖对应字段 |

**归一化流程**：
1. 接收用户输入
2. 提取/推断格式参数
3. 生成标准模板 JSON（符合 `references/template-schema.md`）
4. 展示格式预览（列出样式名→字体/字号/行距/对齐的映射），请用户确认
5. **询问是否持久化**：
   - **不持久化（一次性使用）**：JSON 写入 `.temp/normalized_template.json`，本次生成使用后丢弃
   - **持久化（纳入模板库）**：询问用户模板 ID（小写字母+数字+连字符，如 `math-dept`），保存到 `templates/<id>.json`
6. 继续后续论文生成流程

**设计理由**：归一化有脚本分析+LLM推断的成本，但不是每次都要存储。一次性测试不污染模板库；院系统一格式等反复场景才持久化。

### 第二步：收集需求（必须执行）

**如果用户未提供字数/篇幅要求，必须主动询问，不得跳过。**

需确认的内容：

| 信息项 | 说明 |
|--------|------|
| **字数/篇幅** | 必须明确。如未提供，询问期望字数范围 |
| 论文题目 | 用户已提供则直接使用 |
| 专业方向 | 影响章节结构（计算机 ≠ 物理 ≠ 数学） |
| **个人信息** | 若模板含封面/信息表，必须收集。见下方说明 |
| 提纲/结构 | 用户是否自带大纲？如有则优先遵循 |
| 参考文献 | 用户是否提供了参考文献列表？ |
| **行距/页边距等格式偏好** | 用户是否有特殊格式要求？如 1.5 倍行距、特定页边距等 |
| 输出路径 | 生成的 .docx 保存位置 |

#### 格式冲突检测（必须执行）

收集完需求后，将用户的格式要求与**模板 JSON 中 `styles` 和 `page_setup` 的值**逐项对比。若存在冲突（例如用户要求 1.5 倍行距但模板 JSON 中 `default_line_spacing` 为 1.0），**必须明确告知用户冲突项**，并按以下优先级处理：

| 优先级 | 来源 | 说明 |
|--------|------|------|
| **1（最高）** | 用户明确指定的格式要求 | 用户口头/书面明确要求的格式，覆盖模板默认值 |
| **2** | 模板文件中的默认格式 | 模板 .docx 中定义的样式和页面设置 |
| **3（最低）** | 技能推荐的格式 | 参考文档中的一般性建议 |

在执行后续生成步骤时，始终遵守此优先级。用户未明确要求的格式项，使用模板默认值。

#### 个人信息收集（若模板含封面）

模板 B（课程论文）和部分自定义模板的封面包含个人信息表。此时**主动询问**用户以下信息：

- 姓名
- 学号（如有）
- 院系（如：物理工程学院）
- 专业（如：软件工程）
- 指导教师 + 职称（如：杨玉鹏 教授）
- 课程名称（如：高等数学）

若用户未提供以上信息，封面中的对应位置填入占位符（如「学生姓名」「指导教师姓名」），
并在生成完成后提示用户自行替换。

### 第三步：规划论文结构

根据专业方向规划章节结构。模板 A（理科）和模板 B（课程论文）的默认结构不同，
参见各自格式参考文件中的说明。

**如果用户只给了题目**：根据专业方向生成大纲，**展示给用户确认**后再继续。

**如果用户提供了详细大纲**：遵循用户的大纲，检查合理性并给出建议。

### 第三步附：中文引号写入注意事项（生成脚本前必读）

在生成 .docx 的 Python 脚本中嵌入中文双引号 `"` `"`（U+201C / U+201D）时，Write 工具有可能将这两个字符转换为 ASCII 双引号 `"`（U+0022），导致 Python 字符串被意外截断，引发 `SyntaxError`。

**推荐做法（任选一种）**：

1. **使用 Unicode 转义序列**（最可靠）：
   ```python
   LQ = "“"  # 左双引号 "
   RQ = "”"  # 右双引号 "
   text = f"{LQ}人工智能{RQ}是当前热门研究领域"
   ```

2. **先用角括号占位，生成后全局替换**（备选）：
   - 脚本中用 `「` `」` 代替中文双引号写入正文
   - 生成 .docx 后，解包 → 对所有 `<w:t>` 文本做全局替换 `「`→`"`、`」`→`"`
   - 重新打包

3. **使用变量定义常量**（推荐，便于复用）：
   ```python
   # 在脚本开头定义
   LQ = "“"   # 中文左双引号
   RQ = "”"   # 中文右双引号
   ```

所有包含中文双引号的正文内容（如引用的学术观点、对话等）都必须使用上述方式处理，**严禁直接写 `"..."` 到 Python 字符串中**。

### 第四步：生成论文 .docx

调用 `docx` 技能创建文档。向 `docx` 技能传递已加载的格式规范。

#### 4.1 页面顺序（关键）

文档中各部分的排列顺序**必须严格遵循模板**。不同模板的页面顺序不同：

**模板 A（理科毕业论文）**：
```
第1页: 目录
第2页: 论文中文题目 + 学生/导师信息
第3页: 中文摘要 + 关键词 → 英文摘要 + Keywords
第4页起: 正文各章节（连续，章节间不换页）
最后: 参考文献
```

**模板 B（课程论文）**：
```
第1页: 封面（学校名 → 课程名 → 报告类型 → Logo → 信息表 → 日期）
第2页: 目录
第3页: 论文题目 + 学生/导师信息
第4页: 中文摘要 + 关键词 → 英文摘要 + Keywords
第5页起: 正文各章节（连续）
最后: 参考文献
```

**换页规则（极其重要 — 用户直接反馈）**：

整个文档中**只有一处换页**：目录与论文题目之间。具体规则：

```
目录 → [唯一允许的换页] → 论文题目 → 学生信息 → 摘要 → 关键词
→ Abstract → Key words → 正文各章 → 参考文献
```

即：从论文题目开始，到参考文献结束，**整段区域不插入任何换页符**。
标题页、摘要、正文、参考文献全部连续排版，内容自然流动到下一页。
正文各章节之间通过一级标题自然分隔，不插入 `add_page_break()`。

**绝对禁止**：
- ❌ 题目和摘要之间换页
- ❌ 摘要和正文第一章之间换页
- ❌ 正文最后一章和参考文献之间换页
- ❌ 正文各章节之间换页

**模板 B 额外注意**：封面独占一页。封面→目录之间有换页。目录→题目之间
也有换页。即模板 B 有两处换页（封面→目录、目录→题目），题目开始后整段连续。

**检测文档中的分页符**：python-docx 高层 API（`run.text`/`para.runs`）不能直接检测分页符。分页符可能存在于两个位置，需要同时检查：

```python
from lxml import etree

def count_page_breaks(doc):
    """统计文档中的分页符数量（含段落内和段落间）"""
    count = 0
    nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    body = doc.element.body

    # 1. 段落内的分页符：<w:r><w:br w:type="page"/></w:r>
    for br in body.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br'):
        if br.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type') == 'page':
            count += 1

    # 2. 段落级别的分页属性：<w:pPr><w:pageBreakBefore/></w:pPr>
    for pPr in body.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr'):
        if pPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pageBreakBefore') is not None:
            count += 1

    return count
```

生成完成后使用此函数验证换页符数量是否符合模板预期（模板 A: 1 处；模板 B: 2 处）。

#### 4.1.1 模板 B 封面生成（从模板复制，不从零重建）

模板 B（课程论文）的封面包含学校 Logo 和信息表，结构复杂，**严禁从零重建**。
必须从模板文件 `assets/qfnu-template-course.docx` 直接复制封面部分。

**操作步骤**：
1. 用 python-docx 打开 `assets/qfnu-template-course.docx`
2. **先确认封面边界**：遍历文档前 20 个段落，找到最后一个封面相关段落（含换页符的空段落通常标志封面结束）。模板 B 中封面通常占用段落 0 到段落 16（含：学校名、课程名、报告类型、Logo、信息表、日期、以及末尾带换页符的空段落）。**不要硬编码边界值**——每次执行前确认实际的 `para_count` 边界。
3. 定位并处理封面元素：
   - 学校名段落 → "曲阜师范大学"（华文行楷 42pt 初号，保持不变）
   - 课程名段落 → 将占位文本（如 `<课程名称>` 或模板原有课程名）替换为用户指定的课程名（华文行楷 42pt 初号 bold）
   - 报告类型段落 → "课程论文报告"（黑体 42pt 初号，保持不变）
   - Logo 图片 → 保持不变（1.5英寸 × 1.5英寸，居中）
   - 信息表（5行×4列）→ 填入题目/姓名/院系/专业/指导教师（仿宋_GB2312 12pt）
   - 日期段落 → 填入当前日期（宋体 15pt bold，格式：年   月   日）
4. 封面之后的模板示例内容（从封面结束段落之后到参考文献示例）全部删除
5. 保留 PaperH1/PaperH2/PaperH3 样式定义（TOC内容将在生成后手动填充）
6. 在清理后的文档基础上继续生成目录、题目页、摘要、正文、参考文献

**这样可以保证封面 100% 与模板一致**，包括字体、字号、加粗、对齐、Logo 尺寸、
表格列宽和边框等所有细节。

#### 4.1.1.1 封面信息分类与填充

封面元素分为两类，处理方式不同：

**A. 固定元素（直接复制，绝不改动）**：
- 学校名称、Logo 图片、报告类型文字
- 所有字体、字号、加粗、对齐、间距等格式属性
- 表格结构（行列数、列宽、边框）

**B. 可变信息（替换占位符或填入）**：
| 模板中的占位符/位置 | 替换为 |
|-------------------|--------|
| `<课程名称>` | 用户指定的课程名 |
| 个人信息表各标签后的空白 | 用户提供的姓名/院系/专业等 |
| `年   月   日` | 当前日期 |
| 模板示例标题（如"论文题目"）| 用户实际论文题目 |
| `xxxx专业学生` | 用户专业 |
| `学生姓名` | 用户姓名 |
| `指导教师姓名` | 用户导师姓名 |

**填充格式要求**（以模板 B 为例）：
- 个人信息填入对应标签**同一行、标签之后**
- 字体与标签一致（宋体三号加粗）
- 居中或与标签对齐
- 不要在标签和内容之间加多余的空格或分隔符

填充后的封面效果示例（注意：实际字体为华文行楷/黑体初号42pt，非黑体小号字）：
```
                  曲阜师范大学          ← 华文行楷 42pt (初号)
                    高等数学            ← 华文行楷 42pt bold
                 课程论文报告           ← 黑体 42pt (初号)

                   [学校 Logo]

              题    目    暗物质研究
              姓    名    王政凯       ← 信息表为仿宋_GB2312 12pt
              院    系    物理工程学院
              专    业    软件工程
         指导教师    杨玉鹏    职  称    教授

                 2026年6月30日          ← 宋体 15pt bold
```

#### ⚠️ 封面多 run 段落替换技巧

模板封面的部分段落包含多个 run（分别控制加粗/不加粗），例如指导教师行：

```
Run0: "指导教师"  (bold)
Run1: "                    "  (空格分隔, not bold)
Run2: "职    称"  (bold)
```

如果只替换 Run0 而保留 Run2，会导致输出中出现重复的标签文字。

**正确处理方式**：对包含标签的段落，**先清空该段落所有 run**，再创建一个新 run 写入完整文本：

```python
# 错误做法：只替换第一个 run
para.runs[0].text = "指导教师 郭吉楠  职    称  教授"  # Run2 的"职    称"还在！

# 正确做法：清空后重建
para.clear()  # 删除所有 run
run = para.add_run("指导教师 郭吉楠  职    称  教授")
run.bold = True  # 统一设为加粗
# 如需混合格式，分多个 add_run() 设置各自的 bold 属性
```

对封面信息表中的所有标签段落（「题    目」「姓    名」「院    系」「专    业」「指导教师」），在替换时都遵循此模式：清空 → 写入完整文本 → 设置统一格式。

**自定义模板的封面处理**：
1. 用 `analyze_template.py` 分析模板 → 识别封面区域（文档最前面的段落/表格）
2. 将封面元素列表展示给用户，请用户指出哪些是固定的、哪些需要替换
3. 固定元素原样保留，可变元素按用户提供的信息替换
4. 模板封面中的 Logo 图片等二进制资源需从模板中提取并保留



#### 4.2 目录生成（关键——目录必须立即可见）

目录必须**在生成后立即可见、可读**，不能让用户看到
「请在Word中右键此处，选择更新域以生成目录」的空白提示。采用以下方式：

**首选方式：手动生成目录内容 + 标题设置 outlineLvl（立即可见，且支持域更新）**

**重要**：仅手动写入目录内容不够——用户编辑文档后无法用 Word「更新域」刷新目录。必须同时：
1. 给所有正文标题段落设置 Word 内置 Heading 样式 + `outlineLvl`（模板 JSON 中 `heading_levels[].outline_level` 指定）
2. 在目录页手动生成立即可见的目录条目
3. 在手动目录后插入 TOC 域代码 `TOC \o "1-3" \h \z \u`，用户可在 Word 中右键更新

这样用户打开文档时目录立即可见，编辑后也能用 Word 自动更新。

在整篇论文正文和参考文献都生成完毕后，扫描所有标题段落，手动编写目录：

**步骤：**

1. **扫描标题**：遍历文档中所有段落，识别标题：
   - 一级标题（章）：匹配 `^\d+[\s　]+` 或 `^第[一二三四五六七八九十\d]+章`，outline level 0
   - 二级标题（节）：匹配 `^\d+\.\d+[\s　]+`，outline level 1
   - 三级标题（条）：匹配 `^\d+\.\d+\.\d+[\s　]+`，outline level 2
   - 忽略 TOC 区域内的标题文本（避免重复计入）

   **同时给每个标题段落设置 Word heading 样式和 outlineLvl**（从模板 JSON `heading_levels[].outline_level` 读取）：
   ```python
   from docx.oxml.ns import qn
   
   # 对一级标题段落设置 outlineLvl=0
   pPr = heading_para._element.get_or_add_pPr()
   outline_lvl = parse_xml(f'<w:outlineLvl {qn("w:val")}="0"/>')
   pPr.append(outline_lvl)
   # 同时设置样式为 Heading 1（或模板 JSON 指定的样式名）
   heading_para.style = doc.styles['Heading 1']
   ```
   这会确保用户编辑文档后可在 Word 中右键目录 →「更新域」自动刷新页码和条目。

2. **估算页码**：由于 python-docx 无法获取实际渲染页码，按以下公式估算：
   ```
   每页字符数 = (内容区宽度mm × 内容区高度mm) / (字号mm²) × 填充系数 ÷ 行距系数
   ```
   简化公式（A4纸，25mm/20mm边距，小四宋体）：
   ```
   单倍行距：   每页 ≈ 1500 中文字符
   1.5 倍行距： 每页 ≈ 1000 中文字符
   双倍行距：   每页 ≈ 750 中文字符
   页码 = 1 + 该标题之前的总字符数 / 每页字符数
   ```

   **注意**：此估算值仅供目录结构参考，最终页码以 Word 实际渲染为准。不同字体、字号、图表混排等因素都会影响实际分页位置。

3. **生成目录条目**：对每个标题，在目录页创建一个段落：

   **⚠️ 绝对禁止手动插入固定数量的点号**（如 `标题............1`）。点号数量必须
   由制表位的前导符自动填充，这样无论标题长短，页码始终右对齐。

   python-docx 示例：
   ```python
   from docx.enum.text import WD_TAB_LEADER, WD_TAB_ALIGNMENT
   from docx.shared import Cm

   p = doc.add_paragraph()
   # 标题文字 + 制表符 + 页码
   p.add_run("1  绪论").font.name = "宋体"
   p.add_run("\t1")  # \t 触发制表位，点线自动填充标题到页码之间的空白
   # 在段落属性中设置右对齐制表位 + 点线前导符
   tab_stop = p.paragraph_format.tab_stops.add_tab_stop(
       Cm(14.5),                # 页面右边缘位置
       WD_TAB_ALIGNMENT.RIGHT,  # 右对齐 → 页码顶到最右
       WD_TAB_LEADER.DOTS,      # 点线前导符 → 中间自动填充点号
   )
   ```

   关键参数：
   - 字体：宋体小四（12pt）
   - 制表位位置：页面右边缘（A4 25mm左边距 → 右边缘约 14.5cm）
   - 前导符：`WD_TAB_LEADER.DOTS`（点线）
   - 对齐方式：`WD_TAB_ALIGNMENT.RIGHT`（右对齐）

4. **目录格式**：
   - 目录标题「目  录」：黑体四号(14pt)，居中，两字间空两格
   - 一级标题条目：顶格，宋体小四(12pt)
   - 二级标题条目：缩进2字符，宋体小四(12pt)
   - 所有条目的页码右对齐，由点线前导符连接

5. **插入位置和 TOC 域代码**：目录内容插入到目录页「目  录」标题之后。
   模板中已有的 TOC 占位段落需删除，替换为手动生成的目录条目。
   
   **在手动目录后插入 TOC 域代码**，以便用户编辑文档后可用 Word「更新域」刷新：
   ```python
   from docx.oxml.ns import qn
   
   # 在手动目录段落后插入 TOC 域代码
   toc_para = doc.add_paragraph()
   # 插入域代码：TOC \o "1-3" \h \z \u
   fldChar_begin = parse_xml(f'<w:fldChar {qn("w:fldCharType")}="begin"/>')
   run_begin = toc_para.add_run()
   run_begin._element.append(fldChar_begin)
   
   instrText = parse_xml(f'<w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>')
   run_instr = toc_para.add_run()
   run_instr._element.append(instrText)
   
   fldChar_end = parse_xml(f'<w:fldChar {qn("w:fldCharType")}="end"/>')
   run_end = toc_para.add_run()
   run_end._element.append(fldChar_end)
   ```
   Word 在打开文档时会自动解析此域代码并生成可点击更新的目录。

**备选方式：TOC 域代码 + LibreOffice 自动更新**

如果系统安装了 LibreOffice，可以使用域代码方式
（设置 outline level + 插入 `TOC \o "1-2" \h \z \u` 域代码），
生成后用 `scripts/update_fields.py` 自动更新域：
```bash
python scripts/update_fields.py output.docx
```
此脚本使用 LibreOffice headless 模式打开文档、执行 `.uno:UpdateAllIndexes`
和 `.uno:UpdateFields`、保存并关闭。无需用户手动操作。

**注意**：如果 LibreOffice 不可用（Windows 常见），必须使用首选方式。
域代码方式生成的文件在 Word 中打开后会显示空白目录，直到用户手动「更新域」，
这不满足"目录立即可见"的要求。

生成后验证：目录不为空，且目录中列出的章节与正文中的标题一一对应。

#### 4.3 摘要格式（极其重要 — 用户直接反馈）

**摘要必须使用 inline 格式**：`"摘要："` 标签后紧跟内容，在**同一段落**内，**左对齐**（不居中）。

格式来自模板 JSON 的 `abstract_format` 字段（`"inline"`），以及 `styles.AbstractLabel` + `styles.Abstract` 的字体规格。

**正确格式**（参考理科/课程论文模板实测）：
```
摘要：暗物质是当代天体物理与粒子物理领域中的一大谜题……（五号宋体，200-300字）
关键词：暗物质；宇宙演化；星系团（3-5个，五号宋体）
```

**生成方式**（python-docx）：
```python
# 中文摘要段落：标签加粗黑体 + 内容宋体，同一段落，左对齐
p = doc.add_paragraph()
p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT  # 不居中！
label_run = p.add_run("摘要：")
label_run.bold = True
label_run.font.name = "黑体"
content_run = p.add_run("本文研究了……（摘要正文内容）")
content_run.font.name = "宋体"
# 首行缩进2字符（与模板一致）
p.paragraph_format.first_line_indent = Cm(0.74)  # 五号字2字符
```

**关键词段落**同理：`"关键词："` 标签加粗黑体 + 词列表宋体，同一段落，左对齐。

**英文摘要**同理：
```python
p = doc.add_paragraph()
p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
label = p.add_run("Abstract: ")
label.bold = True
content = p.add_run("Dark matter is a major mystery...")
```

**绝对禁止**：
- ❌ "摘要" 作为独立居中标题段落，内容另起一段
- ❌ 摘要标签和内容分开两个段落
- ❌ 摘要居中排版

#### 4.4 格式规范来源

**所有格式参数均来自第一步加载的模板 JSON 文件。**

模板 JSON 的 `styles` 字段定义了每个语义角色（`BodyText`、`Heading1`、`Title` 等）的完整格式属性（字体、字号、加粗、对齐、行距、缩进等）。生成时直接从 `styles[<role>]` 读取并使用 `docx` 技能创建对应的样式和段落。

**关键**：模板 JSON 是单一真相来源。不参考 `.md` 参考文件，不依赖先验知识。即使你"知道" 中国论文通常用黑体做标题、宋体做正文——如果模板 JSON 中 Heading1 的 `font_east_asia` 是微软雅黑，那就用微软雅黑。

确保以下各项与模板 JSON 一致：
- 页面设置（`page_setup`：纸张、页边距、装订线）
- 各级标题格式（`styles.Heading1/Heading2/Heading3`）
- 正文格式（`styles.BodyText`：字体、行距、段前段后、首行缩进）
- 摘要格式（`styles.Abstract` / `styles.Abstract_EN`）
- 参考文献格式（`styles.Reference`）
- 封面结构（`cover.enabled` → 见 4.1.1；`cover` 为 `null` 或无 `enabled` 则跳过封面）
- 页面顺序和换页规则（`page_order`）
- 目录格式（`toc`：标题样式、层级、制表位）
- **行距一致性**：全文档使用 `styles` 中定义的行距值，除非用户第二步中指定了覆盖值

### 第四点五步：修复字体缺陷（生成后必须执行）

**python-docx 已知缺陷**：

`python-docx` 的 `style.font.name` 只设置 `w:ascii`（拉丁字符）和 `w:hAnsi`（高位ANSI）
字体槽位，**不会设置 `w:eastAsia`（CJK字符）字体槽位**。此外 python-docx 默认模板的
`w:themeFontLang` 使用 `eastAsia="ja-JP"`（日语）。这导致中文字符在 Word 中回退为
日文字体 MS Mincho，而非正确的中文字体。

**生成 .docx 后必须执行以下修复**：

1. 用 `docx` 技能解包 .docx：
   ```
   python scripts/office/unpack.py output.docx unpacked/
   ```
2. 在 `word/styles.xml` 和 `word/document.xml` 中，给所有 `<w:rFonts>` 元素添加
   `w:eastAsia` 属性，值应与 `w:ascii` 相同：
   ```xml
   <!-- 修复前 -->
   <w:rFonts w:ascii="黑体" w:hAnsi="黑体"/>
   <!-- 修复后 -->
   <w:rFonts w:ascii="黑体" w:hAnsi="黑体" w:eastAsia="黑体"/>
   ```
3. 在 `word/settings.xml` 中将 `w:themeFontLang w:eastAsia` 从 `ja-JP` 改为 `zh-CN`
4. 重新打包：`python scripts/office/pack.py unpacked/ output.docx --original output.docx`

当前提供的模板文件（`qfnu-template-science.docx` 和 `qfnu-template-course.docx`）
已经预修复了这两个问题，可直接使用。


### 第五步：质量检查与输出

生成完成后，执行以下检查（**不可跳过**）：

1. **页面顺序检查**：逐页验证文档结构是否符合模板定义的顺序（目录→题目→摘要→
   正文→参考文献，或封面→目录→...）。不仅检查元素是否存在，还要检查它们出现的
   先后顺序是否正确。

2. **目录验证**：
   - 目录页存在且位于正确位置（文档开头附近）
   - 目录内容不为空——列出的章节标题与正文中的标题一一对应
   - 页码右对齐

3. **格式对比**：用 `scripts/analyze_template.py` 分别分析参考模板和产出 .docx，
   逐项对比页面设置、字体层级、段落格式。不一致则修正后重检。

4. **换页检查**：确认文档中只有一处换页（目录→题目之间）。检查题目→摘要、摘要→正文、正文→参考文献之间没有换页符。

5. **内容检查**：确认所有承诺的章节已生成，字数接近目标。

将生成的文件路径告知用户，同时报告检查结果。

## 重要提醒

- **环境先检查**：第零步确认 `docx` 技能和 Python 依赖可用
- **模板 JSON 是唯一权威**：选定模板后加载其 JSON，所有格式参数从 JSON 读取。不要
  查阅 `.md` 参考文件来获取格式值——它们仅供人类参考
- **字数是硬性要求**：用户未提供时必须询问
- **大纲先确认**：自动生成的大纲须经用户确认再动笔
- **顺序不能乱**：页面排列顺序必须与模板 JSON 的 `page_order` 一致。目录在前，题目在后。
  从题目到参考文献，所有内容连续排版无一换页——每一条都直接来自用户反馈，
  不是可选项
- **目录不能空且立即可见**：必须手动生成目录内容（按4.2节首选方式）。不要依赖用户
  「更新域」——用户打开文档时目录就必须能看到完整的章节列表和页码。如果 LibreOffice
  可用，可用 `scripts/update_fields.py` 自动填充 TOC 域代码
  或手动生成带 dot leader 的目录。用户打开文档后目录必须能用
- **章节不换页**：从论文题目到参考文献结束，**整个正文区域不插入任何换页符**。
  换页规则由模板 JSON 的 `page_order` 中 `page_break_after: true` 决定。
  不要在题目/摘要/正文/参考文献之间插入 `add_page_break()`
- **格式不容妥协**：产出必须通过 `analyze_template.py` 与模板的格式对比
- **一个模板一套规则**：不同模板的格式规则以 JSON 文件存放，按需加载，互不干扰
- **归一化先确认**：从非标准输入创建模板时，归一化的 JSON 结果必须经用户确认
  再继续生成

## 参考文件

- `references/template-schema.md` — 标准模板 JSON Schema 定义
- `assets/templates/index.json` — 内置官方模板索引
- `assets/templates/science.json` — 模板 A：理科毕业论文
- `assets/templates/course.json` — 模板 B：课程论文
- `references/format-science.md` — 理科毕业论文格式（人类参考，不驱动生成）
- `references/format-course.md` — 课程论文格式（人类参考，不驱动生成）
- `scripts/analyze_template.py` — 模板分析工具，支持 `--output-template` 输出标准模板 JSON
- `scripts/validate_template.py` — 模板 JSON 校验工具
- `assets/qfnu-template-science.docx` — 模板 A：理科毕业论文 .docx
- `assets/qfnu-template-course.docx` — 模板 B：课程论文 .docx（带封面）
