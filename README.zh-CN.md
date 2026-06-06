# document-format

这个仓库的 README 只描述 `skills/` 目录下的技能。

当前主要技能是 `skills/word-expert-formatting`，它是一个自定义 Claude Code skill，用于格式化中文 Word 内容，并支持把 Markdown 或纯文本本地生成 `.docx`，以及对已有 `.docx` 做原位刷新与核查。

它现在已经同时支持 Mode A 与 Mode B 标题编号，并且 `--auto-toc` 会插入 Word 动态目录域，而不是静态目录文本。

## 技能列表

### `word-expert-formatting`

位置：
- `skills/word-expert-formatting/SKILL.md`
- `skills/word-expert-formatting/scripts/text_to_docx.py`

用途：
- 把原始 Markdown、TXT、HTML 或较粗糙的大纲整理成适合 Word 的结构化格式
- 套用严格的中文正式文档排版模型
- 同时支持 Mode A 与 Mode B 标题编号
- 当输入为 `.md`、`.markdown` 或 `.txt` 时，支持本地生成 `.docx`
- 当输入为已有 `.docx` 时，支持原位规范化与核查

参考：`skills/word-expert-formatting/SKILL.md:1`

## 这个技能定义了什么

`skills/word-expert-formatting/SKILL.md` 定义了：
- 支持的输入形态
- 标题层级规则
- 编号模式以及选择它们的 CLI 参数
- 封面识别规则
- 标题、正文、表格和页脚页码的样式映射
- 结构化格式化结果的输出契约
- 正文段落会先清空原有缩进、行距和对齐，再按标准样式应用两端对齐与首行缩进 2 字符
- 已有 `.docx` 刷新会保留图片、表格、分页和 section 结构，而不是删除正文后重建
- 已有 `.docx` 核查同时覆盖段落版式与结构保真

参考：`skills/word-expert-formatting/SKILL.md:27`

## 本地 DOCX 工作流

本地实现位于：

`skills/word-expert-formatting/scripts/text_to_docx.py`

脚本当前会：
- 接收 `.md`、`.markdown`、`.txt` 和 `.docx`
- 处理标题、段落、列表、表格、代码块、页脚页码和简单封面
- 通过 `--numbering-mode` 支持 Mode A 与 Mode B 标题编号
- 在未检测到封面时，可通过 `--reserve-cover` 预留封面页
- 在未检测到显式目录标题时，可通过 `--auto-toc` 插入 Word 动态目录页
- 支持按单次运行显式指定首页 / 目录决策，并支持手动输入首页文字
- 对已有 `.docx` 采用原位刷新，而不是删除正文后重建
- 在已有 `.docx` 刷新时保留图片、表格、分页和 section 结构
- 当存在封面页时，使用更正式的分页模型：封面无页码，正文节从 1 开始重新编号
- 输出 `.docx` 文件

参考：`skills/word-expert-formatting/scripts/text_to_docx.py:840`

执行方式：

```bash
python3 /Users/huangcn/github/document-format/skills/word-expert-formatting/scripts/text_to_docx.py <input-file> [output.docx] [--reserve-cover] [--auto-toc] [--with-cover|--without-cover] [--cover-text <text>] [--with-toc|--without-toc] [--numbering-mode A|B]
```

如果省略 `output.docx`，脚本会在输入文件旁边生成同名输出文件。

在本仓库通过 skill 驱动本地 DOCX 工作流时，应先询问：
1. 是否生成首页
2. 如果生成首页，首页文字内容是什么（手动输入；默认空）
3. 是否生成目录页

skill 现在会把这些回答当作本次运行的显式决策：
- 生成首页 + 输入非空文字：把第一条非空行渲染为首页标题，后续非空行渲染为居中的公司名 / 日期等元信息
- 生成首页 + 文字为空：生成空白占位首页
- 不生成首页：本次运行同时关闭自动封面识别与占位封面插入
- 目录是否生成：每次都询问，再映射为本次脚本调用参数

参考：
- `skills/word-expert-formatting/SKILL.md:127`
- `skills/word-expert-formatting/scripts/text_to_docx.py:1163`

## 使用前检查

在使用本地 DOCX 工作流前，请先检查：
- 已安装 Python 3
- 已安装 `python-docx`

脚本现在会在启动时主动执行这些检查；如果依赖缺失，会直接给出清晰错误信息并退出。
现在输出文档会从空白 Word 文档直接创建，因此该工作流不再依赖本地模板 `.docx` 文件。

## 运行依赖

- Python 3
- `python-docx`

参考：`skills/word-expert-formatting/scripts/text_to_docx.py:8`

## 当前支持的能力

当前脚本支持：
- Markdown 标题（`#` 到 `######`）
- 普通段落
- 无序列表
- 有序列表
- Markdown 表格
- 围栏代码块
- 简单封面识别
- 面向已有 `.docx` 刷新的多行封面标题块识别
- 对已有 `.docx` 的封面标题，刷新时只修正字体和字号，不改变原有标题对齐方式
- 通过 `--reserve-cover` 预留封面页
- 通过 `--with-cover` 与 `--cover-text` 显式生成首页
- 通过 `--without-cover` 显式禁止生成首页
- 显式目录检测（`目录`、`TOC`、`Table of Contents`）
- 通过 `--auto-toc` 插入 Word 动态目录域
- 通过 `--with-toc` 显式生成目录
- 通过 `--without-toc` 显式禁止生成目录
- 通过 `--numbering-mode` 支持 Mode A / Mode B 标题编号
- 当存在封面页时采用正式分页：封面无页码，正文从 1 开始重新编号
- 保留结构的已有 `.docx` 原位刷新
- 面向已有 `.docx` 的封面、目录、正文、表格、编号与结构计数核查
- TXT 段落块
- 为目录生成提供最小 TXT 标题识别

参考：
- `skills/word-expert-formatting/scripts/text_to_docx.py:695`
- `skills/word-expert-formatting/scripts/text_to_docx.py:822`

## 一个重要实现说明

skill 契约里描述了 Mode A 和 Mode B 两种标题编号模式。

当前 Python 脚本已经真正实现这两种模式：
- Mode A 使用纯十进制标题编号
- Mode B 使用中文一级标题编号（`一、二、三、...`），下级标题仍使用十进制层级编号

另外，`--auto-toc` 现在会插入真正的 Word 动态目录域，而不是静态文本目录项。如果在 Word 中打开后目录未立即更新，请使用 Word 的“更新目录”操作。

新增的显式开关主要用于 skill 驱动运行：
- `--with-cover` / `--without-cover` 会覆盖本次运行的自动封面识别逻辑
- 单独使用 `--cover-text` 时，会隐式视为 `--with-cover`
- `--with-toc` / `--without-toc` 会覆盖本次运行的目录兜底生成逻辑
- `--reserve-cover` 与 `--auto-toc` 仍保留，用于兼容直接 CLI 调用

当文档存在封面页，或者使用了 `--reserve-cover` / `--with-cover` 时，脚本现在会为正文创建独立 section，使封面不显示页码，正文重新从第 1 页开始。

参考：
- `skills/word-expert-formatting/SKILL.md:43`
- `skills/word-expert-formatting/scripts/text_to_docx.py:336`
- `skills/word-expert-formatting/scripts/text_to_docx.py:605`

## 维护规则

如果某条格式规则的变更需要影响最终生成的 `.docx` 输出，请同时更新：
- `skills/word-expert-formatting/SKILL.md`
- `skills/word-expert-formatting/scripts/text_to_docx.py`

参考：`skills/word-expert-formatting/SKILL.md:165`
