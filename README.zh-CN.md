# document-format

这个仓库的 README 只描述 `skills/` 目录下的技能。

当前主要技能是 `skills/word-expert-formatting`，它是一个自定义 Claude Code skill，用于格式化中文 Word 内容，并支持把 Markdown 或纯文本本地生成 `.docx`。

## 技能列表

### `word-expert-formatting`

位置：
- `skills/word-expert-formatting/SKILL.md`
- `skills/word-expert-formatting/scripts/text_to_docx.py`

用途：
- 把原始 Markdown、TXT、HTML 或较粗糙的大纲整理成适合 Word 的结构化格式
- 套用严格的中文正式文档排版模型
- 当输入为 `.md`、`.markdown` 或 `.txt` 时，支持本地生成 `.docx`

参考：`skills/word-expert-formatting/SKILL.md:1`

## 这个技能定义了什么

`skills/word-expert-formatting/SKILL.md` 定义了：
- 支持的输入形态
- 标题层级规则
- 编号模式
- 封面识别规则
- 标题、正文、表格和页脚页码的样式映射
- 结构化格式化结果的输出契约
- 何时使用本地 DOCX 生成脚本

参考：`skills/word-expert-formatting/SKILL.md:27`

## 本地 DOCX 工作流

本地实现位于：

`skills/word-expert-formatting/scripts/text_to_docx.py`

脚本当前会：
- 接收 `.md`、`.markdown` 和 `.txt`
- 处理标题、段落、列表、表格、代码块、页脚页码和简单封面
- 输出 `.docx` 文件

参考：`skills/word-expert-formatting/scripts/text_to_docx.py:840`

执行方式：

```bash
python3 /Users/huangcn/github/document-format/skills/word-expert-formatting/scripts/text_to_docx.py <input-file> [output.docx]
```

如果省略 `output.docx`，脚本会在输入文件旁边生成同名输出文件。

参考：`skills/word-expert-formatting/scripts/text_to_docx.py:857`

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
- TXT 段落块

参考：
- `skills/word-expert-formatting/scripts/text_to_docx.py:695`
- `skills/word-expert-formatting/scripts/text_to_docx.py:822`

## 一个重要实现说明

skill 契约里描述了 Mode A 和 Mode B 两种标题编号模式。

但当前 Python 脚本看起来只真正实现了十进制标题编号。这意味着在 Mode B 支持上，面向人的 skill 契约与实际执行脚本目前还没有完全对齐。

参考：
- `skills/word-expert-formatting/SKILL.md:43`
- `skills/word-expert-formatting/scripts/text_to_docx.py:313`
- `skills/word-expert-formatting/scripts/text_to_docx.py:354`

## 维护规则

如果某条格式规则的变更需要影响最终生成的 `.docx` 输出，请同时更新：
- `skills/word-expert-formatting/SKILL.md`
- `skills/word-expert-formatting/scripts/text_to_docx.py`

参考：`skills/word-expert-formatting/SKILL.md:165`
