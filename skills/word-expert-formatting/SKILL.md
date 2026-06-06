---
name: word-expert-formatting
description: "Use this skill whenever the user wants to format Chinese Word content according to a strict document style guide, especially for 公文, 商务报告, formal reports, or .docx deliverables with required heading levels, cover page text, tables, page numbers, fonts, line spacing, paragraph spacing, first-line indentation, or Chinese numbering patterns. Trigger when the user mentions Word 排版, 文档格式规范, 标题层级, 一级标题/二级标题, 模式A/模式B, 中文编号, 封面, 页脚页码, 表格格式, 宋体, 四号, 小四, 五号, 首行缩进, or asks to turn raw text / Markdown / TXT / HTML into a Word-ready structured format. If the user needs a final `.docx` file from `.md`, `.markdown`, or `.txt` in this repository, use the local script workflow in this skill instead of only describing the format. Do not use this skill for PDFs, spreadsheets, slide decks, or generic writing requests without formatting constraints."
---

# Word expert formatting

## Purpose

Use this skill to translate raw content into a Word-ready structure that follows a strict Chinese formatting specification.

This skill is optimized for inputs such as:
- plain text
- Markdown
- HTML
- rough outlines with heading levels
- partially structured report drafts

This skill is not primarily about producing a `.docx` file by itself. Its main job is to:
1. identify document structure
2. map each block to the required style
3. rebuild heading numbering
4. output a structured formatting result that another Word-generation workflow can apply directly

If the user explicitly needs a finished `.docx`, use this skill to determine the formatting structure first, then hand off to a Word/document-generation workflow.

## Required workflow

Follow this sequence.

### 1. Identify document regions

Classify the input into these regions when present:
- cover title
- cover company name and date
- body headings
- body paragraphs
- tables and table text
- footer / page number requirements

If a region is absent, do not invent it.

### 2. Determine numbering mode

Before formatting, decide which heading numbering mode applies.

- **Mode A: all-numeric hierarchy**
  - level 1 → `1`, `2`, `3`
  - level 2 → `1.1`, `1.2`
  - level 3 → `1.1.1`
  - level 4 → `1.1.1.1`
  - level 5 → `1.1.1.1.1`
- **Mode B: Chinese top-level headings**
  - level 1 → `一、`, `二、`, `三、`
  - level 2 → `1.1`, `1.2`
  - level 3 → `1.1.1`
  - level 4 → `1.1.1.1`
  - level 5 → `1.1.1.1.1`

Default to **Mode A** unless the user explicitly asks for Mode B or clearly uses that pattern.

If the source content has inconsistent numbering, normalize it to the selected mode.

### 3. Rebuild heading hierarchy

Infer the intended heading level from the source structure, wording, and numbering.

When rebuilding hierarchy:
- keep sibling numbering consistent
- ensure child headings match their parent path
- do not skip levels unless the source clearly requires it
- if the source is ambiguous, surface the ambiguity instead of guessing silently

### 4. Apply style mapping

Map every block to the exact style below. Do not improvise font sizes, spacing, or indentation.

## Style matrix

### Heading levels

| Element | Font | Size | Line spacing | Space before | Space after | Other |
|---|---|---:|---|---|---|---|
| Document title | 黑体 | 28号 | single | 10 pt | 10 pt | bold |
| Level 1 heading | 黑体 | 三号 | single | 10 pt | 10 pt | bold, use selected numbering mode |
| Level 2 heading | 黑体 | 小三 | single | 8 pt | 8 pt | bold |
| Level 3 heading | 黑体 | 四号 | 1.5 lines | 0 | 0 | bold |
| Level 4 heading | 黑体 | 小四 | 1.5 lines | 0 | 0 | bold |
| Level 5 heading | 黑体 | 小四 | 1.5 lines | 0 | 0 | bold |
| Level 6 heading | 黑体 | 小四 | 1.5 lines | 0 | 0 | bold |

### Body text

| Element | Font | Size | Line spacing | Space before | Space after | Other |
|---|---|---:|---|---|---|---|
| Standard body paragraph | 宋体 | 小四 | 1.5 lines | 0 | 0 | first-line indent: 2 characters |

### Table text

| Element | Font | Size | Line spacing | Space before | Space after | Other |
|---|---|---:|---|---|---|---|
| Table cell text | 宋体 | 五号 | 1.5 lines | 0 | 0 | no first-line indent |

### Bullet levels

| Element | Font | Size | Other |
|---|---|---:|---|
| Bullet level 1 | 宋体 | 小四 | black circle `●` |
| Bullet level 2 | 宋体 | 小四 | black square `■` |
| Bullet level 3 | 宋体 | 小四 | black diamond `◆` |

### Cover elements

| Element | Font | Size | Line spacing | Space before | Space after | Other |
|---|---|---:|---|---|---|---|
| Cover title | 宋体 | 二号 | 1.5 lines | 0.5 lines | 0.5 lines | centered |
| Company name / date | 宋体 | 四号 | 1.5 lines | 0 | 0 | centered |

### Footer / page number

| Element | Font | Size | Line spacing | Alignment | Other |
|---|---|---:|---|---|---|
| Page number | 宋体 | 五号 | 1.5 lines | centered | apply in footer |

## Local DOCX generation workflow

When the user wants an actual `.docx` file in this repository and the source input is `.md`, `.markdown`, or `.txt`, prefer the built-in local script workflow instead of only returning structured formatting guidance.

Use:

```bash
python3 /Users/huangcn/github/document-format/skills/word-expert-formatting/scripts/text_to_docx.py <input-file> [output.docx]
```

Apply this path when:
- the user wants a final Word document, not just a style spec
- the input is Markdown or TXT
- the content mainly consists of headings, paragraphs, tables, and fenced code blocks

This local script is the default implementation for this repository because it already applies the core formatting rules from this skill, including heading mapping, body formatting, table formatting, page setup, page numbers, first-line indent in Word character units, and basic cover recognition.

### Cover recognition rules

The current script uses **label-first recognition with heuristic fallback**.

Preferred explicit labels:
- `封面标题:` / `封面标题：`
- `公司名称:` / `公司名称：`
- `日期:` / `日期：`
- bracketed forms such as `[封面标题]` and `[封面公司/日期]`

If explicit labels are absent, the script only inspects the first few non-empty lines of the document:
- the first short standalone line is treated as the cover title
- the next 1-2 short lines are treated as company name / date
- detection stops immediately when it encounters a Markdown heading, table, fenced code block, or an obvious long body line
- when a cover block is recognized, the script inserts an automatic page break before the body content
- the script also adds fixed top spacing so the cover content sits in the upper-middle area of the page

Use structured Markdown / JSON output instead only when the user wants the formatting model or wants to feed another generation pipeline.

Boundaries of the script workflow:
- supported well: headings, paragraphs, Markdown tables, fenced code blocks, TXT paragraphs, simple cover blocks
- not a promise yet: complex embedded HTML, images, footnotes, highly customized layouts

## Style source of truth

In this repository, `SKILL.md` is guidance for when and how to use the workflow, but the current script implementation does **not** parse this file dynamically.

That means:
- changing `SKILL.md` alone does not change generated DOCX output
- if you change a formatting rule that must affect output, update both this file and `scripts/text_to_docx.py`

Treat this skill file as the human-readable contract, and the script as the executable implementation. Keep them in sync.

## Output contract

### Default output

By default, output a structured Markdown-style formatting result with explicit labels.

Use labels like these:
- `[封面标题]`
- `[封面公司/日期]`
- `[一级标题-模式A]`
- `[一级标题-模式B]`
- `[二级标题]`
- `[三级标题]`
- `[四级标题]`
- `[五级标题]`
- `[正文]`
- `[表格文字]`
- `[页脚]`

Each labeled block should include the visible content and the applied style summary.

### Structured data output

If the user explicitly asks for machine-readable output for `python-docx`, XML, JSON, or another generation pipeline, output a structure that includes fields such as:
- `font_name`
- `font_size`
- `line_spacing`
- `space_before`
- `space_after`
- `first_line_indent`
- `alignment`
- `heading_level`
- `numbering_mode`

When producing structured data, preserve the same formatting rules as the default output.

## Guardrails

Apply these checks before finalizing the result.

1. **Body indentation must exist**
   - Every normal body paragraph must carry first-line indent of 2 characters.

2. **Table text must not inherit body indentation**
   - Table text must not use first-line indent.

3. **Do not confuse units**
   - Level 1 and level 2 headings use spacing in **pt**.
   - Body text and level 3–6 headings use line spacing in **lines**.
   - Do not collapse these into a single spacing concept.

4. **Do not silently flatten hierarchy**
   - If the source contains nested sections, preserve the hierarchy.

5. **Do not invent missing content**
   - Format only what the user provides.

6. **Call out ambiguity when needed**
   - If a heading could reasonably belong to two different levels, say so and present the chosen interpretation.

7. **Keep cover fallback narrow**
   - When explicit cover labels are absent, inspect only the first few lines.
   - Stop fallback recognition when a heading, table, code fence, or obvious body paragraph appears.

## Response pattern

When formatting content, prefer this response shape:

```markdown
【格式化执行完毕 - 已启用模式A 或 模式B】

[封面标题] ...
[封面公司/日期] ...
[一级标题-模式A] ...
[二级标题] ...
[三级标题] ...
[正文] ...
[表格文字] ...
[页脚] ...
```

If some sections do not exist in the source, omit them.

## Example: Mode B

**Input idea:**
“帮我用模式B格式化这段：公司介绍。1. 历史沿革。1.1 创立期。”

**Output:**

```markdown
【格式化执行完毕 - 已启用模式 B】

[一级标题-模式B] 一、公司介绍
样式：宋体，四号，单倍行距，段前 10 pt，段后 10 pt

[二级标题] 1.1 历史沿革
样式：宋体，小四号，单倍行距，段前 8 pt，段后 8 pt

[三级标题] 1.1.1 创立期
样式：宋体，小四号，1.5 倍行距，段前 0，段后 0

[正文] 正文内容……
样式：宋体，小四号，1.5 倍行距，首行缩进 2 字符

[页脚] 页码
样式：宋体，五号，1.5 倍行距，居中
```

## Example: Mode A

**Input idea:**
“把这份提纲整理成标准 Word 格式：项目背景、实施范围、上线计划。”

**Output pattern:**

```markdown
【格式化执行完毕 - 已启用模式 A】

[一级标题-模式A] 1 项目背景
样式：宋体，四号，单倍行距，段前 10 pt，段后 10 pt

[一级标题-模式A] 2 实施范围
样式：宋体，四号，单倍行距，段前 10 pt，段后 10 pt

[一级标题-模式A] 3 上线计划
样式：宋体，四号，单倍行距，段前 10 pt，段后 10 pt
```

## Handoff guidance

If the user wants a final `.docx` file:
1. use this skill to finalize structure, numbering, and style mapping
2. then generate the document with a Word-capable workflow
3. keep the formatting values from this skill as the source of truth

If the user only wants a formatting specification or a review of whether content conforms to the style guide, stay within this skill and do not force document generation.
