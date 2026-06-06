# document-format

A local repository for Word-formatting assets, sample documents, and a custom Claude Code skill for turning Markdown or plain text into professionally formatted Chinese `.docx` documents.

This is not a conventional application repository. It is primarily a document workspace plus an executable formatting workflow.

## What is in this repository

- Document assets and sample deliverables
- A formatting specification for Chinese formal/business documents
- A custom Claude Code skill: `skills/word-expert-formatting`
- A local Python script that converts `.md` / `.markdown` / `.txt` into `.docx`

## Repository layout

```text
.
├── skills/
│   └── word-expert-formatting/
│       ├── SKILL.md
│       └── scripts/
│           └── text_to_docx.py
├── 天府银行MogDB数据库巡检报告_xc006.md
├── 天府银行MogDB数据库巡检报告_xc006.docx
├── 2025-2026年度数据库运维服务总结_20260522_修改.docx
└── Word 专家级排版排版规范.md
```

## Core components

### 1. Skill contract

The main skill definition lives in `skills/word-expert-formatting/SKILL.md`.

It defines:
- supported input types
- heading hierarchy rules
- cover-page detection rules
- body, heading, table, and footer styles
- the expected structured output contract
- the local DOCX generation workflow

Reference: `skills/word-expert-formatting/SKILL.md:1`

### 2. Local DOCX generator

The executable implementation is:

`skills/word-expert-formatting/scripts/text_to_docx.py`

It currently:
- reads `.md`, `.markdown`, or `.txt`
- loads a template `.docx`
- rebuilds headings, paragraphs, lists, tables, code blocks, footer page numbers, and simple cover sections
- writes a formatted `.docx`

Reference: `skills/word-expert-formatting/scripts/text_to_docx.py:840`

## Requirements

- Python 3
- `python-docx`

The script imports `docx` directly.

Reference: `skills/word-expert-formatting/scripts/text_to_docx.py:8`

## Quick start

Run:

```bash
python3 /Users/huangcn/github/document-format/skills/word-expert-formatting/scripts/text_to_docx.py <input-file> [output.docx]
```

Examples:

```bash
python3 /Users/huangcn/github/document-format/skills/word-expert-formatting/scripts/text_to_docx.py \
  /Users/huangcn/github/document-format/天府银行MogDB数据库巡检报告_xc006.md
```

```bash
python3 /Users/huangcn/github/document-format/skills/word-expert-formatting/scripts/text_to_docx.py \
  input.md output.docx
```

If `output.docx` is omitted, the script writes a file next to the input with the same basename.

Reference: `skills/word-expert-formatting/scripts/text_to_docx.py:857`

## Supported input features

### Markdown

The script currently handles:
- ATX headings (`#` to `######`)
- normal paragraphs
- unordered lists
- ordered lists
- Markdown tables
- fenced code blocks
- simple cover blocks

Reference: `skills/word-expert-formatting/scripts/text_to_docx.py:695`

### Plain text

For `.txt`, the script splits content into paragraph blocks and applies body formatting.

Reference: `skills/word-expert-formatting/scripts/text_to_docx.py:822`

## Cover recognition

The workflow supports explicit labels such as:
- `封面标题:`
- `公司名称:`
- `日期:`
- bracketed forms like `[封面标题]`

If labels are absent, the script uses a narrow heuristic on the first few non-empty lines.

Reference: `skills/word-expert-formatting/SKILL.md:142`
Reference: `skills/word-expert-formatting/scripts/text_to_docx.py:500`

## Formatting model

The formatting contract in `SKILL.md` includes rules for:
- document title
- level 1-6 headings
- body paragraphs
- table text
- cover title and metadata
- centered footer page numbers

Reference: `skills/word-expert-formatting/SKILL.md:78`

## Important implementation note

`SKILL.md` describes two numbering modes:
- Mode A: decimal hierarchy
- Mode B: Chinese top-level headings such as `一、二、三、`

However, the current Python implementation appears to generate decimal heading numbering only. The numbering logic in `text_to_docx.py` uses `decimal` numbering and does not expose a CLI switch for Mode B.

So at the moment:
- `SKILL.md` is the human-facing contract
- `text_to_docx.py` is the executable implementation
- they are not fully in sync for Mode B support yet

References:
- `skills/word-expert-formatting/SKILL.md:43`
- `skills/word-expert-formatting/scripts/text_to_docx.py:313`
- `skills/word-expert-formatting/scripts/text_to_docx.py:354`

## Current limitations

The current local script is good for:
- structured reports
- heading-based documents
- Markdown tables
- basic cover pages
- plain-text to Word conversion

It is not a full general-purpose Word renderer. Based on the skill contract, these areas are not promised yet:
- complex embedded HTML
- images
- footnotes
- highly customized layouts

Reference: `skills/word-expert-formatting/SKILL.md:161`

## Example repository assets

- `天府银行MogDB数据库巡检报告_xc006.md` is a realistic Markdown report sample.
- `天府银行MogDB数据库巡检报告_xc006.docx` is a corresponding generated or edited Word deliverable.
- `2025-2026年度数据库运维服务总结_20260522_修改.docx` is also used as a document asset, and the script currently loads a template file from this repository.

Template reference: `skills/word-expert-formatting/scripts/text_to_docx.py:44`

## Maintenance rule

If you change formatting rules that must affect generated `.docx` output, update both:
- `skills/word-expert-formatting/SKILL.md`
- `skills/word-expert-formatting/scripts/text_to_docx.py`

Reference: `skills/word-expert-formatting/SKILL.md:165`

## Recommended reading order

1. `skills/word-expert-formatting/SKILL.md`
2. `skills/word-expert-formatting/scripts/text_to_docx.py`
3. `Word 专家级排版排版规范.md`
4. `天府银行MogDB数据库巡检报告_xc006.md`

## Status

This repository is best understood as a local document-formatting toolkit and skill workspace, not as a production software service.
