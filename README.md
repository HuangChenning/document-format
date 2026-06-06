# document-format

This repository documents the skills under `skills/`.

At the moment, the primary skill is `skills/word-expert-formatting`, a custom Claude Code skill for formatting Chinese Word content and supporting local `.docx` generation from Markdown or plain text.

## Skills

### `word-expert-formatting`

Location:
- `skills/word-expert-formatting/SKILL.md`
- `skills/word-expert-formatting/scripts/text_to_docx.py`

Purpose:
- turn raw Markdown, TXT, HTML, or rough outlines into a Word-ready structured format
- apply a strict Chinese formal-document style model
- support local `.docx` generation when the input is `.md`, `.markdown`, or `.txt`

Reference: `skills/word-expert-formatting/SKILL.md:1`

## What the skill defines

`skills/word-expert-formatting/SKILL.md` defines:
- supported input shapes
- heading hierarchy rules
- numbering modes
- cover detection rules
- style mapping for headings, body text, tables, and footer page numbers
- output contracts for structured formatting results
- when to use the local DOCX generation script

Reference: `skills/word-expert-formatting/SKILL.md:27`

## Local DOCX workflow

The local implementation lives in:

`skills/word-expert-formatting/scripts/text_to_docx.py`

The script currently:
- accepts `.md`, `.markdown`, and `.txt`
- formats headings, paragraphs, lists, tables, code blocks, footer page numbers, and simple cover sections
- writes a `.docx` file

Reference: `skills/word-expert-formatting/scripts/text_to_docx.py:840`

Run:

```bash
python3 /Users/huangcn/github/document-format/skills/word-expert-formatting/scripts/text_to_docx.py <input-file> [output.docx]
```

If `output.docx` is omitted, the script writes a file next to the input with the same basename.

Reference: `skills/word-expert-formatting/scripts/text_to_docx.py:857`

## Requirements

- Python 3
- `python-docx`

Reference: `skills/word-expert-formatting/scripts/text_to_docx.py:8`

## Supported features

The current script handles:
- Markdown headings (`#` to `######`)
- paragraphs
- unordered lists
- ordered lists
- Markdown tables
- fenced code blocks
- simple cover recognition
- TXT paragraph blocks

References:
- `skills/word-expert-formatting/scripts/text_to_docx.py:695`
- `skills/word-expert-formatting/scripts/text_to_docx.py:822`

## Important implementation note

The skill contract describes both Mode A and Mode B heading numbering.

However, the current Python script appears to implement decimal heading numbering only. That means the human-readable skill contract and the executable script are not fully aligned for Mode B yet.

References:
- `skills/word-expert-formatting/SKILL.md:43`
- `skills/word-expert-formatting/scripts/text_to_docx.py:313`
- `skills/word-expert-formatting/scripts/text_to_docx.py:354`

## Maintenance rule

If a formatting rule change must affect generated `.docx` output, update both:
- `skills/word-expert-formatting/SKILL.md`
- `skills/word-expert-formatting/scripts/text_to_docx.py`

Reference: `skills/word-expert-formatting/SKILL.md:165`
