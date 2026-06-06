# document-format

This repository documents the skills under `skills/`.

At the moment, the primary skill is `skills/word-expert-formatting`, a custom Claude Code skill for formatting Chinese Word content and supporting local `.docx` generation from Markdown or plain text, plus in-place refresh and verification for existing `.docx` files.

It now supports both Mode A and Mode B heading numbering, and `--auto-toc` inserts a Word TOC field instead of static TOC text.

## Skills

### `word-expert-formatting`

Location:
- `skills/word-expert-formatting/SKILL.md`
- `skills/word-expert-formatting/scripts/text_to_docx.py`

Purpose:
- turn raw Markdown, TXT, HTML, or rough outlines into a Word-ready structured format
- apply a strict Chinese formal-document style model
- support both Mode A and Mode B heading numbering
- support local `.docx` generation when the input is `.md`, `.markdown`, or `.txt`
- support in-place normalization and verification when the input is an existing `.docx`

Reference: `skills/word-expert-formatting/SKILL.md:1`

## What the skill defines

`skills/word-expert-formatting/SKILL.md` defines:
- supported input shapes
- heading hierarchy rules
- numbering modes and the CLI switch that selects them
- cover detection rules
- style mapping for headings, body text, tables, and footer page numbers
- output contracts for structured formatting results
- body paragraphs are normalized by clearing inherited indent, line spacing, and alignment before applying justified alignment and a 2-character first-line indent
- existing `.docx` refresh preserves images, tables, page breaks, and section structure instead of deleting and rebuilding the body
- existing `.docx` verification checks both paragraph layout and structure preservation

Reference: `skills/word-expert-formatting/SKILL.md:27`

## Local DOCX workflow

The local implementation lives in:

`skills/word-expert-formatting/scripts/text_to_docx.py`

The script currently:
- accepts `.md`, `.markdown`, `.txt`, and `.docx`
- formats headings, paragraphs, lists, tables, code blocks, footer page numbers, and simple cover sections
- supports both Mode A and Mode B heading numbering through `--numbering-mode`
- can reserve a placeholder cover page when no cover is detected
- can generate a Word TOC field page when no explicit TOC heading is detected
- can take explicit cover / TOC decisions for a run, including manual cover text input
- refreshes existing `.docx` files in place instead of deleting and rebuilding the body
- preserves images, tables, page breaks, and section structure during existing `.docx` refresh
- uses a more formal pagination model when a cover page exists: the cover has no page number and the body section restarts numbering from 1
- writes a `.docx` file

Reference: `skills/word-expert-formatting/scripts/text_to_docx.py:840`

Run:

```bash
python3 /Users/huangcn/github/document-format/skills/word-expert-formatting/scripts/text_to_docx.py <input-file> [output.docx] [--reserve-cover] [--auto-toc] [--with-cover|--without-cover] [--cover-text <text>] [--with-toc|--without-toc] [--numbering-mode A|B]
```

If `output.docx` is omitted, the script writes a file next to the input with the same basename.

For skill-driven runs in this repository, ask first:
1. whether to generate a cover page
2. if yes, what cover text to use (manual input; default empty)
3. whether to generate a TOC page

The skill now treats those answers as explicit per-run decisions:
- generate cover + non-empty text: render the first non-empty line as the cover title and later non-empty lines as centered metadata
- generate cover + empty text: render a blank placeholder cover page
- do not generate cover: suppress automatic cover detection and placeholder insertion for that run
- TOC generation is asked every time and then mapped to the script invocation for that run

References:
- `skills/word-expert-formatting/SKILL.md:127`
- `skills/word-expert-formatting/scripts/text_to_docx.py:1163`

## Before use

Before using the local DOCX workflow, check:
- Python 3 is available
- `python-docx` is installed

The script now performs these checks at startup and exits with a clear error message if a required dependency is missing.
The output document is now created from a blank Word document, so the workflow does not depend on a local template `.docx` file.

## Requirements

- Python 3
- `python-docx`

Reference: `skills/word-expert-formatting/scripts/text_to_docx.py:8`

## Supported features

Current script support includes:
- Markdown headings (`#` to `######`)
- paragraphs
- unordered lists
- ordered lists
- Markdown tables
- fenced code blocks
- simple cover recognition
- multi-line cover title blocks for existing `.docx` cover refresh
- for existing `.docx` cover titles, refresh only corrects font and size; it does not change existing title alignment
- placeholder cover insertion with `--reserve-cover`
- explicit cover generation with `--with-cover` and `--cover-text`
- explicit cover suppression with `--without-cover`
- explicit TOC detection (`目录`, `TOC`, `Table of Contents`)
- Word TOC field insertion with `--auto-toc`
- explicit TOC generation with `--with-toc`
- explicit TOC suppression with `--without-toc`
- Mode A / Mode B heading numbering with `--numbering-mode`
- formal pagination when a cover exists: no page number on the cover, body numbering restarts from 1
- existing `.docx` in-place refresh with structure preservation
- existing `.docx` verification for cover, TOC, body layout, tables, numbering, and structure counts
- TXT paragraph blocks
- minimal TXT heading detection for TOC generation

References:
- `skills/word-expert-formatting/scripts/text_to_docx.py:695`
- `skills/word-expert-formatting/scripts/text_to_docx.py:822`

## Important implementation note

The skill contract describes both Mode A and Mode B heading numbering.

The current Python script now implements both:
- Mode A uses all-decimal heading numbering
- Mode B uses Chinese top-level headings (`一、`, `二、`, `三、`, ...) with decimal numbering for lower levels

Also, `--auto-toc` now inserts a real Word TOC field instead of static text entries. If the table of contents does not appear updated immediately in Word, use Word's update-table action after opening the document.

The newer explicit switches are intended for skill-driven runs:
- `--with-cover` / `--without-cover` override automatic cover detection for that run
- `--cover-text` implies `--with-cover` when used alone
- `--with-toc` / `--without-toc` override fallback TOC generation for that run
- `--reserve-cover` and `--auto-toc` remain available for direct CLI compatibility

When a cover page is present or `--reserve-cover` / `--with-cover` is used, the script now creates a separate body section so the cover stays unnumbered and the body starts again at page 1.

References:
- `skills/word-expert-formatting/SKILL.md:43`
- `skills/word-expert-formatting/scripts/text_to_docx.py:336`
- `skills/word-expert-formatting/scripts/text_to_docx.py:605`

## Maintenance rule

If a formatting rule change must affect generated `.docx` output, update both:
- `skills/word-expert-formatting/SKILL.md`
- `skills/word-expert-formatting/scripts/text_to_docx.py`

Reference: `skills/word-expert-formatting/SKILL.md:165`
