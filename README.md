# document-format

This repository documents the skills under `skills/`.

At the moment, the primary skill is `skills/word-expert-formatting`, a custom Claude Code skill for formatting Chinese Word content and supporting local `.docx` generation from Markdown or plain text, plus refreshed-copy normalization and verification for existing `.docx` files.

It applies all-decimal heading numbering, and `--auto-toc` inserts a Word TOC field instead of static TOC text.

## Skills

### `word-expert-formatting`

Location:
- `skills/word-expert-formatting/SKILL.md`
- `skills/word-expert-formatting/scripts/text_to_docx.py`

Purpose:
- turn raw Markdown, TXT, HTML, or rough outlines into a Word-ready structured format
- apply a strict Chinese formal-document style model
- apply all-decimal heading numbering
- support local `.docx` generation when the input is `.md`, `.markdown`, or `.txt`
- support refreshed-copy normalization and verification when the input is an existing `.docx`

Reference: `skills/word-expert-formatting/SKILL.md:1`

## What the skill defines

`skills/word-expert-formatting/SKILL.md` defines:
- supported input shapes
- heading hierarchy rules
- heading numbering behavior
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
- uses all-decimal heading numbering
- can reserve a placeholder cover page when no cover is detected
- can generate a Word TOC field page when no explicit TOC heading is detected
- can take explicit cover / TOC decisions for a run, including manual cover text input
- refreshes existing `.docx` files into a new output file instead of deleting and rebuilding the body
- preserves images, tables, page breaks, and section structure during existing `.docx` refresh
- uses a more formal pagination model when a cover page exists: the cover has no page number and the body section restarts numbering from 1
- writes a `.docx` file

Reference: `skills/word-expert-formatting/scripts/text_to_docx.py:840`

Run:

```bash
python3 skills/word-expert-formatting/scripts/text_to_docx.py <input-file> [output.docx] [--reserve-cover] [--auto-toc] [--with-cover|--without-cover] [--cover-text <text>] [--with-toc|--without-toc]
```

If `output.docx` is omitted, `.md` / `.markdown` / `.txt` inputs still write a same-basename `.docx`, while `.docx` inputs write `<stem>.refreshed.docx` next to the source file.

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
- explicit cover suppression with `--without-cover` for generated output, while existing `.docx` refresh keeps any cover that is already present unchanged
- explicit TOC detection (`目录`, `TOC`, `Table of Contents`)
- Word TOC field insertion with `--auto-toc`
- explicit TOC generation with `--with-toc`
- explicit TOC suppression with `--without-toc` for generated output, while existing `.docx` refresh keeps any TOC that is already present
- all-decimal heading numbering
- formal pagination when a cover exists: no page number on the cover, body numbering restarts from 1
- refreshed-copy normalization for existing `.docx` with structure preservation
- existing `.docx` verification for cover, TOC, body layout, tables, numbering, and structure counts
- TXT paragraph blocks
- minimal TXT heading detection for TOC generation

References:
- `skills/word-expert-formatting/scripts/text_to_docx.py:695`
- `skills/word-expert-formatting/scripts/text_to_docx.py:822`

## XML debug lane

For tricky existing `.docx` cases where Word rendering does not match the normal refresh result, this skill also includes a local XML debug helper:

```bash
python3 skills/word-expert-formatting/scripts/debug_docx_xml.py unpack <input.docx> <output-dir>
python3 skills/word-expert-formatting/scripts/debug_docx_xml.py validate <output-dir> --original <input.docx>
python3 skills/word-expert-formatting/scripts/debug_docx_xml.py repack <output-dir> <output.docx> --original <input.docx>
```

This path is intended for XML-level debugging and repair, not for normal generation.

## Why this skill exists alongside a generic DOCX skill

| Dimension | Generic DOCX skill | `word-expert-formatting` |
|---|---|---|
| Positioning | Low-level Office/XML toolbox | End-to-end document-production workflow |
| Main job | Unpack, inspect, validate, repack `.docx` | Generate or refresh documents to this repository's formatting contract |
| Best at | XML surgery, schema issues, relationships, tracked changes | Cover, TOC, heading hierarchy, pagination, body and table formatting |
| Existing `.docx` handling | Good for direct XML repair | Good for refreshed-copy normalization with structure preservation |
| Validation focus | XML correctness | Semantic output quality + structure preservation |
| User cost | Higher; better for expert debugging | Lower; better for repeatable daily use |
| Use when | Word rendering and XML disagree, or XML needs surgical fixes | Normal formatting, refresh, and verification work |

In short:
- the generic DOCX skill is the XML surgery toolkit
- `word-expert-formatting` is the repeatable production line for this repository's target document format
- the intended workflow is: use `word-expert-formatting` by default, and fall back to the XML debug lane only for hard low-level cases

## Important implementation note

The skill contract and current Python script use a single heading numbering scheme:
- headings use an all-decimal hierarchy such as `1`, `1.1`, and `1.1.1`
- this is intended for technical documents, project plans, implementation plans, and similar structured reports

Also, `--auto-toc` now inserts a real Word TOC field instead of static text entries. If the table of contents does not appear updated immediately in Word, use Word's update-table action after opening the document.

The newer explicit switches are intended for skill-driven runs:
- `--with-cover` / `--without-cover` override automatic cover detection for that run; on existing `.docx`, `--without-cover` only suppresses creating a new cover and does not remove or restyle one that already exists
- `--cover-text` implies `--with-cover` when used alone
- `--with-toc` / `--without-toc` override fallback TOC generation for that run; on existing `.docx`, `--without-toc` only suppresses creating a new TOC and does not remove one that already exists
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
