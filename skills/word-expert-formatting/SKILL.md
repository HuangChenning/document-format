---
name: word-expert-formatting
description: "Use this skill whenever the user wants to format Chinese Word content according to a strict document style guide, especially for 公文, 商务报告, formal reports, or .docx deliverables with required heading levels, cover page text, tables, page numbers, fonts, line spacing, paragraph spacing, first-line indentation, or Chinese numbering patterns. Trigger when the user mentions Word 排版, 文档格式规范, 标题层级, 一级标题/二级标题, 中文编号, 封面, 页脚页码, 表格格式, 宋体, 四号, 小四, 五号, 首行缩进, or asks to turn raw text / Markdown / TXT / HTML into a Word-ready structured format. If the user needs a final `.docx` file from `.md`, `.markdown`, or `.txt` in this repository, use the local script workflow in this skill instead of only describing the format. Do not use this skill for PDFs, spreadsheets, slide decks, or generic writing requests without formatting constraints."
---

# Word expert formatting

## Purpose

Use this skill to translate raw content into a Word-ready structure that follows a strict Chinese formatting specification.

This skill is optimized for inputs such as:
- plain text
- Markdown
- HTML
- existing `.docx` files that need refreshed-copy normalization
- rough outlines with heading levels
- partially structured report drafts

This skill is not primarily about producing a `.docx` file by itself. Its main job is to:
1. identify document structure
2. map each block to the required style
3. rebuild heading numbering
4. output a structured formatting result that another Word-generation workflow can apply directly

If the user explicitly needs a finished `.docx`, use this skill to determine the formatting structure first, then hand off to a Word/document-generation workflow.

In this repository, the local script workflow applies an all-decimal heading hierarchy, can insert a real Word TOC field when `--auto-toc` is enabled, supports existing `.docx` refreshed-copy normalization, and uses a more formal pagination model when a cover page exists.

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

### 2. Determine heading numbering

Before formatting, use the default all-decimal heading hierarchy.

- **Heading hierarchy**
  - level 1 → `1`, `2`, `3`
  - level 2 → `1.1`, `1.2`
  - level 3 → `1.1.1`
  - level 4 → `1.1.1.1`
  - level 5 → `1.1.1.1.1`

Selection guidance:
- use this hierarchy when the document should use Arabic numerals at every heading level, such as technical documents, project plans, implementation plans, or documents that already use `1` / `1.1` / `1.1.1`

If the source content has inconsistent numbering, normalize it to this hierarchy.

### 3. Rebuild heading hierarchy

Infer the intended heading level from the source structure, wording, and numbering.

When rebuilding hierarchy:
- keep sibling numbering consistent
- ensure child headings match their parent path
- do not skip levels unless the source clearly requires it
- if the source is ambiguous, surface the ambiguity instead of guessing silently

### 4. Apply style mapping

Map every block to the exact style below. Do not improvise font sizes, spacing, or indentation.

### 5. Multi-level heading contract

When the document contains level 3, level 4, level 5, or deeper headings, apply the same contract to every heading level that is present.

For every heading level that exists in the source:
- determine the semantic level first, then apply formatting; do not infer the final style from visual boldness alone
- keep one numbering path per heading level; do not mix manual prefixes, paragraph numbering, and style numbering in the same visible heading
- if the document engine already provides heading numbering, remove or normalize non-standard manual prefixes that would create duplicated visible numbering
- if the source heading text contains a numbering-like prefix that is not part of the selected numbering system, strip or rewrite it before final verification
- if the source uses a deeper heading level, do not silently collapse it upward into a shallower level only because the shallower style looks similar
- if a level is present in the body, the same level must be treated consistently in headings, cross-references, and TOC entries

For numbering behavior:
- use one continuous all-decimal path across every present heading level
- if the source contains deeper levels than level 5, continue the same decimal path instead of inventing a new numbering pattern
- never leave one heading level manually numbered while an adjacent heading level uses automatic numbering unless the user explicitly asked for that mixed behavior

A heading level is considered correctly converted only when all three layers agree:
- the heading style definition matches the intended font, size, boldness, spacing, and outline level
- the paragraph instances assigned to that heading level actually inherit or carry the intended visible formatting
- the numbering definition attached to that level renders numbering with the intended font, size, boldness, and numbering text pattern

If any one of these three layers is inconsistent, treat the heading conversion for that level as failed rather than partially successful.

## Style matrix

### Heading levels

| Element | Font | Size | Line spacing | Space before | Space after | Other |
|---|---|---:|---|---|---|---|
| Document title | 黑体 | 28号 | single | 10 pt | 10 pt | bold |
| Level 1 heading | 黑体 | 三号 | single | 10 pt | 10 pt | bold, use all-decimal heading numbering |
| Level 2 heading | 黑体 | 小三 | single | 8 pt | 8 pt | bold |
| Level 3 heading | 黑体 | 四号 | 1.5 lines | 0 | 0 | bold |
| Level 4 heading | 黑体 | 小四 | 1.5 lines | 0 | 0 | bold |
| Level 5 heading | 黑体 | 小四 | 1.5 lines | 0 | 0 | bold |
| Level 6 heading | 黑体 | 小四 | 1.5 lines | 0 | 0 | bold |

### Body text

| Element | Font | Size | Line spacing | Space before | Space after | Other |
|---|---|---:|---|---|---|---|
| Standard body paragraph | 宋体 | 小四 | 1.5 lines | 0 | 0 | first-line indent: 2 characters; justify; clear existing style-level and paragraph-level left indent, first-line indent, spacing, tabs, numbering, and alignment before applying the standard style |

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
| Cover title | 宋体 | 二号 | 1.5 lines | 0.5 lines | 0.5 lines | centered; when the cover title is split across multiple centered lines, every title line must use the same font, size, boldness, spacing, and alignment; for existing `.docx` refresh, if a cover page already exists, only confirm and correct the cover-title font and size, and do not change the title alignment |
| Company name / date | 宋体 | 四号 | 1.5 lines | 0 | 0 | centered |
| Cover corner metadata | 黑体 | 五号 | single | 0 | 0 | top-left, not bold, no left indent |

### Footer / page number

| Element | Font | Size | Line spacing | Alignment | Other |
|---|---|---:|---|---|---|
| Page number | 宋体 | 五号 | 1.5 lines | centered | apply in footer |

## Local DOCX generation workflow

When the user wants an actual `.docx` file in this repository, prefer the built-in local script workflow instead of only returning structured formatting guidance.

Use this workflow in two modes:
- source `.md`, `.markdown`, or `.txt`: generate a new `.docx`
- source `.docx`: read the existing document and write a refreshed copy by normalizing cover, headings, body paragraphs, and table text without deleting and rebuilding the body

Before launching the local DOCX workflow through this skill, ask these questions in order:
1. whether to generate a cover page
2. if a cover page should be generated, what cover text to use (manual input; default empty)
3. whether to generate a TOC page

Apply the user's answers directly:
- if the user chooses to generate a cover page, treat that as an explicit cover decision and do not rely on automatic cover detection
- if the user provides cover text, pass it through to the script and render the first non-empty line as the cover title, with later non-empty lines rendered as centered metadata
- if the user chooses to generate a cover page but leaves the text empty, generate a blank placeholder cover page
- if the user chooses not to generate a cover page, suppress both automatic cover detection and placeholder-cover insertion for that run
- ask about TOC generation every time, then map the answer to the script invocation for that run

Use:

```bash
python3 skills/word-expert-formatting/scripts/text_to_docx.py <input-file> [output.docx] [--reserve-cover] [--auto-toc] [--with-cover|--without-cover] [--cover-text <text>] [--with-toc|--without-toc]
```


Apply this path when:
- the user wants a final Word document, not just a style spec
- the input is Markdown, TXT, or an existing `.docx`
- the content mainly consists of headings, paragraphs, tables, and fenced code blocks

### Before use

Before running the local DOCX workflow, check:
- Python 3 is available
- `python-docx` is installed

The script now performs these checks at startup and exits with a clear error if a required dependency is missing.

This local script is the default implementation for this repository because it already applies the core formatting rules from this skill, including heading mapping, body formatting, table formatting, page setup, page numbers, first-line indent in Word character units, explicit left-indent clearing, basic cover recognition, and existing `.docx` refreshed-copy normalization.
It now creates the output document from a blank Word document and does not depend on a local template `.docx` file.

Additional options:
- `--reserve-cover`: when no explicit cover decision is provided and no cover is detected, insert a placeholder cover page
- `--auto-toc`: when no explicit TOC decision is provided and no explicit TOC heading is detected, insert a generated Word TOC field page
- `--with-cover`: always generate a cover page for this run and bypass automatic cover detection
- `--without-cover`: never generate a cover page for this run and bypass automatic cover detection; for existing `.docx`, keep any cover that already exists
- `--cover-text <text>`: explicit cover text; the first non-empty line becomes the title and later non-empty lines become centered metadata; implies `--with-cover` when used alone
- `--with-toc`: explicitly request generated TOC insertion for this run
- `--without-toc`: explicitly suppress generated TOC insertion for this run; for existing `.docx`, keep any TOC that already exists

Compatibility behavior:
- the new explicit cover / TOC options are the preferred path for skill-driven runs
- `--reserve-cover` and `--auto-toc` remain supported for direct CLI use
- when an explicit cover or TOC decision is provided, it overrides the older fallback switches

Pagination behavior:
- when a cover page exists, the cover section has no page number
- after the cover, the script starts a new body section and restarts page numbering from 1
- when no cover exists, the document keeps a single section and uses the normal page-number footer directly

TOC detection rules:
- only explicit TOC headings count as an existing TOC
- accepted markers: `目录`, `TOC`, `Table of Contents`
- when an explicit TOC heading is already present, generated TOC insertion must not insert another TOC page
- the inserted TOC is a real Word field, so users can update it inside Word after opening the document

TXT support note:
- the script now includes a minimal TXT heading detector so TXT inputs can also participate in TOC generation
- numeric headings such as `1`, `1.1`, and `1.1.1` are recognized conservatively

### XML debug lane

When an existing `.docx` behaves unexpectedly in Word and the normal `python-docx` refresh path is not enough, use the local XML debug lane instead of guessing at higher-level API behavior.

Use:

```bash
python3 skills/word-expert-formatting/scripts/debug_docx_xml.py unpack <input.docx> <output-dir>
python3 skills/word-expert-formatting/scripts/debug_docx_xml.py validate <output-dir> --original <input.docx>
python3 skills/word-expert-formatting/scripts/debug_docx_xml.py repack <output-dir> <output.docx> --original <input.docx>
```

Use this lane when:
- Word rendering differs from what `python-docx` or raw text inspection suggests
- numbering, TOC, style bindings, or field behavior need direct XML inspection
- you need schema validation before trusting a low-level `.docx` fix

This is a debugging and repair lane, not the default generation path.

#### When to recommend the XML debug lane

Do not jump to XML editing by default. Prefer the normal workflow first, then recommend the XML debug lane only when one of these signals is present.

**Level 1 — stay on the normal workflow**
- ordinary style, spacing, indentation, cover, TOC decision, or table-formatting issues
- first-pass heading mapping issues that can be explained and fixed through the current script workflow

**Level 2 — recommend the XML debug lane**
- verification fails around heading numbering chains, TOC/body consistency, field behavior, or section/page-break preservation
- the generated document passes some checks, but visible Word output still disagrees with the expected result
- an existing `.docx` needs refreshed-copy normalization and the issue appears tied to `numbering.xml`, `styles.xml`, `settings.xml`, or `document.xml.rels`

**Level 3 — strongly recommend the XML debug lane**
- the same rendering/compatibility problem survives more than one targeted fix in the normal workflow
- Word rendering clearly disagrees with both XML inspection and non-Word renderers
- structure preservation is at risk, such as disappearing images, tables, page breaks, sections, comments, or tracked changes

#### Standard escalation messages

When recommending the XML debug lane, say so explicitly and give the reason.

Use wording like:
- `This looks more like a DOCX XML structure or field-binding problem than a normal formatting issue. I recommend switching to the XML debug lane for unpack -> validate -> repack.`
- `The script-level checks are not enough here because Word rendering still disagrees with the expected result. I recommend inspecting the unpacked DOCX XML directly.`
- `This existing .docx appears to need low-level inspection of numbering, style bindings, relationships, or settings. I recommend using the local XML debug lane before making more high-level guesses.`

When useful, immediately provide the next commands:

```bash
python3 skills/word-expert-formatting/scripts/debug_docx_xml.py unpack <input.docx> <output-dir>
python3 skills/word-expert-formatting/scripts/debug_docx_xml.py validate <output-dir> --original <input.docx>
```

### Existing DOCX refresh strategy

For existing `.docx` inputs:
- do not delete and rebuild the body from extracted plain text
- preserve existing images, drawings, embedded objects, tables, page breaks, and section structure
- normalize the document into a refreshed copy after the detected body boundary
- clear residual style-level and paragraph-level left indent before applying the standard body first-line indent
- keep cover corner metadata left-aligned and unindented
- treat a centered multi-line cover title block as one title block and enforce uniform title styling across every title line

### Verification chain

Every conversion or refresh must be checked against the full workflow contract:
- cover presence and cover title block formatting
- cover corner metadata formatting
- TOC presence and TOC/body consistency when a TOC is expected
- body paragraph font, size, alignment, line spacing, paragraph spacing, and first-line indent
- absence of residual body left indent and body numbering
- table text formatting
- heading style definitions, heading paragraph instances, and numbering definitions

For existing `.docx` refresh, verification must also confirm structure preservation:
- image/drawing/object counts must not decrease
- table counts must not decrease
- page-break and section structure must not be lost


The current script uses **label-first recognition with heuristic fallback**.

For existing `.docx`, the cover boundary is inferred more conservatively:
- top-left corner metadata lines (`编号`, `版本号`, `受控状态`, `密级`) are treated as cover corner metadata
- a consecutive centered title block is treated as the cover title block, even when it spans multiple lines
- cover detection stops before later front-matter pages such as change records or approval pages

Preferred explicit labels:
- `封面标题:` / `封面标题：`
- `公司名称:` / `公司名称：`
- `日期:` / `日期：`
- bracketed forms such as `[封面标题]` and `[封面公司/日期]`

If explicit cover decisions are absent, the script only inspects the first few non-empty lines of the document:
- the first short standalone line is treated as the cover title
- the next 1-2 short lines are treated as company name / date
- detection stops immediately when it encounters a Markdown heading, table, fenced code block, or an obvious long body line
- when a cover block is recognized, the script creates a separate body section after the cover so the cover remains unnumbered
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
- `[一级标题]`
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
- `heading_numbering`

When producing structured data, preserve the same formatting rules as the default output.

## Guardrails

Apply these checks before finalizing the result.

1. **Body indentation must exist**
   - Every normal body paragraph must carry first-line indent of 2 characters.
   - Before applying the standard body style, clear any existing paragraph indent, line spacing, and alignment so the final visible first-line indent is exactly 2 characters.
   - Standard body paragraphs must use justified alignment.

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

## Post-conversion verification

After generating a `.docx`, verify the conversion result before reporting success.

### What to verify

1. **Document title / cover title**
   - confirm the visible title text is correct
   - confirm it was not downgraded into body text or the wrong heading level
   - confirm the title font, size, alignment, and spacing follow the selected rule set

2. **Heading hierarchy**
   - inspect every heading level that is actually present in the document, with level 1 through level 3 as the minimum scope
   - when level 4, level 5, or deeper headings are present, include those levels in the same verification pass rather than stopping at level 3
   - confirm numbering matches the default all-decimal heading hierarchy
   - confirm parent / child paths are consistent and no level was silently skipped or flattened
   - confirm no visible heading shows duplicated numbering caused by a manual prefix plus automatic numbering

3. **Heading style integrity by level**
   - for each heading level present, verify the style definition itself: font, size, boldness, spacing, outline level, and style binding
   - for each heading level present, verify representative paragraph instances rather than checking the style definition only
   - for each heading level present, verify the attached numbering definition rather than checking only the heading text after rendering
   - if a heading level uses mixed run formatting inside one visible heading, treat that level as failed conversion unless the mixed formatting is explicitly required by the source

4. **Cover corner metadata**
   - if the cover contains items such as `编号`, `版本号`, `受控状态`, or `密级`, render them in the top-left area of the cover page
   - these cover-corner items must use 黑体、五号、not bold
   - do not merge these items into the centered cover title block

5. **Cover title verification**
   - identify the cover title block from non-empty cover-title paragraphs only; ignore blank placeholder paragraphs when deciding title / title / date roles
   - verify cover-title conversion against the non-empty paragraph sequence rather than raw paragraph indexes alone
   - verify font size consistency run by run inside each cover title paragraph, not only at the paragraph level
   - if a visible cover title paragraph contains mixed or partially unconverted font sizes across runs, report it as a failed conversion

6. **TOC consistency**
   - update fields / refresh the table of contents before final verification when the workflow or environment allows it
   - verify that every TOC entry for every heading level that is present in the document matches the visible body heading text and numbering
   - verify that TOC hierarchy, numbering paths, and page-number references are consistent with the body
   - if the environment cannot refresh TOC fields automatically, say so explicitly and still compare the current TOC entries against the body headings that are present
   - this TOC check is required on every conversion run, not only when the user asks for it

7. **Key styles**
   - confirm body paragraphs keep first-line indent of 2 characters
   - confirm table text uses table styling rather than body indentation
   - confirm page-number footer exists where required
   - confirm cover and TOC pagination behavior matches the workflow rules

### How to verify

Use a practical verification path, not only a code-path assumption.

Preferred checks:
- inspect the generated document structure directly when possible
- if needed, convert the generated `.docx` to Markdown / XML / PDF / page images and inspect the rendered result
- compare the first page, main title area, several representative heading paragraphs, and the TOC page
- compare TOC entries against the actual body headings one by one for every heading level present in the converted document
- for each heading level present, check at least these three things separately: style definition, paragraph instance formatting, and numbering-definition formatting
- if a heading level appears only a few times, verify all of them; if it appears many times, verify all occurrences for numbering continuity and sample multiple occurrences for visible style consistency
- if any title, heading, numbering mark, or TOC entry looks ambiguous after conversion, surface it explicitly instead of claiming success

Mandatory failure conditions:
- duplicated visible numbering remains in any heading
- one heading level is styled correctly in the text runs but incorrectly in numbering marks
- one heading level looks correct in sampled paragraphs but the underlying style or numbering definition is inconsistent
- TOC and body disagree for any verified heading level
- a deeper heading level was present in the source but got flattened, skipped, or reassigned without being explicitly reported

### Verification output expectation

When returning the result, include a brief verification summary that states:
- whether the title conversion met expectations
- whether heading hierarchy and numbering met expectations for every heading level that was present and verified
- whether the TOC entries match the body headings and page references as far as could be verified in the current environment
- whether any style mismatch remains and where it appears
- whether any heading level passed only partially, and if so, which verification layer failed: style definition, paragraph instance, or numbering definition

## Response pattern

When formatting content, prefer this response shape:

```markdown
【格式化执行完毕】

[封面标题] ...
[封面公司/日期] ...
[一级标题] ...
[二级标题] ...
[三级标题] ...
[正文] ...
[表格文字] ...
[页脚] ...
```

If some sections do not exist in the source, omit them.

## Example

**Input idea:**
“把这份提纲整理成标准 Word 格式：项目背景、实施范围、上线计划。”

**Output pattern:**

```markdown
【格式化执行完毕】

[一级标题] 1 项目背景
样式：宋体，四号，单倍行距，段前 10 pt，段后 10 pt

[一级标题] 2 实施范围
样式：宋体，四号，单倍行距，段前 10 pt，段后 10 pt

[一级标题] 3 上线计划
样式：宋体，四号，单倍行距，段前 10 pt，段后 10 pt
```

## Handoff guidance

If the user wants a final `.docx` file:
1. use this skill to finalize structure, numbering, and style mapping
2. then generate the document with a Word-capable workflow
3. keep the formatting values from this skill as the source of truth

If the user only wants a formatting specification or a review of whether content conforms to the style guide, stay within this skill and do not force document generation.
