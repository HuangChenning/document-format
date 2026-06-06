#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

FONT_SONG = "宋体"
FONT_HEI = "黑体"
FONT_MONO = "Menlo"
SIZE_TITLE = Pt(28)
SIZE_H1 = Pt(16)
SIZE_H2 = Pt(15)
SIZE_H3 = Pt(14)
SIZE_H4 = Pt(12)
SIZE_H5 = Pt(12)
SIZE_H6 = Pt(12)
SIZE_BODY = Pt(12)
SIZE_TABLE = Pt(10.5)
SIZE_CODE = Pt(9)
A4_WIDTH = Cm(21)
A4_HEIGHT = Cm(29.7)
PAGE_MARGIN_TOP = Cm(2.54)
PAGE_MARGIN_BOTTOM = Cm(2.54)
PAGE_MARGIN_LEFT = Cm(3.17)
PAGE_MARGIN_RIGHT = Cm(3.17)
TABLE_CONTENT_WIDTH_DXA = 11906 - 1800 - 1800
TABLE_CELL_MARGIN_TOP = 80
TABLE_CELL_MARGIN_BOTTOM = 80
TABLE_CELL_MARGIN_LEFT = 120
TABLE_CELL_MARGIN_RIGHT = 120
CODE_LEFT_INDENT = Cm(0.74)
CODE_RIGHT_INDENT = Cm(0.74)
CODE_SHADE = "F2F2F2"
TABLE_HEADER_SHADE = "D9EAF7"
COVER_TOP_SPACER_COUNT = 6
TEMPLATE_DOCX_PATH = Path('/Users/huangcn/github/document-format/2025-2026年度数据库运维服务总结_20260522_修改.docx')
TITLE_NUMBERING_ABSTRACT_ID = '700'
TITLE_NUMBERING_ID = '701'
BULLET_NUMBERING_ABSTRACT_ID = '702'
BULLET_NUMBERING_ID = '703'


class StyleSpec:
    def __init__(self, font_name: str, font_size, line_spacing: str, before_pt: float, after_pt: float, first_line_chars: int = 0, bold: bool | None = None):
        self.font_name = font_name
        self.font_size = font_size
        self.line_spacing = line_spacing
        self.before_pt = before_pt
        self.after_pt = after_pt
        self.first_line_chars = first_line_chars
        self.bold = bold


TITLE_SPEC = StyleSpec(FONT_HEI, SIZE_TITLE, "single", 10, 10, 0, True)
H1_SPEC = StyleSpec(FONT_HEI, SIZE_H1, "single", 10, 10, 0, True)
H2_SPEC = StyleSpec(FONT_HEI, SIZE_H2, "single", 8, 8, 0, True)
H3_SPEC = StyleSpec(FONT_HEI, SIZE_H3, "one_point_five", 0, 0, 0, True)
H4_SPEC = StyleSpec(FONT_HEI, SIZE_H4, "one_point_five", 0, 0, 0, True)
H5_SPEC = StyleSpec(FONT_HEI, SIZE_H5, "one_point_five", 0, 0, 0, True)
H6_SPEC = StyleSpec(FONT_HEI, SIZE_H6, "one_point_five", 0, 0, 0, True)
COVER_TITLE_SPEC = StyleSpec(FONT_SONG, Pt(22), "one_point_five", 9, 9, 0, None)
COVER_META_SPEC = StyleSpec(FONT_SONG, Pt(14), "one_point_five", 0, 0, 0, None)
BODY_SPEC = StyleSpec(FONT_SONG, SIZE_BODY, "one_point_five", 0, 0, 2, None)
TABLE_SPEC = StyleSpec(FONT_SONG, SIZE_TABLE, "one_point_five", 0, 0, 0, None)
CODE_SPEC = StyleSpec(FONT_MONO, SIZE_CODE, "one_point_five", 6, 6, 0, None)
PAGE_SPEC = StyleSpec(FONT_SONG, Pt(10.5), "one_point_five", 0, 0, 0, None)

HEADING_RE = re.compile(r'^(#{1,6})\s+(.*)$')
BULLET_RE = re.compile(r'^(\s*)[-*+]\s+(.*)$')
ORDERED_RE = re.compile(r'^\s*\d+[.)]\s+(.*)$')
LEADING_HEADING_NUMBER_RE = re.compile(r'^\s*((?:\d+\.)*\d+)\s+')
COVER_LABEL_RE = re.compile(r'^\s*(?:\[(封面标题|封面公司/日期|公司名称|日期)\]|(封面标题|封面公司/日期|公司名称|日期))\s*[:：]\s*(.*)$')
DATE_RE = re.compile(r'\b\d{4}[-/]\d{1,2}[-/]\d{1,2}\b|\b\d{4}年\d{1,2}月\d{1,2}日\b')


def clear_document_body(doc: Document):
    body = doc._element.body
    for child in list(body):
        if child.tag != qn('w:sectPr'):
            body.remove(child)


def get_heading_style_name(level: int) -> str:
    return f'Heading {level}'


def get_heading_style_id(level: int) -> str:
    return {
        1: '11',
        2: '22',
        3: '31',
        4: '41',
        5: '51',
        6: '6',
    }.get(level, '6')


def get_or_create_paragraph_style(doc: Document, style_name: str, style_id: str, spec: StyleSpec):
    if style_id in doc.styles:
        style = doc.styles[style_id]
    elif style_name in doc.styles:
        style = doc.styles[style_name]
    else:
        raise KeyError(f'Style not found: {style_name} / {style_id}')

    style.font.name = spec.font_name
    style.font.size = spec.font_size
    style.font.bold = bool(spec.bold)

    r_pr = style._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement('w:rFonts')
        r_pr.append(r_fonts)
    r_fonts.set(qn('w:ascii'), spec.font_name)
    r_fonts.set(qn('w:hAnsi'), spec.font_name)
    r_fonts.set(qn('w:eastAsia'), spec.font_name)

    return style


def set_run_font(run, font_name: str, font_size):
    run.font.name = font_name
    run.font.size = font_size
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement('w:rFonts')
        r_pr.append(r_fonts)
    r_fonts.set(qn('w:ascii'), font_name)
    r_fonts.set(qn('w:hAnsi'), font_name)
    r_fonts.set(qn('w:eastAsia'), font_name)


def insert_numpr_after_pstyle(p_pr, num_pr):
    inserted = False
    for idx, child in enumerate(list(p_pr)):
        if child.tag == qn('w:pStyle'):
            p_pr.insert(idx + 1, num_pr)
            inserted = True
            break
    if not inserted:
        p_pr.insert(0, num_pr)


def clear_paragraph_indentation(paragraph):
    p_pr = paragraph._p.get_or_add_pPr()
    for child in list(p_pr):
        if child.tag == qn('w:ind'):
            p_pr.remove(child)


def set_first_line_indent_chars(paragraph, chars: int):
    clear_paragraph_indentation(paragraph)
    p_pr = paragraph._p.get_or_add_pPr()
    ind = OxmlElement('w:ind')
    ind.set(qn('w:firstLineChars'), str(chars * 100))
    p_pr.append(ind)


def set_no_first_line_indent(paragraph):
    clear_paragraph_indentation(paragraph)
    p_pr = paragraph._p.get_or_add_pPr()
    ind = OxmlElement('w:ind')
    ind.set(qn('w:firstLineChars'), '0')
    p_pr.append(ind)


def set_paragraph_shading(paragraph, fill: str):
    p_pr = paragraph._p.get_or_add_pPr()
    for child in list(p_pr):
        if child.tag == qn('w:shd'):
            p_pr.remove(child)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), fill)
    p_pr.append(shd)


def apply_line_spacing(paragraph, mode: str):
    fmt = paragraph.paragraph_format
    if mode == 'single':
        fmt.line_spacing = 1.0
        fmt.line_spacing_rule = WD_LINE_SPACING.SINGLE
    elif mode == 'one_point_five':
        fmt.line_spacing = 1.5
        fmt.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    else:
        raise ValueError(f'Unsupported line spacing mode: {mode}')


def apply_paragraph_style(paragraph, spec: StyleSpec):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(spec.before_pt)
    fmt.space_after = Pt(spec.after_pt)
    fmt.left_indent = Cm(0)
    fmt.right_indent = Cm(0)
    apply_line_spacing(paragraph, spec.line_spacing)
    if spec.first_line_chars > 0:
        set_first_line_indent_chars(paragraph, spec.first_line_chars)
    else:
        set_no_first_line_indent(paragraph)

    for run in paragraph.runs:
        set_run_font(run, spec.font_name, spec.font_size)
        if spec.bold is not None:
            run.bold = spec.bold


def apply_code_block_style(paragraph):
    apply_paragraph_style(paragraph, CODE_SPEC)
    paragraph.paragraph_format.left_indent = CODE_LEFT_INDENT
    paragraph.paragraph_format.right_indent = CODE_RIGHT_INDENT
    set_paragraph_shading(paragraph, CODE_SHADE)


def get_heading_spec(level: int) -> StyleSpec:
    if level == 1:
        return H1_SPEC
    if level == 2:
        return H2_SPEC
    if level == 3:
        return H3_SPEC
    if level == 4:
        return H4_SPEC
    if level == 5:
        return H5_SPEC
    return H6_SPEC


def apply_bullet_style(paragraph, level: int):
    apply_paragraph_style(paragraph, BODY_SPEC)
    p_pr = paragraph._p.get_or_add_pPr()
    for child in list(p_pr):
        if child.tag == qn('w:numPr'):
            p_pr.remove(child)

    num_pr = OxmlElement('w:numPr')
    ilvl = OxmlElement('w:ilvl')
    ilvl.set(qn('w:val'), str(max(min(level, 2), 0)))
    num_id = OxmlElement('w:numId')
    num_id.set(qn('w:val'), BULLET_NUMBERING_ID)
    num_pr.append(ilvl)
    num_pr.append(num_id)
    p_pr.append(num_pr)


def set_cell_margins(cell, top=TABLE_CELL_MARGIN_TOP, start=TABLE_CELL_MARGIN_LEFT, bottom=TABLE_CELL_MARGIN_BOTTOM, end=TABLE_CELL_MARGIN_RIGHT):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in('w:tcMar')
    if tc_mar is None:
        tc_mar = OxmlElement('w:tcMar')
        tc_pr.append(tc_mar)
    for tag, value in [('w:top', top), ('w:start', start), ('w:bottom', bottom), ('w:end', end)]:
        node = tc_mar.find(qn(tag))
        if node is None:
            node = OxmlElement(tag)
            tc_mar.append(node)
        node.set(qn('w:w'), str(value))
        node.set(qn('w:type'), 'dxa')


def add_page_number(paragraph):
    run = paragraph.add_run()
    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = ' PAGE '
    fld_sep = OxmlElement('w:fldChar')
    fld_sep.set(qn('w:fldCharType'), 'separate')
    fld_text = OxmlElement('w:t')
    fld_text.text = '1'
    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_sep)
    run._r.append(fld_text)
    run._r.append(fld_end)
    set_run_font(run, PAGE_SPEC.font_name, PAGE_SPEC.font_size)


def ensure_numbering_child_order(numbering, child, kind: str):
    children = list(numbering)
    if kind == 'abstract':
        for idx, existing in enumerate(children):
            if existing.tag in {qn('w:num'), qn('w:numIdMacAtCleanup')}:
                numbering.insert(idx, child)
                return
        numbering.append(child)
        return

    if kind == 'num':
        for idx, existing in enumerate(children):
            if existing.tag == qn('w:numIdMacAtCleanup'):
                numbering.insert(idx, child)
                return
        numbering.append(child)
        return

    raise ValueError(f'Unsupported numbering child kind: {kind}')


def ensure_numbering(doc: Document):
    numbering = doc.part.numbering_part.numbering_definitions._numbering

    title_existing = numbering.xpath(f'./w:abstractNum[@w:abstractNumId="{TITLE_NUMBERING_ABSTRACT_ID}"]')
    if not title_existing:
        title_abstract = OxmlElement('w:abstractNum')
        title_abstract.set(qn('w:abstractNumId'), TITLE_NUMBERING_ABSTRACT_ID)

        nsid = OxmlElement('w:nsid')
        nsid.set(qn('w:val'), '6C6F6E67')
        title_abstract.append(nsid)

        multi_level_type = OxmlElement('w:multiLevelType')
        multi_level_type.set(qn('w:val'), 'multilevel')
        title_abstract.append(multi_level_type)

        level_styles = {
            0: '11',
            1: '22',
            2: '31',
            3: '41',
            4: '51',
            5: '6',
        }
        level_patterns = {
            0: '%1',
            1: '%1.%2',
            2: '%1.%2.%3',
            3: '%1.%2.%3.%4',
            4: '%1.%2.%3.%4.%5',
            5: '%1.%2.%3.%4.%5.%6',
        }

        for ilvl in range(6):
            lvl = OxmlElement('w:lvl')
            lvl.set(qn('w:ilvl'), str(ilvl))

            start = OxmlElement('w:start')
            start.set(qn('w:val'), '1')
            lvl.append(start)

            num_fmt = OxmlElement('w:numFmt')
            num_fmt.set(qn('w:val'), 'decimal')
            lvl.append(num_fmt)

            p_style = OxmlElement('w:pStyle')
            p_style.set(qn('w:val'), level_styles[ilvl])
            lvl.append(p_style)

            suff = OxmlElement('w:suff')
            suff.set(qn('w:val'), 'space')
            lvl.append(suff)

            lvl_text = OxmlElement('w:lvlText')
            lvl_text.set(qn('w:val'), level_patterns[ilvl])
            lvl.append(lvl_text)

            lvl_jc = OxmlElement('w:lvlJc')
            lvl_jc.set(qn('w:val'), 'left')
            lvl.append(lvl_jc)

            p_pr = OxmlElement('w:pPr')
            ind = OxmlElement('w:ind')
            ind.set(qn('w:left'), '0')
            ind.set(qn('w:firstLine'), '0')
            p_pr.append(ind)
            lvl.append(p_pr)

            title_abstract.append(lvl)

        ensure_numbering_child_order(numbering, title_abstract, 'abstract')

    title_num_existing = numbering.xpath(f'./w:num[@w:numId="{TITLE_NUMBERING_ID}"]')
    if not title_num_existing:
        title_num = OxmlElement('w:num')
        title_num.set(qn('w:numId'), TITLE_NUMBERING_ID)
        title_ref = OxmlElement('w:abstractNumId')
        title_ref.set(qn('w:val'), TITLE_NUMBERING_ABSTRACT_ID)
        title_num.append(title_ref)
        ensure_numbering_child_order(numbering, title_num, 'num')

    bullet_existing = numbering.xpath(f'./w:abstractNum[@w:abstractNumId="{BULLET_NUMBERING_ABSTRACT_ID}"]')
    if bullet_existing:
        return

    bullet_abstract = OxmlElement('w:abstractNum')
    bullet_abstract.set(qn('w:abstractNumId'), BULLET_NUMBERING_ABSTRACT_ID)
    bullet_chars = ['●', '■', '◆']
    for ilvl, char in enumerate(bullet_chars):
        lvl = OxmlElement('w:lvl')
        lvl.set(qn('w:ilvl'), str(ilvl))

        start = OxmlElement('w:start')
        start.set(qn('w:val'), '1')
        lvl.append(start)

        num_fmt = OxmlElement('w:numFmt')
        num_fmt.set(qn('w:val'), 'bullet')
        lvl.append(num_fmt)

        lvl_text = OxmlElement('w:lvlText')
        lvl_text.set(qn('w:val'), char)
        lvl.append(lvl_text)

        lvl_jc = OxmlElement('w:lvlJc')
        lvl_jc.set(qn('w:val'), 'left')
        lvl.append(lvl_jc)

        p_pr = OxmlElement('w:pPr')
        ind = OxmlElement('w:ind')
        ind.set(qn('w:left'), '0')
        ind.set(qn('w:hanging'), '0')
        p_pr.append(ind)
        lvl.append(p_pr)

        r_pr = OxmlElement('w:rPr')
        r_fonts = OxmlElement('w:rFonts')
        r_fonts.set(qn('w:ascii'), FONT_HEI)
        r_fonts.set(qn('w:hAnsi'), FONT_HEI)
        r_fonts.set(qn('w:eastAsia'), FONT_HEI)
        r_pr.append(r_fonts)
        sz = OxmlElement('w:sz')
        sz.set(qn('w:val'), str(int(SIZE_BODY.pt * 2)))
        r_pr.append(sz)
        sz_cs = OxmlElement('w:szCs')
        sz_cs.set(qn('w:val'), str(int(SIZE_BODY.pt * 2)))
        r_pr.append(sz_cs)
        lvl.append(r_pr)
        bullet_abstract.append(lvl)

    ensure_numbering_child_order(numbering, bullet_abstract, 'abstract')

    bullet_num = OxmlElement('w:num')
    bullet_num.set(qn('w:numId'), BULLET_NUMBERING_ID)
    bullet_ref = OxmlElement('w:abstractNumId')
    bullet_ref.set(qn('w:val'), BULLET_NUMBERING_ABSTRACT_ID)
    bullet_num.append(bullet_ref)
    ensure_numbering_child_order(numbering, bullet_num, 'num')


def apply_heading_numbering(paragraph, level: int):
    p_pr = paragraph._p.get_or_add_pPr()
    for child in list(p_pr):
        if child.tag == qn('w:numPr'):
            p_pr.remove(child)

    num_pr = OxmlElement('w:numPr')
    ilvl = OxmlElement('w:ilvl')
    ilvl.set(qn('w:val'), str(max(level - 1, 0)))
    num_id = OxmlElement('w:numId')
    num_id.set(qn('w:val'), TITLE_NUMBERING_ID)
    num_pr.append(ilvl)
    num_pr.append(num_id)
    insert_numpr_after_pstyle(p_pr, num_pr)


def strip_heading_number_prefix(text: str) -> str:
    return LEADING_HEADING_NUMBER_RE.sub('', text, count=1).strip()


def detect_heading_base_level(lines: list[str]) -> int:
    levels = []
    in_code = False
    for line in lines:
        stripped = line.strip()
        if stripped.startswith('```'):
            in_code = not in_code
            continue
        if in_code:
            continue
        heading = HEADING_RE.match(line)
        if heading:
            levels.append(len(heading.group(1)))
    return min(levels) if levels else 1


def is_cover_stop_line(line: str) -> bool:
    stripped = line.strip()
    if not stripped:
        return False
    if stripped.startswith('```'):
        return True
    if HEADING_RE.match(line) or is_table_line(line):
        return True
    return len(stripped) > 40


def extract_cover_region(lines: list[str]) -> tuple[dict[str, list[str] | str | None], list[str]]:
    cover = {'title': None, 'meta': []}
    remaining = list(lines)

    explicit_indices = set()
    for idx, line in enumerate(lines[:8]):
        match = COVER_LABEL_RE.match(line)
        if not match:
            continue
        label = (match.group(1) or match.group(2) or '').strip()
        value = match.group(3).strip()
        if not value:
            continue
        explicit_indices.add(idx)
        if label == '封面标题':
            cover['title'] = value
        else:
            cover['meta'].append(value)

    if explicit_indices:
        remaining = [line for idx, line in enumerate(lines) if idx not in explicit_indices]
        return cover, remaining

    probe: list[tuple[int, str]] = []
    for idx, line in enumerate(lines[:6]):
        stripped = line.strip()
        if not stripped:
            continue
        if is_cover_stop_line(line):
            break
        probe.append((idx, stripped))

    if not probe:
        return cover, remaining

    used_indices = set()
    cover['title'] = probe[0][1]
    used_indices.add(probe[0][0])
    for idx, value in probe[1:3]:
        if DATE_RE.search(value) or len(cover['meta']) < 2:
            cover['meta'].append(value)
            used_indices.add(idx)

    remaining = [line for idx, line in enumerate(lines) if idx not in used_indices]
    return cover, remaining


def render_cover(doc: Document, cover: dict[str, list[str] | str | None]):
    title = cover.get('title')
    meta = cover.get('meta') or []
    for _ in range(COVER_TOP_SPACER_COUNT):
        spacer = doc.add_paragraph()
        apply_paragraph_style(spacer, COVER_META_SPEC)
        spacer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if title:
        p = doc.add_paragraph()
        run = p.add_run(str(title))
        set_run_font(run, COVER_TITLE_SPEC.font_name, COVER_TITLE_SPEC.font_size)
        apply_paragraph_style(p, COVER_TITLE_SPEC)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for item in meta:
        p = doc.add_paragraph()
        run = p.add_run(str(item))
        set_run_font(run, COVER_META_SPEC.font_name, COVER_META_SPEC.font_size)
        apply_paragraph_style(p, COVER_META_SPEC)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    page_break = doc.add_paragraph()
    page_break.add_run().add_break(WD_BREAK.PAGE)


def ensure_style_numbering(style, level: int):
    p_pr = style._element.get_or_add_pPr()
    for child in list(p_pr):
        if child.tag == qn('w:numPr'):
            p_pr.remove(child)

    num_pr = OxmlElement('w:numPr')
    ilvl = OxmlElement('w:ilvl')
    ilvl.set(qn('w:val'), str(max(level - 1, 0)))
    num_id = OxmlElement('w:numId')
    num_id.set(qn('w:val'), TITLE_NUMBERING_ID)
    num_pr.append(ilvl)
    num_pr.append(num_id)
    insert_numpr_after_pstyle(p_pr, num_pr)


def configure_document(doc: Document):
    clear_document_body(doc)

    section = doc.sections[0]
    section.page_width = A4_WIDTH
    section.page_height = A4_HEIGHT
    section.top_margin = PAGE_MARGIN_TOP
    section.bottom_margin = PAGE_MARGIN_BOTTOM
    section.left_margin = PAGE_MARGIN_LEFT
    section.right_margin = PAGE_MARGIN_RIGHT

    if 'Normal' in doc.styles:
        normal = doc.styles['Normal']
        normal.font.name = BODY_SPEC.font_name
        normal.font.size = BODY_SPEC.font_size
        r_pr = normal._element.get_or_add_rPr()
        r_fonts = r_pr.rFonts
        if r_fonts is None:
            r_fonts = OxmlElement('w:rFonts')
            r_pr.append(r_fonts)
        r_fonts.set(qn('w:ascii'), BODY_SPEC.font_name)
        r_fonts.set(qn('w:hAnsi'), BODY_SPEC.font_name)
        r_fonts.set(qn('w:eastAsia'), BODY_SPEC.font_name)

    for level, spec in {
        1: H1_SPEC,
        2: H2_SPEC,
        3: H3_SPEC,
        4: H4_SPEC,
        5: H5_SPEC,
        6: H6_SPEC,
    }.items():
        style = get_or_create_paragraph_style(doc, get_heading_style_name(level), get_heading_style_id(level), spec)
        ensure_style_numbering(style, level)

    footer = section.footer
    paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    paragraph.clear()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    apply_line_spacing(paragraph, PAGE_SPEC.line_spacing)
    add_page_number(paragraph)
    ensure_numbering(doc)


def is_table_line(line: str) -> bool:
    s = line.strip()
    return s.startswith('|') and s.endswith('|') and s.count('|') >= 2


def is_separator_line(line: str) -> bool:
    s = line.strip().strip('|').strip()
    if not s:
        return False
    parts = [p.strip() for p in s.split('|')]
    return all(parts) and all(re.fullmatch(r':?-{3,}:?', p) for p in parts)


def split_table_row(line: str) -> list[str]:
    return [c.strip() for c in line.strip().strip('|').split('|')]


def set_table_width(table, col_count: int):
    table.autofit = False
    per_col_dxa = max(int(TABLE_CONTENT_WIDTH_DXA / max(col_count, 1)), 1)
    per_col_cm = TABLE_CONTENT_WIDTH_DXA / 567.0 / max(col_count, 1)
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.first_child_found_in('w:tblW')
    if tbl_w is None:
        tbl_w = OxmlElement('w:tblW')
        tbl_pr.append(tbl_w)
    tbl_w.set(qn('w:w'), str(TABLE_CONTENT_WIDTH_DXA))
    tbl_w.set(qn('w:type'), 'dxa')

    for row in table.rows:
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        for cell in row.cells[:col_count]:
            cell.width = Cm(per_col_cm)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell)
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.first_child_found_in('w:tcW')
            if tc_w is None:
                tc_w = OxmlElement('w:tcW')
                tc_pr.append(tc_w)
            tc_w.set(qn('w:w'), str(per_col_dxa))
            tc_w.set(qn('w:type'), 'dxa')


def set_cell_text(cell, text: str, bold: bool = False):
    cell.text = ''
    p = cell.paragraphs[0]
    run = p.add_run(text)
    run.bold = bold
    apply_paragraph_style(p, TABLE_SPEC)
    set_run_font(run, TABLE_SPEC.font_name, TABLE_SPEC.font_size)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if bold:
        run.bold = True


def shade_cell(cell, fill: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), fill)
    tc_pr.append(shd)


def render_markdown(doc: Document, text: str):
    raw_lines = text.splitlines()
    cover, lines = extract_cover_region(raw_lines)
    if cover.get('title') or cover.get('meta'):
        render_cover(doc, cover)
    heading_base_level = detect_heading_base_level(lines)
    in_code = False
    code_lines: list[str] = []
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if stripped.startswith('```'):
            if not in_code:
                in_code = True
                code_lines = []
            else:
                p = doc.add_paragraph()
                for idx, code_line in enumerate(code_lines):
                    run = p.add_run(code_line)
                    set_run_font(run, CODE_SPEC.font_name, CODE_SPEC.font_size)
                    if idx != len(code_lines) - 1:
                        run.add_break()
                apply_code_block_style(p)
                in_code = False
                code_lines = []
            i += 1
            continue

        if in_code:
            code_lines.append(line)
            i += 1
            continue

        if not stripped:
            i += 1
            continue

        if is_table_line(line):
            table_lines = [line]
            j = i + 1
            while j < len(lines) and is_table_line(lines[j]):
                table_lines.append(lines[j])
                j += 1

            rows = [split_table_row(tl) for tl in table_lines]
            has_header_sep = len(rows) >= 2 and is_separator_line(table_lines[1])
            if has_header_sep:
                header = rows[0]
                body = rows[2:]
            else:
                header = None
                body = rows

            col_count = max(len(r) for r in ([header] if header else []) + body)
            table_rows = (1 if header else 0) + max(len(body), 1)
            table = doc.add_table(rows=table_rows, cols=col_count)
            table.style = 'Table Grid'
            set_table_width(table, col_count)
            row_idx = 0
            if header:
                for col, val in enumerate(header):
                    cell = table.rows[0].cells[col]
                    set_cell_text(cell, val, bold=True)
                    shade_cell(cell, TABLE_HEADER_SHADE)
                row_idx = 1
            body_rows = body if body else [[]]
            for row in body_rows:
                for col in range(col_count):
                    val = row[col] if col < len(row) else ''
                    set_cell_text(table.rows[row_idx].cells[col], val, bold=False)
                row_idx += 1
            i = j
            continue

        heading = HEADING_RE.match(line)
        if heading:
            source_level = len(heading.group(1))
            level = max(source_level - heading_base_level + 1, 1)
            level = min(level, 6)
            content = strip_heading_number_prefix(heading.group(2).strip())
            p = doc.add_paragraph(style=get_heading_style_id(level))
            run = p.add_run(content)
            spec = get_heading_spec(level)
            set_run_font(run, spec.font_name, spec.font_size)
            apply_paragraph_style(p, spec)
            apply_heading_numbering(p, level)
            i += 1
            continue

        bullet = BULLET_RE.match(line)
        if bullet:
            indent = bullet.group(1)
            bullet_level = min(len(indent.replace('\t', '    ')) // 2, 2)
            p = doc.add_paragraph()
            run = p.add_run(bullet.group(2).strip())
            set_run_font(run, BODY_SPEC.font_name, BODY_SPEC.font_size)
            apply_bullet_style(p, bullet_level)
            i += 1
            continue

        ordered = ORDERED_RE.match(line)
        if ordered:
            p = doc.add_paragraph(style='List Number')
            run = p.add_run(ordered.group(1).strip())
            set_run_font(run, BODY_SPEC.font_name, BODY_SPEC.font_size)
            apply_paragraph_style(p, BODY_SPEC)
            i += 1
            continue

        p = doc.add_paragraph()
        run = p.add_run(line)
        set_run_font(run, BODY_SPEC.font_name, BODY_SPEC.font_size)
        apply_paragraph_style(p, BODY_SPEC)
        i += 1

    if in_code and code_lines:
        p = doc.add_paragraph()
        for idx, code_line in enumerate(code_lines):
            run = p.add_run(code_line)
            set_run_font(run, CODE_SPEC.font_name, CODE_SPEC.font_size)
            if idx != len(code_lines) - 1:
                run.add_break()
        apply_code_block_style(p)


def render_txt(doc: Document, text: str):
    raw_lines = text.splitlines()
    cover, body_lines = extract_cover_region(raw_lines)
    if cover.get('title') or cover.get('meta'):
        render_cover(doc, cover)
    body_text = '\n'.join(body_lines)
    blocks = re.split(r'\n\s*\n+', body_text.strip()) if body_text.strip() else []
    for block in blocks:
        p = doc.add_paragraph()
        lines = block.splitlines()
        for idx, line in enumerate(lines):
            run = p.add_run(line)
            set_run_font(run, BODY_SPEC.font_name, BODY_SPEC.font_size)
            if idx != len(lines) - 1:
                run.add_break()
        apply_paragraph_style(p, BODY_SPEC)


def convert(input_path: Path, output_path: Path):
    suffix = input_path.suffix.lower()
    if suffix not in {'.md', '.markdown', '.txt'}:
        raise ValueError(f'Unsupported input type: {suffix}')

    text = input_path.read_text(encoding='utf-8')
    doc = Document(TEMPLATE_DOCX_PATH)
    configure_document(doc)

    if suffix in {'.md', '.markdown'}:
        render_markdown(doc, text)
    else:
        render_txt(doc, text)

    doc.save(output_path)


def parse_args():
    parser = argparse.ArgumentParser(description='Convert Markdown or TXT to formatted DOCX.')
    parser.add_argument('input', help='Input .md, .markdown, or .txt file')
    parser.add_argument('output', nargs='?', help='Optional output .docx path')
    return parser.parse_args()


def main():
    args = parse_args()
    input_path = Path(args.input).expanduser().resolve()
    if not input_path.exists():
        raise SystemExit(f'Input file not found: {input_path}')

    if args.output:
        output_path = Path(args.output).expanduser().resolve()
    else:
        output_path = input_path.with_suffix('.docx')

    convert(input_path, output_path)
    print(output_path)


if __name__ == '__main__':
    main()
