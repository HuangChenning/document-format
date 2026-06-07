import unittest
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))

from text_to_docx import (
    BLACK_FONT_COLOR,
    BODY_SPEC,
    DEBUG_XML_COMMANDS,
    Document,
    OxmlElement,
    TITLE_NUMBERING_ID,
    VerificationContext,
    apply_bullet_style,
    apply_paragraph_style,
    build_xml_debug_lane_guidance,
    chinese_counting_text,
    configure_document,
    find_toc_region,
    get_numbering_level_node,
    qn,
    remove_existing_cover,
    remove_existing_toc,
    render_cover,
    render_toc,
    start_body_section,
    verify_body_paragraph_style,
    verify_cover_presence,
    verify_toc_presence,
)


class VerifyBodyParagraphStyleTests(unittest.TestCase):
    def build_context(self) -> VerificationContext:
        return VerificationContext(
            expected_cover=None,
            expect_cover=False,
            expect_toc=False,
            expected_headings=[],
            source_kind='test',
        )

    def test_bullet_numbering_is_not_reported_as_body_residue(self) -> None:
        doc = Document()
        configure_document(doc)
        paragraph = doc.add_paragraph()
        paragraph.add_run('需关注：第一项')
        apply_bullet_style(paragraph, 0)

        failures = verify_body_paragraph_style(doc, self.build_context())

        self.assertEqual([], failures)

    def test_non_bullet_numbering_still_reports_body_residue(self) -> None:
        doc = Document()
        configure_document(doc)
        paragraph = doc.add_paragraph()
        paragraph.add_run('正文段落')
        apply_paragraph_style(paragraph, BODY_SPEC)
        p_pr = paragraph._p.get_or_add_pPr()
        num_pr = OxmlElement('w:numPr')
        num_id = OxmlElement('w:numId')
        num_id.set(qn('w:val'), '999')
        num_pr.append(num_id)
        p_pr.append(num_pr)

        failures = verify_body_paragraph_style(doc, self.build_context())

        self.assertIn('body numbering residue: 正文段落', failures)

    def test_heading_styles_and_numbering_use_black_font(self) -> None:
        doc = Document()
        configure_document(doc)

        heading_style = doc.styles['Heading 1']
        style_color = heading_style._element.get_or_add_rPr().find(qn('w:color'))
        self.assertIsNotNone(style_color)
        self.assertEqual(BLACK_FONT_COLOR, style_color.get(qn('w:val')))

        numbering_level = get_numbering_level_node(doc, TITLE_NUMBERING_ID, 0)
        self.assertIsNotNone(numbering_level)
        numbering_rpr = numbering_level.find(qn('w:rPr'))
        self.assertIsNotNone(numbering_rpr)
        numbering_color = numbering_rpr.find(qn('w:color'))
        self.assertIsNotNone(numbering_color)
        self.assertEqual(BLACK_FONT_COLOR, numbering_color.get(qn('w:val')))

    def test_blank_document_uses_builtin_heading_style_id(self) -> None:
        doc = Document()
        configure_document(doc)

        numbering_level = get_numbering_level_node(doc, TITLE_NUMBERING_ID, 0)
        self.assertIsNotNone(numbering_level)
        p_style = numbering_level.find(qn('w:pStyle'))
        self.assertIsNotNone(p_style)
        self.assertEqual(doc.styles['Heading 1'].style_id, p_style.get(qn('w:val')))

    def test_chinese_counting_text_keeps_one_in_compound_tens(self) -> None:
        self.assertEqual('一百一十', chinese_counting_text(110))
        self.assertEqual('一百一十五', chinese_counting_text(115))
        self.assertEqual('一千零一十', chinese_counting_text(1010))

    def test_xml_debug_lane_guidance_is_added_for_numbering_failures(self) -> None:
        context = VerificationContext(
            expected_cover=None,
            expect_cover=False,
            expect_toc=False,
            expected_headings=[],
            source_kind='refreshed',
        )

        guidance = build_xml_debug_lane_guidance(['numbering: Heading 1 numbering pattern mismatch'], context)

        self.assertIn('DOCX XML structure or field-binding problem', guidance)
        self.assertIn(str(Path(__file__).resolve().parent / 'debug_docx_xml.py'), DEBUG_XML_COMMANDS)
        self.assertIn(DEBUG_XML_COMMANDS, guidance)

    def test_xml_debug_lane_guidance_is_empty_for_simple_body_failures(self) -> None:
        context = self.build_context()

        guidance = build_xml_debug_lane_guidance(['body: body line spacing mismatch'], context)

        self.assertEqual('', guidance)

    def test_verify_cover_presence_reports_retained_cover_when_not_expected(self) -> None:
        doc = Document()
        configure_document(doc)
        render_cover(doc, {'title': '测试标题', 'title_lines': ['测试标题'], 'meta': ['2026-06-07'], 'corner_meta': []})
        start_body_section(doc)
        paragraph = doc.add_paragraph()
        paragraph.add_run('正文首段')
        apply_paragraph_style(paragraph, BODY_SPEC)

        failures = verify_cover_presence(doc, self.build_context())

        self.assertIn('cover not expected but cover-like paragraphs were retained', failures)

    def test_verify_toc_presence_reports_retained_toc_when_not_expected(self) -> None:
        doc = Document()
        configure_document(doc)
        render_toc(doc, [(1, '第一章')])
        paragraph = doc.add_paragraph()
        paragraph.add_run('正文首段')
        apply_paragraph_style(paragraph, BODY_SPEC)

        failures = verify_toc_presence(doc, self.build_context())

        self.assertIn('TOC not expected but heading was retained', failures)
        self.assertIn('TOC not expected but field was retained', failures)

    def test_remove_existing_cover_keeps_body_content(self) -> None:
        doc = Document()
        configure_document(doc)
        render_cover(doc, {'title': '测试标题', 'title_lines': ['测试标题'], 'meta': ['2026-06-07'], 'corner_meta': []})
        start_body_section(doc)
        paragraph = doc.add_paragraph()
        paragraph.add_run('正文首段')
        apply_paragraph_style(paragraph, BODY_SPEC)

        remove_existing_cover(doc)

        self.assertEqual('正文首段', doc.paragraphs[0].text)

    def test_remove_existing_toc_clears_toc_region(self) -> None:
        doc = Document()
        configure_document(doc)
        render_toc(doc, [(1, '第一章')])
        paragraph = doc.add_paragraph()
        paragraph.add_run('正文首段')
        apply_paragraph_style(paragraph, BODY_SPEC)

        remove_existing_toc(doc)

        toc_heading, toc_field = find_toc_region(doc)
        self.assertIsNone(toc_heading)
        self.assertIsNone(toc_field)
        self.assertEqual('正文首段', doc.paragraphs[0].text)


if __name__ == '__main__':
    unittest.main()
