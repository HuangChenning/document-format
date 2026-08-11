import re
import unittest
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))

from text_to_docx import (
    FONT_MONO,
    AnalysisResult,
    Document,
    MAX_BULLET_LEVEL,
    MERMAID_LANGUAGES,
    apply_bullet_style,
    configure_document,
    create_ordered_numbering_instance,
    ensure_numbering,
    find_drawio_cli,
    fit_code_font_size,
    get_numbering_element,
    is_loose_table_separator,
    is_separator_line,
    parse_fence_language,
    render_blockquote,
    render_markdown,
    render_mermaid_to_png,
    render_mermaid_with_drawio,
    split_table_row,
    strip_inline_emphasis,
    strip_emoji,
)


def build_doc():
    doc = Document()
    configure_document(doc)
    ensure_numbering(doc)
    return doc


def render_body(doc, lines):
    analysis = AnalysisResult(cover=None, body_lines=lines, toc_exists=False, toc_entries=[])
    render_markdown(doc, analysis, reserve_cover=False, auto_toc=False, force_cover=False, force_toc=False)


def paragraph_texts(doc):
    return [p.text for p in doc.paragraphs]


def all_runs(doc):
    runs = []
    for p in doc.paragraphs:
        runs.extend(p.runs)
    return runs


class InlineFormattingTests(unittest.TestCase):
    def test_bold_italic_strike_code_runs(self):
        doc = build_doc()
        render_body(doc, ['包含 **加粗** 与 *斜体* 和 ~~删除线~~ 以及 `代码` 的段落'])

        runs = all_runs(doc)
        self.assertTrue(any(r.bold and r.text == '加粗' for r in runs))
        self.assertTrue(any(r.italic and r.text == '斜体' for r in runs))
        self.assertTrue(any(r.font.strike and r.text == '删除线' for r in runs))
        self.assertTrue(any(r.font.name == FONT_MONO and r.text == '代码' for r in runs))
        full = ''.join(r.text for r in runs)
        for token in ('**', '~~', '`'):
            self.assertNotIn(token, full)
        self.assertNotIn('*斜体*', full)

    def test_nested_bold_italic(self):
        doc = build_doc()
        render_body(doc, ['**外层 *内层* 外层**'])

        runs = all_runs(doc)
        nested = [r for r in runs if r.text == '内层']
        self.assertEqual(1, len(nested))
        self.assertTrue(nested[0].bold)
        self.assertTrue(nested[0].italic)

    def test_underscore_italic_ignores_snake_case(self):
        doc = build_doc()
        render_body(doc, ['变量 snake_case_name 不应斜体 但 _这里_ 应该斜体'])

        runs = all_runs(doc)
        self.assertFalse(any(r.italic and 'snake' in r.text for r in runs))
        self.assertTrue(any(r.italic and r.text == '这里' for r in runs))

    def test_html_entities_decoded(self):
        doc = build_doc()
        render_body(doc, ['实体 &amp; &lt;标签&gt; 解码'])

        full = ''.join(r.text for r in all_runs(doc))
        self.assertIn('实体 & <标签> 解码', full)

    def test_strip_inline_emphasis(self):
        self.assertEqual('纯文本', strip_inline_emphasis('**纯文本**'))
        self.assertEqual('混合', strip_inline_emphasis('*混合*'))


class HyperlinkTests(unittest.TestCase):
    def test_markdown_link_creates_hyperlink(self):
        doc = build_doc()
        render_body(doc, ['访问 [示例](https://example.com) 站点'])

        xml = doc.paragraphs[0]._p.xml
        self.assertIn('<w:hyperlink', xml)
        texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', xml)
        self.assertIn('示例', texts)

    def test_bare_url_creates_hyperlink(self):
        doc = build_doc()
        render_body(doc, ['裸链接 https://www.example.org/path 结束'])

        xml = doc.paragraphs[0]._p.xml
        self.assertIn('<w:hyperlink', xml)
        self.assertNotIn('https://www.example.org/path', ''.join(r.text for r in doc.paragraphs[0].runs))

    def test_bare_url_trailing_punctuation_excluded(self):
        doc = build_doc()
        render_body(doc, ['请看 https://example.com/a. 然后继续'])

        rels = doc.part.rels
        targets = [rel.target_ref for rel in rels.values() if 'hyperlink' in rel.reltype]
        self.assertIn('https://example.com/a', targets)


class TableParsingTests(unittest.TestCase):
    def test_separator_line_relaxed(self):
        self.assertTrue(is_separator_line('| --- | --- |'))
        self.assertTrue(is_separator_line('| - | - |'))
        self.assertTrue(is_separator_line('|:---|---:|'))

    def test_loose_table_separator_without_outer_pipes(self):
        self.assertTrue(is_loose_table_separator('--- | ---'))
        self.assertTrue(is_loose_table_separator('---|---|---'))
        self.assertFalse(is_loose_table_separator('普通文本行'))

    def test_split_table_row_handles_escaped_pipe(self):
        cells = split_table_row(r'| a\|b | c |')
        self.assertEqual(['a|b', 'c'], cells)

    def test_table_without_outer_pipes_renders(self):
        doc = build_doc()
        render_body(doc, ['列A | 列B', '--- | ---', '值1 | 值2'])

        self.assertEqual(1, len(doc.tables))
        self.assertEqual(2, len(doc.tables[0].rows))
        self.assertEqual('值1', doc.tables[0].rows[1].cells[0].text.strip())


class BlockquoteTests(unittest.TestCase):
    def test_consecutive_blockquote_lines_merge_with_border(self):
        doc = build_doc()
        render_body(doc, ['> 第一行', '> 第二行'])

        quote_paras = [p for p in doc.paragraphs if '第一行' in p.text]
        self.assertEqual(1, len(quote_paras))
        self.assertIn('第二行', quote_paras[0].text)
        self.assertIn('<w:left', quote_paras[0]._p.xml)

    def test_render_blockquote_applies_left_border(self):
        doc = build_doc()
        render_blockquote(doc, ['引用内容'])

        para = doc.paragraphs[-1]
        self.assertIn('引用内容', para.text)
        self.assertIn('<w:left', para._p.xml)


class ListTests(unittest.TestCase):
    def test_ordered_lists_get_distinct_numbering_instances(self):
        doc = build_doc()
        render_body(doc, ['1. 甲', '2. 乙', '中断段落', '1. 丙', '2. 丁'])

        num_ids = []
        for p in doc.paragraphs:
            num_ids.extend(int(m) for m in re.findall(r'<w:numId w:val="(\d+)"', p._p.xml))
        ordered_ids = [n for n in num_ids if n >= 9000]
        self.assertEqual(2, len(set(ordered_ids)))

    def test_ordered_list_has_hanging_indent(self):
        doc = build_doc()
        render_body(doc, ['1. 第一项'])

        xml = doc.paragraphs[0]._p.xml
        self.assertIn('w:hanging="480"', xml.replace("'", '"'))

    def test_nested_bullets_support_nine_levels(self):
        doc = build_doc()
        paragraph = doc.add_paragraph()
        apply_bullet_style(paragraph, MAX_BULLET_LEVEL - 1)
        apply_bullet_style(paragraph, MAX_BULLET_LEVEL)

    def test_task_list_marks(self):
        doc = build_doc()
        render_body(doc, ['- [x] 已完成', '- [ ] 未完成'])

        texts = paragraph_texts(doc)
        self.assertTrue(any('☑' in t and '已完成' in t for t in texts))
        self.assertTrue(any('☐' in t and '未完成' in t for t in texts))
        self.assertFalse(any('- [' in t for t in texts))


class FootnoteTests(unittest.TestCase):
    def test_footnote_reference_and_section(self):
        doc = build_doc()
        render_body(doc, ['正文引用[^1]结束', '', '[^1]: 脚注定义内容'])

        texts = paragraph_texts(doc)
        joined = '\n'.join(texts)
        self.assertNotIn('[^1]', joined)
        self.assertIn('[1]', joined)
        self.assertIn('脚注定义内容', joined)

    def test_unknown_footnote_reference_kept_verbatim(self):
        doc = build_doc()
        render_body(doc, ['正文引用[^missing]结束'])

        self.assertIn('[^missing]', ''.join(paragraph_texts(doc)))


class MermaidTests(unittest.TestCase):
    def test_parse_fence_language(self):
        self.assertEqual('mermaid', parse_fence_language('```mermaid'))
        self.assertEqual('python', parse_fence_language('```python'))
        self.assertIsNone(parse_fence_language('```'))
        self.assertEqual('mermaid', parse_fence_language('~~~Mermaid'))

    def test_mermaid_language_set(self):
        self.assertIn('mermaid', MERMAID_LANGUAGES)

    def test_non_mermaid_fence_stays_code_block(self):
        doc = build_doc()
        render_body(doc, ['```python', "print('hi')", '```'])

        self.assertIn("print('hi')", ''.join(paragraph_texts(doc)))
        self.assertEqual(0, len(doc.inline_shapes))

    def test_mermaid_render_failure_falls_back_to_code_block(self):
        import text_to_docx as module

        original_mmdc = module.render_mermaid_with_mmdc
        original_drawio = module.render_mermaid_with_drawio
        original_ink = module.render_mermaid_with_ink
        module.render_mermaid_with_mmdc = lambda source, target: False
        module.render_mermaid_with_drawio = lambda source, target: False
        module.render_mermaid_with_ink = lambda code, target: False
        try:
            doc = build_doc()
            render_body(doc, ['```mermaid', 'graph LR', 'A-->B', '```'])
        finally:
            module.render_mermaid_with_mmdc = original_mmdc
            module.render_mermaid_with_drawio = original_drawio
            module.render_mermaid_with_ink = original_ink

        joined = ''.join(paragraph_texts(doc))
        self.assertIn('graph LR', joined)
        self.assertEqual(0, len(doc.inline_shapes))


class DrawioFallbackTests(unittest.TestCase):
    def test_drawio_renderer_skipped_when_cli_missing(self):
        import text_to_docx as module

        original = module.find_drawio_cli
        module.find_drawio_cli = lambda: None
        try:
            import tempfile
            from pathlib import Path
            with tempfile.TemporaryDirectory() as tmp:
                source = Path(tmp) / 'demo.mmd'
                source.write_text('graph LR\nA-->B', encoding='utf-8')
                result = render_mermaid_with_drawio(source, Path(tmp) / 'demo.png')
        finally:
            module.find_drawio_cli = original

        self.assertFalse(result)

    def test_find_drawio_cli_returns_none_when_nothing_installed(self):
        import text_to_docx as module

        original_which = module.shutil.which
        original_macos = module.DRAWIO_MACOS_CLI
        module.shutil.which = lambda name: None
        module.DRAWIO_MACOS_CLI = module.Path('/nonexistent/draw.io')
        try:
            if module.sys.platform != 'win32':
                self.assertIsNone(find_drawio_cli())
        finally:
            module.shutil.which = original_which
            module.DRAWIO_MACOS_CLI = original_macos

    def test_render_chain_uses_drawio_when_mmdc_fails(self):
        import text_to_docx as module

        calls = []

        def fake_mmdc(source, target):
            calls.append('mmdc')
            return False

        def fake_drawio(source, target):
            calls.append('drawio')
            target.write_bytes(b'\x89PNG fake')
            return True

        def fake_ink(code, target):
            calls.append('ink')
            return False

        original = (module.render_mermaid_with_mmdc, module.render_mermaid_with_drawio, module.render_mermaid_with_ink, module.get_mermaid_cache_dir)
        import tempfile as _tempfile
        isolated_cache = module.Path(_tempfile.mkdtemp())
        module.render_mermaid_with_mmdc = fake_mmdc
        module.render_mermaid_with_drawio = fake_drawio
        module.render_mermaid_with_ink = fake_ink
        module.get_mermaid_cache_dir = lambda: isolated_cache
        try:
            result = render_mermaid_to_png('graph LR\nDrawioChainTest-->B')
        finally:
            module.render_mermaid_with_mmdc, module.render_mermaid_with_drawio, module.render_mermaid_with_ink, module.get_mermaid_cache_dir = original

        self.assertEqual(['mmdc', 'drawio'], calls)
        self.assertIsNotNone(result)

    def test_render_chain_reaches_ink_when_drawio_missing(self):
        import text_to_docx as module

        calls = []

        def fake_mmdc(source, target):
            calls.append('mmdc')
            return False

        def fake_drawio(source, target):
            calls.append('drawio')
            return False

        def fake_ink(code, target):
            calls.append('ink')
            target.write_bytes(b'\xff\xd8 fake')
            return True

        original = (module.render_mermaid_with_mmdc, module.render_mermaid_with_drawio, module.render_mermaid_with_ink, module.get_mermaid_cache_dir)
        import tempfile as _tempfile
        isolated_cache = module.Path(_tempfile.mkdtemp())
        module.render_mermaid_with_mmdc = fake_mmdc
        module.render_mermaid_with_drawio = fake_drawio
        module.render_mermaid_with_ink = fake_ink
        module.get_mermaid_cache_dir = lambda: isolated_cache
        try:
            result = render_mermaid_to_png('graph LR\nInkChainTest-->B')
        finally:
            module.render_mermaid_with_mmdc, module.render_mermaid_with_drawio, module.render_mermaid_with_ink, module.get_mermaid_cache_dir = original

        self.assertEqual(['mmdc', 'drawio', 'ink'], calls)
        self.assertIsNotNone(result)


class EmojiStripTests(unittest.TestCase):
    def test_heading_emoji_removed(self):
        doc = build_doc()
        render_body(doc, ['# 📌 TL;DR（执行摘要）'])
        self.assertEqual('TL;DR（执行摘要）', doc.paragraphs[0].text)

    def test_body_emoji_removed_and_spaces_collapsed(self):
        doc = build_doc()
        render_body(doc, ['状态 ✅ 全版本 ⚠️ 注意 ❌ 无 结束'])
        full = ''.join(r.text for r in all_runs(doc))
        for ch in '✅⚠❌\uFE0F':
            self.assertNotIn(ch, full)
        self.assertIn('状态 全版本 注意 无 结束', full)
        self.assertNotIn('  ', full)

    def test_code_block_preserves_emoji(self):
        doc = build_doc()
        render_body(doc, ['```', '✅ keep emoji', '```'])
        full = ''.join(r.text for r in all_runs(doc))
        self.assertIn('✅ keep emoji', full)

    def test_inline_code_preserves_emoji(self):
        doc = build_doc()
        render_body(doc, ['行内 `✅ code` 结束'])
        runs = all_runs(doc)
        self.assertTrue(any(r.font.name == FONT_MONO and '✅' in r.text for r in runs))

    def test_task_list_mark_preserved(self):
        doc = build_doc()
        render_body(doc, ['- [x] 已完成', '- [ ] 待办'])
        texts = paragraph_texts(doc)
        self.assertTrue(any(t.startswith('☑ 已完成') for t in texts))
        self.assertTrue(any(t.startswith('☐ 待办') for t in texts))

    def test_table_cell_emoji_removed(self):
        doc = build_doc()
        render_body(doc, ['| 引擎 | 支持情况 |', '| --- | --- |', '| Oracle | ✅ 全版本 |'])
        cell_texts = [c.text for t in doc.tables for r in t.rows for c in r.cells]
        self.assertIn('全版本', cell_texts)
        self.assertFalse(any('✅' in t for t in cell_texts))

    def test_glued_emoji_keeps_separator_space(self):
        self.assertEqual('风险 提示', strip_emoji('风险📌 提示'))
        self.assertEqual('全版本', strip_emoji('✅ 全版本'))
        self.assertEqual('状态 全版本', strip_emoji('状态 ✅ 全版本'))
        self.assertEqual('状态', strip_emoji('状态 ✅'))
        self.assertEqual('风险提示', strip_emoji('风险📌提示'))


class CodeBlockFitTests(unittest.TestCase):
    def test_narrow_code_keeps_base_size(self):
        self.assertEqual(9.0, fit_code_font_size(['print("hi")']))
        self.assertEqual(9.0, fit_code_font_size([]))

    def test_wide_code_shrinks_within_bounds(self):
        wide = ['步骤1          步骤2          步骤3          步骤4          步骤5          步骤6']
        size = fit_code_font_size(wide)
        self.assertLess(size, 9.0)
        self.assertGreaterEqual(size, 6.0)


class OrderedNumberingInstanceTests(unittest.TestCase):
    def test_create_ordered_numbering_instance_increments(self):
        doc = build_doc()
        numbering = get_numbering_element(doc)
        first = int(create_ordered_numbering_instance(numbering))
        second = int(create_ordered_numbering_instance(numbering))
        self.assertGreaterEqual(first, 9000)
        self.assertEqual(first + 1, second)


if __name__ == '__main__':
    unittest.main(verbosity=2)
