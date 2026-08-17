import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import text_to_docx as ttd
from text_to_docx import (
    Document,
    Pt,
    STYLE_ENTRY_TO_GLOBAL_NAME,
    build_default_style_config,
    deep_merge_style_config,
    load_style_config_file,
    qn,
    resolve_style_specs,
    validate_style_config,
)

SCRIPT_PATH = Path(__file__).resolve().parent / 'text_to_docx.py'
SHIPPED_CONFIG_PATH = Path(__file__).resolve().parent.parent / 'config' / 'style.json'


def _spec_fields(spec):
    return (
        spec.font_name,
        spec.font_size.pt,
        spec.line_spacing,
        spec.before_pt,
        spec.after_pt,
        spec.first_line_chars,
        spec.bold,
        spec.alignment,
        spec.font_color,
    )


class DefaultStyleConfigMatchesHardcodedSpecsTests(unittest.TestCase):
    """The config machinery must reproduce today's hardcoded StyleSpec values
    exactly, both from its own built-in default and from the shipped JSON
    file, so a missing/absent config file is truly a behavior no-op."""

    def test_resolved_builtin_defaults_match_hardcoded_specs(self) -> None:
        specs = resolve_style_specs(build_default_style_config())
        for entry_id, global_name in STYLE_ENTRY_TO_GLOBAL_NAME.items():
            hardcoded = getattr(ttd, global_name)
            resolved = specs[global_name]
            self.assertEqual(_spec_fields(hardcoded), _spec_fields(resolved), global_name)

    def test_shipped_style_json_matches_hardcoded_defaults(self) -> None:
        raw = load_style_config_file(SHIPPED_CONFIG_PATH)
        validate_style_config(raw)
        merged = deep_merge_style_config(build_default_style_config(), raw)
        specs = resolve_style_specs(merged)
        for entry_id, global_name in STYLE_ENTRY_TO_GLOBAL_NAME.items():
            hardcoded = getattr(ttd, global_name)
            resolved = specs[global_name]
            self.assertEqual(_spec_fields(hardcoded), _spec_fields(resolved), global_name)


class StyleConfigValidationTests(unittest.TestCase):
    """validate_style_config() must fail loud with SystemExit on malformed
    input rather than silently ignoring bad values."""

    def test_non_dict_root_raises(self) -> None:
        with self.assertRaises(SystemExit):
            validate_style_config('not a dict')  # type: ignore[arg-type]

    def test_unknown_top_level_key_raises(self) -> None:
        with self.assertRaises(SystemExit):
            validate_style_config({'styles': {}, 'bogus': 1})

    def test_unknown_style_id_raises(self) -> None:
        with self.assertRaises(SystemExit):
            validate_style_config({'styles': {'h99': {}}})

    def test_unknown_field_in_entry_raises(self) -> None:
        with self.assertRaises(SystemExit):
            validate_style_config({'styles': {'h1': {'weight': 'bold'}}})

    def test_bad_size_pt_type_raises(self) -> None:
        with self.assertRaises(SystemExit):
            validate_style_config({'styles': {'h1': {'size_pt': 'big'}}})

    def test_non_positive_size_pt_raises(self) -> None:
        with self.assertRaises(SystemExit):
            validate_style_config({'styles': {'h1': {'size_pt': 0}}})

    def test_bad_color_format_raises(self) -> None:
        with self.assertRaises(SystemExit):
            validate_style_config({'styles': {'h1': {'color': 'blue'}}})

    def test_bad_line_spacing_raises(self) -> None:
        with self.assertRaises(SystemExit):
            validate_style_config({'styles': {'h1': {'line_spacing': 'double'}}})

    def test_bad_alignment_raises(self) -> None:
        with self.assertRaises(SystemExit):
            validate_style_config({'styles': {'h1': {'alignment': 'top'}}})

    def test_unknown_font_token_reference_raises_on_resolve(self) -> None:
        merged = deep_merge_style_config(
            build_default_style_config(),
            {'styles': {'h1': {'font': 'sans'}}},
        )
        with self.assertRaises(SystemExit):
            resolve_style_specs(merged)


class StyleConfigCliEndToEndTests(unittest.TestCase):
    """Exercise the real CLI via subprocess so the process-global reassignment
    in apply_style_config_to_globals() is verified without mutating this test
    process's own text_to_docx module state (which would risk leaking into
    test_verify_body_paragraph_style.py / test_markdown_features.py depending
    on collection order)."""

    def _run_cli(self, args, cwd) -> subprocess.CompletedProcess:
        return subprocess.run(
            [sys.executable, str(SCRIPT_PATH), *args],
            cwd=cwd,
            capture_output=True,
            text=True,
        )

    def test_override_changes_only_targeted_style(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            md_path = tmp_path / 'sample.md'
            md_path.write_text('# Heading One\n\nBody paragraph text.\n', encoding='utf-8')

            style_override = tmp_path / 'style.json'
            style_override.write_text(
                json.dumps({'styles': {'h1': {'color': '1F4E79', 'size_pt': 20}}}),
                encoding='utf-8',
            )

            default_out = tmp_path / 'default.docx'
            custom_out = tmp_path / 'custom.docx'

            default_result = self._run_cli([str(md_path), str(default_out)], tmp_path)
            self.assertEqual(0, default_result.returncode, default_result.stderr)

            custom_result = self._run_cli(
                [str(md_path), str(custom_out), '--style-config', str(style_override)], tmp_path
            )
            self.assertEqual(0, custom_result.returncode, custom_result.stderr)

            default_doc = Document(str(default_out))
            custom_doc = Document(str(custom_out))

            default_h1_color = default_doc.styles['Heading 1']._element.get_or_add_rPr().find(qn('w:color'))
            custom_h1_color = custom_doc.styles['Heading 1']._element.get_or_add_rPr().find(qn('w:color'))
            self.assertEqual('000000', default_h1_color.get(qn('w:val')))
            self.assertEqual('1F4E79', custom_h1_color.get(qn('w:val')))

            self.assertEqual(Pt(16), default_doc.styles['Heading 1'].font.size)
            self.assertEqual(Pt(20), custom_doc.styles['Heading 1'].font.size)

            # the h1-only override must not bleed into an unrelated style
            self.assertEqual(default_doc.styles['Normal'].font.size, custom_doc.styles['Normal'].font.size)

    def test_missing_explicit_style_config_path_errors(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            md_path = tmp_path / 'sample.md'
            md_path.write_text('# Heading\n\nBody.\n', encoding='utf-8')

            result = self._run_cli(
                [str(md_path), str(tmp_path / 'out.docx'), '--style-config', str(tmp_path / 'missing.json')],
                tmp_path,
            )

            self.assertNotEqual(0, result.returncode)
            self.assertIn('--style-config path not found', result.stderr)

    def test_malformed_style_config_errors(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            md_path = tmp_path / 'sample.md'
            md_path.write_text('# Heading\n\nBody.\n', encoding='utf-8')

            bad_config = tmp_path / 'style.json'
            bad_config.write_text('{not valid json', encoding='utf-8')

            result = self._run_cli(
                [str(md_path), str(tmp_path / 'out.docx'), '--style-config', str(bad_config)],
                tmp_path,
            )

            self.assertNotEqual(0, result.returncode)
            self.assertIn('Invalid style config', result.stderr)


if __name__ == '__main__':
    unittest.main()
