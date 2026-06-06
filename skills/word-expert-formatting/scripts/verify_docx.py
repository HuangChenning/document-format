#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from text_to_docx import (
    Document,
    build_verification_context,
    check_runtime_dependencies,
    find_toc_region,
    infer_cover,
    infer_expected_headings,
    verify_document_workflow,
    NUMBERING_MODE_A,
    NUMBERING_MODE_B,
)


def verify_existing_docx(
    input_path: Path,
    numbering_mode: str,
    expect_cover: bool | None,
    expect_toc: bool | None,
) -> None:
    doc = Document(str(input_path))
    expected_headings = infer_expected_headings(doc)
    inferred_cover = infer_cover(doc)
    toc_heading, toc_field = find_toc_region(doc)
    context = build_verification_context(
        expected_cover=inferred_cover,
        expect_cover=bool(inferred_cover) if expect_cover is None else expect_cover,
        expect_toc=bool(toc_heading or toc_field) if expect_toc is None else expect_toc,
        expected_headings=expected_headings,
        numbering_mode=numbering_mode,
        source_kind='refreshed',
        toc_refresh_state='not_refreshed',
    )
    verify_document_workflow(doc, context)


def parse_args():
    parser = argparse.ArgumentParser(description='Verify an existing DOCX against the word formatting workflow rules.')
    parser.add_argument('input', help='Input .docx file')
    parser.add_argument('--numbering-mode', choices=[NUMBERING_MODE_A, NUMBERING_MODE_B], default=NUMBERING_MODE_A)
    parser.add_argument('--expect-cover', dest='expect_cover', action='store_const', const=True, default=None, help='Require a cover section')
    parser.add_argument('--no-expect-cover', dest='expect_cover', action='store_const', const=False, help='Require no cover section')
    parser.add_argument('--expect-toc', dest='expect_toc', action='store_const', const=True, default=None, help='Require a TOC section')
    parser.add_argument('--no-expect-toc', dest='expect_toc', action='store_const', const=False, help='Require no TOC section')
    return parser.parse_args()


def main():
    args = parse_args()
    input_path = Path(args.input).expanduser().resolve()
    if not input_path.exists():
        raise SystemExit(f'Input file not found: {input_path}')
    if input_path.suffix.lower() != '.docx':
        raise SystemExit(f'Unsupported input type: {input_path.suffix.lower()}. Supported type: .docx')
    check_runtime_dependencies(input_path)
    verify_existing_docx(input_path, args.numbering_mode, args.expect_cover, args.expect_toc)
    print(input_path)


if __name__ == '__main__':
    main()
