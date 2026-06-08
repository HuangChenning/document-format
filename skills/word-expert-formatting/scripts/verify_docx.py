#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from text_to_docx import (
    Document,
    build_verification_context,
    check_runtime_dependencies,
    count_footer_relationships,
    find_toc_region,
    infer_cover,
    infer_expected_headings,
    snapshot_cover_signature,
    snapshot_trailing_sectpr_xml,
    verify_document_workflow,
)


def verify_existing_docx(
    input_path: Path,
    expect_cover: bool | None,
    expect_toc: bool | None,
    freeze_cover: bool,
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
        source_kind='refreshed',
        toc_refresh_state='not_refreshed',
        preserve_cover_format=freeze_cover,
        frozen_cover_signature=snapshot_cover_signature(doc) if freeze_cover else None,
        frozen_section_count=len(doc.sections) if freeze_cover else None,
        frozen_footer_relationship_count=count_footer_relationships(doc) if freeze_cover else None,
        frozen_trailing_sectpr_xml=snapshot_trailing_sectpr_xml(doc) if freeze_cover else None,
    )
    verify_document_workflow(doc, context)


def parse_args():
    parser = argparse.ArgumentParser(description='Verify an existing DOCX against the word formatting workflow rules.')
    parser.add_argument('input', help='Input .docx file')
    parser.add_argument('--expect-cover', dest='expect_cover', action='store_const', const=True, default=None, help='Require a cover section')
    parser.add_argument('--no-expect-cover', dest='expect_cover', action='store_const', const=False, help='Require no cover section')
    parser.add_argument('--expect-toc', dest='expect_toc', action='store_const', const=True, default=None, help='Require a TOC section')
    parser.add_argument('--no-expect-toc', dest='expect_toc', action='store_const', const=False, help='Require no TOC section')
    parser.add_argument('--freeze-cover', action='store_true', help='Verify the existing cover region as fully frozen')
    return parser.parse_args()


def main():
    args = parse_args()
    input_path = Path(args.input).expanduser().resolve()
    if not input_path.exists():
        raise SystemExit(f'Input file not found: {input_path}')
    if input_path.suffix.lower() != '.docx':
        raise SystemExit(f'Unsupported input type: {input_path.suffix.lower()}. Supported type: .docx')
    check_runtime_dependencies(input_path)
    verify_existing_docx(input_path, args.expect_cover, args.expect_toc, args.freeze_cover)
    print(input_path)


if __name__ == '__main__':
    main()
