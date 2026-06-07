#!/usr/bin/env python3
from __future__ import annotations

import argparse
import importlib.util
import subprocess
import sys
from pathlib import Path

DOCX_SKILL_ROOT = Path(
    "/Users/huangcn/.workbuddy/plugins/marketplaces/experts/plugins/document-skills/skills/docx"
)
UNPACK_SCRIPT = DOCX_SKILL_ROOT / "scripts/office/unpack.py"
VALIDATE_SCRIPT = DOCX_SKILL_ROOT / "scripts/office/validate.py"
PACK_SCRIPT = DOCX_SKILL_ROOT / "scripts/office/pack.py"
REQUIRED_MODULES = ("defusedxml", "lxml")


def ensure_tooling() -> None:
    missing = [path for path in (UNPACK_SCRIPT, VALIDATE_SCRIPT, PACK_SCRIPT) if not path.exists()]
    if missing:
        joined = ", ".join(str(path) for path in missing)
        raise SystemExit(f"Required docx tooling not found: {joined}")


def ensure_python_dependencies() -> None:
    module_status = {name: importlib.util.find_spec(name) is not None for name in REQUIRED_MODULES}
    missing = [name for name, present in module_status.items() if not present]
    if not missing:
        return

    status_lines = "\n".join(
        f"- {name}: {'OK' if present else 'MISSING'}" for name, present in module_status.items()
    )
    install_modules = " ".join(missing)
    raise SystemExit(
        "Missing Python dependencies for XML debug lane:\n"
        f"{status_lines}\n"
        "Install the missing modules first, for example with:\n"
        f"! {sys.executable} -m pip install {install_modules}"
    )


def run_python(script: Path, *args: str) -> None:
    command = [sys.executable, str(script), *args]
    subprocess.run(command, check=True)


def cmd_unpack(args: argparse.Namespace) -> None:
    run_python(
        UNPACK_SCRIPT,
        str(args.input),
        str(args.output_dir),
        "--merge-runs",
        "true" if args.merge_runs else "false",
        "--simplify-redlines",
        "true" if args.simplify_redlines else "false",
    )


def cmd_validate(args: argparse.Namespace) -> None:
    command = [str(args.path)]
    if args.original is not None:
        command.extend(["--original", str(args.original)])
    if args.auto_repair:
        command.append("--auto-repair")
    if args.verbose:
        command.append("--verbose")
    run_python(VALIDATE_SCRIPT, *command)


def cmd_repack(args: argparse.Namespace) -> None:
    command = [str(args.input_dir), str(args.output)]
    if args.original is not None:
        command.extend(["--original", str(args.original)])
    command.extend(["--validate", "true" if args.validate else "false"])
    run_python(PACK_SCRIPT, *command)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="DOCX XML debug helpers for unpack -> validate -> repack workflows."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    unpack_parser = subparsers.add_parser("unpack", help="Unpack a DOCX into an editable directory")
    unpack_parser.add_argument("input", type=Path, help="Input .docx file")
    unpack_parser.add_argument("output_dir", type=Path, help="Directory for unpacked XML")
    unpack_parser.add_argument(
        "--no-merge-runs",
        dest="merge_runs",
        action="store_false",
        help="Keep adjacent runs unchanged",
    )
    unpack_parser.add_argument(
        "--no-simplify-redlines",
        dest="simplify_redlines",
        action="store_false",
        help="Keep adjacent tracked changes unchanged",
    )
    unpack_parser.set_defaults(merge_runs=True, simplify_redlines=True, handler=cmd_unpack)

    validate_parser = subparsers.add_parser("validate", help="Validate a DOCX file or unpacked directory")
    validate_parser.add_argument("path", type=Path, help="Packed .docx file or unpacked directory")
    validate_parser.add_argument("--original", type=Path, help="Original .docx used for comparison")
    validate_parser.add_argument("--auto-repair", action="store_true", help="Repair common XML issues before validating")
    validate_parser.add_argument("--verbose", action="store_true", help="Show verbose validation output")
    validate_parser.set_defaults(handler=cmd_validate)

    repack_parser = subparsers.add_parser("repack", help="Pack an unpacked directory back into a DOCX")
    repack_parser.add_argument("input_dir", type=Path, help="Directory created by the unpack step")
    repack_parser.add_argument("output", type=Path, help="Output .docx file")
    repack_parser.add_argument("--original", type=Path, help="Original .docx used for validation comparison")
    repack_parser.add_argument(
        "--skip-validate",
        dest="validate",
        action="store_false",
        help="Pack without running validation",
    )
    repack_parser.set_defaults(validate=True, handler=cmd_repack)

    return parser


def main() -> None:
    ensure_tooling()
    ensure_python_dependencies()
    parser = build_parser()
    args = parser.parse_args()
    args.handler(args)


if __name__ == "__main__":
    main()
