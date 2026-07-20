"""Run one explicit xlvbatools test tier with reliable captured output."""

from __future__ import annotations

import argparse
from pathlib import Path
import subprocess
import sys
import tempfile


SUITE_MARKERS: dict[str, str | None] = {
    "fast": None,
    "unit": "unit",
    "integration": "integration",
    "excel-smoke": "excel and smoke",
    "excel": "excel and not stress",
    "stress": "excel and stress",
    "distribution": "distribution",
    "external": "external",
    "all": "not external",
}


def build_pytest_arguments(
    suite: str,
    *,
    external_workbooks: tuple[str, ...] = (),
    coverage: bool = False,
    extra: tuple[str, ...] = (),
) -> list[str]:
    if suite not in SUITE_MARKERS:
        raise ValueError(f"Unknown test suite: {suite}")
    if suite == "external" and not external_workbooks:
        raise ValueError("external suite requires --external-workbook PATH")
    if external_workbooks and suite != "external":
        raise ValueError("--external-workbook is valid only for the external suite")

    arguments = ["-m", "pytest"]
    marker = SUITE_MARKERS[suite]
    if marker is not None:
        arguments.extend(("-m", marker))
    if suite == "external":
        # An explicit test path lets pytest load tests/conftest.py and register
        # --external-workbook during its initial command-line parse.
        arguments.append("tests/test_external_workbook.py")
    if coverage:
        arguments.extend((
            "--cov=xlvbatools",
            "--cov-report=term-missing",
            "--cov-fail-under=60",
        ))
    for workbook in external_workbooks:
        arguments.extend(("--external-workbook", str(Path(workbook).resolve())))
    arguments.extend(extra)
    return arguments


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Run an explicit xlvbatools test tier. Live-suite output is routed "
            "through a seekable file so Excel cannot retain a console pipe."
        ),
    )
    parser.add_argument(
        "suite",
        nargs="?",
        default="fast",
        choices=tuple(SUITE_MARKERS),
    )
    parser.add_argument("--external-workbook", action="append", default=[])
    parser.add_argument("--coverage", action="store_true")
    parser.add_argument(
        "--log-file",
        type=Path,
        help="also persist captured pytest output and the final exit code",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = _parser()
    options, passthrough = parser.parse_known_args(argv)
    extra = tuple(passthrough)
    if extra[:1] == ("--",):
        extra = extra[1:]
    try:
        arguments = build_pytest_arguments(
            options.suite,
            external_workbooks=tuple(options.external_workbook),
            coverage=options.coverage,
            extra=extra,
        )
    except ValueError as error:
        parser.error(str(error))

    if options.log_file is not None:
        log_file = options.log_file.resolve()
        log_file.parent.mkdir(parents=True, exist_ok=True)
        output_context = log_file.open(
            "w+", encoding="utf-8", errors="replace",
        )
    else:
        output_context = tempfile.TemporaryFile(
            mode="w+", encoding="utf-8", errors="replace",
        )

    with output_context as output:
        completed = subprocess.run(
            [sys.executable, *arguments],
            cwd=Path(__file__).resolve().parents[1],
            stdout=output,
            stderr=subprocess.STDOUT,
        )
        output.write(f"\nTEST_RUNNER_EXIT_CODE={completed.returncode}\n")
        output.flush()
        output.seek(0)
        transcript = output.read()
        sys.stdout.write(transcript)
        sys.stdout.flush()
    return completed.returncode


if __name__ == "__main__":
    sys.exit(main())
