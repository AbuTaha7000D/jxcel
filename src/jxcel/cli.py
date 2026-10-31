"""Command-line interface and pipeline orchestration for jxcel."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Sequence

from . import __version__
from .discovery import DatasetCandidate, discover
from .jsonio import JsonLoadError, load_json
from .planning import (
    DatasetPlan,
    PlanningError,
    PreparedValue,
    build_plan,
    preview_rows,
)
from .reporting import (
    render_export,
    render_listing,
    render_plan,
    render_preview,
)
from .workbook import (
    WorkbookError,
    assign_sheet_names,
    export_workbook,
    validate_output_destination,
)


def _positive_integer(value: str) -> int:
    try:
        number = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            f"invalid integer value: {value!r}"
        ) from exc
    if number < 1:
        raise argparse.ArgumentTypeError("must be a positive integer")
    return number


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="jxcel",
        description=(
            "Convert JSON datasets into a formatted Excel .xlsx workbook."
        ),
    )
    parser.add_argument(
        "input",
        nargs="?",
        metavar="INPUT",
        help="input JSON file (mutually exclusive with -i/--input)",
    )
    parser.add_argument(
        "-i",
        "--input",
        dest="input_option",
        metavar="PATH",
        help="input JSON file (mutually exclusive with positional INPUT)",
    )
    parser.add_argument(
        "-o",
        "--output",
        metavar="PATH",
        help=(
            "output .xlsx path; defaults to the input name with .xlsx "
            "beside the input"
        ),
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="replace an existing output file after validation",
    )
    parser.add_argument(
        "-d",
        "--dataset",
        action="append",
        default=[],
        metavar="JSON_POINTER",
        help=(
            "select a dataset by JSON Pointer; repeat to select more "
            "than one (default: all discovered datasets)"
        ),
    )
    parser.add_argument(
        "--list",
        dest="list_mode",
        action="store_true",
        help="list datasets and unsupported structures without writing",
    )
    parser.add_argument(
        "--preview",
        action="store_true",
        help="show bounded dataset records without writing",
    )
    parser.add_argument(
        "--preview-rows",
        type=_positive_integer,
        default=None,
        metavar="N",
        help="number of preview rows per dataset (default: 5)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="validate and show the complete plan without writing",
    )
    parser.add_argument(
        "--column-map",
        action="append",
        default=[],
        metavar="JSON_PAIR",
        help=(
            "include an exact column using "
            '\'["/dataset/pointer","column"]\'; repeatable'
        ),
    )
    parser.add_argument(
        "--exclude-map",
        action="append",
        default=[],
        metavar="JSON_PAIR",
        help=(
            "exclude an exact column using "
            '\'["/dataset/pointer","column"]\'; repeatable'
        ),
    )
    parser.add_argument(
        "--no-flatten",
        action="store_true",
        help=(
            "use one column per top-level property and keep nested "
            "objects as compact JSON text"
        ),
    )
    verbosity = parser.add_mutually_exclusive_group()
    verbosity.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="add diagnostic detail to reports",
    )
    verbosity.add_argument(
        "-q",
        "--quiet",
        action="store_true",
        help=(
            "suppress the routine success summary; warnings and "
            "findings remain visible"
        ),
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    return parser


def _validate_arguments(
    parser: argparse.ArgumentParser,
    args: argparse.Namespace,
) -> None:
    if args.input is not None and args.input_option is not None:
        parser.error(
            "provide the input either positionally or with -i/--input, "
            "not both"
        )
    if args.input is None and args.input_option is None:
        parser.error("an input JSON file is required")

    if args.list_mode and (args.preview or args.dry_run):
        parser.error(
            "--list cannot be combined with --preview or --dry-run"
        )
    if args.preview_rows is not None and not args.preview:
        parser.error("--preview-rows requires --preview")

    preview_only = args.preview and not args.dry_run
    if args.list_mode or preview_only:
        if args.output is not None:
            parser.error(
                "--output is not valid with --list or preview-only mode"
            )
        if args.overwrite:
            parser.error(
                "--overwrite is not valid with --list or preview-only mode"
            )

    if args.preview and args.preview_rows is None:
        args.preview_rows = 5


def _emit(text: str) -> None:
    if text:
        print(text)


def _fail(message: str) -> None:
    print(f"jxcel: error: {message}", file=sys.stderr)


def _exportable_plans(
    plans: Sequence[DatasetPlan],
) -> tuple[DatasetPlan, ...]:
    return tuple(
        plan
        for plan in plans
        if plan.selected and plan.exportable
    )


def _preview_rows_by_pointer(
    plans: Sequence[DatasetPlan],
    limit: int,
) -> dict[str, tuple[tuple[PreparedValue, ...], ...]]:
    return {
        plan.pointer: tuple(preview_rows(plan, limit))
        for plan in plans
    }


def main(argv: Sequence[str] | None = None) -> int:
    """Run the jxcel command and return its exit status."""
    parser = _build_parser()
    args = parser.parse_args(argv)
    _validate_arguments(parser, args)

    input_path = Path(
        args.input if args.input is not None else args.input_option
    )

    try:
        document = load_json(input_path)
        findings = discover(document)
        candidates = tuple(
            finding
            for finding in findings
            if isinstance(finding, DatasetCandidate)
        )

        schema_only = args.list_mode or (
            args.preview and not args.dry_run
        )
        result = build_plan(
            candidates,
            selected_pointers=args.dataset or None,
            include_maps=args.column_map,
            exclude_maps=args.exclude_map,
            flatten=not args.no_flatten,
            validation_mode="schema" if schema_only else "export",
        )
        plans = result.datasets
        exportable = _exportable_plans(plans)

        if args.list_mode:
            _emit(
                render_listing(
                    plans,
                    findings,
                    input_path=input_path,
                    verbose=args.verbose,
                )
            )
            return 0

        if args.preview and not args.dry_run:
            rows = _preview_rows_by_pointer(
                exportable,
                args.preview_rows,
            )
            _emit(
                render_preview(
                    plans,
                    rows,
                    findings,
                    input_path=input_path,
                    verbose=args.verbose,
                )
            )
            return 0

        if args.output is not None:
            output_path = Path(args.output)
        else:
            output_path = input_path.with_suffix(".xlsx")

        if args.dry_run:
            if not exportable:
                _emit(
                    render_listing(
                        plans,
                        findings,
                        input_path=input_path,
                        verbose=args.verbose,
                    )
                )
                _fail("No selected exportable datasets to write")
                return 1
            validate_output_destination(
                output_path,
                overwrite=args.overwrite,
                input_path=input_path,
            )
            sheet_names = None
            if args.verbose:
                sheet_names = {
                    assignment.pointer: assignment.name
                    for assignment in assign_sheet_names(exportable)
                }
            sections = [
                render_plan(
                    plans,
                    findings,
                    output_path,
                    overwrite=args.overwrite,
                    mapping_outcomes=result.mapping_outcomes,
                    sheet_names=sheet_names,
                    input_path=input_path,
                    verbose=args.verbose,
                )
            ]
            if args.preview:
                rows = _preview_rows_by_pointer(
                    exportable,
                    args.preview_rows,
                )
                sections.append(
                    render_preview(plans, rows, ())
                )
            _emit("\n\n".join(sections))
            return 0

        if not exportable:
            _emit(
                render_listing(
                    plans,
                    findings,
                    input_path=input_path,
                    verbose=args.verbose,
                )
            )
            _fail("No selected exportable datasets to write")
            return 1

        export = export_workbook(
            plans,
            output_path,
            overwrite=args.overwrite,
            input_path=input_path,
        )
        _emit(
            render_export(
                export,
                plans,
                findings,
                input_path=input_path,
                quiet=args.quiet,
                verbose=args.verbose,
            )
        )
        return 0
    except (JsonLoadError, PlanningError, WorkbookError) as exc:
        _fail(str(exc))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
