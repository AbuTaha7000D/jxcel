"""Pure text rendering for jxcel command output."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Mapping, Sequence

from .discovery import Finding, StructureFinding
from .planning import (
    ColumnPlan,
    DatasetPlan,
    MappingOutcome,
    PlanningWarning,
    PreparedValue,
)
from .workbook import WorkbookExport, SheetSummary


def _count(value: int, singular: str, plural: str | None = None) -> str:
    if value == 1:
        return f"1 {singular}"
    return f"{value} {plural or singular + 's'}"


def _escape_controls(text: str) -> str:
    """Escape control characters and surrogates for safe display only."""
    characters: list[str] = []
    for character in text:
        code_point = ord(character)
        if (
            code_point < 0x20
            or 0xD800 <= code_point <= 0xDFFF
            or code_point in (0xFFFE, 0xFFFF)
        ):
            characters.append(f"\\u{code_point:04x}")
        else:
            characters.append(character)
    return "".join(characters)


def _cell_text(value: PreparedValue) -> str:
    if value.kind == "blank":
        return ""
    if value.kind == "boolean":
        return "TRUE" if value.value else "FALSE"
    text = str(value.value)
    if (
        "|" in text
        or text != text.strip()
        or "\n" in text
        or "\r" in text
        or "\t" in text
    ):
        return _escape_controls(
            json.dumps(text, ensure_ascii=False)
        )
    return _escape_controls(text)


def _columns_text(columns: Sequence[ColumnPlan]) -> str:
    names = [column.name for column in columns]
    if not names:
        return "(none)"
    return ", ".join(_escape_controls(name) for name in names)


def _unsupported_findings(
    findings: Sequence[Finding],
) -> list[StructureFinding]:
    return [
        finding
        for finding in findings
        if isinstance(finding, StructureFinding)
    ]


def _unsupported_section(findings: Sequence[Finding]) -> list[str]:
    unsupported = _unsupported_findings(findings)
    if not unsupported:
        return []
    lines = [f"Unsupported findings: {len(unsupported)}"]
    for finding in unsupported:
        lines.append(f"  {finding.pointer!r}: {finding.kind}")
    return lines


def _warning_line(warning: PlanningWarning) -> str:
    parts = [f"  {warning.code} at {warning.dataset_pointer!r}"]
    if warning.column is not None:
        parts.append(f", column {warning.column!r}")
    parts.append(
        f": {_count(warning.occurrence_count, 'occurrence')}"
    )
    if warning.sample_record_indexes:
        samples = ", ".join(
            str(index) for index in warning.sample_record_indexes
        )
        parts.append(f"; sample records {samples}")
    return "".join(parts)


def _warnings_section(plans: Sequence[DatasetPlan]) -> list[str]:
    warnings = [
        warning for plan in plans for warning in plan.warnings
    ]
    if not warnings:
        return []
    lines = [f"Warnings: {len(warnings)}"]
    lines.extend(_warning_line(warning) for warning in warnings)
    return lines


def _no_exportable_note(plans: Sequence[DatasetPlan]) -> list[str]:
    if any(plan.selected and plan.exportable for plan in plans):
        return []
    return [
        "",
        "No exportable datasets; conversion would not produce a workbook.",
    ]


def render_listing(
    plans: Sequence[DatasetPlan],
    findings: Sequence[Finding],
    *,
    input_path: Path | None = None,
    verbose: bool = False,
) -> str:
    """Render the dataset and unsupported-structure listing."""
    selected = [plan for plan in plans if plan.selected]
    lines: list[str] = []
    if verbose and input_path is not None:
        lines.append(f"Input: {str(input_path)!r}")
    lines.append(f"Datasets: {len(selected)}")
    if not selected:
        lines.append("  (none)")
    for plan in selected:
        header = (
            f"  {plan.pointer!r}: "
            f"{_count(len(plan.records), 'record')}, "
            f"{_count(len(plan.columns), 'column')}"
        )
        if not plan.exportable:
            header += (
                f" [non-exportable: {plan.non_exportable_reason}]"
            )
        lines.append(header)
        lines.append(f"    available: {_columns_text(plan.columns)}")
        if plan.effective_columns != plan.columns:
            lines.append(
                f"    effective: {_columns_text(plan.effective_columns)}"
            )

    unsupported = _unsupported_section(findings)
    if unsupported:
        lines.append("")
        lines.extend(unsupported)

    lines.extend(_no_exportable_note(plans))
    return "\n".join(lines)


def render_preview(
    plans: Sequence[DatasetPlan],
    rows_by_pointer: Mapping[
        str, Sequence[tuple[PreparedValue, ...]]
    ],
    findings: Sequence[Finding],
    *,
    input_path: Path | None = None,
    verbose: bool = False,
) -> str:
    """Render bounded dataset previews."""
    selected = [plan for plan in plans if plan.selected]
    lines: list[str] = []
    if verbose and input_path is not None:
        lines.append(f"Input: {str(input_path)!r}")
    if not selected:
        lines.append("No datasets were discovered.")
    for plan in selected:
        if not plan.exportable:
            lines.append(
                f"Preview {plan.pointer!r}: skipped "
                f"(non-exportable: {plan.non_exportable_reason})"
            )
            continue
        rows = rows_by_pointer.get(plan.pointer, ())
        lines.append(
            f"Preview {plan.pointer!r} "
            f"({len(rows)} of {plan.row_count} rows):"
        )
        header = [
            _escape_controls(column.name)
            for column in plan.effective_columns
        ]
        lines.append("  " + " | ".join(header))
        for row in rows:
            lines.append(
                "  " + " | ".join(_cell_text(value) for value in row)
            )
        lines.append("")

    if lines and lines[-1] == "":
        lines.pop()

    unsupported = _unsupported_section(findings)
    if unsupported:
        lines.append("")
        lines.extend(unsupported)

    lines.extend(_no_exportable_note(plans))
    return "\n".join(lines)


def _mapping_line(outcome: MappingOutcome) -> str:
    if outcome.applied:
        status = "applied"
    else:
        reason = (outcome.reason or "not applied").replace("_", " ")
        status = f"validated but not applied ({reason})"
    return (
        f"  {outcome.kind} {outcome.mapping.dataset_pointer!r} "
        f"{outcome.mapping.column!r}: {status}"
    )


def render_plan(
    plans: Sequence[DatasetPlan],
    findings: Sequence[Finding],
    output_path: Path,
    *,
    overwrite: bool,
    mapping_outcomes: Sequence[MappingOutcome] = (),
    sheet_names: Mapping[str, str] | None = None,
    input_path: Path | None = None,
    verbose: bool = False,
) -> str:
    """Render the dry-run plan without creating a workbook."""
    selected = [plan for plan in plans if plan.selected]
    lines = ["Dry run: no workbook will be written."]
    if verbose and input_path is not None:
        lines.append(f"Input: {str(input_path)!r}")
    lines.append(f"Output: {str(output_path)!r}")
    lines.append(f"Overwrite: {'yes' if overwrite else 'no'}")
    lines.append(f"Datasets: {len(selected)}")
    if not selected:
        lines.append("  (none)")
    for plan in selected:
        header = (
            f"  {plan.pointer!r}: "
            f"{_count(plan.row_count, 'row')}, "
            f"{_count(len(plan.effective_columns), 'column')}"
        )
        if not plan.exportable:
            header += (
                f" [non-exportable: {plan.non_exportable_reason}, "
                f"skipped]"
            )
        lines.append(header)
        if plan.exportable:
            lines.append(
                f"    columns: {_columns_text(plan.effective_columns)}"
            )
            if verbose and sheet_names is not None:
                name = sheet_names.get(plan.pointer)
                if name is not None:
                    lines.append(f"    sheet: {name!r}")

    if mapping_outcomes:
        lines.append("")
        lines.append("Mappings:")
        ordered = sorted(
            mapping_outcomes,
            key=lambda outcome: (
                0 if outcome.kind == "include" else 1,
                outcome.mapping.dataset_pointer,
                outcome.mapping.column,
            ),
        )
        lines.extend(_mapping_line(outcome) for outcome in ordered)

    unsupported = _unsupported_section(findings)
    if unsupported:
        lines.append("")
        lines.extend(unsupported)

    warnings = _warnings_section(plans)
    if warnings:
        lines.append("")
        lines.extend(warnings)

    return "\n".join(lines)


def _exported_section(sheets: Sequence[SheetSummary]) -> list[str]:
    lines = [f"Datasets: {len(sheets)}"]
    if not sheets:
        lines.append("  (none)")
    for sheet in sheets:
        lines.append(
            f"  {sheet.pointer!r} -> {sheet.name!r}: "
            f"{_count(sheet.data_rows, 'row')}, "
            f"{_count(sheet.columns, 'column')}"
        )
    return lines


def render_export(
    export: WorkbookExport,
    plans: Sequence[DatasetPlan],
    findings: Sequence[Finding],
    *,
    input_path: Path | None = None,
    quiet: bool = False,
    verbose: bool = False,
) -> str:
    """Render a successful conversion summary and its findings."""
    lines: list[str] = []
    if verbose and input_path is not None:
        lines.append(f"Input: {str(input_path)!r}")
    if not quiet:
        lines.append(f"Exported: {str(export.output_path)!r}")
        lines.extend(_exported_section(export.sheets))

    skipped = [
        plan
        for plan in plans
        if plan.selected and not plan.exportable
    ]
    if skipped:
        if lines:
            lines.append("")
        lines.append(f"Skipped non-exportable datasets: {len(skipped)}")
        for plan in skipped:
            lines.append(
                f"  {plan.pointer!r}: {plan.non_exportable_reason}"
            )

    unsupported = _unsupported_section(findings)
    if unsupported:
        if lines:
            lines.append("")
        lines.extend(unsupported)

    warnings = _warnings_section(plans)
    if warnings:
        if lines:
            lines.append("")
        lines.extend(warnings)

    if export.cleanup_warning is not None:
        if lines:
            lines.append("")
        lines.append(f"Warning: {export.cleanup_warning}")

    return "\n".join(lines)
