"""Deterministic schema planning and value preparation for jxcel."""

from __future__ import annotations

import json
import math
from dataclasses import dataclass, replace
from typing import Any, Iterable, Iterator, Literal, Sequence

from .discovery import DatasetCandidate


EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_DATA_ROWS = EXCEL_MAX_ROWS - 1
EXCEL_MAX_COLUMNS = 16_384
EXCEL_MAX_CELL_UTF16_UNITS = 32_767
EXCEL_MAX_SAFE_INTEGER_EXCLUSIVE = 10**15
EXCEL_MAX_FLOAT = 9.99999999999999e307
EXCEL_MIN_POSITIVE_FLOAT = 2.2250738585072014e-308
WARNING_SAMPLE_LIMIT = 5

ValidationMode = Literal["schema", "export"]
NonExportableReason = Literal["zero_columns", "no_effective_columns"]
PreparedKind = Literal["blank", "text", "integer", "float", "boolean"]
WarningCode = Literal[
    "integer_precision_as_text",
    "float_precision_limited",
    "nested_value_as_json_text",
    "dataset_skipped_zero_columns",
    "dataset_skipped_no_effective_columns",
]


class PlanningError(Exception):
    """A path-aware error raised when an export plan cannot be built."""

    def __init__(
        self,
        message: str,
        *,
        dataset_pointer: str | None = None,
        column: str | None = None,
        record_index: int | None = None,
    ) -> None:
        super().__init__(message)
        self.dataset_pointer = dataset_pointer
        self.column = column
        self.record_index = record_index


@dataclass(frozen=True)
class ColumnPlan:
    name: str
    source_path: tuple[str, ...]


@dataclass(frozen=True)
class PlanningWarning:
    code: WarningCode
    dataset_pointer: str
    column: str | None
    occurrence_count: int
    sample_record_indexes: tuple[int, ...]


@dataclass(frozen=True)
class ColumnMapping:
    dataset_pointer: str
    column: str


@dataclass(frozen=True)
class MappingOutcome:
    kind: Literal["include", "exclude"]
    mapping: ColumnMapping
    applied: bool
    reason: str | None


@dataclass(frozen=True)
class DatasetPlan:
    pointer: str
    records: tuple[dict[str, Any], ...]
    columns: tuple[ColumnPlan, ...]
    effective_columns: tuple[ColumnPlan, ...]
    selected: bool
    non_exportable_reason: NonExportableReason | None
    warnings: tuple[PlanningWarning, ...]
    flatten: bool
    export_validated: bool

    @property
    def exportable(self) -> bool:
        return self.non_exportable_reason is None


@dataclass(frozen=True)
class PlanningResult:
    datasets: tuple[DatasetPlan, ...]
    mapping_outcomes: tuple[MappingOutcome, ...]
    validation_mode: ValidationMode

    @property
    def warnings(self) -> tuple[PlanningWarning, ...]:
        return tuple(
            warning
            for dataset in self.datasets
            for warning in dataset.warnings
        )


@dataclass(frozen=True)
class PreparedValue:
    value: str | int | float | bool | None
    kind: PreparedKind


_MISSING = object()


def _encode_segment(segment: str) -> str:
    return "".join(
        "~0" if character == "~" else "~1" if character == "." else character
        for character in segment
    )


def _decode_segment(segment: str) -> str:
    characters: list[str] = []
    index = 0
    while index < len(segment):
        character = segment[index]
        if character == "~" and index + 1 < len(segment):
            escaped = segment[index + 1]
            if escaped == "0":
                characters.append("~")
                index += 2
                continue
            if escaped == "1":
                characters.append(".")
                index += 2
                continue
        characters.append(character)
        index += 1
    return "".join(characters)


def encode_column_path(path: Sequence[str]) -> str:
    """Encode original key segments as one flattened column name."""
    return ".".join(_encode_segment(segment) for segment in path)


def decode_column_name(name: str) -> tuple[str, ...]:
    """Decode one flattened column name back to original key segments."""
    return tuple(_decode_segment(segment) for segment in name.split("."))


def parse_column_mapping(argument: str) -> ColumnMapping:
    """Parse one JSON-pair column mapping argument."""
    try:
        value = json.loads(argument)
    except json.JSONDecodeError as exc:
        raise PlanningError(
            f"Invalid JSON column mapping {argument!r}: {exc}"
        ) from exc
    if (
        not isinstance(value, list)
        or len(value) != 2
        or not all(isinstance(item, str) for item in value)
    ):
        raise PlanningError(
            f"Column mapping must be a JSON array of exactly two strings: "
            f"{argument!r}"
        )
    return ColumnMapping(dataset_pointer=value[0], column=value[1])


def _iter_record_leaves(
    record: dict[str, Any], flatten: bool
) -> Iterator[tuple[tuple[str, ...], Any]]:
    if not flatten:
        for key, value in record.items():
            yield (key,), value
        return

    stack: list[tuple[tuple[str, ...], Any]] = [((), record)]
    while stack:
        path, value = stack.pop()
        if isinstance(value, dict):
            if not value:
                continue
            for key, child in reversed(list(value.items())):
                stack.append((path + (key,), child))
        else:
            yield path, value


def _add_column(
    columns: list[ColumnPlan],
    registry: dict[str, tuple[str, ...]],
    name: str,
    source_path: tuple[str, ...],
    dataset_pointer: str,
) -> None:
    existing_path = registry.get(name)
    if existing_path is None:
        registry[name] = source_path
        columns.append(ColumnPlan(name=name, source_path=source_path))
        return
    if existing_path != source_path:
        raise PlanningError(
            f"Conflicting flattened column name {name!r} in dataset "
            f"{dataset_pointer!r}: source paths {existing_path!r} and "
            f"{source_path!r}",
            dataset_pointer=dataset_pointer,
            column=name,
        )


def _infer_columns(
    records: Sequence[dict[str, Any]],
    flatten: bool,
    dataset_pointer: str,
) -> tuple[ColumnPlan, ...]:
    columns: list[ColumnPlan] = []
    registry: dict[str, tuple[str, ...]] = {}
    for record in records:
        for source_path, _ in _iter_record_leaves(record, flatten):
            name = (
                encode_column_path(source_path)
                if flatten
                else source_path[0]
            )
            _add_column(
                columns,
                registry,
                name,
                source_path,
                dataset_pointer,
            )
    return tuple(columns)


def _dedupe_mappings(
    arguments: Iterable[str],
) -> tuple[ColumnMapping, ...]:
    seen: set[tuple[str, str]] = set()
    mappings: list[ColumnMapping] = []
    for argument in arguments:
        mapping = parse_column_mapping(argument)
        key = (mapping.dataset_pointer, mapping.column)
        if key in seen:
            continue
        seen.add(key)
        mappings.append(mapping)
    return tuple(mappings)


def _validate_mappings(
    mappings: Iterable[ColumnMapping],
    schemas: dict[str, tuple[ColumnPlan, ...]],
) -> None:
    for mapping in mappings:
        columns = schemas.get(mapping.dataset_pointer)
        if columns is None:
            raise PlanningError(
                f"Column mapping references unknown dataset pointer "
                f"{mapping.dataset_pointer!r}",
                dataset_pointer=mapping.dataset_pointer,
                column=mapping.column,
            )
        if mapping.column not in {column.name for column in columns}:
            raise PlanningError(
                f"Column mapping references unknown column "
                f"{mapping.column!r} for dataset "
                f"{mapping.dataset_pointer!r}",
                dataset_pointer=mapping.dataset_pointer,
                column=mapping.column,
            )


def _resolve_value(
    record: dict[str, Any],
    source_path: tuple[str, ...],
    flatten: bool,
) -> Any:
    value: Any = record
    for segment in source_path:
        if not isinstance(value, dict) or segment not in value:
            return _MISSING
        value = value[segment]
    if flatten and isinstance(value, dict):
        return _MISSING
    return value


def _json_text(value: Any) -> str:
    return json.dumps(
        value,
        ensure_ascii=False,
        separators=(",", ":"),
        allow_nan=False,
    )


def _prepare_value(value: Any, flatten: bool) -> PreparedValue:
    if value is _MISSING or value is None:
        return PreparedValue(value=None, kind="blank")
    if isinstance(value, bool):
        return PreparedValue(value=value, kind="boolean")
    if isinstance(value, int):
        if abs(value) >= EXCEL_MAX_SAFE_INTEGER_EXCLUSIVE:
            return PreparedValue(value=str(value), kind="text")
        return PreparedValue(value=value, kind="integer")
    if isinstance(value, float):
        return PreparedValue(value=value, kind="float")
    if isinstance(value, str):
        return PreparedValue(value=value, kind="text")
    if isinstance(value, list):
        return PreparedValue(value=_json_text(value), kind="text")
    if isinstance(value, dict):
        if flatten:
            return PreparedValue(value=None, kind="blank")
        return PreparedValue(value=_json_text(value), kind="text")
    raise PlanningError(
        f"Unsupported JSON value type: {type(value).__name__}"
    )


def _significant_digit_count(text: str) -> int:
    text = text.lstrip("+-")
    mantissa = text.split("e", 1)[0].split("E", 1)[0]
    digits = mantissa.replace(".", "").lstrip("0").rstrip("0")
    return len(digits)


def _invalid_xml_character(text: str) -> tuple[int, int] | None:
    for index, character in enumerate(text):
        code_point = ord(character)
        if code_point in (0x09, 0x0A, 0x0D):
            continue
        if code_point < 0x20:
            return index, code_point
        if 0xD800 <= code_point <= 0xDFFF:
            return index, code_point
        if code_point in (0xFFFE, 0xFFFF):
            return index, code_point
    return None


def _utf16_length(text: str) -> int:
    return sum(2 if ord(character) > 0xFFFF else 1 for character in text)


def _validate_text(
    text: str,
    dataset_pointer: str,
    column: str,
    record_index: int | None,
) -> None:
    invalid = _invalid_xml_character(text)
    if invalid is not None:
        position, code_point = invalid
        location = (
            "column header"
            if record_index is None
            else f"record {record_index}"
        )
        raise PlanningError(
            f"Invalid Excel character U+{code_point:04X} at position "
            f"{position} in {location} for dataset {dataset_pointer!r}, "
            f"column {column!r}",
            dataset_pointer=dataset_pointer,
            column=column,
            record_index=record_index,
        )
    units = _utf16_length(text)
    if units > EXCEL_MAX_CELL_UTF16_UNITS:
        if record_index is None:
            context = (
                f"Column header for dataset {dataset_pointer!r}, "
                f"column {column!r}"
            )
        else:
            context = (
                f"Cell value for dataset {dataset_pointer!r}, column "
                f"{column!r}, record {record_index}"
            )
        raise PlanningError(
            f"{context} exceeds Excel's "
            f"{EXCEL_MAX_CELL_UTF16_UNITS} UTF-16 code unit limit "
            f"(found {units})",
            dataset_pointer=dataset_pointer,
            column=column,
            record_index=record_index,
        )


def _validate_float(
    value: float,
    dataset_pointer: str,
    column: str,
    record_index: int,
) -> None:
    if not math.isfinite(value):
        raise PlanningError(
            f"Non-finite float value in dataset {dataset_pointer!r}, "
            f"column {column!r}, record {record_index}",
            dataset_pointer=dataset_pointer,
            column=column,
            record_index=record_index,
        )
    magnitude = abs(value)
    if magnitude != 0.0 and (
        magnitude > EXCEL_MAX_FLOAT
        or magnitude < EXCEL_MIN_POSITIVE_FLOAT
    ):
        raise PlanningError(
            f"Float value {value!r} in dataset {dataset_pointer!r}, "
            f"column {column!r}, record {record_index} is outside Excel's "
            f"numeric range",
            dataset_pointer=dataset_pointer,
            column=column,
            record_index=record_index,
        )


def _add_warning(
    counts: dict[tuple[WarningCode, str | None], int],
    samples: dict[tuple[WarningCode, str | None], list[int]],
    code: WarningCode,
    column: str | None,
    record_index: int | None,
) -> None:
    key = (code, column)
    counts[key] = counts.get(key, 0) + 1
    sample_list = samples.setdefault(key, [])
    if (
        record_index is not None
        and len(sample_list) < WARNING_SAMPLE_LIMIT
    ):
        sample_list.append(record_index)


def _finalize_warnings(
    counts: dict[tuple[WarningCode, str | None], int],
    samples: dict[tuple[WarningCode, str | None], list[int]],
    columns: Sequence[ColumnPlan],
    dataset_pointer: str,
) -> tuple[PlanningWarning, ...]:
    column_order = {
        column.name: index for index, column in enumerate(columns)
    }
    ordered_keys = sorted(
        counts,
        key=lambda key: (
            -1
            if key[1] is None
            else column_order.get(key[1], len(column_order)),
            key[0],
        ),
    )
    return tuple(
        PlanningWarning(
            code=code,
            dataset_pointer=dataset_pointer,
            column=column,
            occurrence_count=counts[(code, column)],
            sample_record_indexes=tuple(samples.get((code, column), ())),
        )
        for code, column in ordered_keys
    )


def _validate_export_plan(plan: DatasetPlan) -> tuple[PlanningWarning, ...]:
    if len(plan.records) > EXCEL_MAX_DATA_ROWS:
        raise PlanningError(
            f"Dataset {plan.pointer!r} has {len(plan.records)} records, "
            f"exceeding Excel's {EXCEL_MAX_DATA_ROWS} data-row limit",
            dataset_pointer=plan.pointer,
        )
    if len(plan.effective_columns) > EXCEL_MAX_COLUMNS:
        raise PlanningError(
            f"Dataset {plan.pointer!r} has {len(plan.effective_columns)} "
            f"effective columns, exceeding Excel's {EXCEL_MAX_COLUMNS} "
            f"column limit",
            dataset_pointer=plan.pointer,
        )

    for column in plan.effective_columns:
        _validate_text(
            column.name,
            plan.pointer,
            column.name,
            None,
        )

    counts: dict[tuple[WarningCode, str | None], int] = {}
    samples: dict[tuple[WarningCode, str | None], list[int]] = {}

    for record_index, record in enumerate(plan.records):
        for column in plan.effective_columns:
            value = _resolve_value(
                record,
                column.source_path,
                plan.flatten,
            )
            prepared = _prepare_value(value, plan.flatten)
            if prepared.kind == "text":
                _validate_text(
                    prepared.value,
                    plan.pointer,
                    column.name,
                    record_index,
                )
                if (
                    isinstance(value, int)
                    and not isinstance(value, bool)
                    and abs(value) >= EXCEL_MAX_SAFE_INTEGER_EXCLUSIVE
                ):
                    _add_warning(
                        counts,
                        samples,
                        "integer_precision_as_text",
                        column.name,
                        record_index,
                    )
                elif isinstance(value, (list, dict)):
                    _add_warning(
                        counts,
                        samples,
                        "nested_value_as_json_text",
                        column.name,
                        record_index,
                    )
            elif prepared.kind == "float":
                _validate_float(
                    prepared.value,
                    plan.pointer,
                    column.name,
                    record_index,
                )
                if (
                    _significant_digit_count(repr(prepared.value)) > 15
                ):
                    _add_warning(
                        counts,
                        samples,
                        "float_precision_limited",
                        column.name,
                        record_index,
                    )

    return _finalize_warnings(
        counts,
        samples,
        plan.effective_columns,
        plan.pointer,
    )


def _iter_prepared_rows(
    plan: DatasetPlan,
) -> Iterator[tuple[PreparedValue, ...]]:
    for record in plan.records:
        yield tuple(
            _prepare_value(
                _resolve_value(
                    record,
                    column.source_path,
                    plan.flatten,
                ),
                plan.flatten,
            )
            for column in plan.effective_columns
        )


def prepared_rows(
    plan: DatasetPlan,
) -> Iterator[tuple[PreparedValue, ...]]:
    """Return prepared rows for an export-validated, exportable plan."""
    if not plan.export_validated:
        raise PlanningError(
            f"Dataset plan {plan.pointer!r} has not been validated for export"
        )
    if not plan.selected:
        raise PlanningError(
            f"Dataset plan {plan.pointer!r} is not selected"
        )
    if not plan.exportable:
        raise PlanningError(
            f"Dataset plan {plan.pointer!r} is not exportable"
        )
    return _iter_prepared_rows(plan)


def build_plan(
    candidates: Sequence[DatasetCandidate],
    *,
    selected_pointers: Sequence[str] | None = None,
    include_maps: Sequence[str] = (),
    exclude_maps: Sequence[str] = (),
    flatten: bool = True,
    validation_mode: ValidationMode = "export",
) -> PlanningResult:
    """Build a deterministic schema or export plan."""
    if validation_mode not in ("schema", "export"):
        raise ValueError(
            f"validation_mode must be 'schema' or 'export', "
            f"got {validation_mode!r}"
        )

    candidate_list = tuple(candidates)
    pointers = [candidate.pointer for candidate in candidate_list]

    if selected_pointers is None:
        selected_set = set(pointers)
    else:
        selected_set = set()
        for pointer in selected_pointers:
            if pointer not in pointers:
                raise PlanningError(
                    f"Unknown dataset pointer in selection: {pointer!r}",
                    dataset_pointer=pointer,
                )
            selected_set.add(pointer)

    schemas = {
        candidate.pointer: _infer_columns(
            candidate.records,
            flatten,
            candidate.pointer,
        )
        for candidate in candidate_list
    }

    include_mappings = _dedupe_mappings(include_maps)
    exclude_mappings = _dedupe_mappings(exclude_maps)
    _validate_mappings(
        include_mappings + exclude_mappings,
        schemas,
    )

    include_columns: dict[str, set[str]] = {}
    exclude_columns: dict[str, set[str]] = {}
    for mapping in include_mappings:
        include_columns.setdefault(
            mapping.dataset_pointer, set()
        ).add(mapping.column)
    for mapping in exclude_mappings:
        exclude_columns.setdefault(
            mapping.dataset_pointer, set()
        ).add(mapping.column)

    mapping_outcomes: list[MappingOutcome] = []
    for kind, mappings in (
        ("include", include_mappings),
        ("exclude", exclude_mappings),
    ):
        for mapping in mappings:
            applied = mapping.dataset_pointer in selected_set
            mapping_outcomes.append(
                MappingOutcome(
                    kind=kind,
                    mapping=mapping,
                    applied=applied,
                    reason=None if applied else "dataset_not_selected",
                )
            )

    plans: list[DatasetPlan] = []
    for candidate in candidate_list:
        columns = schemas[candidate.pointer]
        selected = candidate.pointer in selected_set
        reason: NonExportableReason | None = (
            "zero_columns" if not columns else None
        )
        effective_columns: tuple[ColumnPlan, ...] = ()

        if selected and columns:
            included = include_columns.get(candidate.pointer)
            if included:
                effective_columns = tuple(
                    column
                    for column in columns
                    if column.name in included
                )
            else:
                effective_columns = columns
            excluded = exclude_columns.get(candidate.pointer)
            if excluded:
                effective_columns = tuple(
                    column
                    for column in effective_columns
                    if column.name not in excluded
                )
            if not effective_columns:
                reason = "no_effective_columns"

        plans.append(
            DatasetPlan(
                pointer=candidate.pointer,
                records=candidate.records,
                columns=columns,
                effective_columns=effective_columns,
                selected=selected,
                non_exportable_reason=reason,
                warnings=(),
                flatten=flatten,
                export_validated=False,
            )
        )

    final_plans: list[DatasetPlan] = []
    for plan in plans:
        if plan.selected and plan.non_exportable_reason is not None:
            counts: dict[tuple[WarningCode, str | None], int] = {}
            samples: dict[tuple[WarningCode, str | None], list[int]] = {}
            code: WarningCode = (
                "dataset_skipped_zero_columns"
                if plan.non_exportable_reason == "zero_columns"
                else "dataset_skipped_no_effective_columns"
            )
            _add_warning(counts, samples, code, None, None)
            warnings = _finalize_warnings(
                counts,
                samples,
                plan.columns,
                plan.pointer,
            )
            final_plans.append(replace(plan, warnings=warnings))
            continue

        if (
            validation_mode == "export"
            and plan.selected
            and plan.exportable
        ):
            warnings = _validate_export_plan(plan)
            final_plans.append(
                replace(
                    plan,
                    warnings=warnings,
                    export_validated=True,
                )
            )
            continue

        final_plans.append(plan)

    return PlanningResult(
        datasets=tuple(final_plans),
        mapping_outcomes=tuple(mapping_outcomes),
        validation_mode=validation_mode,
    )
