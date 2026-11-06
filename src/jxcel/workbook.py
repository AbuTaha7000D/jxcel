"""Deterministic Excel workbook generation and safe publication for jxcel."""

from __future__ import annotations

import errno
import os
import secrets
import stat
from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO, Sequence

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ERROR_CODES
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .planning import DatasetPlan, PreparedValue, prepared_rows


EXCEL_MAX_SHEET_NAME_UTF16_UNITS = 31
WIDTH_SAMPLE_ROWS = 100
MIN_COLUMN_WIDTH = 10
MAX_COLUMN_WIDTH = 60
_TEMP_FILE_ATTEMPTS = 100
_INVALID_SHEET_NAME_CHARACTERS = frozenset("\\*?:/[]")
_HEADER_FONT = Font(bold=True)
_HEADER_FILL = PatternFill(fill_type="solid", fgColor="D9E1F2")
_HEADER_ALIGNMENT = Alignment(vertical="center")
_EXCEL_ERROR_LITERALS = frozenset(ERROR_CODES)


class WorkbookError(Exception):
    """A path- or workbook-aware failure during generation or publication."""

    def __init__(
        self,
        message: str,
        *,
        output_path: str | os.PathLike[str] | None = None,
        dataset_pointer: str | None = None,
        sheet_name: str | None = None,
        coordinate: str | None = None,
    ) -> None:
        super().__init__(message)
        self.output_path = None if output_path is None else Path(output_path)
        self.dataset_pointer = dataset_pointer
        self.sheet_name = sheet_name
        self.coordinate = coordinate


@dataclass(frozen=True)
class SheetAssignment:
    """A dataset pointer and its final Excel worksheet name."""

    pointer: str
    name: str


@dataclass(frozen=True)
class SheetSummary:
    """A concise description of one written worksheet."""

    pointer: str
    name: str
    data_rows: int
    columns: int


@dataclass(frozen=True)
class WorkbookExport:
    """The result of a successful workbook export."""

    output_path: Path
    sheets: tuple[SheetSummary, ...]
    cleanup_warning: str | None = None


def _utf16_length(text: str) -> int:
    return sum(2 if ord(character) > 0xFFFF else 1 for character in text)


def _truncate_utf16(text: str, limit: int) -> str:
    """Truncate text without splitting a supplementary-plane character."""

    if limit <= 0:
        return ""
    units = 0
    characters: list[str] = []
    for character in text:
        character_units = 2 if ord(character) > 0xFFFF else 1
        if units + character_units > limit:
            break
        characters.append(character)
        units += character_units
    return "".join(characters)


def _decode_pointer_token(token: str) -> str:
    characters: list[str] = []
    index = 0
    while index < len(token):
        character = token[index]
        if character == "~" and index + 1 < len(token):
            escaped = token[index + 1]
            if escaped == "0":
                characters.append("~")
                index += 2
                continue
            if escaped == "1":
                characters.append("/")
                index += 2
                continue
        characters.append(character)
        index += 1
    return "".join(characters)


def _raw_sheet_name(pointer: str) -> str:
    if pointer == "":
        return "root"
    body = pointer[1:] if pointer.startswith("/") else pointer
    return ".".join(
        _decode_pointer_token(token)
        for token in body.split("/")
    )


def _is_forbidden_sheet_character(character: str) -> bool:
    code_point = ord(character)
    return (
        character in _INVALID_SHEET_NAME_CHARACTERS
        or code_point < 0x20
        or 0xD800 <= code_point <= 0xDFFF
        or code_point in (0xFFFE, 0xFFFF)
    )


def _sanitize_sheet_name(value: str) -> str:
    sanitized = "".join(
        "_" if _is_forbidden_sheet_character(character) else character
        for character in value
    ).strip()
    if sanitized.startswith("'"):
        sanitized = "_" + sanitized[1:]
    if sanitized.endswith("'"):
        sanitized = sanitized[:-1] + "_"
    return sanitized


def _base_sheet_name(pointer: str) -> str:
    name = _sanitize_sheet_name(_raw_sheet_name(pointer))
    if not name:
        name = "dataset"
    if name.casefold() == "history":
        name += "_"
    return name


def _fit_sheet_name(name: str, limit: int) -> str:
    fitted = _truncate_utf16(name, limit).strip()
    if fitted.startswith("'"):
        fitted = "_" + fitted[1:]
    if fitted.endswith("'"):
        fitted = fitted[:-1] + "_"
    if not fitted:
        fitted = _truncate_utf16("dataset", limit)
    if not fitted:
        fitted = "_"
    return fitted


def assign_sheet_names(
    plans: Sequence[DatasetPlan],
) -> tuple[SheetAssignment, ...]:
    """Assign deterministic, unique Excel worksheet names to exportable plans."""

    used: set[str] = set()
    assignments: list[SheetAssignment] = []
    for plan in plans:
        if not plan.selected or not plan.exportable:
            continue

        base = _base_sheet_name(plan.pointer)
        candidate = _fit_sheet_name(
            base,
            EXCEL_MAX_SHEET_NAME_UTF16_UNITS,
        )
        suffix_number = 2
        while candidate.casefold() in used:
            suffix = f"_{suffix_number}"
            available = EXCEL_MAX_SHEET_NAME_UTF16_UNITS - len(suffix)
            if available <= 0:
                raise WorkbookError(
                    f"Unable to assign a unique worksheet name for dataset "
                    f"{plan.pointer!r}",
                    dataset_pointer=plan.pointer,
                )
            fitted = _fit_sheet_name(base, available)
            candidate = fitted + suffix
            suffix_number += 1

        if _utf16_length(candidate) > EXCEL_MAX_SHEET_NAME_UTF16_UNITS:
            raise WorkbookError(
                f"Assigned worksheet name exceeds Excel's "
                f"{EXCEL_MAX_SHEET_NAME_UTF16_UNITS} UTF-16 code unit limit",
                dataset_pointer=plan.pointer,
            )
        used.add(candidate.casefold())
        assignments.append(
            SheetAssignment(pointer=plan.pointer, name=candidate)
        )
    return tuple(assignments)


def _comparison_path(path: Path) -> str:
    try:
        resolved = os.path.realpath(path)
    except OSError:
        resolved = os.path.abspath(path)
    return os.path.normcase(resolved)


def _paths_refer_to_same_file(first: Path, second: Path) -> bool:
    try:
        if os.path.exists(first) and os.path.exists(second):
            if os.path.samefile(first, second):
                return True
    except OSError:
        pass
    return _comparison_path(first) == _comparison_path(second)


def validate_output_destination(
    output_path: str | os.PathLike[str],
    *,
    overwrite: bool = False,
    input_path: str | os.PathLike[str] | None = None,
) -> Path:
    """Validate a workbook destination without creating anything."""

    path = Path(output_path)
    parent = path.parent
    if not parent.exists():
        raise WorkbookError(
            f"Output parent directory does not exist: {parent}",
            output_path=path,
        )
    if not parent.is_dir():
        raise WorkbookError(
            f"Output parent path is not a directory: {parent}",
            output_path=path,
        )
    if input_path is not None:
        source = Path(input_path)
        if _paths_refer_to_same_file(path, source):
            raise WorkbookError(
                f"Output path refers to the input file: {path}",
                output_path=path,
            )
    if os.path.lexists(path):
        if path.is_dir():
            raise WorkbookError(
                f"Output path is a directory: {path}",
                output_path=path,
            )
        if not overwrite:
            raise WorkbookError(
                f"Output file already exists (use --overwrite): {path}",
                output_path=path,
            )
    return path


def _selected_plans(
    plans: Sequence[DatasetPlan],
) -> tuple[DatasetPlan, ...]:
    selected = tuple(
        plan
        for plan in plans
        if plan.selected and plan.exportable
    )
    if not selected:
        raise WorkbookError(
            "No selected exportable datasets to write"
        )
    pointers = [plan.pointer for plan in selected]
    if len(set(pointers)) != len(pointers):
        raise WorkbookError(
            "Selected datasets contain duplicate dataset pointers"
        )
    for plan in selected:
        if not plan.export_validated:
            raise WorkbookError(
                f"Dataset plan {plan.pointer!r} has not been validated "
                f"for export",
                dataset_pointer=plan.pointer,
            )
        if plan.validated_rows is None:
            raise WorkbookError(
                f"Dataset plan {plan.pointer!r} has no validated export "
                f"snapshot",
                dataset_pointer=plan.pointer,
            )
        if not plan.effective_columns:
            raise WorkbookError(
                f"Dataset plan {plan.pointer!r} has no effective columns",
                dataset_pointer=plan.pointer,
            )
    return selected


def _create_temp_file(output_path: Path) -> tuple[Path, BinaryIO, int]:
    flags = os.O_CREAT | os.O_EXCL | os.O_RDWR
    parent = output_path.parent
    for _ in range(_TEMP_FILE_ATTEMPTS):
        candidate = parent / f".jxcel-{secrets.token_hex(12)}.tmp.xlsx"
        try:
            descriptor = os.open(candidate, flags, 0o666)
        except FileExistsError:
            continue
        except OSError as exc:
            raise WorkbookError(
                f"Could not create a temporary workbook in {parent}: {exc}",
                output_path=output_path,
            ) from exc
        stream: BinaryIO | None = None
        try:
            creation_mode = stat.S_IMODE(os.fstat(descriptor).st_mode)
            if os.name == "posix":
                # Keep generated workbook contents private while the named
                # temporary file is being written and validated. Restore the
                # intended destination mode only after validation.
                os.fchmod(descriptor, 0o600)
            stream = os.fdopen(descriptor, "w+b")
        except BaseException as exc:
            try:
                if stream is not None:
                    stream.close()
                else:
                    os.close(descriptor)
            except BaseException:
                pass
            try:
                os.unlink(candidate)
            except BaseException:
                pass
            if isinstance(exc, OSError):
                raise WorkbookError(
                    f"Could not prepare the temporary workbook in {parent}: "
                    f"{exc}",
                    output_path=output_path,
                ) from exc
            raise
        return candidate, stream, creation_mode
    raise WorkbookError(
        f"Could not create a unique temporary workbook in {parent}",
        output_path=output_path,
    )


def _display_text(value: PreparedValue) -> str | None:
    if value.kind == "blank":
        return None
    if value.kind == "boolean":
        return "TRUE" if value.value else "FALSE"
    return str(value.value)


def _auto_filter_reference(plan: DatasetPlan) -> str:
    last_column = get_column_letter(len(plan.effective_columns))
    last_row = max(1 + plan.row_count, 1)
    return f"A1:{last_column}{last_row}"


def _write_workbook(
    plans: Sequence[DatasetPlan],
    assignments: Sequence[SheetAssignment],
    target: Path | BinaryIO,
    *,
    reported_path: Path | None = None,
) -> None:
    display_path = (
        target if isinstance(target, Path) else Path("<temporary workbook>")
    ) if reported_path is None else reported_path
    assignment_by_pointer = {
        assignment.pointer: assignment
        for assignment in assignments
    }
    workbook = Workbook()
    workbook.remove(workbook.active)
    try:
        for plan in plans:
            assignment = assignment_by_pointer.get(plan.pointer)
            if assignment is None:
                raise WorkbookError(
                    f"No worksheet assignment for dataset {plan.pointer!r}",
                    output_path=display_path,
                    dataset_pointer=plan.pointer,
                )
            worksheet = workbook.create_sheet(title=assignment.name)
            headers = [column.name for column in plan.effective_columns]
            worksheet.append(headers)
            for cell in worksheet[1]:
                cell.data_type = "s"
                cell.font = _HEADER_FONT
                cell.fill = _HEADER_FILL
                cell.alignment = _HEADER_ALIGNMENT

            widths = [len(header) for header in headers]
            sampled_rows = 0
            for row_number, prepared in enumerate(
                prepared_rows(plan),
                start=2,
            ):
                worksheet.append([value.value for value in prepared])
                for column_number, value in enumerate(prepared, start=1):
                    if (
                        value.kind == "text"
                        and isinstance(value.value, str)
                        and (
                            value.value.startswith("=")
                            or value.value in _EXCEL_ERROR_LITERALS
                        )
                    ):
                        worksheet.cell(
                            row=row_number,
                            column=column_number,
                        ).data_type = "s"
                if sampled_rows < WIDTH_SAMPLE_ROWS:
                    for column_index, value in enumerate(prepared):
                        text = _display_text(value)
                        if text is not None and len(text) > widths[column_index]:
                            widths[column_index] = len(text)
                    sampled_rows += 1

            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = _auto_filter_reference(plan)
            for column_index, width in enumerate(widths, start=1):
                worksheet.column_dimensions[
                    get_column_letter(column_index)
                ].width = min(
                    max(width + 2, MIN_COLUMN_WIDTH),
                    MAX_COLUMN_WIDTH,
                )
        workbook.save(target)
    except WorkbookError:
        raise
    except Exception as exc:
        raise WorkbookError(
            f"Could not write workbook to {display_path}: {exc}",
            output_path=display_path,
        ) from exc
    finally:
        try:
            workbook.close()
        except Exception:
            pass


def _cell_error(
    path: Path,
    assignment: SheetAssignment,
    row: int,
    column: int,
    message: str,
) -> WorkbookError:
    coordinate = f"{get_column_letter(column)}{row}"
    return WorkbookError(
        f"Workbook validation failed for dataset {assignment.pointer!r}, "
        f"worksheet {assignment.name!r}, cell {coordinate}: {message}",
        output_path=path,
        dataset_pointer=assignment.pointer,
        sheet_name=assignment.name,
        coordinate=coordinate,
    )


def _validate_text_cell(
    cell_value: object,
    cell_type: str | None,
    expected: str,
    path: Path,
    assignment: SheetAssignment,
    row: int,
    column: int,
) -> None:
    if cell_type == "f":
        raise _cell_error(
            path,
            assignment,
            row,
            column,
            "expected text but found a formula",
        )
    if expected == "":
        if cell_value not in (None, ""):
            raise _cell_error(
                path,
                assignment,
                row,
                column,
                f"expected blank text, found {cell_value!r}",
            )
        return
    if cell_type not in ("s", "inlineStr") or cell_value != expected:
        raise _cell_error(
            path,
            assignment,
            row,
            column,
            f"expected text {expected!r}, found type {cell_type!r} "
            f"and value {cell_value!r}",
        )


def _validate_prepared_cell(
    cell_value: object,
    cell_type: str | None,
    expected: PreparedValue,
    path: Path,
    assignment: SheetAssignment,
    row: int,
    column: int,
) -> None:
    if cell_type == "f":
        raise _cell_error(
            path,
            assignment,
            row,
            column,
            "expected prepared data but found a formula",
        )
    if expected.kind == "blank":
        if cell_value not in (None, ""):
            raise _cell_error(
                path,
                assignment,
                row,
                column,
                f"expected a blank cell, found {cell_value!r}",
            )
        return
    if expected.kind == "text":
        _validate_text_cell(
            cell_value,
            cell_type,
            expected.value,
            path,
            assignment,
            row,
            column,
        )
        return
    if expected.kind == "boolean":
        if cell_type != "b" or not isinstance(cell_value, bool):
            raise _cell_error(
                path,
                assignment,
                row,
                column,
                f"expected a boolean, found type {cell_type!r} "
                f"and value {cell_value!r}",
            )
        if cell_value != expected.value:
            raise _cell_error(
                path,
                assignment,
                row,
                column,
                f"expected boolean {expected.value!r}, "
                f"found {cell_value!r}",
            )
        return
    if expected.kind in ("integer", "float"):
        if (
            cell_type != "n"
            or isinstance(cell_value, bool)
            or not isinstance(cell_value, (int, float))
        ):
            raise _cell_error(
                path,
                assignment,
                row,
                column,
                f"expected a numeric value, found type {cell_type!r} "
                f"and value {cell_value!r}",
            )
        accepted = (expected.value,)
        if expected.kind == "float":
            accepted += (float(format(expected.value, ".16g")),)
        if cell_value not in accepted:
            raise _cell_error(
                path,
                assignment,
                row,
                column,
                f"expected numeric value {expected.value!r}, "
                f"found {cell_value!r}",
            )
        return
    raise _cell_error(
        path,
        assignment,
        row,
        column,
        f"unsupported prepared value kind {expected.kind!r}",
    )


def _validate_workbook(
    target: Path | BinaryIO,
    plans: Sequence[DatasetPlan],
    assignments: Sequence[SheetAssignment],
    *,
    reported_path: Path | None = None,
) -> None:
    display_path = (
        target if isinstance(target, Path) else Path("<temporary workbook>")
    ) if reported_path is None else reported_path
    try:
        if hasattr(target, "seek"):
            target.seek(0)
        workbook = load_workbook(target, data_only=False)
    except Exception as exc:
        raise WorkbookError(
            f"Could not reopen generated workbook {display_path}: {exc}",
            output_path=display_path,
        ) from exc
    try:
        expected_names = tuple(
            assignment.name for assignment in assignments
        )
        actual_names = tuple(workbook.sheetnames)
        if actual_names != expected_names:
            raise WorkbookError(
                f"Workbook validation failed for {display_path}: expected "
                f"worksheets {expected_names!r}, found {actual_names!r}",
                output_path=display_path,
            )

        for plan, assignment in zip(plans, assignments):
            worksheet = workbook[assignment.name]
            worksheet_max_row = worksheet.max_row
            worksheet_max_column = worksheet.max_column
            headers = [
                column.name for column in plan.effective_columns
            ]
            if worksheet_max_column != len(headers):
                raise WorkbookError(
                    f"Workbook validation failed for worksheet "
                    f"{assignment.name!r}: expected {len(headers)} columns, "
                    f"found {worksheet_max_column}",
                    output_path=display_path,
                    dataset_pointer=plan.pointer,
                    sheet_name=assignment.name,
                )
            planned_last_row = 1 + plan.row_count
            if worksheet_max_row > planned_last_row:
                raise WorkbookError(
                    f"Workbook validation failed for worksheet "
                    f"{assignment.name!r}: found {worksheet_max_row} rows, "
                    f"but the plan has only {planned_last_row}",
                    output_path=display_path,
                    dataset_pointer=plan.pointer,
                    sheet_name=assignment.name,
                )
            if worksheet.freeze_panes != "A2":
                raise WorkbookError(
                    f"Workbook validation failed for worksheet "
                    f"{assignment.name!r}: expected frozen header row, "
                    f"found {worksheet.freeze_panes!r}",
                    output_path=display_path,
                    dataset_pointer=plan.pointer,
                    sheet_name=assignment.name,
                )
            expected_filter = _auto_filter_reference(plan)
            if worksheet.auto_filter.ref != expected_filter:
                raise WorkbookError(
                    f"Workbook validation failed for worksheet "
                    f"{assignment.name!r}: expected auto-filter "
                    f"{expected_filter!r}, found "
                    f"{worksheet.auto_filter.ref!r}",
                    output_path=display_path,
                    dataset_pointer=plan.pointer,
                    sheet_name=assignment.name,
                )

            for column_number, header in enumerate(headers, start=1):
                cell = worksheet.cell(row=1, column=column_number)
                _validate_text_cell(
                    cell.value,
                    cell.data_type,
                    header,
                    display_path,
                    assignment,
                    1,
                    column_number,
                )

            for row_number, prepared in enumerate(
                prepared_rows(plan),
                start=2,
            ):
                for column_number, expected in enumerate(
                    prepared,
                    start=1,
                ):
                    if (
                        row_number <= worksheet_max_row
                        and column_number <= worksheet_max_column
                    ):
                        cell = worksheet.cell(
                            row=row_number,
                            column=column_number,
                        )
                        cell_value = cell.value
                        cell_type = cell.data_type
                    else:
                        cell_value = None
                        cell_type = None
                    _validate_prepared_cell(
                        cell_value,
                        cell_type,
                        expected,
                        display_path,
                        assignment,
                        row_number,
                        column_number,
                    )

            for row in worksheet.iter_rows(
                max_row=worksheet_max_row,
                max_col=worksheet_max_column,
            ):
                for cell in row:
                    if cell.data_type == "f":
                        raise _cell_error(
                            display_path,
                            assignment,
                            cell.row,
                            cell.column,
                            "formula cells are not permitted",
                        )
    finally:
        try:
            workbook.close()
        except Exception:
            pass


def _apply_output_permissions(
    temp_path: Path,
    output_path: Path,
    overwrite: bool,
    creation_mode: int,
    temp_file: BinaryIO | None,
) -> None:
    mode = creation_mode
    if overwrite:
        try:
            existing = os.lstat(output_path)
        except FileNotFoundError:
            existing = None
        except OSError as exc:
            raise WorkbookError(
                f"Could not inspect existing output permissions for "
                f"{output_path}: {exc}",
                output_path=output_path,
            ) from exc
        if existing is not None and stat.S_ISREG(existing.st_mode):
            mode = stat.S_IMODE(existing.st_mode) & 0o777
        elif os.name != "posix":
            return
    if os.name == "posix":
        if temp_file is None:
            raise WorkbookError(
                "Temporary workbook handle is unavailable for permission "
                f"preparation at {output_path}",
                output_path=output_path,
            )
        try:
            os.fchmod(temp_file.fileno(), mode)
        except OSError as exc:
            raise WorkbookError(
                f"Could not prepare permissions for output {output_path}: "
                f"{exc}. The existing output was not replaced.",
                output_path=output_path,
            ) from exc
        return
    if not overwrite or existing is None or not stat.S_ISREG(existing.st_mode):
        return
    try:
        os.chmod(temp_path, stat.S_IMODE(existing.st_mode) & 0o777)
    except OSError as exc:
        raise WorkbookError(
            f"Could not preserve permissions for existing output "
            f"{output_path}: {exc}. The existing output was not replaced.",
            output_path=output_path,
        ) from exc


def _publish_without_overwrite(temp_path: Path, output_path: Path) -> None:
    try:
        os.link(temp_path, output_path)
    except FileExistsError as exc:
        raise WorkbookError(
            f"Output file already exists (use --overwrite): {output_path}",
            output_path=output_path,
        ) from exc
    except OSError as exc:
        if exc.errno == errno.EEXIST:
            raise WorkbookError(
                f"Output file already exists (use --overwrite): "
                f"{output_path}",
                output_path=output_path,
            ) from exc
        raise WorkbookError(
            f"Could not publish workbook to {output_path} without "
            f"overwriting: {exc}. No fallback publication method was "
            f"attempted.",
            output_path=output_path,
        ) from exc


def _publish_with_overwrite(temp_path: Path, output_path: Path) -> None:
    try:
        os.replace(temp_path, output_path)
    except OSError as exc:
        raise WorkbookError(
            f"Could not replace existing workbook at {output_path}: {exc}",
            output_path=output_path,
        ) from exc


def _cleanup_temp_file(path: Path) -> str | None:
    try:
        os.unlink(path)
    except FileNotFoundError:
        return None
    except OSError as exc:
        return f"Temporary file {path} could not be removed: {exc}"
    return None


def export_workbook(
    plans: Sequence[DatasetPlan],
    output_path: str | os.PathLike[str],
    *,
    overwrite: bool = False,
    input_path: str | os.PathLike[str] | None = None,
) -> WorkbookExport:
    """Write, validate, and safely publish a workbook."""

    selected = _selected_plans(plans)
    assignments = assign_sheet_names(selected)
    path = validate_output_destination(
        output_path,
        overwrite=overwrite,
        input_path=input_path,
    )

    temp_path: Path | None = None
    temp_file: BinaryIO | None = None
    try:
        temp_path, temp_file, creation_mode = _create_temp_file(path)
        _write_workbook(
            selected,
            assignments,
            temp_file,
            reported_path=path,
        )
        temp_file.flush()
        _validate_workbook(
            temp_file,
            selected,
            assignments,
            reported_path=path,
        )
        temp_file.flush()
        if os.name != "posix":
            temp_file.close()
            temp_file = None
        _apply_output_permissions(
            temp_path,
            path,
            overwrite,
            creation_mode,
            temp_file,
        )
        if temp_file is not None:
            temp_file.close()
            temp_file = None

        if overwrite:
            _publish_with_overwrite(temp_path, path)
            temp_path = None
            cleanup_warning = None
        else:
            _publish_without_overwrite(temp_path, path)
            cleanup_warning = _cleanup_temp_file(temp_path)
            temp_path = None

        summaries = tuple(
            SheetSummary(
                pointer=assignment.pointer,
                name=assignment.name,
                data_rows=plan.row_count,
                columns=len(plan.effective_columns),
            )
            for plan, assignment in zip(selected, assignments)
        )
        return WorkbookExport(
            output_path=path,
            sheets=summaries,
            cleanup_warning=cleanup_warning,
        )
    except BaseException as exc:
        if temp_file is not None:
            try:
                temp_file.close()
            except BaseException:
                pass
        cleanup_error = None
        if temp_path is not None:
            cleanup_error = _cleanup_temp_file(temp_path)
        if not isinstance(exc, Exception):
            if cleanup_error is not None:
                exc.add_note(cleanup_error)
            raise
        if cleanup_error is not None:
            raise WorkbookError(
                f"{exc}; additionally, {cleanup_error}",
                output_path=path,
            ) from exc
        if isinstance(exc, WorkbookError):
            raise
        raise WorkbookError(
            f"Could not export workbook to {path}: {exc}",
            output_path=path,
        ) from exc
