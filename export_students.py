#!/usr/bin/env python3
"""Export selected student fields from full_data.json to an Excel workbook."""

import argparse
import json
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

COLUMNS = [
    "id",
    "name",
    "email",
    "phone",
    "track",
    "name_ar",
    "group_id",
    "provider",
]

# Fields that must stay text so leading zeros survive and no scientific
# notation is applied by Excel.
TEXT_COLUMNS = {"id", "phone"}

CENTERED_COLUMNS = {"id", "phone", "group_id", "provider"}

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)


def parse_args():
    parser = argparse.ArgumentParser(
        description="Export student fields from a JSON file to Excel."
    )
    parser.add_argument(
        "--input",
        default="full_data.json",
        help="Path to the input JSON file (default: full_data.json)",
    )
    parser.add_argument(
        "--output",
        default="DEPI_Students.xlsx",
        help="Path to the output Excel file (default: DEPI_Students.xlsx)",
    )
    return parser.parse_args()


def load_students(input_path):
    """Read the JSON file and return its top-level students list."""
    path = Path(input_path)
    if not path.is_file():
        raise FileNotFoundError(f"Input file not found: {path}")

    try:
        with path.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
    except json.JSONDecodeError as exc:
        raise ValueError(
            f"Invalid JSON in {path}: line {exc.lineno}, column {exc.colno}: {exc.msg}"
        ) from exc
    except OSError as exc:
        raise OSError(f"Could not read {path}: {exc}") from exc

    if not isinstance(data, dict):
        raise ValueError(f"Invalid JSON structure in {path}: expected a top-level object")

    students = data.get("students")
    if students is None:
        raise ValueError(f"Missing 'students' array in {path}")
    if not isinstance(students, list):
        raise ValueError(
            f"Invalid 'students' value in {path}: expected an array, "
            f"got {type(students).__name__}"
        )

    return students


def cell_value(student, field):
    """Return the raw value for a field, or None when absent/null."""
    if not isinstance(student, dict):
        return None
    value = student.get(field)
    return None if value is None else value


def write_workbook(students, output_path):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Students"

    worksheet.append(COLUMNS)

    for cell in worksheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGNMENT
        cell.border = THIN_BORDER

    for student in students:
        row = []
        for field in COLUMNS:
            value = cell_value(student, field)
            if value is None:
                row.append(None)
            elif field in TEXT_COLUMNS:
                row.append(str(value))
            else:
                row.append(value)
        worksheet.append(row)

    last_row = worksheet.max_row

    # Format every data cell (text fields as text) and apply alignment.
    for row_index in range(2, last_row + 1):
        for column_index, field in enumerate(COLUMNS, start=1):
            cell = worksheet.cell(row=row_index, column=column_index)
            cell.border = THIN_BORDER
            horizontal = "center" if field in CENTERED_COLUMNS else "left"
            cell.alignment = Alignment(horizontal=horizontal, vertical="center")
            if field in TEXT_COLUMNS:
                cell.number_format = "@"

    # Column widths based on the longest displayed value (header included).
    for column_index, field in enumerate(COLUMNS, start=1):
        longest = len(field)
        for row_index in range(2, last_row + 1):
            value = worksheet.cell(row=row_index, column=column_index).value
            if value is not None:
                longest = max(longest, len(str(value)))
        letter = get_column_letter(column_index)
        worksheet.column_dimensions[letter].width = min(max(longest + 2, 10), 60)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{last_row}"
    worksheet.row_dimensions[1].height = 24

    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output)
    return output


def validate(output_path, expected_rows):
    """Reopen the workbook and verify structure and row count."""
    workbook = load_workbook(output_path)
    try:
        if workbook.sheetnames != ["Students"]:
            raise ValueError(
                f"Expected exactly one worksheet named 'Students', "
                f"found: {workbook.sheetnames}"
            )

        worksheet = workbook["Students"]
        header = [cell.value for cell in worksheet[1]][: len(COLUMNS)]
        if header != COLUMNS:
            raise ValueError(f"Unexpected header row: {header}")

        extra_headers = [
            cell.value for cell in worksheet[1][len(COLUMNS):] if cell.value is not None
        ]
        if extra_headers:
            raise ValueError(f"Unexpected extra columns: {extra_headers}")

        data_rows = worksheet.max_row - 1
        if data_rows != expected_rows:
            raise ValueError(
                f"Row count mismatch: exported {data_rows}, expected {expected_rows}"
            )
    finally:
        workbook.close()

    return header, data_rows


def main():
    args = parse_args()

    try:
        students = load_students(args.input)
    except (FileNotFoundError, ValueError, OSError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    try:
        output_path = write_workbook(students, args.output)
        header, data_rows = validate(output_path, len(students))
    except (ValueError, OSError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print("Export completed successfully.")
    print(f"Students exported: {data_rows}")
    print(f"Columns ({len(header)}): {', '.join(header)}")
    print(f"Output file: {output_path.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
