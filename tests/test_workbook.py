import errno
import multiprocessing
import os
import stat
import tempfile
import threading
import unittest
import zipfile
from dataclasses import replace
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

from jxcel import workbook as workbook_module
from jxcel.discovery import DatasetCandidate
from jxcel.planning import (
    EXCEL_MAX_CELL_UTF16_UNITS,
    ColumnPlan,
    DatasetPlan,
    build_plan,
)
from jxcel.workbook import (
    EXCEL_MAX_SHEET_NAME_UTF16_UNITS,
    SheetAssignment,
    WorkbookError,
    assign_sheet_names,
    export_workbook,
    validate_output_destination,
)


def make_plans(specs, **kwargs):
    candidates = [
        DatasetCandidate(pointer=pointer, records=tuple(records))
        for pointer, records in specs
    ]
    result = build_plan(candidates, **kwargs)
    return result.datasets


def make_plan(pointer="/data", records=None, **kwargs):
    if records is None:
        records = ({"a": 1},)
    return make_plans([(pointer, records)], **kwargs)[0]


def _competing_export_worker(directory, worker, rounds, barrier, results):
    for index in range(rounds):
        barrier.wait(timeout=30)
        output = Path(directory) / f"concurrent-{index}.xlsx"
        plan = make_plan(records=({"a": f"worker-{worker}-round-{index}"},))
        try:
            export_workbook([plan], output)
        except WorkbookError as exc:
            result = "exists" if "already exists" in str(exc) else f"error: {exc}"
        else:
            result = "published"
        results.put((index, worker, result))
        barrier.wait(timeout=30)
        barrier.wait(timeout=30)


def naming_plan(pointer, *, selected=True, exportable=True):
    column = ColumnPlan(name="x", source_path=("x",))
    return DatasetPlan(
        pointer=pointer,
        records=({"x": 1},),
        columns=(column,),
        effective_columns=(column,) if exportable else (),
        selected=selected,
        non_exportable_reason=None if exportable else "zero_columns",
        warnings=(),
        flatten=True,
        export_validated=True,
    )


def utf16_length(text):
    return sum(2 if ord(character) > 0xFFFF else 1 for character in text)


def sheet_names(plans):
    return [assignment.name for assignment in assign_sheet_names(plans)]


class SheetNameTests(unittest.TestCase):
    def test_root_and_nested_pointer_names(self):
        plans = [
            naming_plan(""),
            naming_plan("/students"),
            naming_plan("/departments/0/employees"),
        ]
        self.assertEqual(
            sheet_names(plans),
            ["root", "students", "departments.0.employees"],
        )

    def test_pointer_escaping_and_invalid_characters(self):
        plans = [
            naming_plan("/a~1b"),
            naming_plan("/a~0b"),
            naming_plan("/a?b:c*[d]"),
            naming_plan("/a b"),
        ]
        self.assertEqual(
            sheet_names(plans),
            ["a_b", "a~b", "a_b_c__d_", "a b"],
        )

    def test_empty_sanitized_names_become_dataset(self):
        plans = [
            naming_plan("/"),
            naming_plan("/   "),
            naming_plan("/\u00a0"),
        ]
        self.assertEqual(
            sheet_names(plans),
            ["dataset", "dataset_2", "dataset_3"],
        )

    def test_history_is_reserved_case_insensitively(self):
        plans = [
            naming_plan("/History"),
            naming_plan("/history"),
        ]
        names = sheet_names(plans)
        self.assertEqual(names[0], "History_")
        self.assertEqual(names[1], "history__2")
        self.assertEqual(len(set(name.casefold() for name in names)), 2)

    def test_case_insensitive_collisions(self):
        plans = [
            naming_plan("/Users"),
            naming_plan("/users"),
            naming_plan("/USERS"),
        ]
        self.assertEqual(
            sheet_names(plans),
            ["Users", "users_2", "USERS_3"],
        )

    def test_sanitization_collisions(self):
        plans = [
            naming_plan("/a~1b"),
            naming_plan("/a?b"),
        ]
        self.assertEqual(sheet_names(plans), ["a_b", "a_b_2"])

    def test_arabic_sheet_name_is_preserved(self):
        self.assertEqual(sheet_names([naming_plan("/أحمد")]), ["أحمد"])

    def test_names_are_truncated_to_utf16_units(self):
        long_ascii = naming_plan("/" + "a" * 40)
        emoji = naming_plan("/" + "😀" * 31)
        names = sheet_names([long_ascii, emoji])
        self.assertEqual(names[0], "a" * 31)
        self.assertEqual(names[1], "😀" * 15)
        self.assertLessEqual(utf16_length(names[0]), 31)
        self.assertLessEqual(utf16_length(names[1]), 31)

    def test_collisions_after_truncation_reserve_suffix_space(self):
        plans = [
            naming_plan("/" + "a" * 40),
            naming_plan("/" + "a" * 39 + "b"),
        ]
        names = sheet_names(plans)
        self.assertEqual(names[0], "a" * 31)
        self.assertEqual(names[1], "a" * 29 + "_2")
        self.assertLessEqual(utf16_length(names[1]), 31)

    def test_whitespace_apostrophes_and_controls_are_sanitized(self):
        plans = [
            naming_plan("/ 'name' "),
            naming_plan("/''"),
            naming_plan("/a\nb"),
            naming_plan("/a\x00b"),
            naming_plan("/a\ud800b"),
        ]
        self.assertEqual(
            sheet_names(plans),
            ["_name_", "__", "a_b", "a_b_2", "a_b_3"],
        )

    def test_unselected_and_non_exportable_plans_are_skipped(self):
        plans = [
            naming_plan("/keep"),
            naming_plan("/unselected", selected=False),
            naming_plan("/empty", exportable=False),
            naming_plan("/also-keep"),
        ]
        self.assertEqual(sheet_names(plans), ["keep", "also-keep"])

    def test_all_assigned_names_are_valid_and_unique(self):
        plans = [
            naming_plan("/Users"),
            naming_plan("/users"),
            naming_plan("/History"),
            naming_plan("/a~1b"),
            naming_plan("/a?b"),
            naming_plan("/" + "x" * 40),
            naming_plan("/" + "😀" * 31),
        ]
        names = sheet_names(plans)
        folded = [name.casefold() for name in names]
        self.assertEqual(len(folded), len(set(folded)))
        for name in names:
            self.assertTrue(name)
            self.assertLessEqual(utf16_length(name), 31)
            self.assertFalse(any(character in "\\*?:/[]" for character in name))
            self.assertFalse(name.startswith("'"))
            self.assertFalse(name.endswith("'"))


class WorkbookContentTests(unittest.TestCase):
    def setUp(self):
        self._temporary_directory = tempfile.TemporaryDirectory()
        self.addCleanup(self._temporary_directory.cleanup)
        self.directory = Path(self._temporary_directory.name)

    def output(self, name="out.xlsx"):
        return self.directory / name

    def test_writes_worksheets_in_plan_order_and_maps_names(self):
        plans = make_plans(
            [
                ("/first", ({"name": "alpha", "count": 1},)),
                ("/second", ({"name": "beta", "count": 2},)),
            ]
        )
        result = export_workbook(plans, self.output())

        self.assertEqual(
            [summary.pointer for summary in result.sheets],
            ["/first", "/second"],
        )
        self.assertEqual(
            [summary.name for summary in result.sheets],
            ["first", "second"],
        )
        self.assertEqual(
            [(summary.data_rows, summary.columns) for summary in result.sheets],
            [(1, 2), (1, 2)],
        )

        workbook = load_workbook(self.output())
        try:
            self.assertEqual(workbook.sheetnames, ["first", "second"])
            first = workbook["first"]
            self.assertEqual(
                [cell.value for cell in first[1]],
                ["name", "count"],
            )
            self.assertEqual(
                [cell.value for cell in first[2]],
                ["alpha", 1],
            )
        finally:
            workbook.close()

    def test_headers_and_prepared_value_categories_round_trip(self):
        records = (
            {
                "text": "hello",
                "integer": 12,
                "float": 1.0,
                "other_float": 1.5,
                "boolean": True,
                "blank": None,
                "empty": "",
                "nested": {"a": 1},
                "array": [1, 2],
                "arabic": "مرحبا",
                "emoji": "😀",
            },
        )
        plan = make_plan("/data", records)
        export_workbook([plan], self.output())

        workbook = load_workbook(self.output())
        try:
            worksheet = workbook["data"]
            self.assertEqual(
                [cell.value for cell in worksheet[1]],
                [
                    "text",
                    "integer",
                    "float",
                    "other_float",
                    "boolean",
                    "blank",
                    "empty",
                    "nested.a",
                    "array",
                    "arabic",
                    "emoji",
                ],
            )
            row = list(worksheet[2])
            self.assertEqual(row[0].value, "hello")
            self.assertEqual(row[0].data_type, "s")
            self.assertEqual(row[1].value, 12)
            self.assertEqual(row[1].data_type, "n")
            self.assertEqual(row[2].value, 1)
            self.assertEqual(row[2].data_type, "n")
            self.assertEqual(row[3].value, 1.5)
            self.assertEqual(row[3].data_type, "n")
            self.assertIs(row[4].value, True)
            self.assertEqual(row[4].data_type, "b")
            self.assertIsNone(row[5].value)
            self.assertIsNone(row[6].value)
            self.assertEqual(row[7].value, 1)
            self.assertEqual(row[8].value, "[1,2]")
            self.assertEqual(row[9].value, "مرحبا")
            self.assertEqual(row[10].value, "😀")
        finally:
            workbook.close()

    def test_empty_header_is_preserved_as_a_blank_cell(self):
        plan = make_plan("/data", ({"": "value"},))
        export_workbook([plan], self.output())

        workbook = load_workbook(self.output())
        try:
            worksheet = workbook["data"]
            self.assertIsNone(worksheet["A1"].value)
            self.assertEqual(worksheet["A2"].value, "value")
            self.assertEqual(worksheet.max_column, 1)
        finally:
            workbook.close()

    def test_empty_text_and_blank_values_are_equivalent_after_reload(self):
        plan = make_plan(
            "/data",
            ({"a": None, "b": "", "c": "x"},),
        )
        export_workbook([plan], self.output())

        workbook = load_workbook(self.output())
        try:
            worksheet = workbook["data"]
            self.assertIsNone(worksheet["A2"].value)
            self.assertIsNone(worksheet["B2"].value)
            self.assertEqual(worksheet["C2"].value, "x")
            self.assertEqual(worksheet.max_row, 2)
        finally:
            workbook.close()

    def test_trailing_all_blank_records_are_tolerated(self):
        plan = make_plan(
            "/data",
            ({"a": 1}, {"a": None}, {"a": None}),
        )
        result = export_workbook([plan], self.output())
        self.assertEqual(result.sheets[0].data_rows, 3)

        workbook = load_workbook(self.output())
        try:
            worksheet = workbook["data"]
            self.assertEqual(worksheet.max_row, 2)
            self.assertEqual(worksheet.auto_filter.ref, "A1:A4")
            self.assertEqual(worksheet["A2"].value, 1)
        finally:
            workbook.close()

    def test_large_integer_is_written_as_text(self):
        plan = make_plan("/data", ({"n": 10**15},))
        export_workbook([plan], self.output())

        workbook = load_workbook(self.output())
        try:
            cell = workbook["data"]["A2"]
            self.assertIn(cell.data_type, ("s", "inlineStr"))
            self.assertEqual(cell.value, "1000000000000000")
        finally:
            workbook.close()

    def test_cell_at_excel_text_limit_round_trips_without_truncation(self):
        value = "x" * EXCEL_MAX_CELL_UTF16_UNITS
        plan = make_plan("/data", ({"a": value},))
        export_workbook([plan], self.output())

        workbook = load_workbook(self.output())
        try:
            self.assertEqual(workbook["data"]["A2"].value, value)
        finally:
            workbook.close()

    def test_no_flatten_uses_compact_json_text(self):
        plans = make_plans(
            [
                (
                    "/data",
                    (
                        {
                            "obj": {"a": 1},
                            "arr": [1, 2],
                            "empty": {},
                        },
                    ),
                )
            ],
            flatten=False,
        )
        export_workbook(plans, self.output())

        workbook = load_workbook(self.output())
        try:
            worksheet = workbook["data"]
            self.assertEqual(
                [cell.value for cell in worksheet[1]],
                ["obj", "arr", "empty"],
            )
            self.assertEqual(
                [cell.value for cell in worksheet[2]],
                ['{"a":1}', "[1,2]", "{}"],
            )
        finally:
            workbook.close()

    def test_formula_like_text_and_headers_remain_text(self):
        records = (
            {
                "=HEADER": "=1+1",
                "plus": "+1",
                "minus": "-1",
                "at": "@SUM(A1)",
                "tab": "\t=1+1",
                "space": " =1+1",
                "apostrophe": "'=1+1",
                "equals": "==",
                "formula": "=SUM(A1)",
            },
        )
        plan = make_plan("/formulas", records)
        export_workbook([plan], self.output())

        workbook = load_workbook(self.output())
        try:
            worksheet = workbook["formulas"]
            self.assertEqual(worksheet["A1"].value, "=HEADER")
            self.assertEqual(worksheet["A1"].data_type, "s")
            self.assertEqual(worksheet["A2"].value, "=1+1")
            self.assertEqual(worksheet["A2"].data_type, "s")
            self.assertEqual(worksheet["B2"].value, "+1")
            self.assertEqual(worksheet["C2"].value, "-1")
            self.assertEqual(worksheet["D2"].value, "@SUM(A1)")
            self.assertEqual(worksheet["E2"].value, "\t=1+1")
            self.assertEqual(worksheet["F2"].value, " =1+1")
            self.assertEqual(worksheet["G2"].value, "'=1+1")
            self.assertEqual(worksheet["H2"].value, "==")
            self.assertEqual(worksheet["H2"].data_type, "s")
            self.assertEqual(worksheet["I2"].value, "=SUM(A1)")
            for row in worksheet.iter_rows():
                for cell in row:
                    self.assertNotEqual(cell.data_type, "f")
        finally:
            workbook.close()

        with zipfile.ZipFile(self.output()) as archive:
            sheet_xml = [
                archive.read(name).decode("utf-8")
                for name in archive.namelist()
                if name.startswith("xl/worksheets/")
                and name.endswith(".xml")
            ]
        self.assertTrue(sheet_xml)
        self.assertFalse(any("<f" in xml for xml in sheet_xml))

    def test_restrained_formatting_is_applied(self):
        plan = make_plan(
            "/data",
            ({"long": "x" * 100, "short": "y", "empty": ""},),
        )
        export_workbook([plan], self.output())

        workbook = load_workbook(self.output())
        try:
            worksheet = workbook["data"]
            for cell in worksheet[1]:
                self.assertTrue(cell.font.bold)
                self.assertTrue(cell.fill.fgColor.rgb.endswith("D9E1F2"))
                self.assertEqual(cell.alignment.vertical, "center")
            self.assertEqual(worksheet.freeze_panes, "A2")
            self.assertEqual(worksheet.auto_filter.ref, "A1:C2")
            self.assertEqual(worksheet.column_dimensions["A"].width, 60)
            self.assertEqual(worksheet.column_dimensions["B"].width, 10)
            self.assertEqual(worksheet.column_dimensions["C"].width, 10)
        finally:
            workbook.close()

    def test_unselected_and_non_exportable_datasets_are_skipped(self):
        result = build_plan(
            [
                DatasetCandidate("/keep", ({"a": 1},)),
                DatasetCandidate("/skip", ({"b": 2},)),
                DatasetCandidate("/empty", ({},)),
            ],
            selected_pointers=["/keep", "/empty"],
        )
        export = export_workbook(result.datasets, self.output())

        self.assertEqual([summary.pointer for summary in export.sheets], ["/keep"])
        workbook = load_workbook(self.output())
        try:
            self.assertEqual(workbook.sheetnames, ["keep"])
        finally:
            workbook.close()

    def test_no_exportable_dataset_fails_before_creating_files(self):
        plans = make_plans(
            [
                ("/unselected", ({"a": 1},)),
                ("/empty", ({},)),
            ],
            selected_pointers=["/empty"],
        )
        with self.assertRaisesRegex(
            WorkbookError,
            "No selected exportable datasets",
        ):
            export_workbook(plans, self.output())
        self.assertFalse(self.output().exists())
        self.assertEqual(list(self.directory.glob(".jxcel-*.tmp.xlsx")), [])

    def test_schema_only_plan_is_rejected(self):
        plan = make_plan("/data", validation_mode="schema")
        self.assertFalse(plan.export_validated)
        with self.assertRaisesRegex(
            WorkbookError,
            "has not been validated for export",
        ):
            export_workbook([plan], self.output())
        self.assertFalse(self.output().exists())
        self.assertEqual(list(self.directory.glob(".jxcel-*.tmp.xlsx")), [])


class ValidatedSnapshotExportTests(unittest.TestCase):
    def setUp(self):
        self._temporary_directory = tempfile.TemporaryDirectory()
        self.addCleanup(self._temporary_directory.cleanup)
        self.directory = Path(self._temporary_directory.name)

    def output(self, name="snapshot.xlsx"):
        return self.directory / name

    def temp_files(self):
        return sorted(self.directory.glob(".jxcel-*.tmp.xlsx"))

    def data_row(self, path):
        workbook = load_workbook(path)
        try:
            worksheet = workbook[workbook.sheetnames[0]]
            return [(cell.data_type, cell.value) for cell in worksheet[2]]
        finally:
            workbook.close()

    def test_top_level_record_mutation_does_not_publish_new_value(self):
        plan = make_plan("/data", ({"value": 1.0},))

        plan.records[0]["value"] = 1e308

        output = self.output()
        export_workbook([plan], output)
        self.assertEqual(self.data_row(output), [("n", 1.0)])

    def test_nested_dict_mutation_does_not_publish_new_value(self):
        plan = make_plan("/data", ({"outer": {"value": 1.0}},))

        plan.records[0]["outer"]["value"] = 1e308

        output = self.output()
        export_workbook([plan], output)
        self.assertEqual(self.data_row(output), [("n", 1.0)])

    def test_nested_list_mutation_does_not_publish_new_value(self):
        plan = make_plan("/data", ({"values": [1, 2]},))

        values = plan.records[0]["values"]
        values.append(3)
        values[0] = 1e308

        output = self.output()
        export_workbook([plan], output)
        self.assertEqual(self.data_row(output), [("s", "[1,2]")])

    def test_out_of_range_mutation_cannot_be_published(self):
        plan = make_plan("/data", ({"value": 1.0},))
        plan.records[0]["value"] = 1e308

        output = self.output()
        export_workbook([plan], output)

        cells = self.data_row(output)
        self.assertEqual(cells, [("n", 1.0)])
        self.assertNotIn(1e308, [value for _, value in cells])

    def test_validated_plan_without_snapshot_fails_before_publication(self):
        plan = replace(make_plan("/data", ({"a": 1},)), validated_rows=None)
        output = self.output()

        with self.assertRaisesRegex(
            WorkbookError,
            "no validated export snapshot",
        ):
            export_workbook([plan], output)

        self.assertFalse(output.exists())
        self.assertEqual(self.temp_files(), [])

    def test_normal_export_preserves_types_and_values(self):
        plan = make_plan(
            "/data",
            (
                {
                    "i": 1,
                    "f": 1.5,
                    "b": True,
                    "n": None,
                    "s": "text",
                    "big": 10**15,
                    "items": [1, 2],
                },
            ),
        )
        output = self.output()

        export_workbook([plan], output)

        workbook = load_workbook(output)
        try:
            worksheet = workbook["data"]
            self.assertEqual(worksheet["A2"].value, 1)
            self.assertEqual(worksheet["A2"].data_type, "n")
            self.assertEqual(worksheet["B2"].value, 1.5)
            self.assertEqual(worksheet["B2"].data_type, "n")
            self.assertIs(worksheet["C2"].value, True)
            self.assertEqual(worksheet["C2"].data_type, "b")
            self.assertIsNone(worksheet["D2"].value)
            self.assertEqual(worksheet["E2"].value, "text")
            self.assertEqual(worksheet["E2"].data_type, "s")
            self.assertEqual(worksheet["F2"].value, str(10**15))
            self.assertEqual(worksheet["F2"].data_type, "s")
            self.assertEqual(worksheet["G2"].value, "[1,2]")
            self.assertEqual(worksheet["G2"].data_type, "s")
        finally:
            workbook.close()

    def test_dimensions_and_summary_use_snapshot_row_count(self):
        plan = make_plan("/data", ({"a": 1}, {"a": 2}, {"a": 3}))
        changed = replace(plan, records=())

        output = self.output()
        result = export_workbook([changed], output)

        self.assertEqual(result.sheets[0].data_rows, 3)
        workbook = load_workbook(output)
        try:
            worksheet = workbook["data"]
            self.assertEqual(worksheet.max_row, 4)
            self.assertEqual(worksheet.auto_filter.ref, "A1:A4")
            self.assertEqual(
                [worksheet.cell(row=row, column=1).value for row in (2, 3, 4)],
                [1, 2, 3],
            )
        finally:
            workbook.close()

    def test_zero_row_dataset_fails_before_creating_files(self):
        plan = make_plan("/data", ())
        output = self.output()

        with self.assertRaisesRegex(
            WorkbookError,
            "No selected exportable datasets",
        ):
            export_workbook([plan], output)

        self.assertFalse(output.exists())
        self.assertEqual(self.temp_files(), [])


class WorkbookValidationTests(unittest.TestCase):
    def setUp(self):
        self._temporary_directory = tempfile.TemporaryDirectory()
        self.addCleanup(self._temporary_directory.cleanup)
        self.directory = Path(self._temporary_directory.name)

    def write_workbook(self, plan=None):
        plan = plan or make_plan("/data", ({"a": "value"},))
        assignments = assign_sheet_names([plan])
        path = self.directory / "validation.xlsx"
        workbook_module._write_workbook([plan], assignments, path)
        return plan, assignments, path

    def test_written_workbook_passes_validation(self):
        plan, assignments, path = self.write_workbook()
        workbook_module._validate_workbook(path, [plan], assignments)

    def test_validation_rejects_wrong_sheet_name(self):
        plan, assignments, path = self.write_workbook()
        workbook = load_workbook(path)
        workbook["data"].title = "renamed"
        workbook.save(path)
        workbook.close()
        with self.assertRaisesRegex(WorkbookError, "expected worksheets"):
            workbook_module._validate_workbook(path, [plan], assignments)

    def test_validation_rejects_wrong_header(self):
        plan, assignments, path = self.write_workbook()
        workbook = load_workbook(path)
        workbook["data"]["A1"] = "wrong"
        workbook.save(path)
        workbook.close()
        with self.assertRaisesRegex(WorkbookError, "expected text"):
            workbook_module._validate_workbook(path, [plan], assignments)

    def test_validation_error_reports_requested_output_path(self):
        plan, assignments, path = self.write_workbook()
        workbook = load_workbook(path)
        workbook["data"]["A1"] = "wrong"
        workbook.save(path)
        workbook.close()
        reported = self.directory / "requested.xlsx"
        with self.assertRaises(WorkbookError) as context:
            workbook_module._validate_workbook(
                path,
                [plan],
                assignments,
                reported_path=reported,
            )
        self.assertEqual(context.exception.output_path, reported)
        self.assertNotIn(".jxcel-", str(context.exception))

    def test_validation_rejects_wrong_value(self):
        plan, assignments, path = self.write_workbook()
        workbook = load_workbook(path)
        workbook["data"]["A2"] = "wrong"
        workbook.save(path)
        workbook.close()
        with self.assertRaisesRegex(WorkbookError, "expected text"):
            workbook_module._validate_workbook(path, [plan], assignments)

    def test_validation_rejects_text_in_numeric_cell(self):
        plan = make_plan("/data", ({"a": 1},))
        assignments = assign_sheet_names([plan])
        path = self.directory / "numeric.xlsx"
        workbook_module._write_workbook([plan], assignments, path)
        workbook = load_workbook(path)
        workbook["data"]["A2"] = "1"
        workbook.save(path)
        workbook.close()
        with self.assertRaisesRegex(WorkbookError, "expected a numeric value"):
            workbook_module._validate_workbook(path, [plan], assignments)

    def test_validation_rejects_formula_cell(self):
        plan, assignments, path = self.write_workbook()
        workbook = load_workbook(path)
        workbook["data"]["A2"] = "=1+1"
        workbook.save(path)
        workbook.close()
        with self.assertRaisesRegex(WorkbookError, "formula"):
            workbook_module._validate_workbook(path, [plan], assignments)

    def test_validation_rejects_extra_rows(self):
        plan, assignments, path = self.write_workbook()
        workbook = load_workbook(path)
        workbook["data"].append(["extra"])
        workbook.save(path)
        workbook.close()
        with self.assertRaisesRegex(WorkbookError, "rows"):
            workbook_module._validate_workbook(path, [plan], assignments)


class DestinationValidationTests(unittest.TestCase):
    def setUp(self):
        self._temporary_directory = tempfile.TemporaryDirectory()
        self.addCleanup(self._temporary_directory.cleanup)
        self.directory = Path(self._temporary_directory.name)

    def test_returns_path_for_valid_destination(self):
        output = self.directory / "out.xlsx"
        self.assertEqual(
            validate_output_destination(output),
            output,
        )

    def test_missing_parent_fails_with_and_without_overwrite(self):
        output = self.directory / "missing" / "out.xlsx"
        for overwrite in (False, True):
            with self.subTest(overwrite=overwrite):
                with self.assertRaisesRegex(
                    WorkbookError,
                    "parent directory does not exist",
                ):
                    validate_output_destination(
                        output,
                        overwrite=overwrite,
                    )
        self.assertFalse(output.parent.exists())

    def test_parent_that_is_a_file_fails(self):
        parent = self.directory / "not-a-directory"
        parent.write_text("content", encoding="utf-8")
        output = parent / "out.xlsx"
        for overwrite in (False, True):
            with self.subTest(overwrite=overwrite):
                with self.assertRaisesRegex(
                    WorkbookError,
                    "parent path is not a directory",
                ):
                    validate_output_destination(
                        output,
                        overwrite=overwrite,
                    )

    def test_output_directory_is_rejected(self):
        output = self.directory / "folder"
        output.mkdir()
        for overwrite in (False, True):
            with self.subTest(overwrite=overwrite):
                with self.assertRaisesRegex(
                    WorkbookError,
                    "Output path is a directory",
                ):
                    validate_output_destination(
                        output,
                        overwrite=overwrite,
                    )

    def test_existing_output_requires_overwrite(self):
        output = self.directory / "out.xlsx"
        output.write_bytes(b"existing")
        with self.assertRaisesRegex(WorkbookError, "already exists"):
            validate_output_destination(output)
        self.assertEqual(
            validate_output_destination(output, overwrite=True),
            output,
        )

    def test_input_and_output_collision_is_rejected(self):
        input_path = self.directory / "input.json"
        input_path.write_text("{}", encoding="utf-8")
        with self.assertRaisesRegex(WorkbookError, "input file"):
            validate_output_destination(
                input_path,
                overwrite=True,
                input_path=input_path,
            )

    def test_input_collision_is_rejected_without_overwrite(self):
        input_path = self.directory / "input.json"
        input_path.write_text("{}", encoding="utf-8")
        with self.assertRaisesRegex(WorkbookError, "input file"):
            validate_output_destination(
                input_path,
                input_path=input_path,
            )

    def test_input_collision_through_relative_alias_is_rejected(self):
        input_path = self.directory / "input.json"
        input_path.write_text("{}", encoding="utf-8")
        subdirectory = self.directory / "sub"
        subdirectory.mkdir()
        alias = subdirectory / ".." / "input.json"
        with self.assertRaisesRegex(WorkbookError, "input file"):
            validate_output_destination(
                alias,
                overwrite=True,
                input_path=input_path,
            )

    def test_input_collision_through_symlink_is_rejected(self):
        input_path = self.directory / "input.json"
        input_path.write_text("{}", encoding="utf-8")
        alias = self.directory / "alias.json"
        try:
            alias.symlink_to(input_path)
        except (OSError, NotImplementedError) as exc:
            self.skipTest(f"Symbolic links unavailable: {exc}")
        with self.assertRaisesRegex(WorkbookError, "input file"):
            validate_output_destination(
                alias,
                overwrite=True,
                input_path=input_path,
            )

    def test_different_input_and_output_are_allowed(self):
        input_path = self.directory / "input.json"
        input_path.write_text("{}", encoding="utf-8")
        output = self.directory / "out.xlsx"
        self.assertEqual(
            validate_output_destination(output, input_path=input_path),
            output,
        )


class PublicationTests(unittest.TestCase):
    def setUp(self):
        self._temporary_directory = tempfile.TemporaryDirectory()
        self.addCleanup(self._temporary_directory.cleanup)
        self.directory = Path(self._temporary_directory.name)
        self.plan = make_plan("/data", ({"a": 1},))

    def output(self, name="out.xlsx"):
        return self.directory / name

    def temp_files(self):
        return sorted(self.directory.glob(".jxcel-*.tmp.xlsx"))

    def skip_unless_hardlinks_supported(self):
        first = self.directory / "hardlink-probe-a"
        second = self.directory / "hardlink-probe-b"
        first.write_bytes(b"probe")
        try:
            os.link(first, second)
        except OSError as exc:
            unsupported = {
                getattr(errno, "ENOTSUP", -1),
                getattr(errno, "EOPNOTSUPP", -1),
            }
            if exc.errno in unsupported:
                self.skipTest(
                    f"Hard links are unsupported on this filesystem: {exc}"
                )
            raise
        finally:
            for path in (first, second):
                try:
                    path.unlink()
                except FileNotFoundError:
                    pass

    def test_no_overwrite_publication_cleans_temp(self):
        result = export_workbook([self.plan], self.output())
        self.assertIsNone(result.cleanup_warning)
        self.assertTrue(self.output().is_file())
        self.assertEqual(self.temp_files(), [])
        workbook = load_workbook(self.output())
        try:
            self.assertEqual(workbook.sheetnames, ["data"])
        finally:
            workbook.close()

    def test_export_serializes_and_validates_same_stream_while_posix_private(self):
        if os.name != "posix":
            self.skipTest("POSIX descriptor permissions are not available")
        output = self.output()
        observed = {}
        real_save = workbook_module.Workbook.save
        real_validate = workbook_module._validate_workbook

        def save(workbook, stream):
            observed["save_stream"] = stream
            observed["write_mode"] = stat.S_IMODE(os.fstat(stream.fileno()).st_mode)
            real_save(workbook, stream)

        def validate(stream, plans, assignments, **kwargs):
            observed["validation_stream"] = stream
            observed["validation_mode"] = stat.S_IMODE(os.fstat(stream.fileno()).st_mode)
            real_validate(stream, plans, assignments, **kwargs)

        with patch.object(workbook_module.Workbook, "save", autospec=True, side_effect=save), patch.object(
            workbook_module,
            "_validate_workbook",
            side_effect=validate,
        ):
            export_workbook([self.plan], output)

        self.assertIs(observed["save_stream"], observed["validation_stream"])
        self.assertEqual(observed["write_mode"], 0o600)
        self.assertEqual(observed["validation_mode"], 0o600)
        reference = self.directory / "reference-mode"
        descriptor = os.open(reference, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o666)
        os.close(descriptor)
        self.assertEqual(
            stat.S_IMODE(os.stat(output).st_mode),
            stat.S_IMODE(os.stat(reference).st_mode),
        )

    def test_existing_destination_without_overwrite_is_preserved(self):
        output = self.output()
        output.write_bytes(b"existing destination")
        with self.assertRaisesRegex(WorkbookError, "already exists"):
            export_workbook([self.plan], output)
        self.assertEqual(output.read_bytes(), b"existing destination")
        self.assertEqual(self.temp_files(), [])

    def test_real_race_creates_competing_destination_before_link(self):
        self.skip_unless_hardlinks_supported()
        output = self.output()
        real_validate = workbook_module._validate_workbook

        def validate_then_create(path, plans, assignments, **kwargs):
            real_validate(path, plans, assignments, **kwargs)
            output.write_bytes(b"competing content")

        with patch.object(
            workbook_module,
            "_validate_workbook",
            side_effect=validate_then_create,
        ):
            with self.assertRaises(WorkbookError):
                export_workbook([self.plan], output)

        self.assertEqual(output.read_bytes(), b"competing content")
        self.assertEqual(self.temp_files(), [])

    def test_concurrent_process_exports_publish_one_winner_per_destination(self):
        rounds = 10
        context = multiprocessing.get_context("spawn")
        barrier = context.Barrier(3)
        results = context.Queue()
        workers = [
            context.Process(
                target=_competing_export_worker,
                args=(str(self.directory), worker, rounds, barrier, results),
            )
            for worker in (0, 1)
        ]
        for process in workers:
            process.start()

        try:
            for index in range(rounds):
                barrier.wait(timeout=30)
                barrier.wait(timeout=30)
                outcomes = [results.get(timeout=20) for _ in range(2)]
                self.assertEqual({item[0] for item in outcomes}, {index})
                self.assertEqual(
                    sorted(item[2] for item in outcomes),
                    ["exists", "published"],
                    outcomes,
                )
                workbook = load_workbook(
                    self.output(f"concurrent-{index}.xlsx"),
                    data_only=False,
                )
                try:
                    winner = workbook["data"]["A2"].value
                finally:
                    workbook.close()
                self.assertIn(
                    winner,
                    {f"worker-{worker}-round-{index}" for worker in (0, 1)},
                )
                barrier.wait(timeout=30)
        finally:
            for process in workers:
                process.join(timeout=10)
                if process.is_alive():
                    process.terminate()
                    process.join(timeout=5)

        for process in workers:
            self.assertEqual(process.exitcode, 0)

    def test_concurrent_readers_never_observe_partial_overwrite(self):
        output = self.output()
        export_workbook(
            [make_plan(records=({"a": "generation-0"},))],
            output,
        )
        start = threading.Barrier(2)
        ready = threading.Event()
        finished = threading.Event()
        failures = []
        reads = []

        def read_until_done():
            ready.set()
            start.wait(timeout=20)
            while not finished.is_set():
                try:
                    workbook = load_workbook(output, data_only=False)
                    try:
                        value = workbook["data"]["A2"].value
                        if not isinstance(value, str) or not value.startswith("generation-"):
                            failures.append(f"unexpected cell value: {value!r}")
                        else:
                            reads.append(value)
                    finally:
                        workbook.close()
                except PermissionError as exc:
                    # Windows may deny a path-based open while replacement
                    # contends with a reader handle. Final validation below
                    # checks that this was transient and the exact workbook
                    # content remained intact.
                    if os.name != "nt":
                        failures.append(f"reader could not open workbook: {exc!r}")
                except Exception as exc:
                    failures.append(f"reader observed invalid workbook: {exc!r}")

        reader = threading.Thread(target=read_until_done, daemon=True)
        reader.start()
        successful_replacements = 0
        current_value = "generation-0"
        try:
            self.assertTrue(ready.wait(timeout=10))
            start.wait(timeout=20)
            for generation in range(1, 21):
                plan = make_plan(records=({"a": f"generation-{generation}"},))
                try:
                    export_workbook([plan], output, overwrite=True)
                except WorkbookError:
                    # Some platforms may refuse replacing a file opened by a
                    # concurrent reader. The destination must still be readable.
                    try:
                        workbook = load_workbook(output, data_only=False)
                    except PermissionError as exc:
                        if os.name != "nt":
                            raise
                    else:
                        try:
                            self.assertEqual(
                                workbook["data"]["A2"].value,
                                current_value,
                            )
                        finally:
                            workbook.close()
                else:
                    successful_replacements += 1
                    current_value = f"generation-{generation}"
        finally:
            finished.set()
            reader.join(timeout=20)
        self.assertFalse(reader.is_alive())
        workbook = load_workbook(output, data_only=False)
        try:
            self.assertEqual(workbook["data"]["A2"].value, current_value)
        finally:
            workbook.close()
        if os.name == "nt" and successful_replacements == 0:
            # Windows may deny replacement while readers hold handles. Verify
            # overwrite succeeds once the readers close.
            plan = make_plan(records=({"a": "generation-after-readers"},))
            export_workbook([plan], output, overwrite=True)
            successful_replacements += 1
            current_value = "generation-after-readers"
        self.assertGreater(successful_replacements, 0)
        self.assertTrue(reads, "the reader did not observe a published workbook")
        self.assertEqual(failures, [])
        final_workbook = load_workbook(output, data_only=False)
        try:
            self.assertEqual(final_workbook["data"]["A2"].value, current_value)
        finally:
            final_workbook.close()

    def test_link_failure_has_no_unsafe_fallback(self):
        output = self.output()
        with patch(
            "jxcel.workbook.os.link",
            side_effect=OSError(errno.EOPNOTSUPP, "not supported"),
        ), patch(
            "jxcel.workbook.os.replace",
            side_effect=AssertionError("replace fallback attempted"),
        ), patch(
            "jxcel.workbook.os.rename",
            side_effect=AssertionError("rename fallback attempted"),
        ):
            with self.assertRaisesRegex(WorkbookError, "No fallback"):
                export_workbook([self.plan], output)
        self.assertFalse(output.exists())
        self.assertEqual(self.temp_files(), [])

    def test_overwrite_publication_replaces_existing_output(self):
        output = self.output()
        output.write_bytes(b"old content")
        result = export_workbook([self.plan], output, overwrite=True)
        self.assertIsNone(result.cleanup_warning)
        self.assertNotEqual(output.read_bytes(), b"old content")
        self.assertEqual(self.temp_files(), [])
        workbook = load_workbook(output)
        try:
            self.assertEqual(workbook.sheetnames, ["data"])
        finally:
            workbook.close()

    def test_overwrite_preserves_existing_permission_bits(self):
        if os.name != "posix":
            self.skipTest("POSIX permission bits are not available")
        output = self.output()
        output.write_bytes(b"old content")
        os.chmod(output, 0o640)
        export_workbook([self.plan], output, overwrite=True)
        self.assertEqual(stat.S_IMODE(os.stat(output).st_mode), 0o640)

    def test_permission_preservation_failure_aborts_before_replace(self):
        output = self.output()
        output.write_bytes(b"old content")
        permission_function = "os.fchmod" if os.name == "posix" else "os.chmod"
        real_fchmod = os.fchmod if os.name == "posix" else os.chmod

        def fail_when_restoring_mode(*args):
            if os.name == "posix" and args[1] == 0o600:
                return real_fchmod(*args)
            raise OSError(errno.EPERM, "denied")

        with patch(
            f"jxcel.workbook.{permission_function}",
            side_effect=fail_when_restoring_mode,
        ), patch(
            "jxcel.workbook.os.replace",
            side_effect=AssertionError("replace attempted"),
        ):
            with self.assertRaisesRegex(WorkbookError, "permissions"):
                export_workbook([self.plan], output, overwrite=True)
        self.assertEqual(output.read_bytes(), b"old content")
        self.assertEqual(self.temp_files(), [])

    def test_new_output_uses_normal_creation_permissions(self):
        if os.name != "posix":
            self.skipTest("POSIX permission bits are not available")
        reference = self.directory / "reference"
        descriptor = os.open(
            reference,
            os.O_CREAT | os.O_EXCL | os.O_WRONLY,
            0o666,
        )
        os.close(descriptor)
        expected_mode = stat.S_IMODE(os.stat(reference).st_mode)

        output = self.output()
        export_workbook([self.plan], output)
        self.assertEqual(stat.S_IMODE(os.stat(output).st_mode), expected_mode)

    def test_replace_failure_preserves_existing_output(self):
        output = self.output()
        output.write_bytes(b"old content")
        with patch(
            "jxcel.workbook.os.replace",
            side_effect=OSError(errno.EACCES, "denied"),
        ):
            with self.assertRaisesRegex(WorkbookError, "Could not replace"):
                export_workbook([self.plan], output, overwrite=True)
        self.assertEqual(output.read_bytes(), b"old content")
        self.assertEqual(self.temp_files(), [])

    def test_temp_creation_failure_is_reported_without_output(self):
        output = self.output()
        with patch(
            "jxcel.workbook.os.open",
            side_effect=OSError(errno.EACCES, "denied"),
        ):
            with self.assertRaisesRegex(
                WorkbookError,
                "Could not create a temporary workbook",
            ):
                export_workbook([self.plan], output)
        self.assertFalse(output.exists())
        self.assertEqual(self.temp_files(), [])

    def test_interrupt_while_opening_new_temp_stream_cleans_temp(self):
        output = self.output()
        with patch("jxcel.workbook.os.fdopen", side_effect=KeyboardInterrupt):
            with self.assertRaises(KeyboardInterrupt):
                export_workbook([self.plan], output)

        self.assertFalse(output.exists())
        self.assertEqual(self.temp_files(), [])

    def test_write_failure_cleans_temp_and_preserves_output(self):
        output = self.output()
        output.write_bytes(b"old content")
        with patch.object(
            workbook_module.Workbook,
            "save",
            side_effect=OSError(errno.EIO, "write failed"),
        ):
            with self.assertRaisesRegex(WorkbookError, "Could not write") as context:
                export_workbook([self.plan], output, overwrite=True)
        self.assertEqual(context.exception.output_path, output)
        self.assertIn(str(output), str(context.exception))
        self.assertNotIn(".jxcel-", str(context.exception))
        self.assertEqual(output.read_bytes(), b"old content")
        self.assertEqual(self.temp_files(), [])

    def test_interrupt_during_write_cleans_temp(self):
        output = self.output()
        with patch.object(
            workbook_module.Workbook,
            "save",
            side_effect=KeyboardInterrupt,
        ):
            with self.assertRaises(KeyboardInterrupt):
                export_workbook([self.plan], output)
        self.assertFalse(output.exists())
        self.assertEqual(self.temp_files(), [])

    def test_validation_failure_cleans_temp_and_preserves_output(self):
        output = self.output()
        output.write_bytes(b"old content")
        with patch.object(
            workbook_module,
            "_validate_workbook",
            side_effect=WorkbookError("validation failed"),
        ):
            with self.assertRaisesRegex(WorkbookError, "validation failed"):
                export_workbook([self.plan], output, overwrite=True)
        self.assertEqual(output.read_bytes(), b"old content")
        self.assertEqual(self.temp_files(), [])

    def test_corrupt_saved_workbook_fails_validation_and_preserves_output(self):
        output = self.output()
        output.write_bytes(b"old content")

        def write_corrupt_workbook(_workbook, stream):
            stream.seek(0)
            stream.write(b"not a valid xlsx archive")
            stream.truncate()

        with patch.object(
            workbook_module.Workbook,
            "save",
            autospec=True,
            side_effect=write_corrupt_workbook,
        ):
            with self.assertRaisesRegex(
                WorkbookError,
                "Could not reopen generated workbook",
            ):
                export_workbook([self.plan], output, overwrite=True)

        self.assertEqual(output.read_bytes(), b"old content")
        self.assertEqual(self.temp_files(), [])

    def test_temp_file_is_created_in_destination_directory(self):
        output = self.output()
        real_validate = workbook_module._validate_workbook
        observed = []

        def observe(target, plans, assignments, **kwargs):
            observed.append(kwargs["reported_path"].parent)
            real_validate(target, plans, assignments, **kwargs)

        with patch.object(
            workbook_module,
            "_validate_workbook",
            side_effect=observe,
        ):
            export_workbook([self.plan], output)
        self.assertEqual(observed, [self.directory])

    def test_cleanup_warning_after_successful_publication(self):
        output = self.output()
        with patch(
            "jxcel.workbook.os.unlink",
            side_effect=OSError(errno.EACCES, "denied"),
        ):
            result = export_workbook([self.plan], output)
        self.assertIsNotNone(result.cleanup_warning)
        self.assertIn("could not be removed", result.cleanup_warning)
        self.assertTrue(output.is_file())
        for path in self.temp_files():
            path.unlink()

    def test_cleanup_failure_is_reported_with_export_failure(self):
        output = self.output()
        with patch.object(
            workbook_module,
            "_validate_workbook",
            side_effect=WorkbookError("validation failed"),
        ), patch(
            "jxcel.workbook.os.unlink",
            side_effect=OSError(errno.EACCES, "denied"),
        ):
            with self.assertRaisesRegex(
                WorkbookError,
                "additionally.*could not be removed",
            ):
                export_workbook([self.plan], output)
        self.assertFalse(output.exists())
        for path in self.temp_files():
            path.unlink()


if __name__ == "__main__":
    unittest.main()
