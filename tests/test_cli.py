import contextlib
import io
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from jxcel.cli import main


SRC_DIR = Path(__file__).resolve().parents[1] / "src"

SAMPLE = {
    "students": [
        {"id": 1, "user": {"name": "Ada"}, "tags": ["math", "art"]},
        {"id": 2, "user": {"name": "Grace"}, "tags": []},
    ],
    "orders": [
        {"id": 10, "internal": {"status": "open"}},
        {"id": 11, "internal": {"status": "closed"}},
    ],
}


class CliTestCase(unittest.TestCase):
    def setUp(self):
        temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(temp_dir.cleanup)
        self.root = Path(temp_dir.name)

    def write_json(self, content, name="data.json"):
        path = self.root / name
        text = (
            json.dumps(content, ensure_ascii=False)
            if isinstance(content, (dict, list))
            else content
        )
        path.write_text(text, encoding="utf-8")
        return path

    def write_bytes(self, content, name="data.json"):
        path = self.root / name
        path.write_bytes(content)
        return path

    def run_cli(self, *args):
        stdout = io.StringIO()
        stderr = io.StringIO()
        with (
            contextlib.redirect_stdout(stdout),
            contextlib.redirect_stderr(stderr),
        ):
            code = main([str(argument) for argument in args])
        return code, stdout.getvalue(), stderr.getvalue()

    def run_argument_error(self, *args):
        stdout = io.StringIO()
        stderr = io.StringIO()
        with (
            contextlib.redirect_stdout(stdout),
            contextlib.redirect_stderr(stderr),
        ):
            with self.assertRaises(SystemExit) as caught:
                main([str(argument) for argument in args])
        return caught.exception.code, stdout.getvalue(), stderr.getvalue()

    def names(self):
        return sorted(path.name for path in self.root.iterdir())

    def workbook_rows(self, path, sheet):
        workbook = load_workbook(path)
        try:
            worksheet = workbook[sheet]
            return [
                [cell.value for cell in row]
                for row in worksheet.iter_rows()
            ]
        finally:
            workbook.close()


class ArgumentParsingTests(CliTestCase):
    def test_positional_and_input_option_are_equivalent(self):
        path = self.write_json(SAMPLE)

        first = self.run_cli(path, "--list")
        second = self.run_cli("-i", path, "--list")

        self.assertEqual(first[0], 0)
        self.assertEqual(first, second)

    def test_missing_and_conflicting_input_sources_are_usage_errors(self):
        path = self.write_json(SAMPLE)

        code, _, stderr = self.run_argument_error()
        self.assertEqual(code, 2)
        self.assertIn("input JSON file is required", stderr)

        code, _, stderr = self.run_argument_error(path, "-i", path)
        self.assertEqual(code, 2)
        self.assertIn("not both", stderr)

    def test_list_conflicts_with_preview_and_dry_run(self):
        path = self.write_json(SAMPLE)

        for extra in (
            ("--list", "--preview"),
            ("--list", "--dry-run"),
            ("--list", "--preview", "--dry-run"),
        ):
            with self.subTest(extra=extra):
                code, _, stderr = self.run_argument_error(path, *extra)
                self.assertEqual(code, 2)
                self.assertIn("--list cannot be combined", stderr)

    def test_preview_rows_requires_preview_and_a_positive_integer(self):
        path = self.write_json(SAMPLE)

        code, _, stderr = self.run_argument_error(
            path, "--preview-rows", "3"
        )
        self.assertEqual(code, 2)
        self.assertIn("--preview-rows requires --preview", stderr)

        for value in ("0", "-2", "abc"):
            with self.subTest(value=value):
                code, _, stderr = self.run_argument_error(
                    path, "--preview", "--preview-rows", value
                )
                self.assertEqual(code, 2)
                self.assertIn("--preview-rows", stderr)

    def test_quiet_and_verbose_are_mutually_exclusive(self):
        path = self.write_json(SAMPLE)

        code, _, stderr = self.run_argument_error(path, "-q", "-v")

        self.assertEqual(code, 2)
        self.assertIn("not allowed with", stderr)

    def test_output_and_overwrite_are_rejected_in_inspection_modes(self):
        path = self.write_json(SAMPLE)
        output = self.root / "result.xlsx"

        for extra in (
            ("--list", "-o", output),
            ("--list", "--overwrite"),
            ("--preview", "-o", output),
            ("--preview", "--overwrite"),
        ):
            with self.subTest(extra=extra):
                code, _, stderr = self.run_argument_error(path, *extra)
                self.assertEqual(code, 2)
                self.assertIn("not valid", stderr)

        code, _, _ = self.run_cli(
            path, "--preview", "--dry-run", "-o", output
        )
        self.assertEqual(code, 0)

    def test_verbose_inspection_modes_add_input_context(self):
        path = self.write_json(SAMPLE)

        for mode in ("--list", "--preview"):
            with self.subTest(mode=mode):
                code, stdout, _ = self.run_cli(path, mode, "--verbose")
                self.assertEqual(code, 0)
                self.assertIn(f"Input: '{path}'", stdout)

    def test_quiet_inspection_modes_still_show_their_reports(self):
        path = self.write_json(SAMPLE)

        for mode, expected in (
            ("--list", "Datasets: 2"),
            ("--preview", "Preview '/students'"),
            ("--dry-run", "Dry run"),
        ):
            with self.subTest(mode=mode):
                code, stdout, _ = self.run_cli(path, mode, "--quiet")
                self.assertEqual(code, 0)
                self.assertIn(expected, stdout)

    def test_help_version_and_unknown_options(self):
        code, stdout, _ = self.run_argument_error("--help")
        self.assertEqual(code, 0)
        self.assertIn("usage:", stdout)
        self.assertIn("INPUT", stdout)

        code, stdout, _ = self.run_argument_error("--version")
        self.assertEqual(code, 0)
        self.assertIn("jxcel 0.1.0a0", stdout)

        code, _, stderr = self.run_argument_error("--unknown-option")
        self.assertEqual(code, 2)
        self.assertIn("unrecognized arguments", stderr)


class InputFailureTests(CliTestCase):
    def test_missing_input_file_fails_without_output(self):
        path = self.root / "missing.json"

        code, _, stderr = self.run_cli(path)

        self.assertEqual(code, 1)
        self.assertIn(str(path), stderr)
        self.assertIn("Could not read JSON file", stderr)
        self.assertEqual(self.names(), [])

    def test_invalid_json_reports_line_and_column(self):
        path = self.write_json('{"a": }')

        code, _, stderr = self.run_cli(path, "--list")

        self.assertEqual(code, 1)
        self.assertIn("Invalid JSON", stderr)
        self.assertIn("line 1", stderr)
        self.assertIn("column", stderr)
        self.assertEqual(self.names(), ["data.json"])

    def test_duplicate_keys_and_invalid_utf8_are_rejected(self):
        duplicate = self.write_json('{"a": 1, "a": 2}')
        code, _, stderr = self.run_cli(duplicate)
        self.assertEqual(code, 1)
        self.assertIn("Duplicate object key", stderr)

        invalid = self.write_bytes(b"\xff\xfe", name="bad.json")
        code, _, stderr = self.run_cli(invalid)
        self.assertEqual(code, 1)
        self.assertIn("Invalid UTF-8", stderr)

    def test_primitive_root_is_reported_and_cannot_be_converted(self):
        path = self.write_json("42")

        code, stdout, _ = self.run_cli(path, "--list")
        self.assertEqual(code, 0)
        self.assertIn("'': unsupported_root", stdout)
        self.assertIn("No exportable datasets", stdout)

        code, stdout, _ = self.run_cli(path, "--preview")
        self.assertEqual(code, 0)
        self.assertIn("'': unsupported_root", stdout)

        code, stdout, stderr = self.run_cli(path, "--dry-run")
        self.assertEqual(code, 1)
        self.assertIn("No selected exportable datasets", stderr)

        code, _, stderr = self.run_cli(path)
        self.assertEqual(code, 1)
        self.assertIn("No selected exportable datasets", stderr)
        self.assertEqual(self.names(), ["data.json"])

    def test_empty_and_zero_column_datasets_are_informational_for_list(self):
        path = self.write_json({"empty": [], "records": [{}, {}]})

        code, stdout, _ = self.run_cli(path, "--list")
        self.assertEqual(code, 0)
        self.assertIn("'/empty': empty_array", stdout)
        self.assertIn("non-exportable: zero_columns", stdout)

        code, stdout, _ = self.run_cli(path, "--preview")
        self.assertEqual(code, 0)
        self.assertIn("skipped (non-exportable: zero_columns)", stdout)

        code, _, stderr = self.run_cli(path, "--dry-run")
        self.assertEqual(code, 1)
        self.assertIn("No selected exportable datasets", stderr)

        code, _, stderr = self.run_cli(path)
        self.assertEqual(code, 1)
        self.assertIn("No selected exportable datasets", stderr)
        self.assertEqual(self.names(), ["data.json"])


class ListModeTests(CliTestCase):
    def test_list_reports_order_columns_and_findings_without_writing(self):
        content = dict(SAMPLE)
        content["notes"] = [1, 2]
        path = self.write_json(content)

        code, stdout, _ = self.run_cli(path, "--list")

        self.assertEqual(code, 0)
        self.assertLess(stdout.index("'/students'"), stdout.index("'/orders'"))
        self.assertLess(stdout.index("'/orders'"), stdout.index("'/notes'"))
        self.assertIn("available: id, user.name, tags", stdout)
        self.assertIn("'/notes': primitive_array", stdout)
        self.assertEqual(self.names(), ["data.json"])

    def test_dataset_filter_keeps_unsupported_findings_visible(self):
        content = dict(SAMPLE)
        content["notes"] = [1, 2]
        path = self.write_json(content)

        code, stdout, _ = self.run_cli(
            path, "--dataset", "/students", "--list"
        )

        self.assertEqual(code, 0)
        self.assertIn("'/students'", stdout)
        self.assertNotIn("'/orders'", stdout)
        self.assertIn("'/notes': primitive_array", stdout)

    def test_mappings_validate_globally_but_apply_only_to_selection(self):
        path = self.write_json(SAMPLE)

        code, stdout, _ = self.run_cli(
            path,
            "--dataset",
            "/students",
            "--column-map",
            '["/orders","id"]',
            "--list",
        )
        self.assertEqual(code, 0)
        self.assertNotIn("effective:", stdout)

        code, stdout, _ = self.run_cli(
            path,
            "--dataset",
            "/students",
            "--column-map",
            '["/students","id"]',
            "--list",
        )
        self.assertEqual(code, 0)
        self.assertIn("effective: id", stdout)

    def test_unknown_selections_and_mappings_fail_clearly(self):
        path = self.write_json(SAMPLE)

        code, _, stderr = self.run_cli(
            path, "--dataset", "/missing", "--list"
        )
        self.assertEqual(code, 1)
        self.assertIn("Unknown dataset pointer", stderr)

        code, _, stderr = self.run_cli(
            path, "--column-map", '["/missing","id"]', "--list"
        )
        self.assertEqual(code, 1)
        self.assertIn("unknown dataset pointer", stderr)

        code, _, stderr = self.run_cli(
            path, "--column-map", '["/students","missing"]', "--list"
        )
        self.assertEqual(code, 1)
        self.assertIn("unknown column", stderr)

    def test_list_does_not_export_validate_cell_values(self):
        path = self.write_json([{"x": "bad\u0000value"}])

        code, stdout, _ = self.run_cli(path, "--list")

        self.assertEqual(code, 0)
        self.assertIn("available: x", stdout)
        self.assertEqual(self.names(), ["data.json"])

    def test_malformed_mapping_pair_fails_clearly(self):
        path = self.write_json(SAMPLE)

        for argument in ("not json", '["/students"]', '["/students", 1]'):
            with self.subTest(argument=argument):
                code, _, stderr = self.run_cli(
                    path, "--column-map", argument, "--list"
                )
                self.assertEqual(code, 1)
                self.assertIn("column mapping", stderr.lower())
                self.assertEqual(self.names(), ["data.json"])


class PreviewModeTests(CliTestCase):
    def test_preview_defaults_to_five_rows(self):
        path = self.write_json(
            {"data": [{"id": index} for index in range(7)]}
        )

        code, stdout, _ = self.run_cli(path, "--preview")

        self.assertEqual(code, 0)
        self.assertIn("Preview '/data' (5 of 7 rows):", stdout)
        self.assertIn("\n  4\n", "\n" + stdout + "\n")
        self.assertNotIn("\n  5\n", "\n" + stdout + "\n")
        self.assertEqual(self.names(), ["data.json"])

    def test_preview_respects_dataset_and_column_selection(self):
        path = self.write_json(SAMPLE)

        code, stdout, _ = self.run_cli(
            path,
            "--dataset",
            "/students",
            "--column-map",
            '["/students","user.name"]',
            "--preview",
            "--preview-rows",
            "1",
        )

        self.assertEqual(code, 0)
        self.assertIn("Preview '/students' (1 of 2 rows):", stdout)
        self.assertIn("user.name", stdout)
        self.assertNotIn("'/orders'", stdout)
        self.assertNotIn("id |", stdout)

    def test_preview_does_not_export_validate_unshown_rows(self):
        path = self.write_json(
            [
                {"id": 1, "text": "ok"},
                {"id": 2, "text": "bad\u0000value"},
            ]
        )

        code, stdout, _ = self.run_cli(
            path, "--preview", "--preview-rows", "1"
        )
        self.assertEqual(code, 0)
        self.assertIn("ok", stdout)

        code, _, stderr = self.run_cli(path, "--dry-run")
        self.assertEqual(code, 1)
        self.assertIn("Invalid Excel character", stderr)
        self.assertEqual(self.names(), ["data.json"])

    def test_preview_and_dry_run_combine_without_writing(self):
        content = dict(SAMPLE)
        content["notes"] = [1, 2]
        path = self.write_json(content)

        code, stdout, _ = self.run_cli(
            path, "--preview", "--dry-run", "--preview-rows", "1"
        )

        self.assertEqual(code, 0)
        self.assertIn("Dry run: no workbook will be written.", stdout)
        self.assertIn("Preview '/students' (1 of 2 rows):", stdout)
        self.assertIn("nested_value_as_json_text", stdout)
        self.assertEqual(
            stdout.count("'/notes': primitive_array"),
            1,
        )
        self.assertEqual(self.names(), ["data.json"])

    def test_preview_skips_non_exportable_datasets(self):
        path = self.write_json(
            {"good": [{"x": 1}], "empty": [{}, {}]}
        )

        code, stdout, _ = self.run_cli(path, "--preview")

        self.assertEqual(code, 0)
        self.assertIn(
            "Preview '/empty': skipped "
            "(non-exportable: zero_columns)",
            stdout,
        )

    def test_preview_with_dry_run_still_validates_the_destination(self):
        path = self.write_json(SAMPLE)
        output = self.root / "result.xlsx"
        output.write_text("sentinel", encoding="utf-8")

        code, _, stderr = self.run_cli(
            path, "--preview", "--dry-run", "-o", output
        )

        self.assertEqual(code, 1)
        self.assertIn("already exists", stderr)
        self.assertEqual(output.read_text(encoding="utf-8"), "sentinel")


class DryRunTests(CliTestCase):
    def test_dry_run_reports_destination_and_warnings_without_writing(self):
        path = self.write_json(SAMPLE)

        code, stdout, _ = self.run_cli(path, "--dry-run")

        self.assertEqual(code, 0)
        self.assertIn("Dry run: no workbook will be written.", stdout)
        self.assertIn(f"Output: '{self.root / 'data.xlsx'}'", stdout)
        self.assertIn("Datasets: 2", stdout)
        self.assertIn("columns: id, user.name, tags", stdout)
        self.assertIn("nested_value_as_json_text", stdout)
        self.assertEqual(self.names(), ["data.json"])

    def test_verbose_dry_run_adds_input_and_sheet_names(self):
        path = self.write_json(SAMPLE)
        output = self.root / "result.xlsx"

        code, stdout, _ = self.run_cli(
            path, "--dry-run", "-o", output, "--verbose"
        )

        self.assertEqual(code, 0)
        self.assertIn(f"Input: '{path}'", stdout)
        self.assertIn("sheet: 'students'", stdout)
        self.assertIn("sheet: 'orders'", stdout)
        self.assertEqual(self.names(), ["data.json"])

    def test_dry_run_marks_unselected_mappings_as_validated_but_unapplied(self):
        path = self.write_json(SAMPLE)

        code, stdout, _ = self.run_cli(
            path,
            "--dataset",
            "/students",
            "--column-map",
            '["/orders","id"]',
            "--dry-run",
        )

        self.assertEqual(code, 0)
        self.assertIn(
            "include '/orders' 'id': validated but not applied "
            "(dataset not selected)",
            stdout,
        )

    def test_existing_destination_is_checked_without_replacing_it(self):
        path = self.write_json(SAMPLE)
        output = self.root / "result.xlsx"
        output.write_text("sentinel", encoding="utf-8")

        code, _, stderr = self.run_cli(
            path, "--dry-run", "-o", output
        )
        self.assertEqual(code, 1)
        self.assertIn("already exists", stderr)
        self.assertEqual(output.read_text(encoding="utf-8"), "sentinel")

        code, stdout, _ = self.run_cli(
            path, "--dry-run", "-o", output, "--overwrite"
        )
        self.assertEqual(code, 0)
        self.assertIn("Overwrite: yes", stdout)
        self.assertEqual(output.read_text(encoding="utf-8"), "sentinel")

    def test_missing_output_parent_fails_without_creating_it(self):
        path = self.write_json(SAMPLE)
        output = self.root / "missing" / "result.xlsx"

        code, _, stderr = self.run_cli(
            path, "--dry-run", "-o", output
        )

        self.assertEqual(code, 1)
        self.assertIn("does not exist", stderr)
        self.assertFalse((self.root / "missing").exists())

    def test_dry_run_fails_when_nothing_can_be_exported(self):
        path = self.write_json({"empty": [{}, {}]})

        code, stdout, stderr = self.run_cli(path, "--dry-run")

        self.assertEqual(code, 1)
        self.assertIn("non-exportable: zero_columns", stdout)
        self.assertIn("No selected exportable datasets", stderr)
        self.assertEqual(self.names(), ["data.json"])


class ConversionTests(CliTestCase):
    def test_default_output_contains_expected_sheets_and_values(self):
        path = self.write_json(SAMPLE)

        code, stdout, _ = self.run_cli(path)

        self.assertEqual(code, 0)
        self.assertIn("Exported:", stdout)
        output = self.root / "data.xlsx"
        self.assertTrue(output.exists())
        self.assertEqual(self.names(), ["data.json", "data.xlsx"])

        students = self.workbook_rows(output, "students")
        self.assertEqual(students[0], ["id", "user.name", "tags"])
        self.assertEqual(
            students[1:],
            [[1, "Ada", '["math","art"]'], [2, "Grace", "[]"]],
        )
        orders = self.workbook_rows(output, "orders")
        self.assertEqual(orders[0], ["id", "internal.status"])

    def test_input_option_and_custom_output_path(self):
        path = self.write_json(SAMPLE)
        output = self.root / "résultat final.xlsx"

        code, _, _ = self.run_cli("-i", path, "-o", output)

        self.assertEqual(code, 0)
        self.assertTrue(output.exists())
        self.assertEqual(
            self.workbook_rows(output, "students")[1],
            [1, "Ada", '["math","art"]'],
        )

    def test_no_flatten_keeps_top_level_properties(self):
        path = self.write_json(
            [
                {
                    "id": 1,
                    "user": {"name": "Ada"},
                    "tags": ["a"],
                    "extra": {},
                }
            ]
        )

        code, _, _ = self.run_cli(path, "--no-flatten")

        self.assertEqual(code, 0)
        rows = self.workbook_rows(self.root / "data.xlsx", "root")
        self.assertEqual(
            rows,
            [
                ["id", "user", "tags", "extra"],
                [1, '{"name":"Ada"}', '["a"]', "{}"],
            ],
        )

    def test_dataset_selection_and_column_maps_control_the_sheet(self):
        path = self.write_json(SAMPLE)

        code, _, _ = self.run_cli(
            path,
            "--dataset",
            "/students",
            "--column-map",
            '["/students","id"]',
            "--column-map",
            '["/students","user.name"]',
            "--exclude-map",
            '["/students","id"]',
        )

        self.assertEqual(code, 0)
        output = self.root / "data.xlsx"
        workbook = load_workbook(output)
        try:
            self.assertEqual(workbook.sheetnames, ["students"])
        finally:
            workbook.close()
        rows = self.workbook_rows(output, "students")
        self.assertEqual(rows[0], ["user.name"])
        self.assertEqual(rows[1:], [["Ada"], ["Grace"]])

    def test_overwrite_preserves_the_old_file_until_replacement(self):
        path = self.write_json({"data": [{"x": 1}]})

        code, _, _ = self.run_cli(path)
        self.assertEqual(code, 0)
        output = self.root / "data.xlsx"
        self.assertEqual(self.workbook_rows(output, "data")[1], [1])

        path.write_text('{"data": [{"x": 2}]}', encoding="utf-8")
        code, _, stderr = self.run_cli(path)
        self.assertEqual(code, 1)
        self.assertIn("already exists", stderr)
        self.assertEqual(self.workbook_rows(output, "data")[1], [1])

        code, _, _ = self.run_cli(path, "--overwrite")
        self.assertEqual(code, 0)
        self.assertEqual(self.workbook_rows(output, "data")[1], [2])

    def test_input_output_collision_is_rejected(self):
        path = self.write_json({"data": [{"x": 1}]})
        original = path.read_text(encoding="utf-8")

        code, _, stderr = self.run_cli(path, "-o", path)

        self.assertEqual(code, 1)
        self.assertIn("refers to the input file", stderr)
        self.assertEqual(path.read_text(encoding="utf-8"), original)

    def test_missing_output_parent_is_not_created(self):
        path = self.write_json(SAMPLE)
        output = self.root / "missing" / "result.xlsx"

        code, _, stderr = self.run_cli(path, "-o", output)

        self.assertEqual(code, 1)
        self.assertIn("does not exist", stderr)
        self.assertFalse((self.root / "missing").exists())

    def test_conversion_fails_clearly_without_an_exportable_dataset(self):
        path = self.write_json({"empty": [{}, {}]})

        code, stdout, stderr = self.run_cli(path)

        self.assertEqual(code, 1)
        self.assertIn("non-exportable: zero_columns", stdout)
        self.assertIn("No selected exportable datasets", stderr)
        self.assertEqual(self.names(), ["data.json"])

    def test_quiet_keeps_warnings_and_findings_but_hides_summary(self):
        path = self.write_json(
            {"data": [{"id": 10**15}], "notes": [1, 2]}
        )

        code, stdout, _ = self.run_cli(path, "-q")

        self.assertEqual(code, 0)
        self.assertNotIn("Exported:", stdout)
        self.assertNotIn("'/data' -> 'data'", stdout)
        self.assertIn("integer_precision_as_text", stdout)
        self.assertIn("'/notes': primitive_array", stdout)
        self.assertTrue((self.root / "data.xlsx").exists())

    def test_clean_quiet_conversion_prints_nothing(self):
        path = self.write_json({"data": [{"x": 1}]})

        code, stdout, _ = self.run_cli(path, "-q")

        self.assertEqual(code, 0)
        self.assertEqual(stdout, "")
        self.assertTrue((self.root / "data.xlsx").exists())

    def test_verbose_conversion_adds_input_context(self):
        path = self.write_json({"data": [{"x": 1}]})

        code, stdout, _ = self.run_cli(path, "--verbose")

        self.assertEqual(code, 0)
        self.assertIn(f"Input: '{path}'", stdout)
        self.assertIn("Exported:", stdout)

    def test_formula_like_strings_remain_text(self):
        path = self.write_json(
            [{"formula": "=1+1", "other": "+SUM(A1)"}]
        )

        code, _, _ = self.run_cli(path)

        self.assertEqual(code, 0)
        workbook = load_workbook(self.root / "data.xlsx")
        try:
            worksheet = workbook["root"]
            self.assertEqual(worksheet["A2"].value, "=1+1")
            self.assertEqual(worksheet["A2"].data_type, "s")
            self.assertEqual(worksheet["B2"].value, "+SUM(A1)")
        finally:
            workbook.close()

    def test_failed_conversion_leaves_no_temporary_files(self):
        path = self.write_json({"data": [{"x": 1}]})
        output = self.root / "result.xlsx"
        output.write_text("occupied", encoding="utf-8")

        code, _, _ = self.run_cli(path, "-o", output)

        self.assertEqual(code, 1)
        self.assertEqual(
            self.names(),
            ["data.json", "result.xlsx"],
        )
        self.assertEqual(output.read_text(encoding="utf-8"), "occupied")


class EndToEndTests(CliTestCase):
    def test_module_entry_point_process_exit_codes(self):
        path = self.write_json(SAMPLE)
        output = self.root / "result.xlsx"
        environment = os.environ.copy()
        existing = environment.get("PYTHONPATH")
        environment["PYTHONPATH"] = (
            str(SRC_DIR)
            if not existing
            else str(SRC_DIR) + os.pathsep + existing
        )

        completed = subprocess.run(
            [
                sys.executable,
                "-m",
                "jxcel.cli",
                str(path),
                "-o",
                str(output),
            ],
            capture_output=True,
            text=True,
            env=environment,
            timeout=120,
        )
        self.assertEqual(completed.returncode, 0, completed.stderr)
        self.assertIn("Exported:", completed.stdout)
        self.assertTrue(output.exists())

        completed = subprocess.run(
            [sys.executable, "-m", "jxcel.cli"],
            capture_output=True,
            text=True,
            env=environment,
            timeout=120,
        )
        self.assertEqual(completed.returncode, 2)
        self.assertIn("input JSON file is required", completed.stderr)

    def test_relative_input_uses_default_output_beside_it(self):
        self.write_json(SAMPLE)

        with contextlib.chdir(self.root):
            code, _, _ = self.run_cli("data.json")

        self.assertEqual(code, 0)
        self.assertEqual(self.names(), ["data.json", "data.xlsx"])


if __name__ == "__main__":
    unittest.main()
