import unittest
from pathlib import Path

from jxcel.discovery import DatasetCandidate, StructureFinding
from jxcel.planning import build_plan, preview_rows
from jxcel.reporting import (
    render_export,
    render_listing,
    render_plan,
    render_preview,
)
from jxcel.workbook import SheetSummary, WorkbookExport


def build_result(specs, **kwargs):
    candidates = tuple(
        DatasetCandidate(pointer=pointer, records=tuple(records))
        for pointer, records in specs
    )
    return build_plan(candidates, **kwargs)


class ListingTests(unittest.TestCase):
    def test_listing_shows_datasets_effective_columns_and_findings(self):
        result = build_result(
            [
                (
                    "/students",
                    [
                        {"id": 1, "user": {"name": "Ada"}},
                        {"id": 2, "user": {"name": "Grace"}},
                    ],
                ),
                ("/empty", [{}]),
                ("/orders", [{"id": 1, "internal": {"status": "x"}}]),
            ],
            include_maps=('["/students","id"]',),
            validation_mode="schema",
        )
        findings = (
            StructureFinding(pointer="/notes", kind="mixed_array"),
        )

        rendered = render_listing(result.datasets, findings)

        self.assertEqual(
            rendered,
            "\n".join(
                [
                    "Datasets: 3",
                    "  '/students': 2 records, 2 columns",
                    "    available: id, user.name",
                    "    effective: id",
                    "  '/empty': 1 record, 0 columns "
                    "[non-exportable: zero_columns]",
                    "    available: (none)",
                    "  '/orders': 1 record, 2 columns",
                    "    available: id, internal.status",
                    "",
                    "Unsupported findings: 1",
                    "  '/notes': mixed_array",
                ]
            ),
        )

    def test_dataset_filter_hides_rows_but_keeps_unsupported_findings(self):
        result = build_result(
            [
                ("/students", [{"id": 1}]),
                ("/orders", [{"id": 2}]),
            ],
            selected_pointers=("/students",),
            validation_mode="schema",
        )
        findings = (
            StructureFinding(pointer="/notes", kind="mixed_array"),
        )

        rendered = render_listing(result.datasets, findings)

        self.assertIn("Datasets: 1", rendered)
        self.assertIn("'/students'", rendered)
        self.assertNotIn("'/orders'", rendered)
        self.assertIn("'/notes': mixed_array", rendered)

    def test_listing_notes_when_no_dataset_is_exportable(self):
        rendered = render_listing(
            (),
            (StructureFinding(pointer="", kind="unsupported_root"),),
        )

        self.assertEqual(
            rendered,
            "\n".join(
                [
                    "Datasets: 0",
                    "  (none)",
                    "",
                    "Unsupported findings: 1",
                    "  '': unsupported_root",
                    "",
                    "No exportable datasets; conversion would not produce "
                    "a workbook.",
                ]
            ),
        )


    def test_unsupported_findings_keep_discovery_order(self):
        rendered = render_listing(
            (),
            (
                StructureFinding(pointer="/z", kind="empty_array"),
                StructureFinding(pointer="/a", kind="mixed_array"),
            ),
        )

        self.assertLess(
            rendered.index("'/z': empty_array"),
            rendered.index("'/a': mixed_array"),
        )


class PreviewRenderingTests(unittest.TestCase):
    def test_preview_renders_typed_cells_and_quotes_pipes(self):
        result = build_result(
            [
                (
                    "",
                    [
                        {
                            "a": 1,
                            "b": True,
                            "c": None,
                            "d": "x|y",
                            "e": " pad ",
                        },
                        {"a": 2, "b": False, "c": 0, "d": "z", "e": "q"},
                    ],
                )
            ],
            validation_mode="schema",
        )
        plan = result.datasets[0]
        rows = {"": tuple(preview_rows(plan, 1))}

        rendered = render_preview(
            result.datasets,
            rows,
            (),
        )

        self.assertEqual(
            rendered,
            "\n".join(
                [
                    "Preview '' (1 of 2 rows):",
                    "  a | b | c | d | e",
                    '  1 | TRUE |  | "x|y" | " pad "',
                ]
            ),
        )

    def test_preview_shows_skipped_and_unsupported_entries(self):
        result = build_result(
            [("/data", [{"x": 1}]), ("/empty", [{}])],
            validation_mode="schema",
        )
        rows = {
            "/data": tuple(preview_rows(result.datasets[0], 5)),
        }
        findings = (
            StructureFinding(pointer="/notes", kind="primitive_array"),
        )

        rendered = render_preview(result.datasets, rows, findings)

        self.assertIn("Preview '/data' (1 of 1 rows):", rendered)
        self.assertIn(
            "Preview '/empty': skipped "
            "(non-exportable: zero_columns)",
            rendered,
        )
        self.assertIn("'/notes': primitive_array", rendered)
        self.assertNotIn("No exportable datasets", rendered)

    def test_control_characters_and_surrogates_are_escaped_for_display(self):
        result = build_result(
            [("", [{"x": "bad\u0000\ud800value"}])],
            validation_mode="schema",
        )
        rows = {"": tuple(preview_rows(result.datasets[0], 1))}

        rendered = render_preview(result.datasets, rows, ())

        self.assertIn(r"bad\u0000\ud800value", rendered)

    def test_preview_notes_when_no_dataset_exists(self):
        rendered = render_preview(
            (),
            {},
            (StructureFinding(pointer="", kind="unsupported_root"),),
        )

        self.assertIn("No datasets were discovered.", rendered)
        self.assertIn("'': unsupported_root", rendered)
        self.assertIn("No exportable datasets", rendered)


class PlanRenderingTests(unittest.TestCase):
    def make_plan_result(self):
        return build_result(
            [
                ("/students", [{"id": 1}, {"id": 2}]),
                ("/empty", [{}]),
                ("/orders", [{"id": 3}]),
            ],
            selected_pointers=("/students", "/empty"),
            include_maps=(
                '["/students","id"]',
                '["/orders","id"]',
            ),
        )

    def test_dry_run_plan_includes_destination_columns_and_mappings(self):
        result = self.make_plan_result()
        findings = (
            StructureFinding(pointer="/notes", kind="mixed_array"),
        )

        rendered = render_plan(
            result.datasets,
            findings,
            Path("/tmp/out.xlsx"),
            overwrite=False,
            mapping_outcomes=tuple(
                reversed(result.mapping_outcomes)
            ),
        )

        self.assertEqual(
            rendered,
            "\n".join(
                [
                    "Dry run: no workbook will be written.",
                    "Output: '/tmp/out.xlsx'",
                    "Overwrite: no",
                    "Datasets: 2",
                    "  '/students': 2 rows, 1 column",
                    "    columns: id",
                    "  '/empty': 1 row, 0 columns "
                    "[non-exportable: zero_columns, skipped]",
                    "",
                    "Mappings:",
                    "  include '/orders' 'id': validated but not "
                    "applied (dataset not selected)",
                    "  include '/students' 'id': applied",
                    "",
                    "Unsupported findings: 1",
                    "  '/notes': mixed_array",
                    "",
                    "Warnings: 1",
                    "  dataset_skipped_zero_columns at '/empty': "
                    "1 occurrence",
                ]
            ),
        )

    def test_verbose_plan_adds_input_and_sheet_assignment(self):
        result = self.make_plan_result()

        rendered = render_plan(
            result.datasets,
            (),
            Path("out.xlsx"),
            overwrite=True,
            sheet_names={"/students": "students"},
            input_path=Path("data.json"),
            verbose=True,
        )

        self.assertIn("Input: 'data.json'", rendered)
        self.assertIn("Overwrite: yes", rendered)
        self.assertIn("    sheet: 'students'", rendered)
        self.assertNotIn("sheet: 'empty'", rendered)


class ExportRenderingTests(unittest.TestCase):
    def make_export_result(self):
        result = build_result(
            [
                (
                    "/data",
                    [
                        {"id": 10**15, "name": "A"},
                        {"id": 10**15 + 1, "name": "B"},
                    ],
                ),
                ("/empty", [{}]),
            ],
        )
        export = WorkbookExport(
            output_path=Path("out.xlsx"),
            sheets=(
                SheetSummary(
                    pointer="/data",
                    name="data",
                    data_rows=2,
                    columns=2,
                ),
            ),
            cleanup_warning="cleanup failed",
        )
        findings = (
            StructureFinding(pointer="/notes", kind="nested_array"),
        )
        return result, export, findings

    def test_success_summary_includes_sheets_findings_and_cleanup(self):
        result, export, findings = self.make_export_result()

        rendered = render_export(export, result.datasets, findings)

        self.assertIn("Exported: 'out.xlsx'", rendered)
        self.assertIn("'/data' -> 'data': 2 rows, 2 columns", rendered)
        self.assertIn(
            "Skipped non-exportable datasets: 1",
            rendered,
        )
        self.assertIn("'/empty': zero_columns", rendered)
        self.assertIn("'/notes': nested_array", rendered)
        self.assertIn(
            "integer_precision_as_text at '/data', column 'id': "
            "2 occurrences; sample records 0, 1",
            rendered,
        )
        self.assertIn("Warning: cleanup failed", rendered)

    def test_quiet_keeps_warnings_findings_and_cleanup_only(self):
        result, export, findings = self.make_export_result()

        rendered = render_export(
            export,
            result.datasets,
            findings,
            quiet=True,
        )

        self.assertNotIn("Exported:", rendered)
        self.assertNotIn("'/data' -> 'data'", rendered)
        self.assertIn("Skipped non-exportable datasets", rendered)
        self.assertIn("'/notes': nested_array", rendered)
        self.assertIn("integer_precision_as_text", rendered)
        self.assertIn("Warning: cleanup failed", rendered)

    def test_clean_quiet_export_renders_nothing(self):
        result = build_result([("/data", [{"id": 1}])])
        export = WorkbookExport(
            output_path=Path("out.xlsx"),
            sheets=(
                SheetSummary(
                    pointer="/data",
                    name="data",
                    data_rows=1,
                    columns=1,
                ),
            ),
        )

        self.assertEqual(
            render_export(
                export,
                result.datasets,
                (),
                quiet=True,
            ),
            "",
        )


if __name__ == "__main__":
    unittest.main()
