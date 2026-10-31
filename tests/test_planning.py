import copy
import json
import math
import unittest
from dataclasses import FrozenInstanceError, replace
from unittest.mock import patch

from jxcel.discovery import DatasetCandidate
from jxcel.planning import (
    EXCEL_MAX_CELL_UTF16_UNITS,
    WARNING_SAMPLE_LIMIT,
    DatasetPlan,
    PlanningError,
    PreparedValue,
    build_plan,
    decode_column_name,
    encode_column_path,
    parse_column_mapping,
    prepared_rows,
)


def make_dataset(pointer, records):
    return DatasetCandidate(pointer=pointer, records=tuple(records))


def build_one(records, **kwargs):
    result = build_plan([make_dataset("", records)], **kwargs)
    return result.datasets[0]


def column_names(plan):
    return [column.name for column in plan.effective_columns]


def row_values(plan):
    return [tuple(value.value for value in row) for row in prepared_rows(plan)]


def warnings_for(plan):
    return [(warning.code, warning.column) for warning in plan.warnings]


class ColumnEncodingTests(unittest.TestCase):
    def test_encoding_examples_and_round_trips(self):
        cases = (
            (("user", "name"), "user.name"),
            (("user.name",), "user~1name"),
            (("a", "b.c"), "a.b~1c"),
            (("~",), "~0"),
            (("~0",), "~00"),
            (("~1",), "~01"),
            (("",), ""),
            (("", "x"), ".x"),
            (("x", ""), "x."),
            (("a~1b",), "a~01b"),
            (("مستخدم", "اسم"), "مستخدم.اسم"),
            (("😀",), "😀"),
        )
        for path, expected in cases:
            with self.subTest(path=path):
                encoded = encode_column_path(path)
                self.assertEqual(encoded, expected)
                self.assertEqual(decode_column_name(encoded), path)

    def test_encoded_names_do_not_collide_for_apparent_paths(self):
        paths = (
            ("a", "b"),
            ("a.b",),
            ("a~1b",),
            ("~0",),
            ("~1",),
            ("a", "b.c"),
            ("a.b", "c"),
        )
        names = {}
        for path in paths:
            name = encode_column_path(path)
            self.assertNotIn(name, names)
            names[name] = path

    def test_decoding_never_applies_replacements_recursively(self):
        self.assertEqual(decode_column_name("~01"), ("~1",))
        self.assertEqual(decode_column_name("~00"), ("~0",))
        self.assertEqual(decode_column_name("a~2b"), ("a~2b",))
        self.assertEqual(decode_column_name("a~"), ("a~",))

    def test_mapping_parser_accepts_punctuation_and_unicode(self):
        mapping = parse_column_mapping('["/students","order,id=1\\"x"]')
        self.assertEqual(mapping.dataset_pointer, "/students")
        self.assertEqual(mapping.column, 'order,id=1"x')

        mapping = parse_column_mapping('["/طلاب","اسم"]')
        self.assertEqual(mapping.dataset_pointer, "/طلاب")
        self.assertEqual(mapping.column, "اسم")

    def test_mapping_parser_rejects_malformed_arguments(self):
        for argument in (
            "",
            "not json",
            '["a"]',
            '["a","b","c"]',
            '[1,"x"]',
            '{"a":"b"}',
            '"x"',
        ):
            with self.subTest(argument=argument):
                with self.assertRaises(PlanningError):
                    parse_column_mapping(argument)

    def test_build_plan_rejects_malformed_mapping(self):
        with self.assertRaises(PlanningError):
            build_one([{"x": 1}], include_maps=("not json",))


class SchemaTests(unittest.TestCase):
    def test_columns_follow_first_occurrence_and_key_order(self):
        plan = build_one([{"b": 1, "a": 2}, {"c": 3, "a": 4}])

        self.assertEqual(
            [column.name for column in plan.columns],
            ["b", "a", "c"],
        )

    def test_nested_flattening_uses_depth_first_source_order(self):
        plan = build_one(
            [
                {
                    "id": 1,
                    "user": {
                        "name": "Ada",
                        "address": {"city": "Riyadh"},
                    },
                    "tags": ["math"],
                }
            ]
        )

        self.assertEqual(
            [column.name for column in plan.columns],
            ["id", "user.name", "user.address.city", "tags"],
        )

    def test_nested_schema_is_union_across_records(self):
        plan = build_one(
            [
                {"user": {"name": "A"}},
                {"user": {"id": 1}},
            ]
        )

        self.assertEqual(
            [column.name for column in plan.columns],
            ["user.name", "user.id"],
        )

    def test_no_flatten_keeps_one_column_per_top_level_key(self):
        plan = build_one(
            [
                {
                    "id": 1,
                    "user": {"name": "Ada"},
                    "tags": ["math", "art"],
                    "extra": {},
                }
            ],
            flatten=False,
        )

        self.assertEqual(
            [column.name for column in plan.columns],
            ["id", "user", "tags", "extra"],
        )

    def test_empty_nested_objects_produce_no_flattened_column(self):
        plan = build_one([{"id": 1, "extra": {}}])

        self.assertEqual([column.name for column in plan.columns], ["id"])

    def test_all_empty_objects_produce_zero_columns(self):
        plan = build_one([{}, {}])

        self.assertEqual(plan.columns, ())
        self.assertEqual(plan.non_exportable_reason, "zero_columns")
        self.assertTrue(plan.selected)
        self.assertFalse(plan.exportable)
        self.assertEqual(
            warnings_for(plan),
            [("dataset_skipped_zero_columns", None)],
        )

    def test_empty_planned_column_name_is_preserved(self):
        plan = build_one([{"": 1, "a": 2}])

        self.assertEqual([column.name for column in plan.columns], ["", "a"])
        self.assertEqual(row_values(plan), [(1, 2)])

    def test_empty_nested_key_is_preserved(self):
        plan = build_one([{"a": {"": 2}}])

        self.assertEqual([column.name for column in plan.columns], ["a."])
        self.assertEqual(decode_column_name("a."), ("a", ""))
        self.assertEqual(row_values(plan), [(2,)])

    def test_no_flatten_preserves_empty_top_level_key(self):
        plan = build_one([{"": {"x": 1}}], flatten=False)

        self.assertEqual([column.name for column in plan.columns], [""])
        self.assertEqual(row_values(plan), [('{"x":1}',)])


class ValuePreparationTests(unittest.TestCase):
    def test_missing_and_null_values_become_blanks(self):
        plan = build_one([{"a": 1}, {"b": 2}, {"a": None}])

        self.assertEqual(
            row_values(plan),
            [
                (1, None),
                (None, 2),
                (None, None),
            ],
        )

    def test_strings_are_preserved_without_reinterpretation(self):
        plan = build_one(
            [
                {
                    "code": "001",
                    "date": "2026-10-31",
                    "formula": "=SUM(A1:A2)",
                }
            ]
        )

        row = list(prepared_rows(plan))[0]
        self.assertEqual(
            [value.value for value in row],
            ["001", "2026-10-31", "=SUM(A1:A2)"],
        )
        self.assertEqual([value.kind for value in row], ["text"] * 3)

    def test_booleans_before_integers_and_float_preservation(self):
        plan = build_one([{"flag": True, "count": 42, "ratio": 1.25}])

        row = list(prepared_rows(plan))[0]
        self.assertEqual(
            row,
            (
                PreparedValue(value=True, kind="boolean"),
                PreparedValue(value=42, kind="integer"),
                PreparedValue(value=1.25, kind="float"),
            ),
        )

    def test_arabic_and_supplementary_unicode_are_preserved(self):
        plan = build_one([{"name": "أحمد", "emoji": "😀"}])

        row = list(prepared_rows(plan))[0]
        self.assertEqual(row[0].value, "أحمد")
        self.assertEqual(row[1].value, "😀")
        self.assertEqual([value.kind for value in row], ["text", "text"])

    def test_large_integers_become_text_at_the_approved_boundary(self):
        plan = build_one(
            [
                {"safe": 999_999_999_999_999, "large": 10**15},
                {"safe": -(10**15), "large": -(10**15) - 1},
            ]
        )

        rows = list(prepared_rows(plan))
        self.assertEqual(rows[0][0], PreparedValue(999_999_999_999_999, "integer"))
        self.assertEqual(rows[0][1], PreparedValue(str(10**15), "text"))
        self.assertEqual(rows[1][0], PreparedValue(str(-(10**15)), "text"))
        self.assertEqual(rows[1][1], PreparedValue(str(-(10**15) - 1), "text"))
        self.assertEqual(
            warnings_for(plan),
            [("integer_precision_as_text", "safe"), ("integer_precision_as_text", "large")],
        )

    def test_nested_values_serialize_as_compact_deterministic_json(self):
        plan = build_one(
            [
                {
                    "user": {"name": "Ada", "tags": ["math", "art"]},
                    "numbers": [1, 2],
                    "empty_object": {},
                    "empty_array": [],
                    "ordered": {"b": 1, "a": 2},
                    "arabic": {"name": "أحمد"},
                }
            ],
            flatten=False,
        )

        self.assertEqual(
            row_values(plan),
            [
                (
                    '{"name":"Ada","tags":["math","art"]}',
                    "[1,2]",
                    "{}",
                    "[]",
                    '{"b":1,"a":2}',
                    '{"name":"أحمد"}',
                )
            ],
        )

    def test_nested_values_emit_aggregated_warnings(self):
        plan = build_one(
            [
                {"user": {"name": "A"}, "tags": ["x"]},
                {"user": {"name": "B"}, "tags": ["y"]},
            ],
            flatten=False,
        )

        self.assertEqual(
            warnings_for(plan),
            [
                ("nested_value_as_json_text", "user"),
                ("nested_value_as_json_text", "tags"),
            ],
        )
        for warning in plan.warnings:
            self.assertEqual(warning.occurrence_count, 2)
            self.assertEqual(warning.sample_record_indexes, (0, 1))


class FloatPrecisionTests(unittest.TestCase):
    def test_shortest_repr_with_at_most_15_significant_digits_is_not_warned(self):
        values = (
            0.1,
            1.0,
            -0.0,
            0.0,
            1e15,
            123456789012345.0,
            1.23456789012345,
            1e-6,
        )
        for value in values:
            with self.subTest(value=value):
                plan = build_one([{"x": value}])
                self.assertEqual(plan.warnings, ())
                self.assertEqual(row_values(plan), [(value,)])

    def test_more_than_15_significant_digits_is_warned_without_changing_value(self):
        values = (
            1234567890123456.0,
            1.234567890123456,
            1.2345678901234567,
            0.30000000000000004,
            2.2250738585072014e-308,
        )
        for value in values:
            with self.subTest(value=value):
                plan = build_one([{"x": value}])
                self.assertEqual(
                    warnings_for(plan),
                    [("float_precision_limited", "x")],
                )
                row = list(prepared_rows(plan))[0][0]
                self.assertEqual(row, PreparedValue(value=value, kind="float"))
                self.assertEqual(row.value, value)

    def test_float_warning_aggregates_across_records(self):
        plan = build_one(
            [{"x": 0.1}, {"x": 0.30000000000000004}, {"x": 1.2}]
        )

        self.assertEqual(len(plan.warnings), 1)
        warning = plan.warnings[0]
        self.assertEqual(warning.code, "float_precision_limited")
        self.assertEqual(warning.occurrence_count, 1)
        self.assertEqual(warning.sample_record_indexes, (1,))

    def test_float_outside_excel_range_is_fatal(self):
        values = (
            1e308,
            math.nextafter(9.99999999999999e307, math.inf),
            5e-324,
            math.nextafter(2.2250738585072014e-308, 0.0),
            float("nan"),
            float("inf"),
            float("-inf"),
        )
        for value in values:
            with self.subTest(value=value):
                with self.assertRaises(PlanningError) as caught:
                    build_one([{"x": value}])
                self.assertEqual(caught.exception.dataset_pointer, "")
                self.assertEqual(caught.exception.column, "x")
                self.assertEqual(caught.exception.record_index, 0)

    def test_documented_excel_float_boundaries_are_accepted(self):
        for value in (9.99999999999999e307, 2.2250738585072014e-308):
            with self.subTest(value=value):
                plan = build_one([{"x": value}])
                self.assertEqual(row_values(plan), [(value,)])

    def test_float_inside_nested_json_is_text_and_not_range_checked(self):
        plan = build_one([{"x": [1e308]}], flatten=False)

        self.assertEqual(row_values(plan), [("[1e+308]",)])
        self.assertEqual(
            warnings_for(plan),
            [("nested_value_as_json_text", "x")],
        )


class ColumnMappingTests(unittest.TestCase):
    def setUp(self):
        self.records = [
            {"id": 1, "name": "A", "secret": "x"},
            {"id": 2, "name": "B", "secret": "y"},
        ]

    def test_inclusion_follows_schema_order_not_argument_order(self):
        plan = build_one(
            self.records,
            include_maps=('["","name"]', '["","id"]'),
        )

        self.assertEqual(column_names(plan), ["id", "name"])
        self.assertEqual(
            [column.name for column in plan.columns],
            ["id", "name", "secret"],
        )

    def test_exclusion_applies_after_inclusion(self):
        plan = build_one(
            self.records,
            include_maps=('["","id"]', '["","name"]'),
            exclude_maps=('["","name"]',),
        )

        self.assertEqual(column_names(plan), ["id"])

    def test_exclusion_wins_when_column_is_in_both_maps(self):
        plan = build_one(
            self.records,
            include_maps=('["","id"]', '["","name"]'),
            exclude_maps=('["","id"]',),
        )

        self.assertEqual(column_names(plan), ["name"])

    def test_repeated_and_overlapping_maps_are_idempotent(self):
        plan = build_one(
            self.records,
            include_maps=('["","id"]', '["","id"]', '["","name"]'),
            exclude_maps=('["","secret"]', '["","secret"]'),
        )

        self.assertEqual(column_names(plan), ["id", "name"])

    def test_unknown_dataset_pointer_is_fatal_with_context(self):
        with self.assertRaises(PlanningError) as caught:
            build_one(self.records, include_maps=('["/missing","id"]',))

        self.assertEqual(caught.exception.dataset_pointer, "/missing")
        self.assertIn("unknown dataset pointer", str(caught.exception))

    def test_unknown_column_is_fatal_with_context(self):
        with self.assertRaises(PlanningError) as caught:
            build_one(self.records, exclude_maps=('["","missing"]',))

        self.assertEqual(caught.exception.dataset_pointer, "")
        self.assertEqual(caught.exception.column, "missing")
        self.assertIn("unknown column", str(caught.exception))

    def test_unknown_selection_pointer_is_fatal(self):
        with self.assertRaises(PlanningError) as caught:
            build_one(self.records, selected_pointers=("/missing",))

        self.assertEqual(caught.exception.dataset_pointer, "/missing")

    def test_mappings_are_validated_for_unselected_datasets_but_not_applied(self):
        candidates = [
            make_dataset("/a", [{"id": 1}]),
            make_dataset("/b", [{"id": 2, "name": "B"}]),
        ]

        result = build_plan(
            candidates,
            selected_pointers=("/a",),
            include_maps=('["/b","name"]',),
        )

        self.assertEqual(len(result.mapping_outcomes), 1)
        outcome = result.mapping_outcomes[0]
        self.assertFalse(outcome.applied)
        self.assertEqual(outcome.reason, "dataset_not_selected")
        self.assertEqual(result.datasets[1].effective_columns, ())
        self.assertFalse(result.datasets[1].selected)

    def test_selected_mapping_is_marked_applied(self):
        result = build_plan(
            [make_dataset("/a", [{"id": 1, "name": "A"}])],
            include_maps=('["/a","name"]',),
        )

        self.assertTrue(result.mapping_outcomes[0].applied)
        self.assertIsNone(result.mapping_outcomes[0].reason)

    def test_mapping_to_a_punctuation_column_works(self):
        plan = build_one(
            [{"order,id": 1, 'internal=status"x': "ok"}],
            include_maps=('["","order,id"]',),
        )

        self.assertEqual(column_names(plan), ["order,id"])

    def test_mapping_to_unicode_column_and_pointer_works(self):
        candidates = [
            make_dataset("/طلاب", [{"اسم": "أحمد", "عمر": 20}]),
        ]
        result = build_plan(
            candidates,
            include_maps=('["/طلاب","اسم"]',),
        )

        self.assertEqual(
            [column.name for column in result.datasets[0].effective_columns],
            ["اسم"],
        )

    def test_excluding_every_column_makes_selected_dataset_non_exportable(self):
        plan = build_one(
            self.records,
            exclude_maps=(
                '["","id"]',
                '["","name"]',
                '["","secret"]',
            ),
        )

        self.assertEqual(plan.effective_columns, ())
        self.assertEqual(plan.non_exportable_reason, "no_effective_columns")
        self.assertFalse(plan.exportable)
        self.assertEqual(
            warnings_for(plan),
            [("dataset_skipped_no_effective_columns", None)],
        )

    def test_mapping_cannot_make_a_zero_column_dataset_exportable(self):
        with self.assertRaises(PlanningError):
            build_one([{}, {}], include_maps=('["","id"]',))

    def test_invalid_mapping_for_unselected_dataset_still_fails(self):
        candidates = [
            make_dataset("/a", [{"id": 1}]),
            make_dataset("/b", [{"id": 2}]),
        ]

        with self.assertRaises(PlanningError) as caught:
            build_plan(
                candidates,
                selected_pointers=("/a",),
                include_maps=('["/b","missing"]',),
            )

        self.assertEqual(caught.exception.dataset_pointer, "/b")
        self.assertEqual(caught.exception.column, "missing")

    def test_no_flatten_mappings_use_top_level_column_names(self):
        candidates = [make_dataset("", [{"id": 1, "user": {"name": "Ada"}}])]

        result = build_plan(
            candidates,
            flatten=False,
            include_maps=('["","user"]',),
        )
        plan = result.datasets[0]

        self.assertEqual(column_names(plan), ["user"])
        self.assertEqual(row_values(plan), [('{"name":"Ada"}',)])

        with self.assertRaises(PlanningError):
            build_plan(
                candidates,
                flatten=False,
                include_maps=('["","user.name"]',),
            )


class ExcelLimitTests(unittest.TestCase):
    def test_cell_length_boundary_uses_utf16_code_units(self):
        allowed = "a" * EXCEL_MAX_CELL_UTF16_UNITS
        fatal = "a" * (EXCEL_MAX_CELL_UTF16_UNITS + 1)

        plan = build_one([{"x": allowed}])
        self.assertEqual(len(row_values(plan)[0][0]), EXCEL_MAX_CELL_UTF16_UNITS)

        with self.assertRaises(PlanningError) as caught:
            build_one([{"x": fatal}])
        self.assertEqual(caught.exception.column, "x")
        self.assertEqual(caught.exception.record_index, 0)

    def test_supplementary_characters_count_as_two_utf16_units(self):
        allowed = "😀" * 16_383 + "a"
        fatal = "😀" * 16_384

        self.assertEqual(len(allowed), 16_384)
        build_one([{"x": allowed}])

        with self.assertRaises(PlanningError):
            build_one([{"x": fatal}])

    def test_serialized_nested_json_cell_length_is_checked(self):
        allowed = "a" * (EXCEL_MAX_CELL_UTF16_UNITS - 8)
        fatal = "a" * (EXCEL_MAX_CELL_UTF16_UNITS - 7)

        plan = build_one([{"x": {"v": allowed}}], flatten=False)
        self.assertEqual(
            len(row_values(plan)[0][0]),
            EXCEL_MAX_CELL_UTF16_UNITS,
        )

        with self.assertRaises(PlanningError):
            build_one([{"x": {"v": fatal}}], flatten=False)

    def test_header_length_boundary_uses_utf16_code_units(self):
        allowed = "😀" * 16_383 + "a"
        fatal = "😀" * 16_384

        build_one([{allowed: 1}])

        with self.assertRaises(PlanningError) as caught:
            build_one([{fatal: 1}])
        self.assertIsNone(caught.exception.record_index)

    def test_column_limit_is_checked_against_effective_columns(self):
        with patch("jxcel.planning.EXCEL_MAX_COLUMNS", 1):
            with self.assertRaises(PlanningError) as caught:
                build_one([{"a": 1, "b": 2}])
            self.assertEqual(caught.exception.dataset_pointer, "")

            plan = build_one([{"a": 1, "b": 2}], include_maps=('["","a"]',))
            self.assertEqual(column_names(plan), ["a"])

    def test_row_limit_counts_data_rows(self):
        with patch("jxcel.planning.EXCEL_MAX_DATA_ROWS", 1):
            build_one([{"x": 1}])
            with self.assertRaises(PlanningError) as caught:
                build_one([{"x": 1}, {"x": 2}])
            self.assertEqual(caught.exception.dataset_pointer, "")

    def test_oversized_excluded_cell_is_not_validated(self):
        records = [{"id": 1, "bad": "a" * (EXCEL_MAX_CELL_UTF16_UNITS + 1)}]

        with self.assertRaises(PlanningError):
            build_one(records)

        plan = build_one(records, include_maps=('["","id"]',))
        self.assertEqual(column_names(plan), ["id"])
        self.assertEqual(row_values(plan), [(1,)])

    def test_oversized_unselected_dataset_is_not_validated(self):
        candidates = [
            make_dataset("/ok", [{"id": 1}]),
            make_dataset(
                "/bad",
                [{"x": "a" * (EXCEL_MAX_CELL_UTF16_UNITS + 1)}],
            ),
        ]

        result = build_plan(candidates, selected_pointers=("/ok",))

        self.assertEqual(
            [column.name for column in result.datasets[0].effective_columns],
            ["id"],
        )
        self.assertEqual(result.datasets[1].effective_columns, ())

    def test_schema_mode_does_not_fail_on_oversized_cells(self):
        plan = build_one(
            [{"x": "a" * (EXCEL_MAX_CELL_UTF16_UNITS + 1)}],
            validation_mode="schema",
        )

        self.assertEqual([column.name for column in plan.columns], ["x"])
        self.assertFalse(plan.export_validated)


class InvalidCharacterTests(unittest.TestCase):
    def assert_invalid_value(self, value, expected_codepoint):
        with self.assertRaises(PlanningError) as caught:
            build_one([{"x": value}])
        error = caught.exception
        self.assertEqual(error.dataset_pointer, "")
        self.assertEqual(error.column, "x")
        self.assertEqual(error.record_index, 0)
        self.assertIn(f"U+{expected_codepoint:04X}", str(error))
        self.assertIn("position 0", str(error))

    def test_forbidden_xml_controls_are_fatal(self):
        for codepoint in (
            0x00,
            0x01,
            0x08,
            0x0B,
            0x0C,
            0x0E,
            0x1F,
        ):
            with self.subTest(codepoint=codepoint):
                self.assert_invalid_value(chr(codepoint), codepoint)

    def test_tab_line_feed_and_carriage_return_are_allowed(self):
        plan = build_one([{"x": "a\tb\nc\rd"}])

        self.assertEqual(row_values(plan), [("a\tb\nc\rd",)])

    def test_lone_utf16_surrogates_are_fatal(self):
        for codepoint in (0xD800, 0xDBFF, 0xDC00, 0xDFFF):
            with self.subTest(codepoint=codepoint):
                self.assert_invalid_value(chr(codepoint), codepoint)

    def test_noncharacters_fffe_and_ffff_are_fatal(self):
        for codepoint in (0xFFFE, 0xFFFF):
            with self.subTest(codepoint=codepoint):
                self.assert_invalid_value(chr(codepoint), codepoint)

    def test_c1_controls_and_valid_unicode_are_allowed(self):
        text = "\x7f\x80\x9f" + "أحمد" + "😀"
        plan = build_one([{"x": text}])

        self.assertEqual(row_values(plan), [(text,)])

    def test_json_surrogate_pair_decodes_to_a_valid_astral_character(self):
        value = json.loads('"\\ud83d\\ude00"')
        self.assertEqual(value, "😀")

        plan = build_one([{"x": value}])
        self.assertEqual(row_values(plan), [("😀",)])

    def test_invalid_header_character_is_fatal_without_record_index(self):
        with self.assertRaises(PlanningError) as caught:
            build_one([{"bad\x00key": 1}])

        self.assertIsNone(caught.exception.record_index)
        self.assertIn("U+0000", str(caught.exception))

    def test_header_lone_surrogates_and_noncharacters_are_fatal(self):
        for codepoint in (0xD800, 0xFFFE, 0xFFFF):
            with self.subTest(codepoint=codepoint):
                with self.assertRaises(PlanningError) as caught:
                    build_one([{f"h{chr(codepoint)}": 1}])
                self.assertIsNone(caught.exception.record_index)
                self.assertIn(f"U+{codepoint:04X}", str(caught.exception))

    def test_invalid_excluded_value_is_ignored(self):
        plan = build_one(
            [{"id": 1, "bad": "x\x00y"}],
            include_maps=('["","id"]',),
        )

        self.assertEqual(row_values(plan), [(1,)])

    def test_invalid_unselected_value_is_ignored(self):
        candidates = [
            make_dataset("/ok", [{"id": 1}]),
            make_dataset("/bad", [{"x": "x\x00y"}]),
        ]

        result = build_plan(candidates, selected_pointers=("/ok",))

        self.assertEqual(
            [column.name for column in result.datasets[0].effective_columns],
            ["id"],
        )

    def test_control_characters_escaped_inside_nested_json_are_allowed(self):
        plan = build_one([{"x": {"text": "a\x00b"}}], flatten=False)

        self.assertEqual(row_values(plan), [('{"text":"a\\u0000b"}',)])

    def test_lone_surrogate_inside_nested_json_is_fatal(self):
        with self.assertRaises(PlanningError) as caught:
            build_one([{"x": {"text": "\ud800"}}], flatten=False)

        self.assertIn("U+D800", str(caught.exception))

    def test_source_records_are_not_mutated(self):
        records = [{"x": "bad\x00value", "nested": {"y": 1.234567890123456}}]
        original = copy.deepcopy(records)

        with self.assertRaises(PlanningError):
            build_one(records)

        self.assertEqual(records, original)


class ValidationModeTests(unittest.TestCase):
    def test_export_mode_marks_selected_exportable_plan_as_validated(self):
        result = build_plan([make_dataset("", [{"x": 1}])])

        plan = result.datasets[0]
        self.assertEqual(result.validation_mode, "export")
        self.assertTrue(plan.export_validated)
        self.assertTrue(plan.exportable)
        self.assertEqual(row_values(plan), [(1,)])

    def test_schema_mode_skips_value_scanning_and_blocks_prepared_rows(self):
        plan = build_one(
            [{"x": 10**15, "y": 0.30000000000000004}],
            validation_mode="schema",
        )

        self.assertEqual(
            warnings_for(plan),
            [],
        )
        self.assertFalse(plan.export_validated)
        with self.assertRaises(PlanningError) as caught:
            prepared_rows(plan)
        self.assertIn("not been validated for export", str(caught.exception))

    def test_schema_mode_still_reports_non_exportable_datasets(self):
        plan = build_one([{}, {}], validation_mode="schema")

        self.assertEqual(plan.non_exportable_reason, "zero_columns")
        self.assertEqual(
            warnings_for(plan),
            [("dataset_skipped_zero_columns", None)],
        )
        self.assertFalse(plan.export_validated)

    def test_schema_mode_still_validates_mappings(self):
        with self.assertRaises(PlanningError):
            build_one(
                [{"x": 1}],
                include_maps=('["","missing"]',),
                validation_mode="schema",
            )

    def test_schema_mode_tolerates_export_only_header_problems(self):
        plan = build_one(
            [{"h" * (EXCEL_MAX_CELL_UTF16_UNITS + 1): 1}],
            validation_mode="schema",
        )

        self.assertEqual(len(plan.columns), 1)
        self.assertFalse(plan.export_validated)

    def test_schema_mode_skips_invalid_characters_and_float_range(self):
        plan = build_one(
            [{"text": "bad\x00value", "number": 1e308}],
            validation_mode="schema",
        )

        self.assertEqual(
            [column.name for column in plan.columns],
            ["text", "number"],
        )
        self.assertEqual(plan.warnings, ())
        self.assertFalse(plan.export_validated)
        with self.assertRaises(PlanningError):
            prepared_rows(plan)

    def test_unselected_plan_is_not_export_validated_and_rejects_prepared_rows(self):
        candidates = [
            make_dataset("/a", [{"x": 1}]),
            make_dataset("/b", [{"x": 2}]),
        ]

        result = build_plan(candidates, selected_pointers=("/a",))
        unselected = result.datasets[1]

        self.assertFalse(unselected.export_validated)
        with self.assertRaises(PlanningError):
            prepared_rows(unselected)

    def test_non_exportable_plan_rejects_prepared_rows(self):
        plan = build_one([{}, {}])

        with self.assertRaises(PlanningError) as caught:
            prepared_rows(plan)
        self.assertIn("not been validated for export", str(caught.exception))

        with self.assertRaises(PlanningError) as caught:
            prepared_rows(replace(plan, export_validated=True))
        self.assertIn("not exportable", str(caught.exception))

    def test_selected_flag_is_required_for_prepared_rows(self):
        plan = build_one([{"x": 1}])

        with self.assertRaises(PlanningError) as caught:
            prepared_rows(replace(plan, selected=False))
        self.assertIn("not selected", str(caught.exception))

    def test_invalid_validation_mode_is_rejected(self):
        with self.assertRaises(ValueError):
            build_one([{"x": 1}], validation_mode="preview")


class WarningAggregationTests(unittest.TestCase):
    def test_many_occurrences_store_one_warning_and_bounded_samples(self):
        records = [{"id": 10**15 + index} for index in range(10)]
        plan = build_one(records)

        self.assertEqual(
            warnings_for(plan),
            [("integer_precision_as_text", "id")],
        )
        warning = plan.warnings[0]
        self.assertEqual(warning.occurrence_count, 10)
        self.assertEqual(
            warning.sample_record_indexes,
            tuple(range(WARNING_SAMPLE_LIMIT)),
        )

    def test_warning_order_follows_schema_column_order(self):
        plan = build_one(
            [
                {
                    "z": 10**15,
                    "a": 0.30000000000000004,
                }
            ]
        )

        self.assertEqual(
            warnings_for(plan),
            [
                ("integer_precision_as_text", "z"),
                ("float_precision_limited", "a"),
            ],
        )

    def test_warning_codes_for_one_column_are_sorted(self):
        plan = build_one(
            [{"x": 10**15}, {"x": ["a"]}],
            flatten=False,
        )

        self.assertEqual(
            warnings_for(plan),
            [
                ("integer_precision_as_text", "x"),
                ("nested_value_as_json_text", "x"),
            ],
        )
        self.assertEqual(
            [warning.occurrence_count for warning in plan.warnings],
            [1, 1],
        )
        self.assertEqual(
            [warning.sample_record_indexes for warning in plan.warnings],
            [(0,), (1,)],
        )

    def test_result_warnings_follow_dataset_order(self):
        candidates = [
            make_dataset("/first", [{"x": 10**15}]),
            make_dataset("/second", [{"x": 10**15}]),
        ]

        result = build_plan(candidates)

        self.assertEqual(
            [warning.dataset_pointer for warning in result.warnings],
            ["/first", "/second"],
        )

    def test_repeated_builds_produce_identical_warnings(self):
        records = [
            {"a": 10**15, "b": 0.30000000000000004},
            {"a": 10**15 + 1, "b": 1.234567890123456},
        ]

        first = build_one(records)
        second = build_one(records)

        self.assertEqual(first, second)
        self.assertEqual(first.warnings, second.warnings)

    def test_unselected_datasets_do_not_emit_value_warnings(self):
        candidates = [
            make_dataset("/a", [{"x": 1}]),
            make_dataset("/b", [{"x": 10**15}]),
        ]

        result = build_plan(candidates, selected_pointers=("/a",))

        self.assertEqual(result.datasets[1].warnings, ())
        self.assertEqual(
            [warning.dataset_pointer for warning in result.warnings],
            [],
        )


class PreparedRowTests(unittest.TestCase):
    def test_prepared_rows_follow_effective_column_order(self):
        plan = build_one(
            [{"id": 1, "name": "A", "secret": "x"}],
            include_maps=('["","name"]', '["","id"]'),
        )

        self.assertEqual(row_values(plan), [(1, "A")])

    def test_prepared_values_expose_kind_and_value(self):
        plan = build_one(
            [
                {
                    "text": "a",
                    "number": 2,
                    "float": 2.5,
                    "flag": False,
                    "blank": None,
                }
            ]
        )

        row = list(prepared_rows(plan))[0]
        self.assertEqual(
            row,
            (
                PreparedValue("a", "text"),
                PreparedValue(2, "integer"),
                PreparedValue(2.5, "float"),
                PreparedValue(False, "boolean"),
                PreparedValue(None, "blank"),
            ),
        )

    def test_prepared_rows_is_repeatable(self):
        plan = build_one([{"x": 1}, {"x": 2}])

        first = row_values(plan)
        second = row_values(plan)

        self.assertEqual(first, second)
        self.assertEqual(first, [(1,), (2,)])

    def test_zero_record_candidate_is_non_exportable(self):
        plan = build_one([])

        self.assertEqual(plan.non_exportable_reason, "zero_columns")
        with self.assertRaises(PlanningError):
            prepared_rows(plan)

    def test_validated_snapshot_remains_authoritative_when_records_are_replaced(self):
        plan = replace(build_one([{"x": 1}]), records=())

        self.assertEqual(
            list(prepared_rows(plan)),
            [(PreparedValue(1, "integer"),)],
        )
        self.assertEqual(plan.row_count, 1)


class ValidatedSnapshotTests(unittest.TestCase):
    def test_top_level_record_mutation_does_not_change_prepared_rows(self):
        plan = build_one([{"value": 1.0}])

        plan.records[0]["value"] = 1e308

        self.assertEqual(row_values(plan), [(1.0,)])

    def test_nested_dict_mutation_does_not_change_prepared_rows(self):
        plan = build_one([{"outer": {"value": 1.0}}])

        plan.records[0]["outer"]["value"] = 1e308

        self.assertEqual(row_values(plan), [(1.0,)])

    def test_nested_list_mutation_does_not_change_prepared_rows(self):
        plan = build_one([{"values": [1, 2]}])

        values = plan.records[0]["values"]
        values.append(3)
        values[0] = 1e308

        self.assertEqual(row_values(plan), [("[1,2]",)])

    def test_mutation_through_plan_records_cannot_change_export_data(self):
        records = [{"value": 1.0}]
        plan = build_one(records)

        self.assertIs(plan.records[0], records[0])
        plan.records[0]["value"] = 1e308

        self.assertEqual(records[0]["value"], 1e308)
        self.assertEqual(row_values(plan), [(1.0,)])

    def test_validated_snapshot_is_deeply_immutable(self):
        plan = build_one([{"value": 1.0}, {"value": 2.0}])

        self.assertIsInstance(plan.validated_rows, tuple)
        rows = plan.validated_rows
        self.assertIsInstance(rows[0], tuple)
        with self.assertRaises(FrozenInstanceError):
            rows[0][0].value = 3.0
        with self.assertRaises(TypeError):
            rows[0][0] = PreparedValue(3.0, "float")

    def test_replacing_records_does_not_change_validated_rows(self):
        plan = build_one([{"x": 1}])

        changed = replace(plan, records=({"x": 1e308},))

        self.assertEqual(row_values(changed), [(1,)])
        self.assertEqual(changed.row_count, 1)

    def test_export_validated_plan_without_snapshot_fails_safely(self):
        plan = replace(build_one([{"x": 1}]), validated_rows=None)

        with self.assertRaisesRegex(
            PlanningError,
            "no validated export snapshot",
        ):
            prepared_rows(plan)
        with self.assertRaisesRegex(
            PlanningError,
            "no validated export snapshot",
        ):
            _ = plan.row_count

    def test_snapshot_preserves_prepared_kinds_and_values(self):
        plan = build_one(
            [
                {
                    "text": "a",
                    "number": 2,
                    "float": 2.5,
                    "flag": False,
                    "blank": None,
                    "big": 10**15,
                    "items": [1, 2],
                }
            ]
        )

        self.assertEqual(
            list(prepared_rows(plan))[0],
            (
                PreparedValue("a", "text"),
                PreparedValue(2, "integer"),
                PreparedValue(2.5, "float"),
                PreparedValue(False, "boolean"),
                PreparedValue(None, "blank"),
                PreparedValue(str(10**15), "text"),
                PreparedValue("[1,2]", "text"),
            ),
        )

    def test_warnings_are_unchanged_by_post_plan_mutation(self):
        plan = build_one(
            [
                {
                    "big": 10**15,
                    "number": 1.234567890123456,
                    "items": [1],
                }
            ]
        )
        original_warnings = plan.warnings

        plan.records[0]["big"] = 1
        plan.records[0]["number"] = 1.0
        plan.records[0]["items"].append(2)

        self.assertEqual(
            warnings_for(plan),
            [
                ("integer_precision_as_text", "big"),
                ("float_precision_limited", "number"),
                ("nested_value_as_json_text", "items"),
            ],
        )
        self.assertEqual(plan.warnings, original_warnings)
        self.assertEqual(
            [warning.occurrence_count for warning in plan.warnings],
            [1, 1, 1],
        )
        self.assertEqual(
            [warning.sample_record_indexes for warning in plan.warnings],
            [(0,), (0,), (0,)],
        )

    def test_row_count_uses_snapshot_for_export_plans(self):
        plan = build_one([{"x": 1}, {"x": 2}, {"x": 3}])

        self.assertEqual(plan.row_count, 3)
        self.assertEqual(plan.row_count, len(plan.validated_rows))

    def test_zero_record_dataset_has_no_snapshot(self):
        plan = build_one([])

        self.assertEqual(plan.row_count, 0)
        self.assertIsNone(plan.validated_rows)
        self.assertFalse(plan.export_validated)

    def test_schema_only_plan_has_no_snapshot_and_uses_record_count(self):
        plan = build_one([{"x": 1}, {"x": 2}], validation_mode="schema")

        self.assertEqual(plan.row_count, 2)
        self.assertIsNone(plan.validated_rows)


class PlanModelTests(unittest.TestCase):
    def test_plan_keeps_full_schema_and_effective_columns(self):
        plan = build_one(
            [{"id": 1, "name": "A"}],
            include_maps=('["","id"]',),
        )

        self.assertIsInstance(plan, DatasetPlan)
        self.assertEqual(
            [column.name for column in plan.columns],
            ["id", "name"],
        )
        self.assertEqual(column_names(plan), ["id"])

    def test_unselected_zero_column_dataset_has_reason_but_no_warning(self):
        candidates = [
            make_dataset("/a", [{"x": 1}]),
            make_dataset("/empty", [{}]),
        ]

        result = build_plan(candidates, selected_pointers=("/a",))
        empty = result.datasets[1]

        self.assertFalse(empty.selected)
        self.assertEqual(empty.non_exportable_reason, "zero_columns")
        self.assertFalse(empty.exportable)
        self.assertEqual(empty.warnings, ())


if __name__ == "__main__":
    unittest.main()
