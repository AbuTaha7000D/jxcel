import math
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from jxcel.jsonio import JsonLoadError, load_json


class LoadJsonTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.root = Path(self.temp_dir.name)

    def write_bytes(self, name, content):
        path = self.root / name
        path.write_bytes(content)
        return path

    def write_text(self, name, content):
        path = self.root / name
        path.write_text(content, encoding="utf-8")
        return path

    def test_valid_utf8_and_unicode_text_are_preserved(self):
        path = self.write_text("unicode.json", '{"text":"مرحبا — café"}')

        self.assertEqual(load_json(path), {"text": "مرحبا — café"})

    def test_utf8_bom_is_accepted(self):
        path = self.write_bytes("bom.json", b'\xef\xbb\xbf{"ok": true}')

        self.assertEqual(load_json(path), {"ok": True})

    def test_empty_and_json_whitespace_only_files_fail_clearly(self):
        for name, content in (("empty.json", b""), ("space.json", b" \t\r\n")):
            with self.subTest(name=name):
                path = self.write_bytes(name, content)
                with self.assertRaises(JsonLoadError) as caught:
                    load_json(path)
                self.assertIn(str(path), str(caught.exception))
                self.assertIn("empty", str(caught.exception).lower())

    def test_malformed_json_reports_path_line_and_column(self):
        path = self.write_text("malformed.json", '{\n  "items": [1,]\n}')

        with self.assertRaises(JsonLoadError) as caught:
            load_json(path)

        message = str(caught.exception)
        self.assertIn(str(path), message)
        self.assertIn("line 2", message)
        self.assertIn("column", message)
        self.assertIsNotNone(caught.exception.__cause__)

    def test_excessive_nesting_fails_with_path_aware_loader_error(self):
        nesting = 100_000
        path = self.write_text(
            "deeply-nested.json",
            "[" * nesting + "0" + "]" * nesting,
        )

        with self.assertRaises(JsonLoadError) as caught:
            load_json(path)

        self.assertIn(str(path), str(caught.exception))
        self.assertIn("nesting", str(caught.exception).lower())
        self.assertIsNone(caught.exception.__cause__)

    def test_brackets_inside_json_strings_do_not_count_as_nesting(self):
        path = self.write_text(
            "brackets-in-string.json",
            '["' + "[{}]" * 2_000 + '"]',
        )

        self.assertEqual(load_json(path), ["[{}]" * 2_000])

    def test_array_nesting_at_supported_limit_is_accepted(self):
        depth = 500
        path = self.write_text(
            "arrays-at-depth-limit.json",
            "[" * depth + "0" + "]" * depth,
        )

        value = load_json(path)
        for _ in range(depth):
            self.assertIsInstance(value, list)
            self.assertEqual(len(value), 1)
            value = value[0]
        self.assertEqual(value, 0)

    def test_array_nesting_above_supported_limit_is_rejected(self):
        depth = 501
        path = self.write_text(
            "arrays-over-depth-limit.json",
            "[" * depth + "0" + "]" * depth,
        )

        with self.assertRaises(JsonLoadError) as caught:
            load_json(path)

        self.assertIn(str(path), str(caught.exception))
        self.assertIn("500", str(caught.exception))

    def test_object_nesting_at_supported_limit_is_accepted(self):
        depth = 500
        path = self.write_text(
            "objects-at-depth-limit.json",
            '{"a":' * depth + "0" + "}" * depth,
        )

        value = load_json(path)
        for _ in range(depth):
            self.assertIsInstance(value, dict)
            value = value["a"]
        self.assertEqual(value, 0)

    def test_object_nesting_above_supported_limit_is_rejected(self):
        depth = 501
        path = self.write_text(
            "objects-over-depth-limit.json",
            '{"a":' * depth + "0" + "}" * depth,
        )

        with self.assertRaises(JsonLoadError) as caught:
            load_json(path)

        self.assertIn(str(path), str(caught.exception))
        self.assertIn("500", str(caught.exception))

    def test_trailing_content_is_rejected(self):
        path = self.write_text("trailing.json", "{} {}")

        with self.assertRaises(JsonLoadError) as caught:
            load_json(path)

        self.assertIn(str(path), str(caught.exception))
        self.assertIn("line 1", str(caught.exception))

    def test_invalid_utf8_reports_path_and_byte_offset(self):
        path = self.write_bytes("invalid-utf8.json", b'{"x":"\xff"}')

        with self.assertRaises(JsonLoadError) as caught:
            load_json(path)

        self.assertIn(str(path), str(caught.exception))
        self.assertIn("byte offset", str(caught.exception))
        self.assertIsNotNone(caught.exception.__cause__)

    def test_duplicate_keys_at_root_are_rejected(self):
        self.assert_duplicate_rejected('{"x": 1, "x": 2}', "x")

    def test_duplicate_keys_in_nested_objects_are_rejected(self):
        self.assert_duplicate_rejected('{"outer": {"x": 1, "x": 2}}', "x")

    def test_duplicate_keys_inside_objects_in_arrays_are_rejected(self):
        self.assert_duplicate_rejected('[{"x": 1}, {"y": 2, "y": 3}]', "y")

    def test_escaped_equivalent_keys_are_duplicates(self):
        self.assert_duplicate_rejected('{"a": 1, "\\u0061": 2}', "a")

    def test_same_key_in_separate_objects_is_allowed(self):
        path = self.write_text("separate.json", '[{"x": 1}, {"x": 2}]')

        self.assertEqual(load_json(path), [{"x": 1}, {"x": 2}])

    def assert_duplicate_rejected(self, document, key):
        path = self.write_text("duplicate.json", document)
        with self.assertRaises(JsonLoadError) as caught:
            load_json(path)
        self.assertIn(str(path), str(caught.exception))
        self.assertIn(repr(key), str(caught.exception))
        self.assertIsNotNone(caught.exception.__cause__)

    def test_nonstandard_numeric_constants_are_rejected(self):
        for token in ("NaN", "Infinity", "-Infinity"):
            for document in (token, '{"nested": [' + token + "]}"):
                with self.subTest(token=token, document=document):
                    path = self.write_text("constant.json", document)
                    with self.assertRaises(JsonLoadError) as caught:
                        load_json(path)
                    self.assertIn(str(path), str(caught.exception))
                    self.assertIn(repr(token), str(caught.exception))
                    self.assertIn("constant", str(caught.exception).lower())

    def test_all_valid_root_categories_are_returned(self):
        cases = (
            ('{"x": 1}', {"x": 1}, dict),
            ('[1, 2]', [1, 2], list),
            ('"hello"', "hello", str),
            ('42', 42, int),
            ('1.25', 1.25, float),
            ('true', True, bool),
            ('false', False, bool),
            ('null', None, type(None)),
        )
        for document, expected_value, expected_type in cases:
            with self.subTest(document=document):
                path = self.write_text("root.json", document)
                value = load_json(path)
                self.assertIs(type(value), expected_type)
                self.assertEqual(value, expected_value)

    def test_nested_values_and_strings_are_not_reinterpreted(self):
        path = self.write_text(
            "nested.json",
            '{"array": [true, null, {"date": "2024-01-02", "number": "12"}]}',
        )

        value = load_json(path)

        self.assertEqual(
            value,
            {"array": [True, None, {"date": "2024-01-02", "number": "12"}]},
        )
        self.assertIs(type(value["array"][0]), bool)
        self.assertIsNone(value["array"][1])
        self.assertIs(type(value["array"][2]["number"]), str)

    def test_integer_tokens_decode_as_python_int(self):
        path = self.write_text("integer.json", "123456789012345678901234567890")

        value = load_json(path)

        self.assertEqual(value, 123456789012345678901234567890)
        self.assertIs(type(value), int)

    def test_integer_conversion_limit_is_respected_when_enabled(self):
        get_limit = getattr(sys, "get_int_max_str_digits", None)
        if get_limit is None or get_limit() == 0:
            self.skipTest("Python integer-string conversion limit is disabled or unavailable")

        limit = get_limit()
        accepted_path = self.write_text("integer-at-limit.json", "1" * limit)
        accepted = load_json(accepted_path)
        self.assertIs(type(accepted), int)

        rejected_path = self.write_text("integer-over-limit.json", "1" * (limit + 1))
        with self.assertRaises(JsonLoadError) as caught:
            load_json(rejected_path)

        message = str(caught.exception).lower()
        self.assertIn(str(rejected_path), str(caught.exception))
        self.assertIn("integer", message)
        self.assertIn("conversion limit", message)
        self.assertNotIn("invalid json", message)
        self.assertNotIn("unsupported numeric range", message)
        self.assertIsInstance(caught.exception.__cause__, ValueError)

    def test_ordinary_real_number_decodes_as_python_float(self):
        path = self.write_text("real.json", "1.234567890123456789")

        value = load_json(path)

        self.assertIs(type(value), float)
        self.assertEqual(value, float("1.234567890123456789"))

    def test_largest_approved_example_is_finite_float(self):
        path = self.write_text("large-finite.json", "1e308")

        value = load_json(path)

        self.assertIs(type(value), float)
        self.assertTrue(math.isfinite(value))
        self.assertEqual(value, 1e308)

    def test_overflowing_real_number_tokens_are_rejected_as_unsupported_range(self):
        for token in ("1e309", "1e999", "-1e309", "-1e999"):
            with self.subTest(token=token):
                path = self.write_text("overflow.json", token)
                with self.assertRaises(JsonLoadError) as caught:
                    load_json(path)
                message = str(caught.exception)
                self.assertIn(str(path), message)
                self.assertIn(repr(token), message)
                self.assertIn("unsupported numeric range", message.lower())
                self.assertNotIn("invalid json", message.lower())

    def test_nonzero_underflow_tokens_are_rejected(self):
        for token in ("1e-324", "1e-325", "-1e-324", "-1e-325"):
            with self.subTest(token=token):
                path = self.write_text("underflow.json", token)
                with self.assertRaises(JsonLoadError) as caught:
                    load_json(path)
                message = str(caught.exception)
                self.assertIn(str(path), message)
                self.assertIn(repr(token), message)
                self.assertIn("unsupported numeric range", message.lower())

    def test_finite_subnormal_float_is_accepted(self):
        path = self.write_text("subnormal.json", "1e-320")

        value = load_json(path)

        self.assertIs(type(value), float)
        self.assertTrue(math.isfinite(value))
        self.assertNotEqual(value, 0.0)

    def test_zero_with_large_exponent_is_accepted_including_negative_zero(self):
        for token in ("0e999", "-0e999"):
            with self.subTest(token=token):
                path = self.write_text("zero-exponent.json", token)
                value = load_json(path)
                self.assertEqual(value, 0.0)
                self.assertIs(type(value), float)
                self.assertEqual(math.copysign(1.0, value), -1.0 if token.startswith("-") else 1.0)

    def test_missing_file_has_path_and_preserves_io_cause(self):
        path = self.root / "missing.json"

        with self.assertRaises(JsonLoadError) as caught:
            load_json(path)

        self.assertIn(str(path), str(caught.exception))
        self.assertIsInstance(caught.exception.__cause__, FileNotFoundError)

    def test_permission_error_has_path_and_preserves_cause(self):
        path = self.root / "unreadable.json"
        error = PermissionError("permission denied")
        with patch.object(Path, "open", side_effect=error):
            with self.assertRaises(JsonLoadError) as caught:
                load_json(path)

        self.assertIn(str(path), str(caught.exception))
        self.assertIn("permission denied", str(caught.exception))
        self.assertIs(caught.exception.__cause__, error)

    def test_other_io_error_has_path_and_preserves_cause(self):
        path = self.root / "io-error.json"
        error = OSError("device read failure")
        with patch.object(Path, "open", side_effect=error):
            with self.assertRaises(JsonLoadError) as caught:
                load_json(path)

        self.assertIn(str(path), str(caught.exception))
        self.assertIn("device read failure", str(caught.exception))
        self.assertIs(caught.exception.__cause__, error)


if __name__ == "__main__":
    unittest.main()
