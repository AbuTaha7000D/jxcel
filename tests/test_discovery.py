import copy
import unittest

from jxcel.discovery import DatasetCandidate, StructureFinding, discover


def pairs(findings):
    return [(finding.kind, finding.pointer) for finding in findings]


class DiscoveryTests(unittest.TestCase):
    def test_empty_root_array_is_reported_once(self):
        findings = discover([])

        self.assertIsInstance(findings, tuple)
        self.assertEqual(pairs(findings), [("empty_array", "")])

    def test_all_object_root_array_is_dataset_at_empty_pointer(self):
        document = [{"x": 1}, {"x": 2}]

        findings = discover(document)

        self.assertEqual(pairs(findings), [("dataset", "")])
        self.assertIsInstance(findings[0], DatasetCandidate)
        self.assertEqual(findings[0].records, ({"x": 1}, {"x": 2}))

    def test_all_empty_object_records_remain_a_dataset_candidate(self):
        findings = discover([{}, {}])

        self.assertEqual(pairs(findings), [("dataset", "")])
        self.assertEqual(findings[0].records, ({}, {}))

    def test_classification_precedence_and_descent(self):
        cases = (
            ([{}], [("dataset", "")]),
            ([{}, 2], [("mixed_array", "")]),
            ([{}, None], [("mixed_array", "")]),
            ([{}, True], [("mixed_array", "")]),
            ([{}, []], [("mixed_array", ""), ("empty_array", "/1")]),
            (
                [{}, [{"x": 1}]],
                [("mixed_array", ""), ("dataset", "/1")],
            ),
            ([[1], 2], [("nested_array", ""), ("primitive_array", "/0")]),
            (
                [[], []],
                [
                    ("nested_array", ""),
                    ("empty_array", "/0"),
                    ("empty_array", "/1"),
                ],
            ),
            ([1, 2], [("primitive_array", "")]),
            ([None, True, "x"], [("primitive_array", "")]),
        )
        for document, expected in cases:
            with self.subTest(document=document):
                self.assertEqual(pairs(discover(document)), expected)

    def test_root_primitive_values_are_unsupported_root(self):
        for value in ("text", 42, 1.25, True, False, None):
            with self.subTest(value=value):
                findings = discover(value)
                self.assertEqual(pairs(findings), [("unsupported_root", "")])
                self.assertIsInstance(findings[0], StructureFinding)

    def test_root_object_is_a_container_and_not_a_finding(self):
        findings = discover({"data": [{"x": 1}]})

        self.assertEqual(pairs(findings), [("dataset", "/data")])

    def test_root_objects_without_arrays_have_no_findings(self):
        for document in ({}, {"a": 1}, {"a": {"b": 2}}, {"a": {}}):
            with self.subTest(document=document):
                self.assertEqual(discover(document), ())

    def test_nested_objects_are_traversed(self):
        document = {"a": {"b": {"c": [{"x": 1}]}}}

        self.assertEqual(pairs(discover(document)), [("dataset", "/a/b/c")])

    def test_empty_array_nested_under_an_object_is_reported(self):
        self.assertEqual(
            pairs(discover({"a": []})),
            [("empty_array", "/a")],
        )

    def test_roadmap_example_preserves_exact_preorder(self):
        document = {"a": [{"x": 1}], "b": [1, {"nested": [{"y": 2}]}]}

        findings = discover(document)

        self.assertEqual(
            pairs(findings),
            [
                ("dataset", "/a"),
                ("mixed_array", "/b"),
                ("dataset", "/b/1/nested"),
            ],
        )
        self.assertEqual(findings[0].records, ({"x": 1},))
        self.assertEqual(findings[2].records, ({"y": 2},))

    def test_source_key_order_is_preserved(self):
        document = {
            "z": [{"x": 1}],
            "a": [{"x": 2}],
            "m": [{"x": 3}],
        }

        self.assertEqual(
            pairs(discover(document)),
            [
                ("dataset", "/z"),
                ("dataset", "/a"),
                ("dataset", "/m"),
            ],
        )

    def test_array_preorder_is_preserved(self):
        document = {"items": [[{"x": 1}], [{"y": 2}]]}

        findings = discover(document)

        self.assertEqual(
            pairs(findings),
            [
                ("nested_array", "/items"),
                ("dataset", "/items/0"),
                ("dataset", "/items/1"),
            ],
        )
        self.assertEqual(findings[1].records, ({"x": 1},))
        self.assertEqual(findings[2].records, ({"y": 2},))

    def test_dataset_boundary_stops_traversal(self):
        document = {
            "data": [
                {
                    "nested": [{"x": 1}],
                    "empty": [],
                    "primitive": [1],
                    "mixed": [1, {}],
                    "objects": [{}],
                }
            ]
        }

        self.assertEqual(pairs(discover(document)), [("dataset", "/data")])

    def test_mixed_root_array_descends_into_object_elements(self):
        document = [1, {"nested": [{"y": 2}]}]

        self.assertEqual(
            pairs(discover(document)),
            [
                ("mixed_array", ""),
                ("dataset", "/1/nested"),
            ],
        )

    def test_nested_array_descent_finds_dataset(self):
        findings = discover([[{"x": 1}]])

        self.assertEqual(
            pairs(findings),
            [
                ("nested_array", ""),
                ("dataset", "/0"),
            ],
        )
        self.assertEqual(findings[1].records, ({"x": 1},))

    def test_pointer_escaping(self):
        document = {
            "a/b~c": [{"x": 1}],
            "~": [{"x": 2}],
            "/": [{"x": 3}],
            "": [{"x": 4}],
            "بيانات": [{"x": 5}],
        }

        self.assertEqual(
            pairs(discover(document)),
            [
                ("dataset", "/a~1b~0c"),
                ("dataset", "/~0"),
                ("dataset", "/~1"),
                ("dataset", "/"),
                ("dataset", "/بيانات"),
            ],
        )

    def test_nested_empty_key_produces_a_trailing_slash_pointer(self):
        findings = discover({"a": {"": [{"x": 1}]}})

        self.assertEqual(pairs(findings), [("dataset", "/a/")])

    def test_pointer_escaping_is_injective(self):
        document = {
            "a/b": [{"x": 1}],
            "a~1b": [{"x": 2}],
            "~0": [{"x": 3}],
            "~00": [{"x": 4}],
        }

        self.assertEqual(
            pairs(discover(document)),
            [
                ("dataset", "/a~1b"),
                ("dataset", "/a~01b"),
                ("dataset", "/~00"),
                ("dataset", "/~000"),
            ],
        )

    def test_finding_pointers_are_unique(self):
        document = {
            "a": [{"x": 1}],
            "b": [1, {"nested": [{"y": 2}]}],
            "c": [[{"z": 3}]],
        }

        findings = discover(document)
        pointers = [finding.pointer for finding in findings]

        self.assertEqual(len(pointers), len(set(pointers)))

    def test_dataset_records_preserve_identity_and_key_order(self):
        records = [{"z": 1, "a": 2}, {"a": 3}]

        findings = discover({"data": records})

        candidate = findings[0]
        self.assertIsInstance(candidate, DatasetCandidate)
        self.assertIs(candidate.records[0], records[0])
        self.assertIs(candidate.records[1], records[1])
        self.assertEqual(list(candidate.records[0]), ["z", "a"])

    def test_discovery_does_not_mutate_source_document(self):
        document = {"data": [{"nested": [{"x": 1}], "empty": []}]}
        snapshot = copy.deepcopy(document)

        discover(document)

        self.assertEqual(document, snapshot)

    def test_findings_are_repeatable(self):
        document = {"a": [{"x": 1}], "b": [1, {"nested": [{"y": 2}]}]}

        self.assertEqual(discover(document), discover(document))


if __name__ == "__main__":
    unittest.main()
