"""Deterministic dataset discovery for parsed JSON documents."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Literal


StructureKind = Literal[
    "empty_array",
    "mixed_array",
    "nested_array",
    "primitive_array",
    "unsupported_root",
    "scalar_value",
    "empty_object",
]


@dataclass(frozen=True)
class DatasetCandidate:
    """An array whose direct elements are all JSON objects."""

    pointer: str
    records: tuple[dict[str, Any], ...]

    @property
    def kind(self) -> Literal["dataset"]:
        return "dataset"


@dataclass(frozen=True)
class StructureFinding:
    """An array or root value that cannot be treated as a dataset."""

    pointer: str
    kind: StructureKind


Finding = DatasetCandidate | StructureFinding


def _escape_pointer_token(token: str) -> str:
    """Escape one JSON Pointer reference token."""
    return token.replace("~", "~0").replace("/", "~1")


def _object_child_pointer(pointer: str, key: str) -> str:
    return f"{pointer}/{_escape_pointer_token(key)}"


def _array_child_pointer(pointer: str, index: int) -> str:
    return f"{pointer}/{index}"


def _classify_array(
    values: list[Any],
) -> StructureKind | Literal["dataset"]:
    """Classify one array using the approved precedence exactly."""
    if not values:
        return "empty_array"
    if all(isinstance(item, dict) for item in values):
        return "dataset"
    if any(isinstance(item, dict) for item in values):
        return "mixed_array"
    if any(isinstance(item, list) for item in values):
        return "nested_array"
    return "primitive_array"


def discover(document: object) -> tuple[Finding, ...]:
    """Return deterministic discovery findings for a parsed JSON document.

    Traversal is depth-first and pre-order. Object values are visited in source
    key order and array elements by ascending index. Dataset arrays are recorded
    with their original record dictionaries and are not traversed further.
    """
    if not isinstance(document, (dict, list)):
        return (StructureFinding(pointer="", kind="unsupported_root"),)

    findings: list[Finding] = []
    # The boolean tracks whether an ancestor unsupported-array finding already
    # accounts for values in this branch. Continue traversing those branches to
    # find independent arrays/datasets, but avoid redundant scalar/object
    # findings beneath the array finding.
    stack: list[tuple[object, str, bool]] = [(document, "", False)]

    while stack:
        value, pointer, covered_by_array_finding = stack.pop()

        if isinstance(value, dict):
            if not value:
                if not covered_by_array_finding:
                    findings.append(
                        StructureFinding(pointer=pointer, kind="empty_object")
                    )
                continue
            for key, child in reversed(list(value.items())):
                stack.append(
                    (
                        child,
                        _object_child_pointer(pointer, key),
                        covered_by_array_finding,
                    )
                )
            continue

        if not isinstance(value, list):
            if not covered_by_array_finding:
                findings.append(
                    StructureFinding(pointer=pointer, kind="scalar_value")
                )
            continue

        kind = _classify_array(value)
        if kind == "dataset":
            findings.append(DatasetCandidate(pointer=pointer, records=tuple(value)))
            continue

        findings.append(StructureFinding(pointer=pointer, kind=kind))
        for index in range(len(value) - 1, -1, -1):
            stack.append(
                (
                    value[index],
                    _array_child_pointer(pointer, index),
                    True,
                )
            )

    return tuple(findings)
