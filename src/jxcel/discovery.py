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
    stack: list[tuple[object, str]] = [(document, "")]

    while stack:
        value, pointer = stack.pop()

        if isinstance(value, dict):
            for key, child in reversed(list(value.items())):
                stack.append((child, _object_child_pointer(pointer, key)))
            continue

        if not isinstance(value, list):
            # Scalar values outside the root never produce findings.
            continue

        kind = _classify_array(value)
        if kind == "dataset":
            findings.append(DatasetCandidate(pointer=pointer, records=tuple(value)))
            continue

        findings.append(StructureFinding(pointer=pointer, kind=kind))
        for index in range(len(value) - 1, -1, -1):
            stack.append((value[index], _array_child_pointer(pointer, index)))

    return tuple(findings)
