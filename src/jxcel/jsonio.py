"""Strict whole-document JSON loading for jxcel."""

from __future__ import annotations

import json
import math
import os
from pathlib import Path
from typing import Any


class JsonLoadError(Exception):
    """A path-aware error raised when a JSON document cannot be loaded."""


class _DuplicateKeyError(Exception):
    def __init__(self, key: str) -> None:
        self.key = key


class _NonStandardConstantError(Exception):
    def __init__(self, token: str) -> None:
        self.token = token


class _UnsupportedNumericRangeError(Exception):
    def __init__(self, token: str) -> None:
        self.token = token


class _IntegerConversionLimitError(Exception):
    def __init__(self, cause: ValueError) -> None:
        self.cause = cause


def _has_nonzero_significand(token: str) -> bool:
    """Return whether a JSON real-number token has a nonzero significand."""
    significand = token.lower().split("e", 1)[0]
    return any("1" <= character <= "9" for character in significand)


def _parse_float(token: str) -> float:
    """Parse a JSON real number without silently producing infinity or zero."""
    value = float(token)
    if not math.isfinite(value):
        raise _UnsupportedNumericRangeError(token)
    if value == 0.0 and _has_nonzero_significand(token):
        raise _UnsupportedNumericRangeError(token)
    return value


def _parse_int(token: str) -> int:
    """Use Python's configured integer conversion limit without overriding it."""
    try:
        return int(token)
    except ValueError as exc:
        # The JSON decoder has already validated the integer token grammar.
        # Preserve Python's configured digit limit and report it distinctly.
        raise _IntegerConversionLimitError(exc) from exc


def _reject_constant(token: str) -> None:
    raise _NonStandardConstantError(token)


def _unique_object(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise _DuplicateKeyError(key)
        result[key] = value
    return result


def load_json(path: str | os.PathLike[str]) -> object:
    """Load and validate one complete UTF-8 JSON document from *path*.

    A UTF-8 BOM is accepted. All valid JSON root categories are returned
    without discovery or exportability checks. JSON real-number tokens must
    convert to finite Python floats; overflow and nonzero-to-zero underflow
    are reported as unsupported numeric range.
    """
    input_path = Path(path)
    try:
        with input_path.open("r", encoding="utf-8-sig") as handle:
            document = handle.read()
    except UnicodeDecodeError as exc:
        raise JsonLoadError(
            f"Invalid UTF-8 in {input_path} at byte offset {exc.start}: {exc.reason}"
        ) from exc
    except OSError as exc:
        raise JsonLoadError(f"Could not read JSON file {input_path}: {exc}") from exc

    if not document.strip(" \t\r\n"):
        raise JsonLoadError(f"JSON file is empty: {input_path}")

    try:
        return json.loads(
            document,
            object_pairs_hook=_unique_object,
            parse_constant=_reject_constant,
            parse_float=_parse_float,
            parse_int=_parse_int,
        )
    except json.JSONDecodeError as exc:
        raise JsonLoadError(
            f"Invalid JSON in {input_path}: line {exc.lineno}, "
            f"column {exc.colno}: {exc.msg}"
        ) from exc
    except _DuplicateKeyError as exc:
        raise JsonLoadError(
            f"Duplicate object key {exc.key!r} in JSON file {input_path}"
        ) from exc
    except _NonStandardConstantError as exc:
        raise JsonLoadError(
            f"Non-standard JSON numeric constant {exc.token!r} in {input_path}"
        ) from exc
    except _UnsupportedNumericRangeError as exc:
        raise JsonLoadError(
            f"Unsupported numeric range for JSON number {exc.token!r} "
            f"in {input_path}"
        ) from exc
    except _IntegerConversionLimitError as exc:
        raise JsonLoadError(
            f"JSON integer exceeds Python's configured integer-string conversion "
            f"limit in {input_path}: {exc.cause}"
        ) from exc.cause
    except RecursionError as exc:
        raise JsonLoadError(
            f"JSON nesting exceeds the parser limit in {input_path}"
        ) from exc
    except ValueError as exc:
        raise JsonLoadError(
            f"Could not decode a JSON value in {input_path}: {exc}"
        ) from exc
