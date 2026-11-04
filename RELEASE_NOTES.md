# jxcel 0.1.0 Release Notes

Version: `0.1.0`

Status: **initial release**. This version has not been published as a package distribution to PyPI or any other package index.

## Summary

`jxcel` is a cross-platform command-line application for converting JSON files into formatted Excel `.xlsx` workbooks. It discovers tabular object arrays in general JSON documents and has no fixed columns or domain-specific assumptions.

This release includes:

- Strict JSON loading with path-aware errors for malformed JSON, invalid UTF-8, duplicate object keys, and non-standard numeric constants.
- Deterministic dataset discovery, including reporting of unsupported structures and non-exportable datasets.
- Schema planning with nested-object flattening (or `--no-flatten`), column selection, and reversible column-name encoding.
- Dataset and column selection with repeatable `--dataset`, `--column-map`, and `--exclude-map` options.
- Inspection modes: `--list`, `--preview` with `--preview-rows`, and `--dry-run`, including `--preview --dry-run`.
- Data-preserving value handling with aggregated warnings and enforcement of Excel row, column, and cell-text limits.
- Workbook generation with worksheet naming, header formatting, frozen headers, auto-filters, and column widths.
- Validated publication: temporary workbook first, no-overwrite publication via a same-directory hard link, and `--overwrite` via the operating system's replace operation.
- User-facing output through `stdout` and errors through `stderr`, with exit statuses `0` (success), `1` (processing or publication failure), and `2` (usage error).

## Installation

No distribution is published. Install from a local checkout:

```bash
python -m pip install .
```

Or install in an isolated environment with `pipx`:

```bash
pipx install .
```

Both commands install a `jxcel` command that points at `jxcel.cli:main`. The build uses isolation, so it needs access to a package index or a populated build-dependency cache.

## Validation evidence

- 299 unit tests pass with the standard-library `unittest` suite.
- Linux: Python 3.14.8, exercised locally.
- Windows and macOS: Python 3.14.10 via GitHub Actions run `41666703215` at commit `f351669e6ed5ef072eb888e9a2f2f85401832210`.
- 25/25 acceptance checks passed on both Windows and macOS.
- The Windows run skipped 2 expected POSIX permission-bit tests; the macOS run had no skips.

## Known limitations

- The complete JSON document and the export snapshot are held in memory; jxcel is not a streaming converter.
- No-overwrite publication relies on hard links, so it depends on filesystem support. Filesystems without hard-link support fail with an explicit error rather than falling back.
- Python 3.11–3.13 are declared as supported by package metadata but were not tested. Only Python 3.14 has been exercised.
- There is no graphical user interface; jxcel is a command-line tool only.
- The package is not published, so there is no supported installation from a package index.
- Workbook contents were validated by reopening the generated file with openpyxl; opening the workbook in Excel or LibreOffice was not verified.
- No crash-durability guarantee stronger than the operating system and filesystem provide is claimed.

## License

jxcel is licensed under the MIT License. See [LICENSE](LICENSE).

## Legacy script

The former `export_students.py` legacy exporter was removed during project cleanup. It is not part of jxcel or its distributions.
