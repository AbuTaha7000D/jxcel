# jxcel

`jxcel` is a cross-platform command-line application for converting JSON files into formatted Excel `.xlsx` workbooks. It discovers tabular object arrays in general JSON documents; it has no fixed columns or domain-specific assumptions.

## Development status

The converter is functional. JSON loading and validation, dataset discovery, schema planning and value preparation, workbook generation with validated publication, and command-line integration are implemented in `src/jxcel/`. The package has not been published, and Windows and macOS behavior has not been verified.

Development checks have run on Linux with Python 3.14.8. See [ROADMAP.md](ROADMAP.md) for the specification, phase status, and known limitations.

## Requirements

- Python 3.11 or newer.
- Runtime dependencies: the Python standard library and `openpyxl`.
- Linux, Windows, and macOS are intended targets; only Linux has been exercised so far.

## Installation

A package distribution is not published yet. From a local checkout:

```bash
python -m pip install .
```

An isolated installation from the checkout also works with `pipx`:

```bash
pipx install .
```

Both commands install a `jxcel` command that points at `jxcel.cli:main`.

## Usage

Exactly one input source is required: positional `INPUT` or `-i/--input PATH`.

```bash
jxcel data.json
jxcel -i data.json
jxcel /path/to/data.json
jxcel -i /path/to/data.json -o result.xlsx
```

By default, the output is written beside the input with an `.xlsx` suffix, so `data.json` produces `data.xlsx`. The output parent directory must already exist; jxcel never creates directories.

Existing output files are preserved unless `--overwrite` is supplied:

```bash
jxcel data.json --overwrite
```

Without `--overwrite`, an existing destination is an error. With it, jxcel builds and validates a temporary workbook first, then replaces the destination through the approved publication path. jxcel does not claim crash durability or stronger atomicity than the operating system and filesystem provide.

## Inspection modes

The following modes never create a workbook, temporary workbook, or output directory. They are safe to run against an existing destination.

```bash
jxcel data.json --list
jxcel data.json --preview
jxcel data.json --preview --preview-rows 10
jxcel data.json --dry-run
jxcel data.json --preview --dry-run
```

- `--list` builds schemas, reports dataset pointers, record counts, available columns, effective columns after mappings, non-exportable reasons, and unsupported structures. It performs schema-only validation and does not scan cell values for exportability.
- `--preview` displays bounded records, respecting dataset and column selection. It prepares only the displayed rows and does not export-validate values outside the preview window. `--preview-rows` defaults to 5 and must be a positive integer.
- `--dry-run` performs full export validation and shows the selected datasets, effective columns, destination, warnings, and unsupported structures. It validates the destination read-only, so an existing file without `--overwrite` or a missing parent directory fails with exit status 1.
- `--preview --dry-run` combines both: the full export plan is built once and preview rows come from its validated snapshot.

`--list` cannot be combined with `--preview` or `--dry-run`. `--preview-rows` requires `--preview`. `--output` and `--overwrite` are rejected with `--list` or preview-only mode because they have no effect there; they are valid with `--dry-run` and normal conversion.

## Dataset and column selection

`--dataset` is repeatable and takes a canonical JSON Pointer:

```bash
jxcel data.json --dataset /students
jxcel data.json --dataset /students --dataset /orders
```

Unsupported structures remain visible even when only some datasets are selected. Scalar values and empty objects outside dataset boundaries are listed with their JSON Pointers, so input content that is not exported is visible.

`--column-map` and `--exclude-map` are repeatable and take a JSON array of exactly two strings: the dataset pointer and the exact planned column name. Use shell quoting appropriate to your platform.

Bash:

```bash
jxcel data.json \
  --column-map '["/students","user.name"]' \
  --column-map '["/orders","order,id"]' \
  --exclude-map '["/orders","internal=status"]'
```

PowerShell:

```powershell
jxcel data.json `
  --column-map '["/students","user.name"]' `
  --exclude-map '["/orders","internal=status"]'
```

Mappings are validated against every discovered dataset, but applied only to selected datasets. Include mappings select columns; exclude mappings remove columns afterward and take precedence. In `--dry-run`, a mapping for a valid but unselected dataset is reported as validated but unapplied. Unknown pointers and columns fail clearly.

## Nested values and flattening

Nested objects flatten by default. For example, `user.name` becomes a column named `user.name`. An empty nested object contributes no column in this mode and produces an `empty_object_omitted` warning with its path and record index. Use `--no-flatten` for one column per top-level property:

```bash
jxcel data.json --no-flatten
```

With `--no-flatten`, nested objects and arrays are written as compact JSON text. Flattening and `--no-flatten` use the documented reversible column-name encoding; conflicting flattened names fail rather than overwriting data.

## Output and exit status

- `stdout`: listings, previews, dry-run plans, warnings, unsupported findings, and the successful conversion summary.
- `stderr`: errors and argument-parser diagnostics.
- If the console cannot encode a character, jxcel escapes it; a closed output pipe exits with status 1 without a traceback.
- Exit `0`: successful help, version, listing, preview, dry-run, or conversion.
- Exit `1`: input, JSON, planning, mapping, destination, workbook, or publication failure.
- Exit `2`: command-line argument or usage error.

`--quiet` suppresses only the routine successful conversion summary. Errors, warnings, unsupported findings, and non-exportable dataset notices remain visible. `--quiet` and `--verbose` are mutually exclusive. `--verbose` adds context such as the input path and worksheet assignments.

`--list` and `--preview` treat the absence of an exportable dataset as informational and exit 0 after reporting it. `--dry-run` and normal conversion exit 1 because no workbook can be produced.

## Data preservation

- Strings remain strings; numeric-looking and date-like text is not reinterpreted.
- Integers with absolute value at least `10**15` are written as text with a warning.
- Floats remain numeric and are never rounded; values with more than 15 significant decimal digits produce a warning.
- Values outside Excel's supported numeric range, forbidden XML characters, and cells beyond Excel's text limit fail planning rather than being truncated.
- Formula-like strings are written as text, never as formulas.
- Null and missing values become blank cells.

Warnings are aggregated by code, dataset, and effective column, with an occurrence count and at most five sample record indexes.

## Limitations

- The complete JSON document and the export snapshot are held in memory; jxcel is not a streaming converter.
- Concurrent mutation of source records while a plan is being built is unsupported. Constructing or forging internal plan objects directly can bypass validation and is outside the supported contract.
- Windows and macOS command behavior, path handling, and publication behavior have not been verified.
- See [ROADMAP.md](ROADMAP.md) for deferred features and remaining hardening work.

## Project guidance

- [AGENTS.md](AGENTS.md) contains repository-wide instructions for coding agents.
- [ROADMAP.md](ROADMAP.md) is the detailed product and implementation specification.

The existing `export_students.py` is a domain-specific legacy script and is not part of the general-purpose package.
