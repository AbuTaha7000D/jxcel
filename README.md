# jxcel

`jxcel` is a cross-platform command-line application for converting JSON files into formatted Excel `.xlsx` workbooks. It discovers tabular object arrays in general JSON documents; it has no fixed columns or domain-specific assumptions.

## Development status

The converter is functional. JSON loading and validation, dataset discovery, schema planning and value preparation, workbook generation with validated publication, and command-line integration are implemented in `src/jxcel/`. The package has not been published.

Historical Phase 8 cross-platform validation passed on commit `f351669e6ed5ef072eb888e9a2f2f85401832210`: GitHub Actions run [41666703215](https://github.com/AbuTaha7000D/jxcel/actions/runs/41666703215) passed on Windows (`windows-2025`, X64) and macOS (`macos-15`, ARM64), with 25 acceptance checks on each. Those results predate the current local audit changes; current cross-platform release verification is pending.

Release notes for version `0.1.0` are in [RELEASE_NOTES.md](https://github.com/AbuTaha7000D/jxcel/blob/main/RELEASE_NOTES.md).

## Requirements

- Python 3.11 or newer is declared. Only Python 3.14 has been exercised so far: 3.14.8 on Linux and 3.14.10 on the historical Windows and macOS CI runners.
- Runtime dependencies: the Python standard library and `openpyxl`.
- Linux, Windows, and macOS are intended targets. All three have been exercised: Linux locally, and Windows and macOS on GitHub-hosted runners in Phase 8.

## Installation

A package distribution is not published yet. From a local checkout:

```bash
python -m pip install .
```

This command was verified on Linux in an isolated virtual environment and on the Windows and macOS CI runners in Phase 8. The build uses isolation, so it needs access to a package index or a populated build-dependency cache; an offline install fails if the required `setuptools` build dependency is unavailable.

An isolated installation from the checkout was also verified with `pipx`: 1.18.11 on Linux and 1.18.12 on the Windows and macOS CI runners.

```bash
pipx install .
```

Both commands install a `jxcel` command that points at `jxcel.cli:main`. Command availability, the isolated `pipx` command path, and execution from outside the checkout were verified on Linux, Windows, and macOS.

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

No-overwrite publication creates the destination with a same-directory hard link, so the temporary and destination files stay on one filesystem. Some filesystems do not support hard links, including FAT32/exFAT and some network or container mounts. On such filesystems, no-overwrite publication fails with exit status 1 and an explicit error. jxcel does not fall back to rename or replace, because that fallback could destroy a destination created concurrently. Overwrite publication uses the operating system's replace operation and depends on that filesystem's support. Hard-link behavior was exercised on Linux `tmpfs` and on the Windows and macOS runner temporary filesystems; other filesystem types, including FAT32/exFAT and network mounts, were not tested.

## Input validation

jxcel reads the input as UTF-8 and accepts an optional UTF-8 BOM. Empty or whitespace-only files, malformed JSON, invalid UTF-8, and trailing content after the JSON value are rejected with exit status 1 and a path-aware error. Malformed JSON reports the line and column; invalid UTF-8 reports the byte offset.

Object keys must be unique at every level. Duplicate keys are rejected even when one is written with an escape such as `\u0061`, because it decodes to the same key; the same key in separate objects is allowed. `NaN`, `Infinity`, and `-Infinity` are rejected as non-standard JSON constants. JSON real numbers must be finite: tokens that overflow a Python float or underflow a nonzero value to zero are rejected. Integers decode as Python integers, subject to Python's configured integer-string conversion limit.

The loader does not restrict the JSON root category: objects, arrays, strings, numbers, booleans, and `null` are all accepted, subject to the validation above. Dataset discovery then treats a root that is not an object or array as `unsupported_root`. Discovery reports an empty array outside dataset records as `empty_array`, and an empty object outside arrays as `empty_object`; neither is an exportable dataset. `--list` and `--preview` report these findings and exit 0; `--dry-run` and conversion exit 1 when no workbook can be produced.

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

Pointer tokens use JSON Pointer escaping: `~` is written as `~0` and `/` as `~1`. For example, a dataset under the key `a/b~c` has the pointer `/a~1b~0c`.

Unsupported structures remain visible even when only some datasets are selected. Scalar values and empty objects outside dataset boundaries are listed with their JSON Pointers, so input content that is not exported is visible.

`--column-map` and `--exclude-map` are repeatable and take a JSON array of exactly two strings: the dataset pointer and the exact planned column name. Use shell quoting appropriate to your platform.

Bash:

```bash
jxcel data.json \
  --column-map '["/students","user.name"]' \
  --column-map '["/orders","order,id"]' \
  --exclude-map '["/orders","internal=status"]'
```

PowerShell (verified on Windows in Phase 8):

```powershell
jxcel data.json `
  --column-map '["/students","user.name"]' `
  --exclude-map '["/orders","internal=status"]'
```

Mappings are validated against every discovered dataset, but applied only to selected datasets. Include mappings select columns; exclude mappings remove columns afterward and take precedence. In `--dry-run`, a mapping for a valid but unselected dataset is reported as validated but unapplied. Unknown pointers and columns fail clearly.

## Nested values and flattening

Nested objects flatten by default. For example, `user.name` becomes a column named `user.name`. An empty nested object outside an array contributes no column in this mode and produces an `empty_object_omitted` warning with its column path and sample record indexes. Use `--no-flatten` for one column per top-level property:

```bash
jxcel data.json --no-flatten
```

With `--no-flatten`, nested objects, arrays, and empty objects are written as compact JSON text (`{}` for an empty object, `[]` for an empty array); an empty object no longer produces an `empty_object_omitted` warning. Arrays never expand into rows or worksheets in either mode, and empty objects inside arrays are preserved as JSON text. Flattened column names use a reversible encoding of the original key path; conflicting flattened names fail rather than overwriting data. With `--no-flatten`, top-level keys are used as column names unchanged.

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
- Direct integer values with absolute value at least `10**15` are written as text with an `integer_precision_as_text` warning; smaller direct integers remain numeric.
- Direct float values remain numeric and are never rounded; values with more than 15 significant decimal digits produce a `float_precision_limited` warning, and direct finite floats outside Excel's numeric range fail export validation.
- Values written as JSON text (arrays in either mode, and nested objects under `--no-flatten`) produce a `nested_value_as_json_text` warning; numbers inside that text are not converted or range-checked individually.
- Export validation enforces Excel's limits: at most 1,048,575 data rows per dataset (the header uses row 1 of Excel's 1,048,576 rows), at most 16,384 columns, and at most 32,767 UTF-16 code units per cell or column header. Cell text and column headers reject control characters below `U+0020` other than tab, line feed, and carriage return, lone surrogates, and `U+FFFE`/`U+FFFF`; control characters inside JSON-serialized nested values are escaped. These failures are reported rather than truncated.
- Formula-like strings are written as text, never as formulas.
- Null and missing values become blank cells.

Warnings are aggregated by code, dataset, and effective column, with an occurrence count and at most five sample record indexes.

## Workbook output

Each selected exportable dataset becomes one worksheet, in plan order. Worksheet names are derived from dataset pointers: the root dataset is `root`, and pointer tokens are joined with dots. Names are sanitized for Excel, truncated to 31 UTF-16 code units, and made unique case-insensitively with `_2`, `_3`, and so on. The name `History` is reserved case-insensitively.

The header row is bold, filled, and vertically centered, with the header row frozen and an auto-filter over the data range. Column widths are based on the header and the first 100 data rows, with padding, and are clamped to 10–60 characters.

The workbook is written and validated through one open temporary-file stream before publication. On POSIX, the temporary file is mode `0600` while it is written and validated; its intended final mode is applied only after validation. With `--overwrite`, an existing regular file's POSIX permission bits are preserved; new files use the platform's normal file-creation mode. A permission failure before publication leaves the existing destination unchanged. Windows temporary files inherit the output directory's ACL; Windows `chmod` does not preserve general ACLs.

Choose an output directory whose access controls are appropriate for the workbook data. jxcel protects against accidental concurrent publishers: no-overwrite publication uses a same-directory hard link, and overwrite uses the operating system's replace operation. It does not defend against a malicious process running as the same user or another user/process with permission to change entries in the output directory. The temporary pathname is still used for publication and cleanup, so a hostile writable output directory is outside the security guarantee.

## Platform validation

Historical Phase 8 results are available in [GitHub Actions run 41666703215](https://github.com/AbuTaha7000D/jxcel/actions/runs/41666703215), validating commit `f351669e6ed5ef072eb888e9a2f2f85401832210`. That evidence predates the current audit changes.

| Platform | Status | Evidence |
|---|---|---|
| Linux (Fedora 44, kernel `7.2.13-200.fc44.x86_64`, Python 3.14.8, pip 26.0.1) | **Tested; exercised checks passed** | Isolated `python -m pip install .` and isolated `pipx 1.18.11 install .`; `jxcel` available on PATH; execution outside the checkout; `--version`/`--help`; relative, absolute, and Unicode paths; default and explicit output paths; existing-output and `--overwrite` behavior; missing/invalid output parents; input/output collision; repeated `--dataset`; `--column-map`/`--exclude-map`; preview, dry-run, and preview+dry-run; exit statuses 0, 1, and 2; no-write inspection modes; publication-failure cleanup. |
| Windows (`windows-2025`, X64, Python 3.14.10, pip 26.2.2, pipx 1.18.12) | **Tested; 25/25 acceptance checks passed** | Isolated virtual environment and `pipx` installs; command available outside the checkout; relative, absolute, and Unicode paths; `--version`/`--help`; existing-output and `--overwrite` behavior; missing/invalid output parents; input/output collision and symlink alias; repeated `--dataset`; `--column-map`/`--exclude-map`; preview, dry-run, and list; exit statuses 0, 1, and 2; no-write inspection modes; hard-link publication; publication race and temporary-file cleanup; Bash, PowerShell, and CMD JSON-pair quoting; native console encoding. 299 unit tests passed with 2 expected POSIX permission-bit skips. |
| macOS (`macos-15`, ARM64, Python 3.14.10, pip 26.2.2, pipx 1.18.12) | **Tested; 25/25 acceptance checks passed** | Shared cross-platform checks: isolated virtual environment and `pipx` installs; command available outside the checkout; relative, absolute, and Unicode paths; `--version`/`--help`; existing-output and `--overwrite` behavior; missing/invalid output parents; input/output collision and symlink alias; repeated `--dataset`; `--column-map`/`--exclude-map`; preview, dry-run, and list; exit statuses 0, 1, and 2; no-write inspection modes; hard-link publication; publication race and temporary-file cleanup; Bash JSON-pair quoting. macOS also verified preservation of POSIX mode `0o640` on overwrite. The Windows-only PowerShell, CMD, and native-console encoding checks were not run on macOS by design. 299 unit tests passed with no skips. |

Reproducible Linux install and smoke commands, with temporary directories:

```bash
python -m venv /tmp/jxcel-venv
/tmp/jxcel-venv/bin/python -m pip install .
cd /tmp
/tmp/jxcel-venv/bin/jxcel --version
/tmp/jxcel-venv/bin/jxcel --help
```

The Linux, Windows, and macOS Python matrix is automated by `.github/workflows/phase8-cross-platform.yml`, which is dispatched manually.

The full test suite uses the standard library. Bash:

```bash
PYTHONDONTWRITEBYTECODE=1 PYTHONPATH=src python -m unittest discover -s tests
```

PowerShell:

```powershell
$env:PYTHONDONTWRITEBYTECODE = "1"
$env:PYTHONPATH = "src"
python -m unittest discover -s tests
```

Windows CMD:

```bat
set PYTHONDONTWRITEBYTECODE=1
set PYTHONPATH=src
python -m unittest discover -s tests
```

## PyPI release process

The manually dispatched [PyPI workflow](.github/workflows/publish.yml) builds a wheel and source distribution. Its `publish` input defaults to `false`; publishing runs only when it is explicitly set to `true` on `main`. Do not enable that input without separate authorization to publish this release.

The workflow uses [PyPI Trusted Publishing](https://docs.pypi.org/trusted-publishers/) and does not store a PyPI token. Before first publication, configure the PyPI trusted publisher for owner `AbuTaha7000D`, repository `jxcel`, workflow `.github/workflows/publish.yml`, and GitHub environment `pypi`. Configure that GitHub environment with the required release reviewers. The project currently has no PyPI registration; PyPI's [trusted publisher project creation](https://docs.pypi.org/trusted-publishers/creating-a-project-through-oidc/) flow can register it during an explicitly authorized first publish.

## Limitations

- The complete JSON document and the export snapshot are held in memory; jxcel is not a streaming converter.
- JSON documents nested more than 500 object/array levels are rejected with a path-aware error. This fixed limit avoids interpreter-dependent parser behavior and excessive discovery paths; brackets inside strings do not count.
- Concurrent mutation of source records while a plan is being built is unsupported. Constructing or forging internal plan objects directly can bypass validation and is outside the supported contract.
- Windows and macOS behavior for the prior Phase 8 code was verified on GitHub-hosted `windows-2025` and `macos-15` runners; the current audit changes have not yet been rerun there.
- Hard-link publication depends on filesystem support and was exercised on Linux `tmpfs` and the Windows and macOS runner temporary filesystems; other filesystem types were not tested.
- Only Python 3.14 has been exercised: 3.14.8 on Linux and 3.14.10 on the Windows and macOS CI runners. Python 3.11–3.13 are declared as supported by package metadata but were not tested.
- The detailed roadmap is maintained locally and is not included in Git or package distributions.

## License

jxcel is licensed under the MIT License. See [LICENSE](LICENSE) for the full text.

## Project guidance

- [AGENTS.md](AGENTS.md) contains repository-wide instructions for coding agents.
- The detailed product roadmap is maintained locally and may not be present in a checkout.

The former `export_students.py` legacy exporter was removed during project cleanup. It is not part of jxcel or its distributions.
