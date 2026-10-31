# jxcel

`jxcel` is being prepared as a cross-platform command-line application for converting JSON files into formatted Excel `.xlsx` workbooks. It is intended for general JSON datasets, without fixed columns or domain-specific assumptions.

## Development status

**The converter is not functional yet.** The repository contains the project specification and a minimal command bootstrap. Strict JSON loading and validation are implemented in `src/jxcel/jsonio.py` (Phase 2), deterministic dataset discovery is implemented in `src/jxcel/discovery.py` (Phase 3), and schema planning and value preparation are implemented in `src/jxcel/planning.py` (Phase 4), but none are integrated into the command. Workbook generation and user-facing conversion options remain unimplemented. The declared command currently reports that the application is not implemented.

## Planned capabilities

- Discover tabular object arrays in JSON and report unsupported structures.
- Export datasets to separate worksheets with deterministic schemas.
- Preview datasets and inspect a conversion plan before writing.
- Protect existing output files and report data-preservation decisions.

These are planned features, not currently working functionality. See the [implementation roadmap](ROADMAP.md) for the complete specification and phase acceptance criteria.

## Platforms and dependencies

Linux, Windows, and macOS are intended targets. Cross-platform compatibility has not yet been verified. The planned runtime dependencies are the Python standard library and `openpyxl`. Python 3.11 or newer is required by the initial package metadata.

## Planned installation and usage

Once the CLI is functional and a package distribution is available, the recommended isolated installation method will be `pipx`. Publication has not occurred and package-name availability is not asserted. A local checkout may eventually be installed with `pipx install .`.

Planned command examples (not functional yet):

```text
jxcel data.json
jxcel -i data.json --output result.xlsx
jxcel data.json --list
jxcel data.json --dry-run
```

## Project guidance

- [AGENTS.md](AGENTS.md) contains repository-wide instructions for coding agents.
- [ROADMAP.md](ROADMAP.md) is the detailed product and implementation specification.

The existing `export_students.py` is a domain-specific legacy script and is not part of the planned general-purpose package.
