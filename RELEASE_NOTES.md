# JXCEL release notes

## 0.1.2

Performance, compatibility, and release-safety release.

- Fixed quadratic workbook validation scaling; dimension reads are now constant per worksheet.
- Excel error-literal strings (`#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`) are preserved as text instead of being rejected.
- High-precision floats are accepted as numeric cells using `openpyxl`'s 16-significant-digit serialization. The `float_precision_limited` warning remains, and `--dry-run` agrees with real export.
- CI and release safety: the cross-platform workflow now runs on pull requests and pushes to `main`, supports reusable invocation, reads the expected version from project metadata, and pins third-party actions to commit SHAs. Publishing now requires a preflight (dispatch from `main`, tag exists, tag is an ancestor of `main`, tag version matches both version files) and a successful cross-platform verification before build and publish.
- Documentation: clarified memory characteristics and the incomparable historical measurements, documented the float precision contract, fixed the LICENSE link, and removed the hard-coded version from the CLI test.

## 0.1.1

Documentation and packaging metadata maintenance release. The README now explains installation, dataset discovery, flattening, selection, inspection modes, output behavior, and key limits for first-time users. The PyPI long description now reflects the updated README. No runtime behavior changes are included.

## 0.1.0

Initial published release of JXCEL, a cross-platform command-line tool for converting JSON datasets into formatted Excel workbooks. It introduced strict JSON validation, deterministic dataset discovery, nested-object flattening, dataset and column selection, list/preview/dry-run modes, Excel-aware value validation, and validated workbook publication.

JXCEL is distributed under the MIT License.
