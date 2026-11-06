# jxcel

`jxcel` converts JSON files containing tabular records into formatted Excel `.xlsx` workbooks. It finds datasets in general JSON documents, so you do not need to define a schema or write an exporter.

## Requirements and installation

JXCEL 0.1.2 requires Python 3.11 or later. Linux, Windows, and macOS are intended platforms. CI is configured to test Python 3.11–3.14 on all three platforms; this describes configured coverage, not passing results for a particular release. `openpyxl` is installed automatically.

```bash
python -m pip install jxcel==0.1.2
```

## Quick start

Save this as `orders.json`:

```json
{
  "orders": [
    {
      "id": "A-100",
      "customer": {"name": "Amina", "email": "amina@example.com"},
      "total": 42.5
    },
    {
      "id": "A-101",
      "customer": {"name": "Omar", "email": "omar@example.com"},
      "total": 18.0
    }
  ]
}
```

Convert it with:

```bash
jxcel orders.json
```

JXCEL writes `orders.xlsx` beside the input. The workbook contains an `orders` worksheet with columns such as `id`, `customer.name`, `customer.email`, and `total`. Its header is styled and frozen, and the worksheet has an auto-filter.

## What JXCEL discovers

A dataset is a non-empty array whose direct elements are all JSON objects. For example, the `/orders` array above is a dataset. JXCEL also searches nested objects and unsupported arrays for independent datasets. Once it finds a dataset, it treats its elements as records and does not search inside those records for more datasets.

Arrays that mix objects with other values, arrays of primitives, nested arrays, and empty arrays are reported as unsupported structures rather than exported as tables. A dataset made only of empty objects is discovered, but cannot be exported because it has no columns.

Each selected, exportable dataset becomes one worksheet. By default, `{"customer":{"name":"Amina"}}` produces a `customer.name` column. With `--no-flatten`, it stays in one `customer` column as compact JSON text: `{"name":"Amina"}`. In a literal key, `.` is encoded as `~1` and `~` as `~0`; this keeps literal names distinct from nested paths. This is JXCEL's column-name format, not JSON Pointer. Arrays inside records are written as compact JSON text, not expanded into additional rows.

Missing properties and `null` values become blank cells. Strings are not converted to numbers or dates, and formula-like strings remain text. Excel error literals such as `#N/A` and `#DIV/0!` are also preserved as text.

## Inspect and select data

Use these commands to review an input before creating a workbook:

```bash
jxcel orders.json --list
jxcel orders.json --preview --preview-rows 1
jxcel orders.json --dry-run -o planned-orders.xlsx
jxcel orders.json --preview --dry-run -o preview-plan.xlsx
```

- `--list` shows discovered datasets, columns, and unsupported structures without writing a workbook.
- `--preview` shows up to five records per dataset by default. Set `--preview-rows N` to change the limit.
- `--dry-run` validates and displays the conversion plan without writing an `.xlsx` file. It includes the output path, selected datasets and columns, and warnings.
- `--preview --dry-run` shows preview rows as part of the complete plan. `--list` cannot be combined with `--preview` or `--dry-run`.

Select one or more datasets with `--dataset` and a JSON Pointer. For example, select `/orders` and include only the customer name:

```bash
jxcel orders.json --dataset /orders --column-map '["/orders","customer.name"]' -o customer-names.xlsx
```

`--column-map` and `--exclude-map` each take a JSON array with the dataset pointer and exact column name. Both options can be repeated. Exclusions are applied after inclusions and take precedence over them:

```bash
jxcel orders.json --exclude-map '["/orders","customer.email"]' -o orders-without-email.xlsx
```

JSON Pointer tokens escape `~` as `~0` and `/` as `~1`. The mapping values are JSON strings, so punctuation and Unicode in column names can be represented without splitting the pair.

## Output and data limits

By default, the output uses the input filename with an `.xlsx` suffix. Use `-o PATH` to choose another path. JXCEL does not create output directories. If the output already exists, conversion fails unless you pass `--overwrite`:

```bash
jxcel orders.json -o report.xlsx --overwrite
```

JXCEL writes and validates a temporary workbook before replacing an existing output. Without `--overwrite`, publication requires hard-link support on the destination filesystem. If hard links are unavailable, JXCEL fails rather than using a method that could overwrite a file created concurrently. Replacement and durability behavior depend on the operating system and filesystem. Use an output directory you trust; JXCEL does not protect against a hostile local process that can change files in that directory.

Input must be UTF-8 JSON; an optional UTF-8 BOM is accepted. Malformed JSON, duplicate object keys, and non-standard values such as `NaN` and `Infinity` are rejected. Floating-point values outside Python's finite range are rejected, and very long integer tokens are subject to Python's configured conversion limit. Documents deeper than 500 nested arrays or objects are rejected.

Excel limits apply: a worksheet can contain at most 1,048,575 data rows and 16,384 columns, and a cell or column name can contain at most 32,767 UTF-16 code units. JXCEL reports invalid or oversized values instead of truncating them. Integers with an absolute value of at least `10^15` are written as text to avoid losing digits. Floats are written as numeric cells. `openpyxl` stores them with at most 16 significant digits, so values with more than 15 significant digits produce a `float_precision_limited` warning and the stored value may differ slightly from the JSON value. Values outside Excel's supported numeric range are rejected.

JXCEL loads the complete JSON document into memory, and export temporarily holds multiple representations of the data and workbook during writing and validation. Memory use depends on record count, field count, and value sizes; workbooks near Excel's row limit may not be practical on memory-constrained machines. Historical and current measurements used different environments and datasets, so they are not a controlled comparison and do not establish expected memory requirements or a regression result.

## License

JXCEL is distributed under the MIT License. See [LICENSE](https://github.com/AbuTaha7000D/jxcel/blob/main/LICENSE).
