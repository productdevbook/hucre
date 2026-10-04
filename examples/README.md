# Excel workbooks

Download the files in [`workbooks/`](workbooks/) to inspect ordinary import
and export scenarios. These are synthetic, project-authored examples;
they contain no customer data or personal contact details. Their content
uses the repository's MIT licence.

| File                                                   | Scenario                                               |
| ------------------------------------------------------ | ------------------------------------------------------ |
| [invoice.xlsx](workbooks/invoice.xlsx)                 | Quantities, prices, cached formulas and totals         |
| [budget.xlsx](workbooks/budget.xlsx)                   | Budget, actuals, zero and negative variance            |
| [inventory.xlsx](workbooks/inventory.xlsx)             | Text SKUs, booleans, a table, filtering and validation |
| [sales.xlsx](workbooks/sales.xlsx)                     | Revenue, profit and percentage formulas                |
| [employees.xlsx](workbooks/employees.xlsx)             | Unicode names, dates and booleans                      |
| [projects.xlsx](workbooks/projects.xlsx)               | List validation and conditional formatting             |
| [attendance.xlsx](workbooks/attendance.xlsx)           | True, false, blank and zero                            |
| [expenses.xlsx](workbooks/expenses.xlsx)               | Currency precision, refunds and multiline text         |
| [unicode.xlsx](workbooks/unicode.xlsx)                 | Several scripts, emoji, whitespace and XML characters  |
| [dates-and-times.xlsx](workbooks/dates-and-times.xlsx) | Leap day, timestamps and fractional days               |
| [merged-report.xlsx](workbooks/merged-report.xlsx)     | Merged heading, empty row and totals                   |
| [multi-sheet.xlsx](workbooks/multi-sheet.xlsx)         | Cross-sheet formulas and escaped sheet names           |

## Provenance and expectations

These bytes were exported with `@oai/artifact-tool` 2.8.71 from the
literal inputs in [`scenarios.json`](scenarios.json). The export identifies
openpyxl in its document properties. They were not written by hucre and
have not been opened in Microsoft Excel. Every sheet was rendered for
visual review, formulas were recalculated and checked against the authored
results, and the saved files were read independently with openpyxl 3.1.5.

The **Microsoft Excel-written** XLSX, XLS and XLSB fixtures remain in
[`test/fixtures/`](../test/fixtures/PROVENANCE.md). ExcelJS, SheetJS and
LibreOffice fixtures live in its [`third-party/`](../test/fixtures/third-party/README.md)
subdirectory. Producer names are kept distinct so generated examples do
not claim native Excel verification.

[`manifest.json`](manifest.json) pins SHA-256 checksums. Expectations are
authored data, not captured reader output. The same scenarios drive the
normal reader, format detection, streaming, `writeXlsx` and `saveXlsx`
tests in [`examples.test.ts`](../test/integration/examples.test.ts).

## Regeneration

CI uses the committed files and never installs a producer. Regenerate
deliberately in an environment with the pinned artifact-tool producer:

```sh
node scripts/fixtures/make-examples.mjs /absolute/path/to/hucre /tmp/hucre-previews
bun run test:integration
```

The producer is optional tooling, not a runtime or development dependency
of hucre. In Codex it is provided by the bundled workspace runtime; run a
copy of the builder beside that runtime's `node_modules` symlink. The
builder takes the repository path as its first argument and a temporary
preview directory as its second. It also updates the checksum manifest.
Review all changed bytes and previews before committing. Export timestamps
can change, so byte-for-byte regeneration is not promised.
