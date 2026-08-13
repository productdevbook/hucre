# Third-party fixtures

Not one of these was written by hucre.

That is the whole point. Before they existed, the suite parsed nothing
but hucre's own output, so a writer bug the reader mirrored was
invisible — and three of the defects fixed in the #439 round were exactly
that shape. See #464.

Two producers, two formats:

|            | `*.xlsx`                                        | `sheetjs-*.ods`                                        |
| ---------- | ----------------------------------------------- | ------------------------------------------------------ |
| Written by | ExcelJS (MIT)                                   | SheetJS Community Edition, the `xlsx` package (Apache-2.0) |
| Generator  | `scripts/fixtures/make-exceljs-fixtures.mjs`    | `scripts/fixtures/make-sheetjs-ods-fixtures.mjs`       |
| Read by    | `test/third-party-fixtures.test.ts`             | `test/ods-third-party.test.ts`                         |

The XLSX corpus came first, and left the ODS reader as the one that had
still never parsed a byte it did not write. The SheetJS half closed that.

## What they are, and what they are not

Both are independent implementations with their own element ordering,
their own defaults, and their own idea of what a minimal document
contains. Neither is **Excel, LibreOffice or Google Sheets**, and #464
asks for those too. What these give is the class of divergence a
golden-model test needs — bytes hucre did not write — from producers that
run anywhere and whose output is reproducible.

Two things SheetJS specifically will **not** emit, both of which a
LibreOffice corpus still would:

- `table:number-columns-repeated`, which LibreOffice uses for every run
  of like cells and is the sharpest trap in the format.
- error cells. SheetJS writes an error as an empty
  `<table:table-cell/>`, so there is no error in the file to read.

So the ODS half narrows the gap rather than closing it.

### It found one

The ODS corpus earned itself on the first run. A multi-line cell has two
spellings in ODF: `<text:line-break/>` inside one paragraph, which is
what hucre writes, or two `<text:p>` elements, which is what SheetJS and
LibreOffice write. `streamOdsRows` handled the first and ran the second
together — `"linebreak"` for a cell `readOds` read as `"line\nbreak"`.

A suite that only ever parsed hucre's own output could not see it,
because hucre never writes the spelling that breaks.

## Licensing

ExcelJS is MIT; SheetJS Community Edition is Apache-2.0. Every value
inside these files was written in the generator scripts in this
repository. Nothing is scraped and no third-party document is
redistributed.

## Regenerating

Neither producer is a devDependency — these are committed bytes, so
neither CI nor a contributor needs them installed:

```bash
mkdir -p /tmp/gen && cd /tmp/gen && npm init -y && npm i exceljs xlsx
node /path/to/hucre/scripts/fixtures/make-exceljs-fixtures.mjs /tmp/gen/node_modules
node /path/to/hucre/scripts/fixtures/make-sheetjs-ods-fixtures.mjs /tmp/gen/node_modules
```

Regenerating changes the bytes (timestamps, producer version). The tests
assert on _content_, not on bytes, so that is fine — but do not
regenerate casually, because the value of a fixture is that it stopped
moving.
