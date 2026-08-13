# Third-party fixtures

Not one of these was written by hucre.

That is the whole point. Before they existed, the suite parsed nothing
but hucre's own output, so a writer bug the reader mirrored was
invisible — and three of the defects fixed in the #439 round were exactly
that shape. See #464.

Two producers, two formats:

|            | `*.xlsx`                       | `sheetjs-*.ods`                                            | `sheetjs-*.{xlsb,xls}`             |
| ---------- | ------------------------------ | ---------------------------------------------------------- | ---------------------------------- |
| Written by | ExcelJS (MIT)                  | SheetJS Community Edition, the `xlsx` package (Apache-2.0) | SheetJS                            |
| Generator  | `make-exceljs-fixtures.mjs`    | `make-sheetjs-ods-fixtures.mjs`                            | `make-sheetjs-binary-fixtures.mjs` |
| Read by    | `third-party-fixtures.test.ts` | `ods-third-party.test.ts`                                  | `xlsb-short-records.test.ts`       |

(Generators live in `scripts/fixtures/`, tests in `test/`.)

The XLSX corpus came first, and left the ODS reader as the one that had
still never parsed a byte it did not write. The SheetJS ODS half closed
that, and the binary half closed the last of it — `PROVENANCE.md` had
named the XLS and XLSB readers "the sharp end", and openpyxl writes
`.xlsx` only, so they had exactly one source until these.

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

### It found two

The ODS corpus earned itself on the first run. A multi-line cell has two
spellings in ODF: `<text:line-break/>` inside one paragraph, which is
what hucre writes, or two `<text:p>` elements, which is what SheetJS and
LibreOffice write. `streamOdsRows` handled the first and ran the second
together — `"linebreak"` for a cell `readOds` read as `"line\nbreak"`.

A suite that only ever parsed hucre's own output could not see it,
because hucre never writes the spelling that breaks.

The binary half earned itself on its first file too, and more sharply.
The XLSB reader handled the full-form cell records and none of the
`BrtShort*` forms — the same records with the column omitted, which is
the previous cell's plus one. Excel writes the full form every time;
SheetJS writes the short form for every cell after the first in a row.
A twelve-column sheet read back **one column wide**, with no error. See
`test/xlsb-short-records.test.ts`.

Regenerating the binary fixtures: SheetJS converts a `Date` to a serial
through _local_ time, so `sheetjs-dates.*` carries the offset of the
machine that wrote it. The tests compare the two readers against each
other rather than against an absolute instant, so regenerating elsewhere
changes the bytes without breaking anything.

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
node /path/to/hucre/scripts/fixtures/make-sheetjs-binary-fixtures.mjs /tmp/gen/node_modules
```

Regenerating changes the bytes (timestamps, producer version). The tests
assert on _content_, not on bytes, so that is fine — but do not
regenerate casually, because the value of a fixture is that it stopped
moving.
