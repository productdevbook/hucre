# test/fixtures — workbooks hucre did not write

These twenty binaries exist because of [#464][]. Every other binary input
under `test/` is assembled byte-by-byte by the test that reads it, which
is a closed loop: a reader that misunderstands a record is checked
against a hand-built record that misunderstands it identically, and the
suite stays green. The XLS and XLSB readers are the sharp end — they
exist only to consume other tools' output and, until this directory, had
never seen any.

They are read by [`test/real-files.test.ts`](../real-files.test.ts).

[#464]: https://github.com/productdevbook/hucre/issues/464

## Producers

Two, on purpose. One producer is a single point of agreement: if hucre
and Excel happened to share a misreading of the spec, a corpus made only
of Excel output would agree with itself about it and stay green. openpyxl
was written by different people from an independent reading of ECMA-376,
and it emits four shapes Excel never does — formulas with no cached
result, ISO-8601 `t="d"` date cells, `date1904`, and inline strings with
no shared string table.

|             | Excel fixtures                                                                                              | openpyxl fixtures                                                               |
| ----------- | ----------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------- |
| Application | Microsoft Excel 16.0 (Microsoft 365, Windows 11 x64)                                                        | openpyxl 3.1.5 on CPython 3.14                                                  |
| Files       | `excel-*.{xlsx,xls,xlsb}`                                                                                   | `openpyxl-*.xlsx`                                                               |
| Driven by   | [`make-fixtures.vbs`](../../scripts/fixtures/make-fixtures.vbs) via `cscript.exe`, late-bound COM           | [`make-openpyxl-fixtures.py`](../../scripts/fixtures/make-openpyxl-fixtures.py) |
| Content     | synthetic, written for this corpus by the contributor — no third-party or confidential document is involved | same                                                                            |
| Licence     | same as the repository (MIT)                                                                                | same                                                                            |

openpyxl writes `.xlsx` only, so the `.xls` and `.xlsb` readers still see
Excel output alone.

Every Excel file was produced by Excel's own `SaveAs`. Nothing here was
post-processed, repacked or hand-edited; what is committed is the byte
stream each producer emitted.

## Regenerating

### The openpyxl half

```sh
python3 scripts/fixtures/make-openpyxl-fixtures.py test/fixtures
```

No Excel, no COM, no printer. Document timestamps are pinned so
regenerating does not churn `docProps/core.xml` — except
`dcterms:modified`, which openpyxl stamps at save time and gives the
caller no way to pin. It is not part of any golden model.

### The Excel half

```sh
# From WSL. Excel refuses to SaveAs to a \\wsl.localhost\ path, so write
# to a C:\ directory and copy afterwards.
mkdir -p /mnt/c/hucre-fixtures
cp scripts/fixtures/make-fixtures.vbs /mnt/c/hucre-fixtures/
cscript.exe //Nologo 'C:\hucre-fixtures\make-fixtures.vbs' 'C:\hucre-fixtures'
cp /mnt/c/hucre-fixtures/*.xls /mnt/c/hucre-fixtures/*.xlsx /mnt/c/hucre-fixtures/*.xlsb test/fixtures/
```

Two things bite:

- **The default printer must be reachable.** Every `PageSetup` write goes
  through the printer driver. With an offline default printer Excel
  raises 1004 on each one and pops a modal dialog that `DisplayAlerts`
  does not suppress, and `excel-pagesetup.xlsx` comes out with nothing
  but its print area. `Application.ActivePrinter` cannot be reassigned
  from inside Excel when this happens; change the OS default first:
  `(Get-WmiObject Win32_Printer -Filter "Name='Microsoft Print to PDF'").SetDefaultPrinter()`.
  The script reports any `MakePageSetup` failure at the end of the run.
- **`cscript` reads a `.vbs` in the system ANSI codepage.** A UTF-8
  source file gets its non-ASCII string literals mangled before Excel
  ever sees them — which once produced a fixture full of mojibake that
  hucre then read back perfectly faithfully. The script is therefore
  7-bit clean and builds every non-ASCII character with `ChrW`.

Regenerating does not reproduce the committed bytes: Excel stamps
timestamps and revision GUIDs on every save. That is fine — the fixtures
are committed artifacts and the script is documentation of how they were
made. CI never runs it; CI has no Excel.

## Scrubbing

openpyxl is the easy half: it stamps `openpyxl` as the creator, and
`make-openpyxl-fixtures.py` blanks `creator` and `lastModifiedBy`
outright. No personal identity is ever involved.

Excel is the awkward half. `Workbook.RemovePersonalInformation = True`
is set on every workbook before `SaveAs`. It blanks `dc:creator` and `cp:lastModifiedBy` in the
xlsx/xlsb `docProps/core.xml` and the author fields of the BIFF
`SummaryInformation` stream.

`Application.UserName` — the obvious lever, and the one the plan for this
work assumed — **does not do it**. Excel 16 takes the author from the
signed-in Office identity, not from that property; setting it still
produced `<dc:creator>` with a real name.

Verify after regenerating:

```sh
cd test/fixtures
for f in *.xlsx *.xlsb; do unzip -p "$f" docProps/core.xml | grep -o '<dc:creator>[^<]*'; done
for f in *.xls; do strings -el "$f"; strings "$f"; done | grep -i '<your name>'
```

`test/real-files.test.ts` also asserts `hasAuthor: false` for every
fixture — but that only bites for the `.xlsx` files. `readXls` and
`readXlsb` do not surface workbook properties at all (see the XLS/XLSB
table in `docs/PARITY.md`), so for the two `.xls` and two `.xlsb`
fixtures the assertion is vacuous and an authored file would still pass
it. Those four are covered by the `strings` grep above, which is a
manual step. Run it.

## The corpus

| file                    | what it is for                                                                                                                                                                                                            |
| ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `excel-basic.xlsx`      | strings, numbers, a negative, a leap day, booleans, formulas with cached numeric/string/error results                                                                                                                     |
| `excel-basic.xls`       | the same sheet as BIFF8                                                                                                                                                                                                   |
| `excel-basic.xlsb`      | the same sheet as XLSB — three containers, one authored sheet, so a reader that disagrees with its siblings shows up                                                                                                      |
| `excel-strings.xlsx`    | shared-string table: leading/trailing whitespace needing `xml:space="preserve"` (#441), a reused entry, XML metacharacters, an embedded newline, tab and NBSP, several scripts and a non-BMP codepoint, and a rich string |
| `excel-strings.xlsb`    | the same strings through the XLSB string table                                                                                                                                                                            |
| `excel-styled.xlsx`     | fonts, a fill, borders, alignment, number formats, and a format on a column that has no cells                                                                                                                             |
| `excel-layout.xlsx`     | merges across and down, a frozen pane, a conditional-formatting rule, a hidden row and a hidden column                                                                                                                    |
| `excel-pagesetup.xlsx`  | A3 landscape, fit-to-pages, a non-default margin, print area and print titles                                                                                                                                             |
| `excel-styleonly.xlsx`  | whole-row formatting with no values out to the right — the #492 shape                                                                                                                                                     |
| `excel-dates.xls`       | BIFF number formats: built-ins, a time, a date-time, thousands, percent, text-formatted digits, and CJK date codes                                                                                                        |
| `excel-empty.xlsx`      | one sheet, nothing in it                                                                                                                                                                                                  |
| `excel-chartsheet.xlsx` | a chart on its own tab, listed in `<sheets>` _before_ the worksheet — a sheet whose relationship type is not `worksheet` (#499)                                                                                           |
| `excel-sparse.xlsx`     | ~30 values placed out to column 15,312 — a 30.6M-slot bounding box from a 9 KB file (#501)                                                                                                                                |
| `excel-features.xlsx`   | named ranges (workbook- and sheet-scoped), an autofilter, a list validation, sheet protection and a cell comment — the features real workbooks use most                                                                   |

And from openpyxl, the second producer:

| file                           | what it is for                                                                                                                                     |
| ------------------------------ | -------------------------------------------------------------------------------------------------------------------------------------------------- |
| `openpyxl-basic.xlsx`          | the value cells of `excel-basic.xlsx`, authored again by the other producer — the two are read and compared against each other                     |
| `openpyxl-formulas.xlsx`       | formulas with **no cached result**, and an empty `<v/>` rather than an absent one. Excel always caches, so this shape is unreachable from Excel    |
| `openpyxl-isodates.xlsx`       | `t="d"` cells holding an ISO-8601 date (ECMA-376 `ST_CellType` 'd'), with a serial-number date under the same format as an in-fixture control      |
| `openpyxl-1904.xlsx`           | `date1904="1"`. The two epochs are 1,462 days apart, so a reader that drops the flag cannot fail quietly                                           |
| `openpyxl-inline-strings.xlsx` | `t="inlineStr"` with `<is><t>` and **no `sharedStrings.xml` at all** — the #441 whitespace case from a producer that actually emits inline strings |
| `openpyxl-styled.xlsx`         | the same style facets as `excel-styled.xlsx`, through openpyxl's different `styles.xml` ordering and defaults                                      |

192 KB of binaries, 303 KB for the whole directory including the golden
models and this file. Small on purpose: the point is coverage of shapes,
not of size. The two `.xls` files are 26 KB each and account for most of
it — BIFF8 has a floor no amount of trimming gets under; the six openpyxl
files are about 5 KB each.

### A second real corpus, and what it did not find

After the first triage, the same treatment was applied to 270 real
_business_ workbooks — invoices, offers, planning sheets — a completely
different population from the instrument data. Structure-only reporting
again: no cell values, no sheet names, nothing committed.

**All 270 read successfully. No new defects.** The only anomaly in the
whole set was six more `.xls` files with #494's `undefined` rows, two of
them with 79 holes each — evidence that went onto that issue and has
since been fixed by #509, which also settled the ragged-rows half.

What that corpus did surface is a different kind of gap. By frequency it
uses named ranges (43% of files), autofilter (30%), images (28%), cell
comments (18%), data validation (16%) and sheet protection (14%) — and
until `excel-features.xlsx` **not one of those had an Excel-written
fixture here.** hucre reads them all correctly; the point is that a
regression in any of them would have been caught only by files hucre
wrote itself, which is the exact hole #464 exists to close. Aiming a
corpus at the defects you already know about is a different thing from
aiming it at what people's files actually contain.

That is also why `excel-features.xlsx` is not tied to an open issue. It
is the only fixture here that documents something working.

### What it found

Six reader defects, in code that was at 98.8% coverage. All six are
fixed; each entry names the change that closed it:

- **[#493][]** — `readXlsx` does not apply XML line-ending normalization,
  so a newline Excel wrote as a literal CRLF comes back as `\r\n`. The
  same authored cell read from the XLSB fixture gives `\n`.
- **[#494][]** — `readXls` and `readXlsb` return ragged rows where
  `readXlsx` pads to the sheet width, for one sheet saved three ways, and
  leave `undefined` holes for rows with no cell records.
- **[#496][]** — ISO-8601 `t="d"` date cells are read as strings;
  `ST_CellType` 'd' has no case in the reader's cell-type switch.
  `openpyxl-isodates.xlsx` carries its own control: the same day as a
  serial, under the same number format, comes back as a `Date`.
- **[#497][]** — cached formula results that are strings or errors are
  dropped, so `readXlsx` → `writeXlsx` loses them. Only numeric results
  survive. Found by asserting the cached results in `excel-basic.xlsx`,
  which the openpyxl formula fixture made worth projecting.
- **[#499][]** — a workbook containing a chart sheet cannot be read _at
  all_: every ordinary worksheet in it is unreachable too.
- **[#501][]** — a sparse sheet is refused outright. `Sheet.rows` is a
  dense rectangle, so a read costs the bounding box rather than the cell
  count, and ~30 values placed out to column 15,312 describe 30.6M slots.

Each was recorded here first with `it.fails` and its issue number rather
than fixed in the corpus change, and each then got its own fix with its
own failing test, per `CONTRIBUTING.md`. Those markers are all gone now —
the assertions they guarded are ordinary passing tests, which is what a
`knownDefects` entry is for: it retires itself.

#493 and #494 came from the authored fixtures. **#499 came from pointing
hucre at a corpus of real instrument-exported workbooks** — 538 files
from battery-test equipment, private, never committed and never leaving
the machine they were read on. 61 of them failed to read, and 52 of those
61 were this one bug. The other nine were legitimately unreadable: six
password-protected, three with damaged ZIP central directories that
Info-ZIP and Python's `zipfile` reject as well. #501 came from the same
corpus, from a sheet Excel declares as `A1:VPX19959` holding 76,277
values across 507 scattered columns.

That corpus also confirmed #494's `undefined` holes are not a synthetic
edge case — 19 real `.xls` files from three different instrument vendors
produce them.

The lesson is worth writing down: the fixtures here are _authored_, so
they only contain shapes someone thought to author. A chart sheet is
completely ordinary and nobody thought of it. Pointing the reader at real
files, then reproducing whatever breaks as a new synthetic fixture, is
the loop that found it — and `make-fixtures.vbs <dir> <one-file-name>`
regenerates a single fixture so a new shape does not rewrite the other
twelve.

[#493]: https://github.com/productdevbook/hucre/issues/493
[#494]: https://github.com/productdevbook/hucre/issues/494
[#496]: https://github.com/productdevbook/hucre/issues/496
[#497]: https://github.com/productdevbook/hucre/issues/497
[#499]: https://github.com/productdevbook/hucre/issues/499
[#501]: https://github.com/productdevbook/hucre/issues/501

### What this corpus still does not cover

- **The CJK built-in BIFF format ids (27–36, 50–58) that #444 is about.**
  `excel-dates.xls` carries Japanese, Japanese-era and Korean date codes,
  but a non-CJK Windows Excel writes them as _custom_ `FORMAT` records
  with ids ≥ 164 — checked by dumping the BIFF `FORMAT` records of a
  probe file, which showed 164–169 and no built-in in the CJK range.
  Reaching those ids needs an Excel running under a CJK locale.
- **LibreOffice and Google Sheets**, both named in #464. Adding a third
  producer costs one more script and a handful of goldens; nothing about
  the layout here has to change to take one.
- **`.xls` and `.xlsb` from anything but Excel.** openpyxl writes `.xlsx`
  only, so the two binary readers still have exactly one producer between
  them. That is the thinnest part of the corpus, and LibreOffice — which
  writes `.xls` — is the obvious way to thicken it.
