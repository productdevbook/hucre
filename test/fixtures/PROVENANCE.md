# test/fixtures — workbooks hucre did not write

These eleven binaries exist because of [#464][]. Every other binary input
under `test/` is assembled byte-by-byte by the test that reads it, which
is a closed loop: a reader that misunderstands a record is checked
against a hand-built record that misunderstands it identically, and the
suite stays green. The XLS and XLSB readers are the sharp end — they
exist only to consume other tools' output and, until this directory, had
never seen any.

They are read by [`test/real-files.test.ts`](../real-files.test.ts).

[#464]: https://github.com/productdevbook/hucre/issues/464

## Producer

|             |                                                                                                                    |
| ----------- | ------------------------------------------------------------------------------------------------------------------ |
| Application | Microsoft Excel 16.0 (Microsoft 365, Windows 11 x64)                                                               |
| Driven by   | [`scripts/fixtures/make-fixtures.vbs`](../../scripts/fixtures/make-fixtures.vbs) via `cscript.exe`, late-bound COM |
| Content     | synthetic, written for this corpus by the contributor — no third-party or confidential document is involved        |
| Licence     | same as the repository (MIT)                                                                                       |

Every file was produced by Excel's own `SaveAs`. Nothing here was
post-processed, repacked or hand-edited; what is committed is the byte
stream Excel emitted.

## Regenerating

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

`Workbook.RemovePersonalInformation = True` is set on every workbook
before `SaveAs`. It blanks `dc:creator` and `cp:lastModifiedBy` in the
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

| file                   | what it is for                                                                                                                                                                                                            |
| ---------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `excel-basic.xlsx`     | strings, numbers, a negative, a leap day, booleans, formulas with cached numeric/string/error results                                                                                                                     |
| `excel-basic.xls`      | the same sheet as BIFF8                                                                                                                                                                                                   |
| `excel-basic.xlsb`     | the same sheet as XLSB — three containers, one authored sheet, so a reader that disagrees with its siblings shows up                                                                                                      |
| `excel-strings.xlsx`   | shared-string table: leading/trailing whitespace needing `xml:space="preserve"` (#441), a reused entry, XML metacharacters, an embedded newline, tab and NBSP, several scripts and a non-BMP codepoint, and a rich string |
| `excel-strings.xlsb`   | the same strings through the XLSB string table                                                                                                                                                                            |
| `excel-styled.xlsx`    | fonts, a fill, borders, alignment, number formats, and a format on a column that has no cells                                                                                                                             |
| `excel-layout.xlsx`    | merges across and down, a frozen pane, a conditional-formatting rule, a hidden row and a hidden column                                                                                                                    |
| `excel-pagesetup.xlsx` | A3 landscape, fit-to-pages, a non-default margin, print area and print titles                                                                                                                                             |
| `excel-styleonly.xlsx` | whole-row formatting with no values out to the right — the #492 shape                                                                                                                                                     |
| `excel-dates.xls`      | BIFF number formats: built-ins, a time, a date-time, thousands, percent, text-formatted digits, and CJK date codes                                                                                                        |
| `excel-empty.xlsx`     | one sheet, nothing in it                                                                                                                                                                                                  |

135 KB of binaries, 149 KB for the whole directory including the golden
models and this file. Small on purpose: the point is coverage of shapes,
not of size. The two `.xls` files are 26 KB each and account for most of
it — BIFF8 has a floor no amount of trimming gets under.

### What it found

Two reader defects, on the first pass, in code that was at 98.8% coverage:

- **[#493][]** — `readXlsx` does not apply XML line-ending normalization,
  so a newline Excel wrote as a literal CRLF comes back as `\r\n`. The
  same authored cell read from the XLSB fixture gives `\n`.
- **[#494][]** — `readXls` and `readXlsb` return ragged rows where
  `readXlsx` pads to the sheet width, for one sheet saved three ways.

Both are recorded in `test/real-files.test.ts` with `it.fails` and their
issue number, not fixed there. Each fix belongs in its own change with
its own failing test first, per `CONTRIBUTING.md`.

[#493]: https://github.com/productdevbook/hucre/issues/493
[#494]: https://github.com/productdevbook/hucre/issues/494

### What this corpus still does not cover

- **The CJK built-in BIFF format ids (27–36, 50–58) that #444 is about.**
  `excel-dates.xls` carries Japanese, Japanese-era and Korean date codes,
  but a non-CJK Windows Excel writes them as _custom_ `FORMAT` records
  with ids ≥ 164 — checked by dumping the BIFF `FORMAT` records of a
  probe file, which showed 164–169 and no built-in in the CJK range.
  Reaching those ids needs an Excel running under a CJK locale.
- **LibreOffice and Google Sheets**, both named in #464. Everything here
  is Excel. A second producer is the obvious next contribution, and the
  layout of this directory does not have to change to take one.
