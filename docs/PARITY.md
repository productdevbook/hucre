# Read/write parity

What hucre reads, what it writes, and where those differ.

This exists because a gap you discover is worse than a gap you were told
about. v1 makes the public API a stability commitment, and part of that
commitment is being honest about the edges: everything below is either at
parity or listed here as an exception.

If you hit a loss that is **not** on this page, that is a bug — please
open an issue. That is the whole point of writing it down.

## The one thing to read first

XLSX has **two** write paths, and they have different fidelity. Most
confusion about what hucre preserves comes from conflating them.

(Both, and the streaming writers, now serialize a cell through one
implementation — so an error value, `xml:space` handling and formula
result typing mean the same thing whichever writer you use. What differs
between the paths is what the _model_ can express, not how a cell is
written.)

|                | entry points             | behaviour                                                                                                                         |
| -------------- | ------------------------ | --------------------------------------------------------------------------------------------------------------------------------- |
| **Authoring**  | `readXlsx` / `writeXlsx` | the workbook is rebuilt from the model. Only what `WriteSheet` and `WriteOptions` describe comes out                              |
| **Round-trip** | `openXlsx` / `saveXlsx`  | parts hucre does not regenerate are copied byte-for-byte, with relationships and content types re-declared so they stay reachable |

So `readXlsx` → `writeXlsx` on a workbook with charts, macros or pivot
tables gives you a workbook without them. `openXlsx` → `saveXlsx` on the
same file keeps them, whether or not hucre understands them.

Pick the round-trip path when you are **editing someone's file**. Pick the
authoring path when you are **producing a new one**.

### Getting from one model to the other

`readXlsx` returns a `Workbook` and `writeXlsx` takes `WriteOptions`, and
neither is assignable to the other — `Chart` is not `SheetChart`, and
`PivotTable` is not `WritePivotTable`. `toWriteOptions` converts, drops
what the authoring model has no field for, and tells you what went:

```ts
import { readXlsx, toWriteOptions, writeXlsx } from "hucre"

const wb = await readXlsx(bytes)
wb.sheets[0].rows[0][0] = "edited"

const out = await writeXlsx(
  toWriteOptions(wb, {
    onDrop: ({ field, sheet, reason }) => console.warn(`dropped ${field}`, sheet, reason),
  }),
)
```

It drops exactly the fields listed below — `slicers`, `timelines`,
`threadedComments`, `charts`, `pivotTables` per sheet, and `themeColors`,
`externalLinks`, `cellImages`, `persons`, `pivotCaches`, `slicerCaches`,
`timelineCaches` per workbook — and `test/write-model.test.ts` derives
that set from the types, so a new field with no counterpart fails until
someone decides.

## XLSX

### Read + write, at parity

Cell values and types, formulas (shared, array and dynamic), rich text,
hyperlinks with tooltips, comments, checkboxes, the full cell style model
(fonts, pattern and gradient fills, borders including diagonal, every
alignment field, number formats, protection), merges, data validations,
all 15 conditional-rule types **including their dxf styles**, auto-filters
with value filters, freeze and split panes, sheet protection, every
attribute of `CT_PageSetup` bar one — print areas, print titles, every
OOXML paper size (by name where hucre has one, by code otherwise) plus
custom `paperWidth` / `paperHeight` for the sizes that have none, page
order, first page number, draft and black-and-white, comment and error
printing, copies, DPI and printer defaults; the exception is `r:id`,
which points at a binary printer-settings part and has no portable
meaning — headers and footers, sheet views
and tab colours, hidden and very-hidden sheets, tables, row and column
definitions, manual page breaks, outline properties, sparklines, text
boxes, background images, images, document properties, named ranges, the
1904 date system, workbook protection, theme colours, Excel 2024
checkboxes, and encryption.

### Formula results are write-only in practice

`writeXlsx` writes the `formulaResult` you give it into `<v>`, so the
value is genuinely in the file. But hucre has no formula engine, so it
cannot know whether that cached value still matches the formula — and a
stale `<v>` is a workbook that shows the wrong number. Every workbook
hucre writes therefore carries `calcPr fullCalcOnLoad="1"`, which tells
Excel to recalculate everything on open and discard the cached results.

The practical consequence: **`formulaResult` survives a hucre round-trip
and does not survive being opened in Excel.** Set it when something other
than Excel will read the file — a second tool, a diff, hucre itself — and
do not rely on it as the number a person will see.

### Numbers are written at full precision, not Excel's 15 digits

`1e21` is written `1e+21` and `0.1 + 0.2` is written
`0.30000000000000004` — seventeen significant digits, where Excel writes
`1E+21` and caps its _display_ at fifteen. Both spellings are conformant
`xsd:double`; the lexical space is `[Ee](\+|-)?[0-9]+`, so a lowercase
`e` is correct rather than merely tolerated.

The precision is the part that matters. `String(value)` is the only
spelling that round-trips a double exactly, and every extreme does:
`1e-7`, `1e300`, `Number.MIN_VALUE`, `Number.EPSILON`,
`Number.MAX_SAFE_INTEGER`. Rounding to fifteen digits to look more like a
file Excel wrote would turn `0.1 + 0.2` into `0.3` — a number the caller
did not write. A library for moving data faithfully should not make that
trade, and `test/number-serialisation.test.ts` pins it so the change
cannot land quietly.

The one loss is negative zero: `String(-0)` is `"0"`, and Excel has no
signed zero either, so the sign goes.

CSV has the same guarantee. It prefers the plain decimal form where Excel
would otherwise show `1E-07` — but only when that form is the same
number, which it was not for the smallest values (`Number.MIN_VALUE` used
to come out as `0.0`).

### Ranges: two spellings, one meaning

Ranges are A1 strings on `DataValidation.range`, `ConditionalRule.range`,
`AutoFilter.range`, `TableDefinition.range`, `NamedRange.range`,
`PageSetup.printArea` and `ReadOptions.range`, and coordinate objects on
`MergeRange`, `SheetImage.anchor` and the sparkline fields. There was no
rule to hold in your head about which a field wanted.

The authoring surfaces that take a rectangle now take either —
`WriteSheet.merges`, `XlsxStreamWriter`'s `merges`, and `copyRange` — and
`toRange` / `toRanges` are exported for anywhere else. `Sheet.merges`
stays coordinates, because that is what the reader produces and a read
model with two spellings would push the normalising onto every consumer.

`SheetImage.anchor` is a different shape — a corner plus an optional
second corner, not a rectangle — and is unchanged.

`test/xlsx-write-read-parity.test.ts` holds every field of `WriteSheet`
and `WriteOptions` in a register typed over `keyof Required<…>`. Adding a
field to either interface fails `tsc` until it is registered — as a probe
that round-trips, or as a one-way entry with its reason. That register,
not this list, is the thing that stays current.

### Styles read out of a file are shared, not copied

`readXlsx(…, { readStyles: true })` gives every cell its own `CellStyle`
wrapper, but the `font`, `fill`, `border`, `alignment` and `protection`
objects inside it are **the parsed records themselves** — one per distinct
format in `xl/styles.xml`, referenced by every cell that uses it. So

```ts
cells.get("0,0").style.font === cells.get("5,3").style.font // true, same format
```

and writing through one changes every cell that shares it. Copying per
cell nearly doubles peak memory on a styled read — 407 MB against 787 MB
over 720,000 styled cells — for a guarantee most callers never need,
since a resolved style is normally read and not written through.

Use `cloneCellStyle` before editing one cell's format:

```ts
import { cloneCellStyle } from "hucre"

const mine = cloneCellStyle(cell.style!)
mine.font!.bold = true
cell.style = mine
```

The same holds for a conditional rule's `style`, which is the workbook's
`<dxf>` record shared by every rule pointing at it.

### Read and round-trip only — no authoring API

These are parsed into the model and preserved through `openXlsx` →
`saveXlsx`, but there is no way to create one from scratch:

|                                         | model field                                                                 |
| --------------------------------------- | --------------------------------------------------------------------------- |
| Slicers and their caches                | `Sheet.slicers`, `Workbook.slicerCaches`                                    |
| Timeline filters and their caches       | `Sheet.timelines`, `Workbook.timelineCaches`                                |
| Threaded comments and their person list | `Sheet.threadedComments`, `Workbook.persons`                                |
| External workbook links                 | `Workbook.externalLinks`                                                    |
| WPS DISPIMG cell images                 | `Workbook.cellImages`                                                       |
| Theme colours from the file             | `Workbook.themeColors` — `writeXlsx` always emits the standard Office theme |

`WriteSheet` has no fields for these, deliberately: a typed field that is
silently discarded is worse than no field at all, which is why
`WriteSheet.threadedComments` was removed rather than left in place.

### Charts

16 chart kinds are read and round-tripped; **7 can be authored** (bar,
column, line, pie, scatter, area, doughnut). `bar3D`, `line3D`, `pie3D`,
`area3D`, `bubble`, `radar`, `surface`, `surface3D`, `stock` and `ofPie`
read and survive `saveXlsx`, but `SheetChart` cannot express them.

That last sentence has one exception, and it follows from the format: a
worksheet carries exactly one `<drawing>` element, so images and charts on
a sheet have to share one drawing part. When a sheet has hucre-managed
images, hucre regenerates that part — and can only put back the charts it
can author. So on **a sheet that also has images**, a chart of one of the
nine unauthorable kinds is dropped; the seven authorable kinds survive
with their anchors. A chart on a sheet hucre leaves alone is preserved as
raw bytes regardless of kind, which is the common case. See #465.

### Pivot tables

Read and written, but **the two types are disjoint**. `PivotTable` (what
the reader produces: layout, fields, cache id) is not accepted by the
writer, and `WritePivotTable` (source range plus rows/columns/pages/values)
is not what the reader returns. A pivot table therefore cannot be
round-tripped _through the model_ — `openXlsx`/`saveXlsx` preserves it as
a raw part instead. The writer also emits the pivot structure without
pre-computed value cells; Excel computes them on open.

`Workbook.pivotCaches` follows from that: it is the read model's view of
the workbook-level caches, and has no write counterpart because
`WritePivotTable` builds its own cache from the source range.

### Write-only, by design

| field                                                     | why                                                                                                                                                                        |
| --------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `WriteOptions.vbaProject`                                 | attaches a macro project; `Workbook` has no counterpart, so `readXlsx` → `writeXlsx` **drops macros silently**. Use `openXlsx`/`saveXlsx` to edit a macro-enabled workbook |
| `WriteOptions.encryption`                                 | a property of the container, not of the workbook. Read back with `ReadOptions.password`                                                                                    |
| `WriteOptions.stringMode`                                 | an encoding choice with nothing to surface on read                                                                                                                         |
| `WriteSheet.data`                                         | the object form of `rows`; comes back as `rows`                                                                                                                            |
| `WriteSheet.a11y`                                         | authoring metadata. `a11y.summary` is promoted to `properties.description` and does survive; `a11y.headerRow` has no cell to live in                                       |
| `SheetProtection.password`, `workbookProtection.password` | the file holds a one-way digest, never the password. The digest is read; the password cannot be                                                                            |

### Losses inside fields that otherwise round-trip

Three cases where a field is carried but a particular _value_ is not.

**A cell whose text is literally `_x0041_` reads back as `A`.** OOXML uses
`_xHHHH_` to encode characters XML cannot hold, and hucre decodes it on
read. Excel disambiguates by escaping a leading underscore as `_x005F_`;
hucre deliberately does not, because doing so would mangle the far more
common case of ordinary text that happens to contain an underscore. The
ambiguity is accepted, and now written down.

**Serial 60 collapses onto 59.** Serial 60 in the 1900 system is the
Lotus 1-2-3 phantom 29 February 1900, a date that does not exist. It has
no instant to map to, so `serialToDate(60)` gives 28 February 1900 —
the same as serial 59 — and writing that back produces 59. A workbook
containing serial 60 shifts by one day on rewrite. Excel produces it only
for files inherited from Lotus.

**A `Date` you build with local components is converted as an instant.**
`dateToSerial(new Date(2024, 0, 15))` in UTC+3 is 45305.875, not 45306,
because local midnight is 21:00 the previous day in UTC. hucre reads UTC
components everywhere, which is what keeps the readers, the writers and
`formatValue` consistent; it does not infer a calendar day from a
timezone. Use `Date.UTC(...)` when you mean a day.

## ODS

ODS reads and writes the same narrow model, so **ODS → ODS is lossless**.
The loss is in conversion _into_ ODS from a format that models more.

Carried: cell values, formulas (including cross-sheet references), merges,
hyperlinks on any cell type, rich text, multi-section number formats,
document properties (six fields), and six style facets — bold, italic,
font size, font colour, background colour, number format.

Not modelled in **either** direction: borders, alignment, font name,
underline, strikethrough, column widths, row heights, hidden rows and
columns, freeze and split panes, data validation, conditional formatting,
auto-filter, named ranges, tables, images, page setup, sheet protection,
tab colour, hidden sheets, and `time` cells (which read back as the raw
ISO duration string).

The reader opens `content.xml` and `meta.xml`. It does not open
`styles.xml` or `settings.xml`, which is where LibreOffice keeps named and
default cell styles and all page setup — so a LibreOffice-authored file
reads back with its direct formatting only.

See [What ODS carries](../README.md#what-ods-carries) for the consequences
worth knowing before relying on it.

## `Sheet.rows` is not guaranteed rectangular

Every reader returns `rows: CellValue[][]`, and they do not agree on the
shape of an empty row. `readXlsx` pads to the sheet's bounding box, so an
all-empty row comes back as `[null, null, …]`; `readOds` returns `[]` for
it, and `parseCsv` returns whatever the file had, so a short line stays
short.

Code that walks a sheet generically — `sheetToObjects`, `toHtml`,
`toMarkdown`, `a11y.audit`, the schema validator — has to read
`row[i] ?? null` rather than assume a slot exists. That is what they all
do; it is written down here because nothing said so.

The streaming readers _do_ agree: `streamXlsxRows` and `streamOdsRows`
both skip an entirely empty row and keep the true index on `StreamRow`,
so a gap in the indexes is the signal.

## The readers are lenient, and will tell you

A corrupt reference does not throw. A cell pointing at a shared string
the file does not have reads as `null`; a cell naming a format that is
not there reads unstyled. That is the right default for a format you
receive rather than control — half a sheet is usually more useful than an
exception.

It used to be the only mode, which made a damaged file indistinguishable
from a clean one. `ReadOptions.onWarning` is the other half:

```ts
const warnings: ReadWarning[] = []
const wb = await readXlsx(bytes, { onWarning: (w) => warnings.push(w) })

// unresolved-shared-string — Cell A1 points at shared string 9999, which
// the file does not have (3 present). Read as empty.
```

Each warning carries a `code`, a sentence, and the sheet and cell it came
from. Nothing changes when the option is omitted, and a file hucre wrote
produces none.

| code                       | what silently went missing                                    |
| -------------------------- | ------------------------------------------------------------- |
| `unresolved-shared-string` | a cell's text; reads as empty                                 |
| `unresolved-style`         | a cell's format; reads unstyled                               |
| `unresolved-dxf`           | a conditional rule's formatting; the rule keeps its condition |
| `unresolved-hyperlink`     | a link's target; reads as an empty target                     |
| `unusable-paper-size`      | the sheet's paper size; reads as unset                        |

Each is a place where leniency produces something _indistinguishable from
correct_ — an empty cell, an unstyled cell, a rule that paints nothing, a
link that goes nowhere. Where a reader drops something that is visibly
absent instead, there is no warning, because none is needed.

Structural damage still throws: a missing `xl/workbook.xml`, or a
worksheet part a sheet declares and the archive does not contain, is a
`ParseError`. The difference is whether the answer would be _short_ or
_wrong_.

## Read options, per reader

`ReadOptions` is one interface for `readXlsx`, `readOds`, `readXlsb`,
`readXls` and `read`. Not every option means something to every format:

| option                 | `readXlsx` | `readOds` | `readXlsb` | `readXls` |
| ---------------------- | :--------: | :-------: | :--------: | :-------: |
| `maxInputBytes`        |    yes     |    yes    |    yes     |    yes    |
| `maxTotalCells`        |    yes     |    yes    |     —      |    yes    |
| `maxDecompressedBytes` |    yes     |    yes    |    yes     |    n/a    |
| `maxSpinCount`         |    yes     |    n/a    |    yes     |    n/a    |
| `sheets`               |    yes     |    yes    |     —      |     —     |
| `readStyles`           |    yes     |    yes    |    n/a     |    n/a    |
| `dateSystem`           |    yes     |    n/a    |    yes     |    yes    |
| `password`             |    yes     |     —     |    yes     |     —     |
| `maxRows`              |    yes     |    yes    |     —      |     —     |
| `range`                |    yes     |    yes    |     —      |     —     |

`n/a` means the option cannot apply: ODS stores ISO date strings, so
there is no 1900/1904 system to pick, neither legacy reader surfaces
styles at all, `.xls` is a CFB container rather than a ZIP, and ODS
encryption is not implemented (#156). A `—` is a gap, not a decision.

### Resource limits

The bounds in `src/limits.ts` are exported from the root, so a caller can
quote `MAX_TOTAL_CELLS` in their own message instead of hard-coding
20,000,000. Three of them are also `ReadOptions` fields, per the table
above; the defaults do not change.

Two are still constants only, because both clamp rather than throw — a
file over the bound is read with the excess trimmed, not rejected, so
there is nothing for a caller to rescue:

| bound              | where                         |
| ------------------ | ----------------------------- |
| `MAX_REPEAT_COUNT` | ODS `text:c` / decimal places |
| `MAX_SPAN_CELLS`   | HTML `rowspan` x `colspan`    |

`MAX_SPAN_CELLS` also belongs to `fromHtml`, which takes its own options
type rather than `ReadOptions`.

## XLS and XLSB — read only

No writer exists for either. What the readers surface is narrower than
XLSX, which matters because converting to XLSX can only carry what was
read:

|                                                 | XLS (BIFF8) | XLSB |
| ----------------------------------------------- | ----------- | ---- |
| Sheet names, strings, numbers, booleans, errors | yes         | yes  |
| Dates, honouring the file's 1900/1904 flag      | yes         | yes  |
| Formula **values** (never the formula text)     | yes         | yes  |
| Merges                                          | yes         | yes  |
| Everything else on `Sheet` / `Workbook`         | no          | no   |

So **XLS/XLSB → XLSX is a values-and-names conversion**. Every formula
becomes a hard-coded value, styles and dimensions are dropped, hidden
sheets become visible, and workbook properties and named ranges are lost.

BIFF5 and BIFF7 are rejected outright rather than misread — only BIFF8
(Excel 97-2003) is supported.

## CSV / TSV

Symmetric: delimiter (including tab auto-detection), quote character, line
separator, header handling, `skipHeaderRow`, type inference, leading-zero
preservation, comment lines, and formula-injection escaping — which now
has an inverse in `unescapeFormulae`.

**Encoding is read, not guessed.** The readers take bytes and honour a
byte-order mark — UTF-8, UTF-16LE, UTF-16BE — and fall back to UTF-8.
Anything else has to be named through `encoding`, because an encoding like
windows-1254 leaves no trace in the file and distinguishing it from
windows-1252 by byte frequency is a guess. The write side emits UTF-8
only; `bom: true` is the option that makes Excel read it correctly on
every locale.

One-way, and documented on each option:

| option               | why                                                                                                                          |
| -------------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| `nullValue`          | CSV has no null. Any token you choose reads back as that string, and treating `""` as null would guess for every empty field |
| custom `dateFormat`  | the reader recognises ISO only, and `03/04/2024` cannot be disambiguated. The ISO default does round-trip                    |
| `quoteStyle: "none"` | produces output no read option can reliably parse back — inherent to the mode                                                |

`CsvReadOptions.escape` is honoured on read with no writer counterpart:
hucre can read a backslash-escaped dialect it does not write.

There is no `parseTsv`. TSV reads through `parseCsv` because tab is an
auto-detect candidate — which is a guess, and a single-column TSV or one
whose values contain commas can lose to `,`. Pass `delimiter: "\t"` when
you know.

## JSON / NDJSON

The most symmetric formats in the library. Values and types round-trip in
both the whole-string and streaming readers and writers.

ISO dates round-trip **under `typeInference`**, which is off by default.
JSON already carries numbers and booleans, so the only type it genuinely
cannot express is a date — and reviving one means deciding that a string
which looks like a date was meant as one, which is the caller's call:

```ts
parseJson(json).data[0].when // "2024-01-15T00:00:00.000Z"
parseJson(json, { typeInference: true }).data[0].when // Date
```

Same option, same default, same reasoning as `CsvReadOptions`.

Nesting is the exception worth understanding. `flatten` turns
`{user:{name}}` into `{"user.name"}` by default on read, and
`unflatten: true` reverses it on write — but not exactly, and it cannot
be made exact: `flatten` does not escape dots that were already in a key,
so `{"a.b": 1}` and `{a:{b:1}}` are indistinguishable by the time the
inverse runs. Every dot is treated as a separator. Primitive arrays are
joined into one cell and are not recoverable at all.

`workbookToJson` emits a bare array for a one-sheet workbook and a
sheet-keyed object otherwise; `shape: "sheets"` pins the keyed form, and
`jsonToWorkbook` reads either. `parseJson` on a multi-sheet document
throws rather than returning one row of stringified sheets.

`format: "arrays"` and `format: "columns"` (in `src/export/json.ts`) are
write-only — handoffs to a charting library or dataframe, with no reader.

## XML

Reads and writes nested data, and its writer is the one that un-flattens
dot-paths. Its asymmetry is the mirror image of JSON's: `readXml` returns
every value as a string, with no type-inference option.

## HTML

`fromHtml` reconstructs values, types, merges, header structure
(`<thead>` or an all-`<th>` row → `a11y.headerRow`), and `<caption>` →
`a11y.summary`. Type CSS classes are honoured as declarations, so a
string `"42"` written by `toHtml` comes back a string.

Named HTML entities are decoded from the HTML 4.01 set plus `&apos;`; a
reference outside it is left as written rather than guessed at. `<br>`
becomes a newline. `<script>` and `<style>` are skipped as raw text, the
way HTML5 parses them. `<tfoot>` is placed last however it was declared,
matching where a browser renders it. A document with several tables reads
the first unless `tableIndex` says otherwise.

Not reconstructed: inline styles, the `<style>` block, `role` and
`aria-label`. There is nowhere in `Sheet` for CSS to live, and inventing
a place for it would be a bigger lie than saying HTML export is
presentation output. Cell text is trimmed on read and preserved on write —
indentation in markup is not data.

## Markdown — write only

There is no `fromMarkdown`, and none is planned. `toMarkdown` is a
terminal output format, not an interchange one: it truncates any cell over
50 characters by default and converts newlines to `<br>` with no inverse.

---

Every gap on this page is either a deliberate scope decision or has an
open issue. Anything else is a bug.
