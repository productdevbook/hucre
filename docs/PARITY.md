# Read/write parity

What hucre reads, what it writes, and where those differ.

This exists because a gap you discover is worse than a gap you were told
about. v1 makes the public API a stability commitment, and part of that
commitment is being honest about the edges: everything below is either at
parity or listed here as an exception.

If you hit a loss that is **not** on this page, that is a bug — please
open an issue. That is the whole point of writing it down.

This page is written by hand and says what hucre _does_.
[`SPEC-COVERAGE.md`](SPEC-COVERAGE.md) is generated and says what the
**formats** define that hucre does not — derived from the ECMA-376 and
OASIS schemas rather than from anyone's memory, and crossed with the
fixture corpus so that a gap real files actually contain is separated
from one that is merely theoretical.

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

### The string ceiling, and where it still is

One JavaScript string cannot hold an arbitrarily large part. V8's limit
is 536,870,888 characters — about 512 MB — so a part above it cannot be
turned into a string at all, however much memory the machine has.
Instrument logs at Excel's row limit reach it: in a corpus of ~600, 14
workbooks had a worksheet part between 588 MB and 1.13 GB. See #503.

**`readXlsx` reads a worksheet part of any size.** When the part is over
the ceiling it is parsed from the ZIP entry's stream instead of from a
string, by the same SAX handlers the buffered parse uses — so the `Sheet`
is the same `Sheet`, and no option selects between them. Which path runs
is decided from the size the ZIP declares, before anything is
decompressed; when the ZIP declares no size, the buffered read runs and
the ceiling error it raises is the signal to retry as a stream.

Two things this does not do:

- **The other buffered readers still have the ceiling.** `readXlsb` and
  `readOds` are unchanged, as is every non-worksheet part of an
  XLSX — `sharedStrings.xml`, `styles.xml`, a drawing. Those still fail
  with the `ParseError` from #514: it names the part, the size, the
  bound, and `streamXlsxRows`, and it says the workbook is not damaged,
  because everything else a reader throws means the file is wrong and
  this one means it is large. (None of the ~600 has a non-worksheet part
  anywhere near the ceiling; a workbook with a 512 MB shared-string
  table would still fail.)
- **It does not lift the cell bound.** A part over the ceiling is
  usually also a sheet with a very large bounding box, so clearing the
  ceiling often lands on `maxTotalCells` instead — of the 14, ten read
  and four hit that limit. See the next section.

Two trades worth knowing:

- The streaming ZIP path has no whole entry to check, so a worksheet read
  this way is **not CRC-32 verified**. That is the same trade
  `streamXlsxRows` has always made. It is bounded on the way in — a
  declared size the compressed body could not have produced (DEFLATE tops
  out at 1032:1) is not believed, so a small entry cannot claim to be
  enormous and skip the checksum that way.
- A truncated part is still an error. The row streamers have no error
  contract for one — they drop the unfinished construct and let the
  caller notice the missing rows — but the worksheet reader asks the
  streaming parser for `strict`, so a part that ends mid-tag throws the
  `XmlError` the buffered parser throws rather than returning a short
  `Sheet`.

**`streamXlsxRows` still has no ceiling either** and remains the lower-cost
answer when you only need to walk the rows: it never builds a `Workbook`,
so it is bounded by the row you are on rather than by the sheet.

### A sparse sheet has a way out

`Sheet.rows` being a dense rectangle means a read costs the bounding box
rather than the cell count. That is right for almost every sheet and
wrong for a sparse one: a real workbook with 82,000 values scattered over
a 305,612,208-slot box — 0.03% filled — could not be read at all, and
none of the three options the error named would have helped. Raising
`maxTotalCells` trades a clean error for a multi-gigabyte allocation,
`range` needs you to already know where the data is, and `maxRows` bounds
rows when the problem is columns. See #501.

Two answers, and the error now names both:

- **`streamXlsxRows`** already read that file, a row at a time, and the
  message never said so. It is the better answer when you only need to
  walk the rows once.
- **`readXlsx(input, { sparse: true })`** returns `cells` keyed
  `"row,col"` and leaves `rows` empty. Memory tracks the values rather
  than the box, the bounding-box limit does not apply because nothing
  dense is built, and you get a `Workbook` — which is what streaming
  cannot give you.

The error also reports how full the box actually is, which is what turns
"your sheet is too large" into "your sheet is mostly nothing" — a
different problem with a different answer.

`sparse` is XLSX-only and off by default. With it on, anything that reads
`Sheet.rows` — `sheetToObjects`, `readObjects`, the writers — sees an
empty sheet; `cells` is the whole answer.

**`sparse` has a ceiling of its own: 16,777,216 filled cells.** `cells`
is a `Map`, and V8 caps a `Map` at 2^24 entries. That is not a bound
hucre chose and it cannot raise it without `Sheet.cells` ceasing to be a
`Map`, which would break every caller using it as one.

It matters because the two ways a sheet can be over the bounding-box
limit want different answers, and only one of them is `sparse`:

| the sheet is…             | example                    | use              |
| ------------------------- | -------------------------- | ---------------- |
| a large box, mostly empty | 82k values over 305M slots | `sparse: true`   |
| genuinely dense and large | 28.4M filled of 30.2M      | `streamXlsxRows` |

For the second, the cell count that blew the box limit is the same count
that blows the `Map`, so `sparse` cannot work by construction. The
oversize error now says which case it is looking at and stops offering
`sparse` when the filled count is already past what a `Map` holds; going
over anyway is a `ParseError` naming the sheet, not a raw
`RangeError: Map maximum size exceeded`. `streamXlsxRows` has no such
bound. See #527.

### A style-only cell does not widen the sheet

Excel writes a self-closing `<c r="WVF45" s="3"/>` for every position
formatting was ever applied to. A real packing-list workbook had 145,315
of them against 197 values, so `rows` came back 45 x 16,126 and
`writeCsv` of that was 727 KB — 99.75% bare commas — from 1.8 KB of data.

Under the default `readStyles: false` those cells contribute nothing:
their styles are not read, so their only effect was null padding a caller
cannot tell from never-written cells. They no longer extend the sheet, and
the row they sit in is not allocated either. Fixed in #492.

With `readStyles: true` nothing changes — there the styles _are_ the
information, and the full box is what the caller asked for.

Only the **trailing** box shrinks. Interior positions are untouched: a
value at D1 still sits at index 3 behind three nulls, and an empty row
between two populated ones is still there.

### `Sheet.rows` is a dense rectangle

Every row is an array, every row is the same length, and no element is
`undefined`. That is what makes `rows[r][c]` safe without guarding either
index, and it is what the readers' limits are sized against — the cost of
a sheet is its bounding box rather than its cell count, which is why
`maxTotalCells` bounds the product.

The contract went unwritten and two readers did not hold it. `readXls`
and `readXlsb` padded a row only to its _own_ last cell and never
allocated a row the file left empty, so one authored sheet saved three
ways came back three shapes, and a gap row came back as `undefined` —
which `CellValue` cannot express. Fixed in #494; it is now stated on
`Sheet.rows` as well as here.

### Line endings are normalized on the way in

XML 1.0 §2.11 requires a processor to turn a literal CRLF, and a literal
lone CR, into a single LF before the application sees the content.
hucre's writer knew this — it escapes a deliberate CR as `&#13;` — and
the parser did not, so the two disagreed.

Excel writes a multi-line cell with a literal CRLF inside `<t>`, so the
same authored workbook read as `.xlsx` gave `"line one\r\nline two"` and
as `.xlsb` — which stores a bare LF — gave `"line one\nline two"`. Fixed
in #493.

A character reference is not a literal line ending: `&#13;` still comes
through as CR, which is what makes a deliberate one distinguishable from
a line break and what keeps hucre's own round trip exact.

### Cached formula results, whatever their type

`Cell.formulaResult` was assigned in one place — the numeric arm of the
cell-type switch — so a cached result survived only when it happened to
be a number. A formula whose result is a string, an error or a boolean
lost it.

That was a round-trip loss rather than a missing field: the _writer_ has
always been able to write those back, so `readXlsx` → `writeXlsx` emitted
`<f>` with no `<v>`, and anything opening the result without
recalculating saw an empty cell where Excel showed `#DIV/0!`. Fixed in
#497.

One behaviour changed with it. A formula cell now reports
`type: "formula"` whatever its cached result is; it used to report
`"error"` on the way in and `"formula"` on the way back out, and both
cannot be right. `value` still holds the error token, so spotting an
error by its value is unaffected, and a _hard-coded_ error cell — one
with no formula — still reports `"error"`.

### ISO-8601 date cells

`ST_CellType` (§18.18.11) has seven members and the reader's switch had
six: `d`, "cell containing a date in the ISO 8601 format", fell through
to the numeric branch and came back as a **string**. openpyxl writes it
whenever `iso_dates=True`, so the same day under the same number format
read as a `Date` when stored as a serial and as text when stored as ISO.
Fixed in #496.

The parse is deliberately strict — `new Date(text)` accepts a great deal
that is not ISO 8601 — and two cases stay text on purpose: a bare time
(`13:45:30`, which openpyxl writes for a `datetime.time`) has no day to
anchor it, and anything that is not an ISO date is left as the file wrote
it. An unqualified time is read as UTC, and `date1904` is **not** applied:
the value is an instant, not an offset from an epoch.

### Not every tab is a worksheet

`xl/workbook.xml`'s `<sheets>` lists every tab whatever its kind — chart
sheets, dialog sheets and macro sheets included — and the _relationship
type_ is what tells them apart. A chart sheet used to make the whole
workbook unreadable: its rId was not in the worksheet map, so the lookup
missed and the reader threw, taking every ordinary worksheet beside it
down too. On one corpus of real instrument-exported files that was 52
failures out of 538. See #499.

A non-worksheet tab now reads as an empty `Sheet` carrying
`kind: "chartsheet"` (or `"dialogsheet"`). It is kept rather than skipped
so `sheets: [2]` still selects Excel's third tab — renumbering would be a
quieter kind of wrong. `kind` is **read-only**: hucre writes worksheets,
and `toWriteOptions` reports it as a drop.

`streamXlsxRows` on a non-worksheet tab yields nothing rather than
throwing. A missing worksheet _part_ is still a `ParseError`, because
that is damage rather than a different kind of tab.

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

**An indexed colour outside the palette keeps its index and no RGB.**
A colour may name a palette index rather than an RGB. hucre resolves those
against the file's `<indexedColors>` when it overrides the palette and
against the ECMA-376 §18.8.27 defaults otherwise, so `ColorSpec` carries
both the `indexed` it was given and the `rgb` it stands for.

Two indices are not colours: 64 is the system foreground and 65 the system
background, and the specification gives neither an ARGB. Those keep the
index and get no `rgb`, because the colour depends on the reader's own
theme and inventing one would be answering a question the file did not
ask. The same is true of any index past the end of the palette — Excel
writes 81 for tooltip text.

**A cell whose text is literally `_x0041_` reads back as `A` — in XLSX.**
OOXML uses `_xHHHH_` to encode characters XML cannot hold, and hucre
decodes it on read. Excel disambiguates by escaping a leading underscore
as `_x005F_`; hucre deliberately does not, because doing so would mangle
the far more common case of ordinary text that happens to contain an
underscore. The ambiguity is accepted, and now written down.

It is an OOXML convention and the loss is OOXML's alone. ODF has no
`_xHHHH_`, so the same string round-trips through ODS unchanged. The ODS
writer used to borrow the spelling for carriage returns, which meant
`"a\r\nb"` came back as `"a_x000D_\nb"` — and LibreOffice showed those
seven characters too, since to anything but Excel that is all they are.
ODS now writes `&#13;`, which is what XML gives you for this: a character
reference is not subject to the end-of-line normalisation of §2.11, so it
survives any conforming parse.

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
auto-filter, tables, images, page setup, sheet protection, tab colour,
hidden sheets, and `time` cells (which read back as the raw ISO duration
string).

Named ranges left that list: `<table:named-expressions>` carries them
both ways. Only workbook-level ones — ODF scopes a name to a sheet by
putting the block inside that `<table:table>`, which `NamedRange.scope`
could drive and does not yet, so a scoped name is written as a workbook
one. A `<table:named-expression>` — a formula rather than a range — has
no field in `NamedRange` to land in and is skipped rather than
half-read.

Number format is carried, but not every Excel code has an ODF spelling.
These do not survive intact. Every row was measured through the round
trip rather than reasoned about:

| code              | comes back as | why                                                                                             |
| ----------------- | ------------- | ----------------------------------------------------------------------------------------------- |
| `0.00_);(0.00)`   | `0.00;(0.00)` | `_)` reserves the width of a character. ODF has no equivalent, so the padding is dropped        |
| `[mm]:ss`, `[ss]` | `mm:ss`, `ss` | the elapsed marker survives on hours (`[hh]:mm`) but not on minutes or seconds alone            |
| `0.00;[Red]-0.00` | `0.00;-0.00`  | colour tags are dropped on purpose — it is what stops `[White]0.00` being read as a time format |
| `0.00;;`          | `0.00`        | empty trailing sections have nothing to write                                                   |

`General` returning no `numFmt` is not in that list: General _is_ the
absence of a data style, so there is nothing lost.

Three rows have left that list since it was written, all for the same
reason and all found the same way. Engineering notation — `##0.0E+0`,
where the exponent steps in threes — was listed here as lost, with
`number:exponent-interval` named as its ODF spelling _in the same
sentence_: the attribute existed all along and the writer simply did not
emit it. It does now, along with `number:forced-exponent-sign`, which is
what keeps `0.00E-00` from becoming `0.00E+00`. `#` versus `0` — an optional digit
against a mandatory one — was described here as a distinction ODF could
not express; it expresses it with `number:min-decimal-places` and
`number:min-integer-digits`, which LibreOffice writes on every number
style. `#,###` was worse than documented: it lost its thousands separator
as well, because the writer's grouping test only recognised `#,##0` and
`0,000` literally.

`@` — the text format — used to be on that list, described here as having
"no data style to write". That was wrong. ODF spells it
`<number:text-style>` and LibreOffice writes one into every document it
saves; hucre carries it now. The error survived because it was written
down: a documented loss is one nobody looks at again. It was found by
crossing the ODF grammar with a LibreOffice file
([`SPEC-COVERAGE.md`](SPEC-COVERAGE.md)), not by anyone re-reading this
page.

Ordinary scientific formats — `0.00E+00` and its widths — do round-trip,
through `<number:scientific-number>`. They did not until it was written:
the code fell through to the plain-number branch, so `0.00E+00` became
`0.00` and the file displayed a plain decimal in LibreOffice too.

The reader opens `content.xml` and `meta.xml`. It does not open
`styles.xml` or `settings.xml`, which is where LibreOffice keeps named and
default cell styles and all page setup — so a LibreOffice-authored file
reads back with its direct formatting only.

A column's `table:default-cell-style-name` _is_ read, and fills in for
cells that name no style of their own. That matters because LibreOffice
puts a column's number format there rather than on the cells: without it
a LibreOffice document came back with its values and none of its formats.
A column pointing at a named style — `"Default"`, which it writes on the
columns it did not format — still resolves to nothing, for the reason
above.

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

They do **not** agree on how many sheets they walk. `streamXlsxRows`
yields one sheet — the first, unless `sheet` names another —
while `streamOdsRows` walks every sheet in the document and tags each row
with `sheetIndex`. So the same loop over a three-sheet workbook gives you
one sheet of it as `.xlsx` and all three as `.ods`, silently. Pass
`sheet` when you mean one, and read `row.sheetIndex` when you mean all.

### `streamXmlRows` gives you a row's own keys, not a rectangle

XML rows are objects rather than arrays, and the two XML readers differ
in one way worth knowing before you swap one for the other.

`readXml` returns a **rectangle**: it collects the union of every row's
keys and fills the gaps, so a record with no `<note>` still comes back
with `note: null` and an empty `<record/>` is a row of nulls.

Knowing that union means having read the last row, so `streamXmlRows`
cannot do it — each row carries **only the keys it had**, and an empty
`<record/>` yields `{}`. Read `values.note ?? null` rather than
`values.note` when moving code across; `undefined` is where `null` used
to be.

It is the same cause as the `rowTag` difference already documented on
that function: what a streaming reader gives up is everything that
depends on the end of the document.

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

| code                       | what silently went missing                                     |
| -------------------------- | -------------------------------------------------------------- |
| `unresolved-shared-string` | a cell's text; reads as empty                                  |
| `unresolved-style`         | a cell's format; reads unstyled                                |
| `unresolved-dxf`           | a conditional rule's formatting; the rule keeps its condition  |
| `unresolved-hyperlink`     | a link's target; reads as an empty target                      |
| `unusable-paper-size`      | the sheet's paper size; reads as unset                         |
| `malformed-cell-ref`       | one cell whose `r` is not a reference; the sheet is still read |

Each is a place where leniency produces something _indistinguishable from
correct_ — an empty cell, an unstyled cell, a rule that paints nothing, a
link that goes nowhere. Where a reader drops something that is visibly
absent instead, there is no warning, because none is needed.

Structural damage still throws: a missing `xl/workbook.xml`, or a
worksheet part a sheet declares and the archive does not contain, is a
`ParseError`. The difference is whether the answer would be _short_ or
_wrong_.

## Read options, per reader

Each reader has its own options type, and the type is the statement of
what it honours: `XlsxReadOptions`, `OdsReadOptions`, `XlsbReadOptions`,
`XlsReadOptions`, all extending `ReadOptionsBase` (`maxInputBytes`,
`maxTotalCells`). `read()` takes `ReadOptions`, the widest of them, because
it does not know the format until it has looked at the bytes.

Passing a reader an option it does not honour is a compile error rather
than a silent no-op — `readXls(bytes, { password })` used to type-check
and do nothing. `test/read-options-per-reader.test.ts` reads each
reader's source and fails if a declared field is never looked at, so the
type cannot drift from the behaviour.

What is absent is absent for a reason: ODS stores ISO date strings, so
there is no 1900/1904 system to pick; neither legacy reader surfaces
styles at all; `.xls` is a CFB container rather than a ZIP; ODS encryption
is not implemented (#156); and the two binary readers read every sheet,
so there is no `sheets`, `maxRows` or `range`.

### Resource limits

The bounds in `src/limits.ts` are exported from the root, so a caller can
quote `MAX_TOTAL_CELLS` in their own message instead of hard-coding
20,000,000. Three of them are also read-option fields; the defaults do not change.

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

Two things a `writeCsv` → `parseCsv` round-trip does not carry, both
facts about the format rather than defects. Property testing found them
(#473); they are here so nobody has to find them again.

**A final row holding a single empty cell is lost.** It renders as
nothing after the preceding line's terminator, and a file ending in a
terminator is universally read as having no record after it. RFC 4180
leaves the trailing CRLF optional and says nothing about the ambiguity.
A trailing row of _two_ empty cells survives — it renders as a bare
delimiter — and so does an empty row anywhere but the end.

**Delimiter auto-detection is a guess, and a file can defeat it.**
`[["with,comma"], ["with\ttab"]]` written with the default comma quotes
its one comma, so the only unquoted separator character left in the file
is a tab — and `parseCsv` is right to read it as tab-separated. Pass
`delimiter` when you know it; there is no reading of those bytes that
recovers the intent.

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
