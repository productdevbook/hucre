# Changelog

Releases are tagged; the notes for a breaking release live in
[`MIGRATION.md`](MIGRATION.md), which is written for the person upgrading
rather than as a list of commits.

## Unreleased — 2.0.0

Buffered and streamed authoring share header/object projection and Excel grid
preflight. Blank headers and empty keys keep their positions; inherited
object fields are excluded. XLSX column defaults, collapsed groups and
frozen panes share serialization. Incremental XLSX now honors `autoWidth`
with the same display-width collector as buffered output; true streaming
uses `XlsxStreamColumn` and rejects automatic widths without an explicit
width. Rollover caps must be integers within Excel's sheet limit or Infinity;
Infinity disables splitting, not physical grid validation.

Spreadsheet writers share full inline cell resolution, including explicit
formula caches (null included), rich text and `link()` in rows or object data.
XLSX streams retain shared/array/dynamic formulas, checkboxes, links and
comments with their package parts; incremental filters and conditional
formats are honored. ODS uses the shared serializer and reports unsupported
cell fields through `onDrop` with a cell reference. True ODS streaming
remains unstyled. Streaming XLSX retains current-sheet link/comment metadata.

Buffered writers share one authoring row-source boundary. Object data works
in text outputs, omitted columns are inferred across all records, and
competing `rows`/`data` sources throw. ODS keeps data hyperlinks and column
styles; XLSX pivot sources share effective worksheet headers and caches.

Object readers, sheet projections and JSON/HTML/Markdown/text exports share
effective metadata and formula-cache values. Sparse object reads no longer
return an empty result. Projection keeps physical transform indexes,
disambiguates duplicate headers, preserves prototype-looking data keys and
checks output limits before padding. Accessibility coalesces sparse blank-row
gaps into one finding. Object options/results share their definitions.

Search, replacement and templates share one value traversal for dense and
sparse cells. Value edits update cell types and formula caches together,
remove obsolete rich-text runs and retain styles and formula text. Partial
metadata preserves inline cell fields; explicit null overrides clear values.

Dense ODS/text normalization checks the growing rectangle before repeat
expansion or padding. ODS row repeats carry formulas, cached values,
hyperlinks, styles and horizontal merges; `maxRows` bounds repeats before
allocation. `read()` forwards `maxTotalCells` to text readers, and JSON,
NDJSON, XML and HTML accept the same limit at their normalization boundary.

Breaking. See [Migrating to v2](MIGRATION.md#migrating-to-v2): deprecated
names removed, per-reader read options, `Color` on every colour field,
`CellError` for error cells, rectangular `Sheet.rows` from every reader,
one `StreamRow` and one writer surface across formats, one name per
option, `serializeWorkbook` removed. New entry points `hucre/cell`,
`hucre/format` and `hucre/a11y`. `read()` refuses a ZIP that is not a
spreadsheet by name; every error the library throws is a `HucreError`;
`moveSheet` / `removeSheet` check their indexes.

Buffered writers now accept the read workbook directly through derived
`WorkbookInput` / `SheetInput` types; container options move to a second
argument and `onDrop` reports unsupported metadata. Caller-side
`toWriteOptions` / `toWriteSheet` converters are removed. Builders share
one authoring state, so `set()` values are honored. Empty formulas retain
cached values of every cell type, and `openXlsx` reads styles by default
so saving preserves number formats.

Twelve independently generated XLSX business examples live under
`examples/`, with authored expectations and SHA-256 checksums. Disk-backed
integration tests and shared fixture helpers are separated from unit
tests; every test is typechecked in `tsconfig.test.json`. CI checks the
`v2` branch as well as `main`.

`Sheet.cells` now uses a plain numeric `CellStore`, with six coordinate
helpers on `hucre/cell`, `hucre/xlsx` and the root. Blocks remove the
single-Map filled-cell ceiling and coordinate strings, while preserving
structured cloning. Readers, writers, builders and sheet operations share
this storage; structural edits share one metadata remapping path.

Row and column insertions/deletions share an editing engine and reference
rewriter. Own-sheet qualified references move, multi-area metadata keeps
surviving areas, and edits validate geometry before mutation. Large row
insertions avoid the runtime function argument ceiling. Integration tests
edit the independent invoice and verify both save paths.

## 1.1.0

See the v1.1.0 tag.

## 1.0.0

See [Migrating to v1](MIGRATION.md#migrating-to-v1).
