# Changelog

Releases are tagged; the notes for a breaking release live in
[`MIGRATION.md`](MIGRATION.md), which is written for the person upgrading
rather than as a list of commits.

## Unreleased — 2.0.0

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

## 1.1.0

See the v1.1.0 tag.

## 1.0.0

See [Migrating to v1](MIGRATION.md#migrating-to-v1).
