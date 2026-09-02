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

## 1.1.0

See the v1.1.0 tag.

## 1.0.0

See [Migrating to v1](MIGRATION.md#migrating-to-v1).
