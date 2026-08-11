# Third-party fixtures

Every one of these was written by **ExcelJS**, not by hucre.

That is the whole point. Before they existed, 9,934 tests parsed nothing
but hucre's own output, so a writer bug the reader mirrored was
invisible — and three of the defects fixed in the #439 round were exactly
that shape. See #464.

## What they are, and what they are not

ExcelJS is an independent implementation with its own element ordering,
its own defaults, and its own idea of what a minimal workbook contains.
It is **not** Excel, LibreOffice or Google Sheets, and #464 asks for
those too. What these give is the class of divergence a golden-model test
needs — bytes hucre did not write — from a producer that can be run in
CI-less environments and whose output is reproducible.

Files from real Excel and LibreOffice remain worth adding. They need
someone with those tools; this needed only `npm i exceljs`.

## Licensing

ExcelJS is MIT. Every value inside these files was written in
`scripts/make-fixtures.mjs` in this repository. Nothing is scraped and no
third-party document is redistributed.

## Regenerating

ExcelJS is deliberately **not** a devDependency — these are committed
bytes, so neither CI nor a contributor needs it:

```bash
mkdir -p /tmp/gen && cd /tmp/gen && npm init -y && npm i exceljs
node /path/to/hucre/scripts/make-fixtures.mjs /tmp/gen/node_modules
```

Regenerating changes the bytes (timestamps, ExcelJS version). The tests
assert on _content_, not on bytes, so that is fine — but do not
regenerate casually, because the value of a fixture is that it stopped
moving.
