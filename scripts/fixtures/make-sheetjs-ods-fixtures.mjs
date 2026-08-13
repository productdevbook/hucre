#!/usr/bin/env node
// ── Third-party ODS fixture generation ──────────────────────────────
//
// #464 asks for files hucre did not write. The XLSX side got them —
// Excel, openpyxl, ExcelJS — and the **ODS reader got none**. Every ODS
// byte the suite parsed was one hucre had just produced, which is the
// closed loop the issue is about: a reader that misunderstands
// `office:value-type` is checked against a writer that misunderstands it
// identically and the suite stays green.
//
// These are written by **SheetJS** (the `xlsx` package, Apache-2.0), an
// independent implementation whose ODS output differs from hucre's in
// element order, in style naming, and in what a minimal document
// contains.
//
// What SheetJS is not: it is not LibreOffice. Two things it will not
// produce, both of which matter and both of which #464 still wants from
// a LibreOffice corpus:
//
//   * `table:number-columns-repeated`, which LibreOffice uses for every
//     run of like cells and is the single sharpest ODS reader trap.
//   * error cells — SheetJS writes `t: "e"` as an empty
//     `<table:table-cell/>` in ODS, so an error is simply not in the
//     file to be read.
//
// So this narrows the gap rather than closing it. What it buys is the
// thing that did not exist before: an ODS document in the suite that
// hucre's writer had no hand in.
//
// SheetJS is deliberately NOT a devDependency. The fixtures are
// committed bytes, so neither CI nor a contributor needs it:
//
//   mkdir -p /tmp/gen && cd /tmp/gen && npm init -y && npm i xlsx
//   node /path/to/hucre/scripts/fixtures/make-sheetjs-ods-fixtures.mjs /tmp/gen/node_modules
//
// Licensing: SheetJS Community Edition is Apache-2.0, and every value in
// these files was written here. Nothing is scraped and no third-party
// document is redistributed.

import { mkdirSync, writeFileSync } from "node:fs"
import { join } from "node:path"
import { fileURLToPath } from "node:url"
import { createRequire } from "node:module"

const modulesDir = process.argv[2]
if (!modulesDir) {
  console.error(
    "usage: node scripts/fixtures/make-sheetjs-ods-fixtures.mjs <path-to-node_modules-with-xlsx>",
  )
  process.exit(1)
}

const require = createRequire(join(modulesDir, "index.js"))
const XLSX = require("xlsx")

const outDir = fileURLToPath(new URL("../../test/fixtures/third-party", import.meta.url))
mkdirSync(outDir, { recursive: true })

/** Written to disk and asserted against in test/ods-third-party.test.ts. */
const FIXTURES = []

function fixture(name, build) {
  FIXTURES.push({ name, build })
}

/** A sheet from rows, appended under `name`. */
function sheet(wb, name, rows) {
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows, { cellDates: true }), name)
}

// ── The shapes #464 asks for ─────────────────────────────────────────

fixture("sheetjs-basic.ods", (wb) => {
  // Strings, numbers, dates, booleans — the four ODF value types hucre
  // has to tell apart from `office:value-type` alone.
  sheet(wb, "Values", [
    ["text", "number", "date", "boolean"],
    ["Widget", 42, new Date(Date.UTC(2024, 0, 15)), true],
    ["Gadget", -7.25, new Date(Date.UTC(2023, 11, 31)), false],
  ])
})

fixture("sheetjs-dates.ods", (wb) => {
  // SheetJS writes `office:date-value` with a trailing `Z`, which
  // LibreOffice does not — the ODF grammar is a plain ISO 8601 date and
  // the zone designator is SheetJS's own habit. A reader that anchors on
  // LibreOffice's spelling drops all of these.
  //
  // 1899-12-30 is the serial-0 epoch: a reader doing its own serial
  // arithmetic rather than reading the literal date lands off by one or
  // two here, which is the #415 family of bug.
  sheet(wb, "Dates", [
    [new Date(Date.UTC(2024, 2, 17))],
    [new Date(Date.UTC(2024, 2, 17, 13, 45, 30))],
    [new Date(Date.UTC(1899, 11, 30))],
    [new Date(Date.UTC(1900, 0, 1))],
    [new Date(Date.UTC(2000, 1, 29))],
  ])
})

fixture("sheetjs-whitespace.ods", (wb) => {
  // The #441 shape, in ODF's spelling. XLSX needs `xml:space="preserve"`
  // and ODF does not — text in `<text:p>` keeps its spaces by the format
  // rather than by an attribute the writer might forget. A reader that
  // trims is wrong here with nothing to blame it on.
  sheet(wb, "Whitespace", [
    ["  leading"],
    ["trailing  "],
    ["  both  "],
    ["inner   gap"],
    ["line\nbreak"],
  ])
})

fixture("sheetjs-unicode.ods", (wb) => {
  // Astral-plane and combining characters through a second UTF-8
  // encoder. The last two are written as escapes on purpose: a
  // decomposed `e` + U+0301 and a zero-width space are both invisible
  // in a source file, and an editor that helpfully precomposes — or a
  // reader that normalises — would otherwise change them unseen.
  sheet(wb, "Unicode", [
    ["ünïcödé"],
    ["日本語のテキスト"],
    ["🎉 emoji 🚀"],
    ["Ω≈ç√∫˜µ"],
    ["e\u0301 combining"],
    ["\u200bzero width"],
  ])
})

fixture("sheetjs-sparse.ods", (wb) => {
  // Gaps written as bare `<table:table-cell/>` — SheetJS's spelling of a
  // hole. A whole empty row in the middle, and a value far to the right
  // of the row above it.
  const ws = XLSX.utils.aoa_to_sheet([
    ["a", null, null, null, null, null, null, null, null, null, "far"],
    [1, null, 2],
    [],
    [null, null, null, "island"],
  ])
  ws["!ref"] = "A1:K4"
  XLSX.utils.book_append_sheet(wb, ws, "Sparse")
})

fixture("sheetjs-multi-sheet.ods", (wb) => {
  // Three sheets, one with a name needing escaping and one non-ASCII, to
  // check the sheet list survives the round trip in order.
  sheet(wb, "First", [["one"], [1]])
  sheet(wb, "İkinci Sayfa", [["two"], [2]])
  sheet(wb, "Third & Last", [["three"], [3]])
})

fixture("sheetjs-numbers.ods", (wb) => {
  // `office:value` is a decimal string, so the reader's parse is the only
  // thing between the file and the double. The awkward ones: the
  // seventeen-significant-digit float, the exponent forms, and the
  // subnormal that a `toFixed`-based writer flattens to zero (the defect
  // #485 found in the CSV writer).
  sheet(wb, "Numbers", [
    [0.1 + 0.2],
    [1e21],
    [1e-7],
    [Number.MAX_SAFE_INTEGER],
    [-0.000001],
    [12345678.9],
  ])
})

fixture("sheetjs-formulas.ods", (wb) => {
  // ODF formulas carry a `of:`-prefixed namespace and `[.B2]` cell
  // references — nothing like XLSX's `B2*2`. What matters to a reader
  // with no formula engine is that the *cached value* still arrives, so
  // these each have one.
  const ws = XLSX.utils.aoa_to_sheet([
    ["n", "doubled"],
    [21, null],
    [50, null],
  ])
  ws["B2"] = { t: "n", f: "A2*2", v: 42 }
  ws["B3"] = { t: "n", f: "A3*2", v: 100 }
  XLSX.utils.book_append_sheet(wb, ws, "Formulas")
})

fixture("sheetjs-empty.ods", (wb) => {
  // A document with a sheet and nothing in it. The degenerate case that
  // tends to throw rather than return an empty sheet.
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([]), "Empty")
})

// ── Write them ───────────────────────────────────────────────────────

for (const { name, build } of FIXTURES) {
  const wb = XLSX.utils.book_new()
  build(wb)
  // `type: "buffer"` rather than `writeFile`: the ESM build of SheetJS
  // has no `fs` bound and throws "cannot save file" if asked to write.
  const bytes = XLSX.write(wb, { bookType: "ods", type: "buffer" })
  writeFileSync(join(outDir, name), bytes)
  console.log(`${name}  ${bytes.length} bytes`)
}

console.log(`\n${FIXTURES.length} ODS files written to ${outDir}`)
