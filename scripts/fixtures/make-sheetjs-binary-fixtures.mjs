#!/usr/bin/env node
// ── Third-party XLSB and XLS fixture generation ─────────────────────
//
// `test/fixtures/PROVENANCE.md` says it plainly: "The XLS and XLSB
// readers are the sharp end — they exist only to consume other tools'
// output and, until this directory, had never seen any." That directory
// fixed half of it. The files in it are Excel's, and openpyxl — the
// second producer that keeps Excel honest — writes `.xlsx` only, so the
// binary readers still had exactly one source.
//
// SheetJS writes both. It found a defect in the XLSB reader on the first
// file: hucre handled the full-form cell records (`BrtCellRk`,
// `BrtCellSt`, …) and none of the `BrtShort*` forms, which are the same
// records with the column omitted — it is the previous cell's plus one.
// Excel writes the full form every time; SheetJS writes the short form
// for every cell after the first in a row. Every one of them vanished —
// a twelve-column sheet read back one column wide, with no error.
//
// BIFF5 is deliberately not generated: `readXls` supports BIFF8 only and
// says so with a typed error, which `test/real-files.test.ts` covers.
//
// SheetJS is deliberately NOT a devDependency. The fixtures are
// committed bytes, so neither CI nor a contributor needs it:
//
//   mkdir -p /tmp/gen && cd /tmp/gen && npm init -y && npm i xlsx
//   node /path/to/hucre/scripts/fixtures/make-sheetjs-binary-fixtures.mjs /tmp/gen/node_modules
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
    "usage: node scripts/fixtures/make-sheetjs-binary-fixtures.mjs <path-to-node_modules-with-xlsx>",
  )
  process.exit(1)
}

const require = createRequire(join(modulesDir, "index.js"))
const XLSX = require("xlsx")

const outDir = fileURLToPath(new URL("../../test/fixtures/third-party", import.meta.url))
mkdirSync(outDir, { recursive: true })

/** Written to disk and asserted against in test/xlsb-short-records.test.ts. */
const FIXTURES = []

function fixture(stem, build) {
  FIXTURES.push({ stem, build })
}

function sheet(wb, name, rows) {
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows, { cellDates: true }), name)
}

// ── The shapes that exercise each cell record ────────────────────────

fixture("sheetjs-values", (wb) => {
  // One of each cell type, so every `BrtShort*` variant appears: a
  // string, an integer that fits an RK, a float that does not, a
  // boolean, and a blank.
  sheet(wb, "Values", [
    ["text", "int", "float", "bool", "blank"],
    ["Widget", 42, -7.25, true, null],
    ["Gadget", 0, 0.1 + 0.2, false, null],
    ["", 1000000, 1e-7, true, null],
  ])
})

fixture("sheetjs-wide", (wb) => {
  // Twelve columns, so a reader that keeps only the first cell of a row
  // is off by eleven rather than by one.
  const header = Array.from({ length: 12 }, (_, i) => `col${i + 1}`)
  const body = Array.from({ length: 4 }, (_, r) =>
    Array.from({ length: 12 }, (_, c) => (c % 2 === 0 ? `r${r}c${c}` : r * 12 + c)),
  )
  sheet(wb, "Wide", [header, ...body])
})

fixture("sheetjs-unicode", (wb) => {
  sheet(wb, "Unicode", [
    ["ünïcödé"],
    ["日本語のテキスト"],
    ["🎉 emoji 🚀"],
    ["  padded  "],
    ['quote"inside'],
  ])
})

fixture("sheetjs-dates", (wb) => {
  // Note for whoever regenerates these: SheetJS converts a `Date` to a
  // serial through *local* time, so the bytes below carry the offset of
  // the machine that wrote them. The tests compare the two readers
  // against each other rather than against an absolute instant, so a
  // regeneration in another zone changes the bytes without breaking
  // anything — but do not add an assertion that pins one.
  sheet(wb, "Dates", [
    [new Date(Date.UTC(2024, 2, 17))],
    [new Date(Date.UTC(1900, 0, 1))],
    [new Date(Date.UTC(2000, 1, 29))],
  ])
})

fixture("sheetjs-sparse", (wb) => {
  const ws = XLSX.utils.aoa_to_sheet([
    ["a", null, null, null, "far"],
    [1, null, 2],
    [],
    [null, null, null, "island"],
  ])
  ws["!ref"] = "A1:E4"
  XLSX.utils.book_append_sheet(wb, ws, "Sparse")
})

// ── Write them, in both binary formats ───────────────────────────────

for (const { stem, build } of FIXTURES) {
  for (const bookType of ["xlsb", "xls"]) {
    const wb = XLSX.utils.book_new()
    build(wb)
    // `type: "buffer"` rather than `writeFile`: the ESM build of SheetJS
    // has no `fs` bound and throws "cannot save file" if asked to write.
    const bytes = XLSX.write(wb, { bookType, type: "buffer" })
    const name = `${stem}.${bookType}`
    writeFileSync(join(outDir, name), bytes)
    console.log(`${name}  ${bytes.length} bytes`)
  }
}

console.log(`\n${FIXTURES.length * 2} binary files written to ${outDir}`)
