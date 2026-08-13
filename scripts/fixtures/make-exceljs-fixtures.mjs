#!/usr/bin/env node
// ── Third-party fixture generation ──────────────────────────────────
//
// 9,934 tests and not one byte produced by another tool was ever parsed:
// every assertion was hucre writing something and hucre reading it back.
// A writer bug the reader mirrors is invisible that way, and three of the
// defects fixed in the #439 round were exactly that shape. See #464.
//
// These fixtures are written by **ExcelJS** — an independent
// implementation with its own element ordering, its own defaults, and its
// own idea of what a minimal workbook contains. It is not Excel, and the
// test file says so; what it is, is bytes hucre did not write.
//
// ExcelJS is deliberately NOT a devDependency. The fixtures are committed
// bytes, so neither CI nor a contributor needs it — regenerating is a
// once-in-a-while thing:
//
//   mkdir -p /tmp/gen && cd /tmp/gen && npm init -y && npm i exceljs
//   node /path/to/hucre/scripts/fixtures/make-exceljs-fixtures.mjs /tmp/gen/node_modules
//
// Licensing: ExcelJS is MIT, and every value in these files was written
// here. Nothing is scraped and no third-party document is redistributed.

import { mkdirSync, writeFileSync } from "node:fs"
import { join } from "node:path"
import { fileURLToPath } from "node:url"
import { createRequire } from "node:module"

const modulesDir = process.argv[2]
if (!modulesDir) {
  console.error(
    "usage: node scripts/fixtures/make-exceljs-fixtures.mjs <path-to-node_modules-with-exceljs>",
  )
  process.exit(1)
}

const require = createRequire(join(modulesDir, "index.js"))
const ExcelJS = require("exceljs")

const outDir = fileURLToPath(new URL("../../test/fixtures/third-party", import.meta.url))
mkdirSync(outDir, { recursive: true })

/** Written to disk and asserted against in test/third-party-fixtures.test.ts. */
const FIXTURES = []

function fixture(name, build) {
  FIXTURES.push({ name, build })
}

// ── The shapes #464 asks for ─────────────────────────────────────────

fixture("basic-values.xlsx", (wb) => {
  // Strings, numbers, dates, booleans, and a formula with a cached
  // result — the five things every reader has to get right.
  const ws = wb.addWorksheet("Values")
  ws.addRow(["text", "number", "date", "boolean", "formula"])
  ws.addRow(["Widget", 42, new Date(Date.UTC(2024, 0, 15)), true, { formula: "B2*2", result: 84 }])
  ws.addRow([
    "Gadget",
    -7.25,
    new Date(Date.UTC(2023, 11, 31)),
    false,
    { formula: "B3*2", result: -14.5 },
  ])
})

fixture("whitespace-strings.xlsx", (wb) => {
  // The #441 shape. A reader that trims agrees with a writer that drops
  // `xml:space`, and only a third tool notices.
  const ws = wb.addWorksheet("Whitespace")
  ws.addRow(["  leading"])
  ws.addRow(["trailing  "])
  ws.addRow(["  both  "])
  ws.addRow(["inner   gap"])
  ws.addRow(["line\nbreak"])
  ws.addRow([" "])
})

fixture("styled.xlsx", (wb) => {
  const ws = wb.addWorksheet("Styled")
  const header = ws.addRow(["bold", "italic", "filled", "bordered", "formatted"])
  header.getCell(1).font = { bold: true }
  header.getCell(2).font = { italic: true, size: 14, name: "Georgia" }
  header.getCell(3).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } }
  header.getCell(4).border = {
    top: { style: "thin" },
    left: { style: "medium", color: { argb: "FFFF0000" } },
    bottom: { style: "double" },
    right: { style: "thin" },
  }
  const row = ws.addRow([1, 2, 3, 4, 1234.5])
  row.getCell(5).numFmt = "#,##0.00"
  ws.getColumn(1).width = 22
})

fixture("layout.xlsx", (wb) => {
  // Merges, a frozen pane, and a column-level format.
  const ws = wb.addWorksheet("Layout", {
    views: [{ state: "frozen", xSplit: 1, ySplit: 2 }],
  })
  ws.addRow(["Report", null, null])
  ws.addRow(["name", "qty", "price"])
  ws.addRow(["Widget", 3, 9.99])
  ws.mergeCells("A1:C1")
  ws.getColumn(3).numFmt = '"$"#,##0.00'
})

fixture("conditional.xlsx", (wb) => {
  const ws = wb.addWorksheet("Rules")
  ws.addRow(["value"])
  for (const n of [1, 50, 99]) ws.addRow([n])
  ws.addConditionalFormatting({
    ref: "A2:A4",
    rules: [
      {
        type: "cellIs",
        operator: "greaterThan",
        formulae: ["10"],
        priority: 1,
        style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFF0000" } } },
      },
    ],
  })
})

fixture("multi-sheet.xlsx", (wb) => {
  wb.addWorksheet("First").addRow(["a", 1])
  wb.addWorksheet("Second").addRow(["b", 2])
  const hidden = wb.addWorksheet("Hidden")
  hidden.addRow(["c", 3])
  hidden.state = "hidden"
})

fixture("errors-and-blanks.xlsx", (wb) => {
  const ws = wb.addWorksheet("Edge")
  ws.addRow(["before", null, "after"])
  ws.getCell("A2").value = { error: "#DIV/0!" }
  ws.getCell("C2").value = { error: "#N/A" }
  ws.getCell("A4").value = "gap above"
})

fixture("hyperlinks-and-comments.xlsx", (wb) => {
  const ws = wb.addWorksheet("Links")
  ws.addRow(["site"])
  ws.getCell("A2").value = { text: "example", hyperlink: "https://example.com" }
  ws.getCell("A3").value = "commented"
  ws.getCell("A3").note = "a note from another tool"
})

fixture("wide-and-tall.xlsx", (wb) => {
  // Past the single-letter column boundary, so `AA`/`AB` references are
  // exercised rather than assumed.
  const ws = wb.addWorksheet("Grid")
  ws.addRow(Array.from({ length: 30 }, (_, i) => `c${i + 1}`))
  for (let r = 0; r < 5; r++) {
    ws.addRow(Array.from({ length: 30 }, (_, i) => r * 30 + i))
  }
})

fixture("unicode.xlsx", (wb) => {
  const ws = wb.addWorksheet("Unicode")
  ws.addRow(["Türkçe", "şehir"])
  ws.addRow(["日本語", "テスト"])
  ws.addRow(["Ελληνικά", "δοκιμή"])
  ws.addRow(["emoji", "😀🎉"])
  ws.addRow(["rtl", "مرحبا"])
})

fixture("properties.xlsx", (wb) => {
  wb.creator = "hucre fixture generator"
  wb.created = new Date(Date.UTC(2024, 5, 1, 12, 0, 0))
  wb.modified = new Date(Date.UTC(2024, 5, 2, 12, 0, 0))
  wb.title = "Third-party fixture"
  wb.subject = "Testing"
  wb.addWorksheet("Props").addRow(["x", 1])
})

// ── Write them ───────────────────────────────────────────────────────

for (const { name, build } of FIXTURES) {
  const wb = new ExcelJS.Workbook()
  build(wb)
  const buffer = await wb.xlsx.writeBuffer()
  writeFileSync(join(outDir, name), Buffer.from(buffer))
  console.log(`  ${name.padEnd(30)} ${String(buffer.byteLength).padStart(7)} bytes`)
}

console.log(`\n  ${FIXTURES.length} fixtures written to test/fixtures/third-party/`)
