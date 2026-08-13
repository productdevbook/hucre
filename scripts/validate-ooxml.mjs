#!/usr/bin/env node
// ── Does hucre write valid SpreadsheetML? ────────────────────────────
//
// The XLSX counterpart of `validate-odf.mjs`, and for the same reason:
// every XLSX test in this repository reads the file back with hucre, so
// none of them can tell a valid workbook from one that merely round
// trips. That gap is not theoretical — the ODF half of this pair found
// three defects on its first run, one of them a day old.
//
// This one found none. All three writers — `writeXlsx`,
// `writeXlsxStream`, `XlsxStreamWriter` — produce parts that validate
// against the ECMA-376 Transitional schema, including with tables,
// sparklines, comments, data validations, conditional formatting,
// auto-filters, freeze panes, page setup and named ranges. Worth
// recording as a checked fact rather than an assumption, and worth
// having the tool for the next feature.
//
// Two things to know before reading its output.
//
// **Transitional, not Strict.** Excel writes the Transitional namespace
// (`…/spreadsheetml/2006/main`) and so does hucre; the Strict schemas
// shipped with ECMA-376 Part 1 declare a different one, so validating
// against those reports only that it cannot find `worksheet`. The
// Transitional set is in Part 4.
//
// **Excel's own files do not validate.** They carry MCE extension
// attributes — `x14ac:dyDescent` on every row — which a plain XSD cannot
// know about. Six errors on `excel-styled.xlsx`, all of them that. So
// this is a check on hucre's output, not a general conformance oracle.
//
// Neither the schemas nor a JDK are vendored:
//
//   curl -sSLO https://www.ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip
//   unzip -o ECMA-376-4_5th_edition_december_2016.zip OfficeOpenXML-XMLSchema-Transitional.zip
//   unzip -o OfficeOpenXML-XMLSchema-Transitional.zip -d xsd-t
//
//   node scripts/validate-ooxml.mjs --schema xsd-t/sml.xsd
//
// `--schema` must point at `sml.xsd` inside the extracted directory: it
// imports its siblings by relative path.

import { execFileSync } from "node:child_process"
import { mkdtempSync, mkdirSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { dirname, join } from "node:path"
import { fileURLToPath } from "node:url"

const args = new Map()
for (let i = 2; i < process.argv.length; i += 2) {
  args.set(process.argv[i].replace(/^--/, ""), process.argv[i + 1])
}
const schema = args.get("schema")
const dist = args.get("dist") ?? new URL("../dist/index.mjs", import.meta.url).href

if (!schema) {
  console.error("usage: node scripts/validate-ooxml.mjs --schema <xsd-t/sml.xsd>")
  console.error("       (see the header of this file for where to get it)")
  process.exit(1)
}

const hucre = await import(dist)

// ── Compile the validator ────────────────────────────────────────────

const work = mkdtempSync(join(tmpdir(), "hucre-ooxml-"))
const javaSrc = fileURLToPath(new URL("./xsd/XsdValidate.java", import.meta.url))
try {
  execFileSync("javac", ["-d", work, javaSrc], { stdio: "pipe" })
} catch (error) {
  console.error("could not compile the validator — is a JDK installed?")
  console.error(String(error.stderr ?? error))
  process.exit(1)
}

// ── The documents to check ───────────────────────────────────────────

const STYLE = {
  font: { bold: true, italic: true, size: 14, color: { rgb: "FF0000" } },
  fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FFFF00" } },
  alignment: { horizontal: "center", wrapText: true },
  border: { top: { style: "thin" }, bottom: { style: "double" } },
}

async function drain(stream) {
  const chunks = []
  const reader = stream.getReader()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    chunks.push(value)
  }
  const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0))
  let at = 0
  for (const c of chunks) {
    out.set(c, at)
    at += c.length
  }
  return out
}

async function documents() {
  const out = []

  const cells = new Map([
    ["0,0", { value: "Name", style: STYLE }],
    ["1,1", { value: 42, style: { numFmt: "#,##0.00" } }],
    ["2,0", { value: new Date(Date.UTC(2024, 2, 17)), style: { numFmt: "yyyy-mm-dd" } }],
    ["3,0", { value: 1, formula: "B2+1", formulaResult: 43 }],
    ["4,0", { value: "link", hyperlink: { target: "https://example.test", tooltip: "t" } }],
  ])

  out.push([
    "writeXlsx (styles, formats, formula, hyperlink, validation, rules, page setup)",
    await hucre.writeXlsx({
      sheets: [
        {
          name: "Data",
          rows: [
            ["Name", "Qty"],
            ["Widget", 42],
            [new Date(Date.UTC(2024, 2, 17)), null],
            [1, null],
            ["link", null],
          ],
          cells,
          columns: [{ width: 20 }, { width: 10 }],
          merges: ["A1:B1"],
          freezePane: { rows: 1 },
          autoFilter: { range: "A1:B1" },
          dataValidations: [{ type: "list", range: "B2:B5", formula1: '"x,y"' }],
          conditionalRules: [
            { type: "cellIs", range: "B2:B5", operator: "greaterThan", formula: "1", priority: 1 },
          ],
          pageSetup: { orientation: "landscape", paperSize: "a4" },
          protection: { sheet: true },
        },
      ],
      properties: { title: "T", creator: "C" },
      namedRanges: [{ name: "R", range: "Data!$A$1:$B$2" }],
    }),
  ])

  out.push([
    "writeXlsxStream (inline strings)",
    await drain(
      hucre.writeXlsxStream(
        [
          ["a", "b"],
          [1, 2],
        ],
        { name: "S" },
      ),
    ),
  ])

  const incremental = new hucre.XlsxStreamWriter({
    name: "S",
    columns: [{ header: "A", width: 12 }, { header: "B" }],
  })
  incremental.addRow(["x", { value: 1, style: { font: { bold: true }, numFmt: "0.00" } }])
  incremental.addRow([2, 3])
  out.push(["XlsxStreamWriter (buffered, styled)", await incremental.finish()])

  out.push([
    "writeXlsx (table, sparkline, comment, header/footer, row heights)",
    await hucre.writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [
            ["c", "d"],
            [1, 2],
          ],
          cells: new Map([["0,0", { value: "c", comment: { text: "note", author: "me" } }]]),
          tables: [
            {
              name: "T1",
              displayName: "T1",
              range: "A1:B2",
              columns: [{ name: "c" }, { name: "d" }],
            },
          ],
          sparklines: [{ type: "line", location: "C1", dataRange: "A2:B2" }],
          rowDefs: new Map([[0, { height: 22 }]]),
          headerFooter: { oddHeader: "&Ctitle" },
        },
      ],
    }),
  ])

  return out
}

// ── Validation ───────────────────────────────────────────────────────

/** Every SpreadsheetML part in the package, by path. */
function smlParts(dir) {
  const out = execFileSync("unzip", ["-Z1", dir], { encoding: "utf8" })
    .split("\n")
    .map((s) => s.trim())
    .filter(Boolean)
  // The parts sml.xsd describes. Drawings, charts and the relationship
  // files have their own schemas and are not this script's business.
  return out.filter(
    (p) =>
      /^xl\/(workbook|styles|sharedStrings)\.xml$/.test(p) ||
      /^xl\/(worksheets|tables)\/[^/]+\.xml$/.test(p),
  )
}

let failures = 0

for (const [label, bytes] of await documents()) {
  console.log(`\n${label}  (${bytes.length} bytes)`)

  const pkg = join(work, `pkg-${failures}-${label.length}.xlsx`)
  writeFileSync(pkg, bytes)
  const extracted = join(work, `x-${label.length}`)
  mkdirSync(extracted, { recursive: true })
  execFileSync("unzip", ["-o", "-q", pkg, "-d", extracted])

  for (const part of smlParts(pkg)) {
    const file = join(extracted, part)
    let output
    try {
      output = execFileSync("java", ["-cp", work, "XsdValidate", schema, file], {
        encoding: "utf8",
        stdio: "pipe",
      })
    } catch (error) {
      output = `${error.stdout ?? ""}${error.stderr ?? ""}`
    }

    if (output.includes("VALID")) {
      console.log(`  ${part.padEnd(30)} valid`)
    } else {
      failures++
      console.log(`  ${part.padEnd(30)} INVALID`)
      for (const line of output.trim().split("\n").slice(0, 4)) console.log(`    ${line}`)
    }
  }
}

console.log(failures === 0 ? "\nAll parts valid." : `\n${failures} invalid part(s).`)
process.exit(failures === 0 ? 0 : 1)
