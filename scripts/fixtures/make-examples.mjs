// Independent XLSX producer. CI consumes the committed files; hucre is
// never imported here and never supplies the expected results.
import fs from "node:fs/promises"
import path from "node:path"
import { createHash } from "node:crypto"
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool"

const root = path.resolve(process.argv[2] ?? ".")
const scenarios = JSON.parse(await fs.readFile(path.join(root, "examples/scenarios.json"), "utf8"))
const outputDir = path.join(root, "examples/workbooks")
const previewDir = process.argv[3]
await fs.mkdir(outputDir, { recursive: true })
if (previewDir) await fs.mkdir(previewDir, { recursive: true })
const manifest = []

for (const scenario of scenarios) {
  const workbook = Workbook.create()
  for (const spec of scenario.sheets) workbook.worksheets.add(spec.name)
  for (const spec of scenario.sheets) {
    const sheet = workbook.worksheets.getItem(spec.name)
    const rows = spec.rows.map((row) =>
      row.map((value) =>
        typeof value === "string" && value.startsWith("D:") ? new Date(value.slice(2)) : value,
      ),
    )
    const range = sheet.getRangeByIndexes(0, 0, rows.length, rows[0].length)
    range.values = rows
    range.format.font = { name: "Arial", size: 11 }
    range.format.columnWidth = 22
    range.format.rowHeight = 23
    range.format.verticalAlignment = "center"
    sheet.showGridLines = false
    const header = sheet.getRangeByIndexes((spec.headerRow ?? 1) - 1, 0, 1, rows[0].length)
    header.format.fill = "#17365D"
    header.format.font = { name: "Arial", size: 11, bold: true, color: "#FFFFFF" }
    for (const [ref, formula] of Object.entries(spec.formulas ?? {}))
      sheet.getRange(ref).formulas = [[`=${formula}`]]
    for (const [ref, format] of Object.entries(spec.formats ?? {}))
      sheet.getRange(ref).setNumberFormat(format)
    for (const ref of spec.merges ?? []) sheet.getRange(ref).merge()
    for (const [col, width] of (spec.widths ?? []).entries())
      sheet.getRangeByIndexes(0, col, rows.length, 1).format.columnWidth = width
    if (spec.freezeRows) sheet.freezePanes.freezeRows(spec.freezeRows)
    if (spec.wrap) sheet.getRange(spec.wrap).format.wrapText = true
    if (spec.table) sheet.tables.add(range.address, true, spec.table)
    if (spec.validation)
      sheet.getRange(spec.validation.range).dataValidation = {
        rule: {
          type: "whole",
          operator: "between",
          formula1: spec.validation.min,
          formula2: spec.validation.max,
        },
      }
    if (spec.statusRange)
      sheet.getRange(spec.statusRange).dataValidation = {
        rule: { type: "list", values: ["Not started", "In progress", "Done"] },
      }
    if (spec.progressRange)
      sheet.getRange(spec.progressRange).conditionalFormats.add("cellIs", {
        operator: "equal",
        formula: 1,
        format: { fill: "#E2F0D9" },
      })
  }
  workbook.recalculate()
  // Expectations are authored in scenarios.json, not captured from this
  // engine or from hucre. Check this engine against them before export.
  for (const spec of scenario.sheets) {
    const sheet = workbook.worksheets.getItem(spec.name)
    for (const [ref] of Object.entries(spec.formulas ?? {})) {
      const [, letters, rowText] = /^([A-Z]+)(\d+)$/.exec(ref)
      const col = [...letters].reduce((n, letter) => n * 26 + letter.charCodeAt(0) - 64, 0) - 1
      const expected = spec.rows[Number(rowText) - 1][col]
      const actual = sheet.getRange(ref).values[0][0]
      if (Math.abs(actual - expected) > 1e-10)
        throw new Error(`${scenario.file}:${ref}: ${actual} != ${expected}`)
    }
    if (previewDir) {
      const preview = await workbook.render({
        sheetName: spec.name,
        range: `A1:${String.fromCharCode(64 + spec.rows[0].length)}${spec.rows.length}`,
        scale: 1.5,
        format: "png",
      })
      await fs.writeFile(
        path.join(previewDir, `${scenario.file}-${spec.name.replaceAll(/[^\w-]/g, "_")}.png`),
        new Uint8Array(await preview.arrayBuffer()),
      )
    }
  }
  const file = path.join(outputDir, scenario.file)
  await (await SpreadsheetFile.exportXlsx(workbook)).save(file)
  await fs.rm(`${file}.inspect.ndjson`, { force: true })
  const bytes = await fs.readFile(file)
  manifest.push({ file: scenario.file, sha256: createHash("sha256").update(bytes).digest("hex") })
  console.log(`${scenario.file}: ${bytes.length} bytes`)
}
await fs.writeFile(
  path.join(root, "examples/manifest.json"),
  `${JSON.stringify(manifest, null, 2)}\n`,
)
