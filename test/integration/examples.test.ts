import { getCell, cellEntries, createCellStore } from "../../src/cell-store"
import { writeOds, readOds } from "../../src/ods"
import type { Workbook } from "../../src/_types"
import { read, readObjects, write } from "../../src/defter"
import { sheetToArrays, sheetToObjects } from "../../src/sheet-utils"
import { workbookToJson } from "../../src/json/writer"
import { createHash } from "node:crypto"
import { readFileSync, readdirSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { readXlsx, streamXlsxRows, writeXlsx, openXlsx, saveXlsx } from "../../src/xlsx"
import { flat } from "../support/workbook-model"
import { insertRows, insertColumns, replaceCells } from "../../src/sheet-ops"
import { parseCellRef, toRange } from "../../src/cell-utils"

it("projects the independent invoice through sparse object and export paths", async () => {
  const input = new Uint8Array(
    readFileSync(new URL("../../examples/workbooks/invoice.xlsx", import.meta.url)),
  )
  const headers = ["Item", "Quantity", "Unit price", "Amount"]
  const data = [
    { Item: "Keyboard", Quantity: 2, "Unit price": 75, Amount: 150 },
    { Item: "Mouse", Quantity: 3, "Unit price": 20, Amount: 60 },
    { Item: "Total", Quantity: null, "Unit price": null, Amount: 210 },
  ]
  const workbook = await readXlsx(input, { sparse: true })
  expect(workbook.sheets[0].rows).toEqual([])
  expect(await readObjects(input, { sparse: true })).toEqual({ headers, data })
  expect(sheetToObjects(workbook.sheets[0])).toEqual({ headers, data })
  expect(sheetToArrays(workbook.sheets[0]).data).toEqual([
    ["Keyboard", 2, 75, 150],
    ["Mouse", 3, 20, 60],
    ["Total", null, null, 210],
  ])
  expect(JSON.parse(workbookToJson(workbook))).toEqual(data)
  expect(new TextDecoder().decode(await write(workbook, { format: "csv" }))).toBe(
    "Item,Quantity,Unit price,Amount\r\nKeyboard,2,75,150\r\nMouse,3,20,60\r\nTotal,,,210",
  )
})

it("rebuilds the independent invoice from object data with shared header and formula resolution", async () => {
  const input = new Uint8Array(
    readFileSync(new URL("../../examples/workbooks/invoice.xlsx", import.meta.url)),
  )
  const source = (await readXlsx(input, { readStyles: true })).sheets[0]
  const data = sheetToObjects(source).data
  // Preserve only formula metadata. Values and headers must come from the
  // object records, so a numeric store cannot mask a dropped data source.
  const cells = createCellStore(
    [...cellEntries(source.cells)]
      .filter(([, , cell]) => cell.formula !== undefined)
      .map(([row, col, cell]) => [
        row,
        col,
        {
          formula: cell.formula,
          formulaResult: cell.formulaResult,
          style: cell.style,
        },
      ]),
  )
  const sheet = { name: "Invoice", data, cells }
  const before = structuredClone(sheet)
  const expected = [
    ["Item", "Quantity", "Unit price", "Amount"],
    ["Keyboard", 2, 75, 150],
    ["Mouse", 3, 20, 60],
    ["Total", null, null, 210],
  ]
  for (const [writer, reader] of [
    [writeXlsx, readXlsx],
    [writeOds, readOds],
  ] as const) {
    const result = (await reader(await writer({ sheets: [sheet] }))).sheets[0]
    expect(result.rows).toEqual(expected)
    expect(getCell(result.cells, 3, 3)?.formulaResult).toBe(210)
  }
  expect(new TextDecoder().decode(await write({ sheets: [sheet] }, { format: "csv" }))).toBe(
    "Item,Quantity,Unit price,Amount\r\nKeyboard,2,75,150\r\nMouse,3,20,60\r\nTotal,,,210",
  )
  expect(sheet).toEqual(before)
})

interface ScenarioSheet {
  name: string
  rows: Array<Array<string | number | boolean | null>>
  formulas?: Record<string, string>
  formats?: Record<string, string>
  merges?: string[]
  freezeRows?: number
  table?: string
  validation?: { range: string }
  statusRange?: string
  progressRange?: string
}
interface Scenario {
  file: string
  description: string
  sheets: ScenarioSheet[]
}
interface ManifestEntry {
  file: string
  sha256: string
}

const root = new URL("../../examples/", import.meta.url)
const json = <T>(path: string): T => JSON.parse(readFileSync(new URL(path, root), "utf8")) as T
const scenarios = json<Scenario[]>("scenarios.json")
const manifest = json<ManifestEntry[]>("manifest.json")
const bytes = (file: string): Uint8Array =>
  new Uint8Array(readFileSync(new URL(`workbooks/${file}`, root)))

function expectScenario(workbook: Workbook, scenario: Scenario): void {
  expect(workbook.sheets.map((s) => s.name)).toEqual(scenario.sheets.map((s) => s.name))
  for (const spec of scenario.sheets) {
    const sheet = workbook.sheets.find((s) => s.name === spec.name)!
    expect(sheet.rows.map((row) => row.map(flat))).toEqual(spec.rows)
    for (const [ref, formula] of Object.entries(spec.formulas ?? {})) {
      const { row, col } = parseCellRef(ref)
      expect(getCell(sheet.cells, row, col)?.formula, ref).toBe(formula)
      expect(flat(getCell(sheet.cells, row, col)?.formulaResult), ref).toBe(spec.rows[row][col])
    }
    for (const [range, format] of Object.entries(spec.formats ?? {})) {
      const { startRow, startCol, endRow, endCol } = toRange(range)
      for (let r = startRow; r <= endRow; r++)
        for (let c = startCol; c <= endCol; c++) {
          expect(getCell(sheet.cells, r, c)?.style?.numFmt, `${range}:${r},${c}`).toBe(format)
        }
    }
    expect(sheet.merges ?? []).toEqual((spec.merges ?? []).map(toRange))
    if (spec.freezeRows) expect(sheet.freezePane?.rows).toBe(spec.freezeRows)
    if (spec.table) expect(sheet.tables?.map((t) => t.name)).toContain(spec.table)
    if (spec.validation || spec.statusRange)
      expect(sheet.dataValidations?.length).toBeGreaterThan(0)
    if (spec.progressRange) expect(sheet.conditionalRules?.[0].range).toBe(spec.progressRange)
  }
}

describe("independently authored Excel examples", () => {
  it("registers every workbook exactly once, with a pinned checksum", () => {
    const files = readdirSync(new URL("workbooks/", root))
      .filter((f) => f.endsWith(".xlsx"))
      .sort()
    expect(scenarios.map((s) => s.file).sort()).toEqual(files)
    expect(manifest.map((s) => s.file).sort()).toEqual(files)
    expect(new Set(files).size).toBeGreaterThanOrEqual(12)
    for (const entry of manifest) {
      expect(createHash("sha256").update(bytes(entry.file)).digest("hex"), entry.file).toBe(
        entry.sha256,
      )
    }
  })

  for (const scenario of scenarios) {
    describe(`${scenario.file}: ${scenario.description}`, () => {
      it("reads authored values, formulas and layout", async () => {
        const workbook = await readXlsx(bytes(scenario.file), { readStyles: true })
        expectScenario(workbook, scenario)
      })

      it("detects the format and streams the same values", async () => {
        const workbook = await read(bytes(scenario.file))
        expect(workbook.sheets.map((s) => s.rows.map((r) => r.map(flat)))).toEqual(
          scenario.sheets.map((s) => s.rows),
        )
        for (const spec of scenario.sheets) {
          const rows = new Map<number, unknown[]>()
          for await (const row of streamXlsxRows(bytes(scenario.file), { sheet: spec.name })) {
            rows.set(row.index, row.values.map(flat))
          }
          for (let r = 0; r < spec.rows.length; r++) {
            const actual = rows.get(r) ?? []
            expect(
              spec.rows[r].map((_, c) => actual[c] ?? null),
              `${spec.name}:${r}`,
            ).toEqual(spec.rows[r])
          }
        }
      })

      it("survives both authoring and part-preserving saves", async () => {
        const input = bytes(scenario.file)
        const workbook = await readXlsx(input, { readStyles: true })
        const outputs = [await writeXlsx(workbook), await saveXlsx(await openXlsx(input))]
        for (const output of outputs) {
          const again = await readXlsx(output, { readStyles: true })
          expectScenario(again, scenario)
        }
      })
    })
  }
})

it.each(["authoring", "preserved"] as const)(
  "edits the independent invoice and saves through the %s path",
  async (mode) => {
    const input = bytes("invoice.xlsx")
    const opened = mode === "preserved" ? await openXlsx(input) : undefined
    const workbook = opened ?? (await readXlsx(input, { readStyles: true }))
    insertRows(workbook.sheets[0], 1, 1)
    insertColumns(workbook.sheets[0], 1, 1)
    const output = opened ? await saveXlsx(opened) : await writeXlsx(workbook)
    const sheet = (await readXlsx(output, { readStyles: true })).sheets[0]
    expect(sheet.rows).toEqual([
      ["Item", null, "Quantity", "Unit price", "Amount"],
      [null, null, null, null, null],
      ["Keyboard", null, 2, 75, 150],
      ["Mouse", null, 3, 20, 60],
      ["Total", null, null, null, 210],
    ])
    expect(getCell(sheet.cells, 2, 4)?.formula).toBe("C3*D3")
    expect(getCell(sheet.cells, 3, 4)?.formula).toBe("C4*D4")
    expect(getCell(sheet.cells, 4, 4)?.formula).toBe("SUM(E3:E4)")
    expect(getCell(sheet.cells, 2, 3)?.style?.numFmt).toBe("#,##0.00")
  },
)

it.each(["authoring", "preserved"] as const)(
  "edits independent invoice values and cached results through the %s path",
  async (mode) => {
    const input = bytes("invoice.xlsx")
    const opened = mode === "preserved" ? await openXlsx(input) : undefined
    const workbook = opened ?? (await readXlsx(input, { readStyles: true }))
    const sheet = workbook.sheets[0]
    expect(replaceCells(sheet, "Keyboard", "Keys")).toBe(1)
    expect(replaceCells(sheet, 150, 151)).toBe(1)
    const output = opened ? await saveXlsx(opened) : await writeXlsx(workbook)
    const again = (await readXlsx(output, { readStyles: true })).sheets[0]
    expect(again.rows).toEqual([
      ["Item", "Quantity", "Unit price", "Amount"],
      ["Keys", 2, 75, 151],
      ["Mouse", 3, 20, 60],
      ["Total", null, null, 210],
    ])
    expect(getCell(again.cells, 1, 3)).toMatchObject({ formula: "B2*C2", formulaResult: 151 })
    expect(getCell(again.cells, 1, 3)?.style?.numFmt).toBe("#,##0.00")
  },
)
