import { getCell } from "../../src/cell-store"
import type { Workbook } from "../../src/_types"
import { read } from "../../src/defter"
import { createHash } from "node:crypto"
import { readFileSync, readdirSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { readXlsx, streamXlsxRows, writeXlsx, openXlsx, saveXlsx } from "../../src/xlsx"
import { flat } from "../support/workbook-model"
import { insertRows, insertColumns, replaceCells } from "../../src/sheet-ops"
import { parseCellRef, toRange } from "../../src/cell-utils"

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
