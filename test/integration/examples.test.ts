import type { Workbook } from "../../src/_types"
import { read } from "../../src/defter"
import { createHash } from "node:crypto"
import { readFileSync, readdirSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { readXlsx, streamXlsxRows, writeXlsx, openXlsx, saveXlsx } from "../../src/xlsx"
import { flat } from "../support/workbook-model"
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
      expect(sheet.cells?.get(`${row},${col}`)?.formula, ref).toBe(formula)
      expect(flat(sheet.cells?.get(`${row},${col}`)?.formulaResult), ref).toBe(spec.rows[row][col])
    }
    for (const [range, format] of Object.entries(spec.formats ?? {})) {
      const { startRow, startCol, endRow, endCol } = toRange(range)
      for (let r = startRow; r <= endRow; r++)
        for (let c = startCol; c <= endCol; c++) {
          expect(sheet.cells?.get(`${r},${c}`)?.style?.numFmt, `${range}:${r},${c}`).toBe(format)
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
