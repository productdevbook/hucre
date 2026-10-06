import { describe, expect, it, vi } from "vitest"
// Scale the shared hard bounds down so overflow tests need five rows,
// not a million. The actual constants are pinned by existing limit tests.
vi.mock("../src/limits", async (original) => ({
  ...(await original<typeof import("../src/limits")>()),
  MAX_ROW_INDEX: 3,
  MAX_COL_INDEX: 3,
}))
import { XlsxStreamWriter, writeXlsxStream, writeXlsxStreamSheets } from "../src/xlsx/stream-writer"
import { writeOdsStream } from "../src/ods/stream-writer"
import { OdsStreamWriter } from "../src/ods/incremental-writer"
import { readXlsx, writeXlsx } from "../src/xlsx"
import { readOds, writeOds } from "../src/ods"
import { InvalidArgumentError } from "../src/errors"
import { drain } from "./support/writer-paths"
import type { CellInput, ColumnDef } from "../src/_types"

const wide: CellInput[] = [1, 2, 3, 4, { value: 5, comment: { text: "outside" } }]
for (const [name, create, read] of [
  ["XLSX", () => new XlsxStreamWriter({ name: "S", maxRowsPerSheet: Infinity }), readXlsx],
  ["ODS", () => new OdsStreamWriter({ name: "S" }), readOds],
] as const) {
  describe(`${name} incremental preflight`, () => {
    it("rejects a wide row atomically before consuming its first cells", async () => {
      const writer = create()
      expect(() => writer.addRow(wide)).toThrow(InvalidArgumentError)
      writer.addRow([42])
      expect((await read(await writer.finish())).sheets[0].rows).toEqual([[42]])
    })
    it("rejects past the last physical row while retaining exact-bound output", async () => {
      const writer = create()
      for (let i = 0; i < 4; i++) writer.addRow([i])
      expect(() => writer.addRow([4])).toThrow(InvalidArgumentError)
      expect((await read(await writer.finish())).sheets[0].rows).toEqual([[0], [1], [2], [3]])
    })
  })
}

for (const [name, write] of [
  [
    "XLSX",
    (rows: CellInput[][]) => drain(writeXlsxStream(rows, { name: "S", maxRowsPerSheet: Infinity })),
  ],
  ["ODS", (rows: CellInput[][]) => drain(writeOdsStream(rows, { name: "S" }))],
] as const) {
  it(`${name} true stream rejects an oversized row and closes its source`, async () => {
    let closed = false
    function* rows() {
      try {
        yield wide
      } finally {
        closed = true
      }
    }
    const output =
      name === "XLSX"
        ? writeXlsxStream(rows(), { name: "S" })
        : writeOdsStream(rows(), { name: "S" })
    await expect(drain(output)).rejects.toThrow(InvalidArgumentError)
    expect(closed).toBe(true)
  })
  it(`${name} true stream refuses a physical row past the hard limit`, async () => {
    await expect(write([[0], [1], [2], [3], [4]])).rejects.toThrow(InvalidArgumentError)
  })
}

it("default XLSX rollover uses the same row bound and retains every input row", async () => {
  const w = new XlsxStreamWriter({ name: "S", repeatHeaders: false })
  for (let i = 0; i < 5; i++) w.addRow([i])
  const wb = await readXlsx(await w.finish())
  expect(wb.sheets.map((sheet) => sheet.rows)).toEqual([[[0], [1], [2], [3]], [[4]]])
})

for (const [name, create] of [
  ["XLSX class", (columns: ColumnDef[]) => new XlsxStreamWriter({ name: "S", columns })],
  ["ODS class", (columns: ColumnDef[]) => new OdsStreamWriter({ name: "S", columns })],
  ["XLSX stream", (columns: ColumnDef[]) => writeXlsxStream([], { name: "S", columns })],
  [
    "XLSX multi-stream",
    (columns: ColumnDef[]) => writeXlsxStreamSheets([{ name: "S", rows: [], columns }]),
  ],
  ["ODS stream", (columns: ColumnDef[]) => writeOdsStream([], { name: "S", columns })],
] as const) {
  it(`${name} rejects out-of-grid declared columns before output`, () => {
    expect(() => create(Array.from({ length: 5 }, () => ({})))).toThrow(InvalidArgumentError)
  })
}

for (const [name, write] of [
  ["XLSX", writeXlsx],
  ["ODS", writeOds],
] as const) {
  it(`${name} buffered output rejects positional rows outside the grid`, async () => {
    await expect(write({ sheets: [{ name: "S", rows: [wide] }] })).rejects.toThrow(
      InvalidArgumentError,
    )
    await expect(
      write({ sheets: [{ name: "S", rows: [[0], [1], [2], [3], [4]] }] }),
    ).rejects.toThrow(InvalidArgumentError)
  })
  it(`${name} checks object expansion before accessing record values`, async () => {
    let accessed = false
    const record = {
      get value() {
        accessed = true
        return 1
      },
    }
    await expect(
      write({
        sheets: [
          { name: "S", data: [record, record, record, record], columns: [{ header: "value" }] },
        ],
      }),
    ).rejects.toThrow(InvalidArgumentError)
    expect(accessed).toBe(false)
  })
}
