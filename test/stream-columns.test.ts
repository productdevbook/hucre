import { describe, expect, it } from "vitest"
import type { CellInput, ColumnDef } from "../src/_types"
import { getCell } from "../src/cell-store"
import { InvalidArgumentError } from "../src/errors"
import { writeXlsx, readXlsx } from "../src/xlsx"
import { writeOds } from "../src/ods"

import { XlsxStreamWriter, writeXlsxStream, writeXlsxStreamSheets } from "../src/xlsx/stream-writer"
import { drain, writerPaths, objectWriterPaths } from "./support/writer-paths"

// Buffered positional rows are already complete; generated headers belong
// to its object-data path. Streams generate them from columns themselves.
function generatedHeaders(
  path: (typeof writerPaths)[number],
  rows: CellInput[][],
  columns: ColumnDef[],
) {
  if (!path.name.startsWith("buffered")) return path.write(rows, columns)
  const data = rows.map((row) =>
    Object.fromEntries(
      columns.flatMap((col, c) => {
        const key = col.key ?? col.header
        return key === undefined ? [] : [[key, row[c] ?? null]]
      }),
    ),
  )
  return path.name.includes("XLSX")
    ? writeXlsx({ sheets: [{ name: "S", data, columns }] })
    : writeOds({ sheets: [{ name: "S", data, columns }] })
}

for (const path of writerPaths) {
  describe(path.name, () => {
    it("retains blank headers, key fallbacks and unnamed column positions", async () => {
      const columns: ColumnDef[] = [{ header: "", key: "id" }, { key: "other" }, {}]
      const sheet = (await path.read(await generatedHeaders(path, [[1, 2]], columns))).sheets[0]
      expect(sheet.rows).toEqual([
        ["", "other", ""],
        [1, 2, null],
      ])
    })
    it("key-only columns do not synthesize a header", async () => {
      const sheet = (await path.read(await generatedHeaders(path, [[7]], [{ key: "id" }])))
        .sheets[0]
      expect(sheet.rows).toEqual([[7]])
    })
    if (path.name !== "streamed ODS") {
      it("generated headers keep the column font without its data number format", async () => {
        const columns: ColumnDef[] = [
          { header: "Amount", numFmt: "0.00", style: { font: { bold: true } } },
        ]
        const before = structuredClone(columns)
        const sheet = (
          await path.read(await generatedHeaders(path, [[42]], columns), { readStyles: true })
        ).sheets[0]
        expect(getCell(sheet.cells, 0, 0)?.style?.font?.bold).toBe(true)
        expect(getCell(sheet.cells, 0, 0)?.style?.numFmt ?? "General").toBe("General")
        expect(getCell(sheet.cells, 1, 0)?.style?.numFmt).toBe("0.00")
        expect(columns).toEqual(before)
      })
    }
    if (path.name.includes("XLSX")) {
      it("writes collapsed groups and default column styles beyond populated cells", async () => {
        const columns: ColumnDef[] = [
          { collapsed: true },
          { style: { font: { bold: true }, numFmt: "0.00" } },
        ]
        const sheet = (await path.read(await path.write([[1]], columns), { readStyles: true }))
          .sheets[0]
        expect(sheet.columns?.[0]?.collapsed).toBe(true)
        expect(sheet.columns?.[1]?.style?.font?.bold).toBe(true)
        expect(sheet.columns?.[1]?.style?.numFmt).toBe("0.00")
      })
    }
  })
}

for (const path of objectWriterPaths) {
  it(`${path.name} projects only own fields with key/header fallbacks, including empty keys`, async () => {
    const row: Record<string, CellInput> = Object.assign(
      Object.create({ inherited: 99 }) as Record<string, CellInput>,
      { "": 7, id: 1, Amount: 42 },
    )
    Object.defineProperty(row, "__proto__", { value: "own", enumerable: true })
    const columns: ColumnDef[] = [
      { key: "", header: "Blank" },
      { header: "Amount" },
      { key: "__proto__" },
      { key: "inherited" },
      {},
    ]
    const sheet = (await path.read(await path.write([row], columns))).sheets[0]
    expect(sheet.rows).toEqual([
      ["Blank", "Amount", "__proto__", "inherited", ""],
      [7, 42, "own", null, null],
    ])
    expect(row[""]).toBe(7)
    expect(Object.getPrototypeOf(row)).toEqual({ inherited: 99 })
  })
}

for (const [name, create] of [
  ["incremental", (cap: number) => new XlsxStreamWriter({ name: "S", maxRowsPerSheet: cap })],
  ["single-stream", (cap: number) => writeXlsxStream([[1]], { name: "S", maxRowsPerSheet: cap })],
  [
    "multi-stream",
    (cap: number) => writeXlsxStreamSheets([{ name: "S", rows: [[1]], maxRowsPerSheet: cap }]),
  ],
] as const) {
  it.each([NaN, 2.5, 1_048_577, -Infinity])(
    `${name} rejects invalid rollover cap %s before output`,
    (cap) => {
      expect(() => create(cap)).toThrow(InvalidArgumentError)
    },
  )
}

it("incremental XLSX computes auto-width while retaining no source cells", async () => {
  const columns: ColumnDef[] = [{ header: "Name", autoWidth: true }]
  const w = new XlsxStreamWriter({ name: "S", columns })
  w.addRow(["a considerably longer report title"])
  const sheet = (await readXlsx(await w.finish(), { readStyles: true })).sheets[0]
  expect(sheet.columns?.[0]?.width).toBeGreaterThan(25)
  expect(columns[0].width).toBeUndefined()
})
for (const multi of [false, true]) {
  it(`${multi ? "multi" : "single"} true XLSX streaming rejects auto-width before consuming rows`, () => {
    let consumed = false
    function* rows() {
      consumed = true
      yield [1]
    }
    const columns: ColumnDef[] = [{ autoWidth: true }]
    expect(() =>
      multi
        ? writeXlsxStreamSheets([{ name: "S", rows: rows(), columns }])
        : writeXlsxStream(rows(), { name: "S", columns }),
    ).toThrow(InvalidArgumentError)
    expect(consumed).toBe(false)
  })
}

for (const path of writerPaths.filter(
  (path) => path.name === "buffered XLSX" || path.name === "incremental XLSX",
)) {
  it(`${path.name} auto-width measures rich-text display without a scalar cache`, async () => {
    const columns: ColumnDef[] = [{ autoWidth: true }]
    const sheet = (
      await path.read(
        await path.write(
          [[{ richText: [{ text: "a considerably longer report title" }] }]],
          columns,
        ),
      )
    ).sheets[0]
    expect(sheet.columns?.[0]?.width).toBeGreaterThan(25)
  })
  it(`${path.name} auto-width uses the explicit style number format`, async () => {
    const columns: ColumnDef[] = [{ autoWidth: true, numFmt: "0", style: { numFmt: "0.000000" } }]
    const sheet = (await path.read(await path.write([[1]], columns))).sheets[0]
    expect(sheet.columns?.[0]?.width).toBe(11)
  })
}

for (const [name, write] of [
  [
    "buffered",
    (rows: number) => writeXlsx({ sheets: [{ name: "S", rows: [[1]], freezePane: { rows } }] }),
  ],
  [
    "incremental",
    async (rows: number) => {
      const w = new XlsxStreamWriter({ name: "S", freezePane: { rows } })
      w.addRow([1])
      return w.finish()
    },
  ],
  [
    "single-stream",
    (rows: number) => drain(writeXlsxStream([[1]], { name: "S", freezePane: { rows } })),
  ],
  [
    "multi-stream",
    (rows: number) =>
      drain(writeXlsxStreamSheets([{ name: "S", rows: [[1]], freezePane: { rows } }])),
  ],
] as const) {
  it.each([NaN, -1, 1.5, 1_048_576])(
    `${name} rejects invalid frozen row coordinate %s`,
    async (rows) => {
      await expect(write(rows)).rejects.toThrow(InvalidArgumentError)
    },
  )
}
