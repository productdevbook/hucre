import { describe, expect, it } from "vitest"
import { writeXlsxObjects, readXlsxObjects } from "../src/xlsx/objects"
import { writeOdsObjects, readOdsObjects } from "../src/ods/objects"
import { writeCsvObjects } from "../src/csv/writer"
import { parseCsvObjects } from "../src/csv/reader"
import { writeObjects, readObjects } from "../src/defter"
import type { CellValue } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #439 — every object writer took its column set from `data[0]`, so a key
// absent from the first record was dropped, and a record made only of
// such keys became an empty row. The union helper (`collectHeaders`)
// already existed and was called by nothing but the JSON reader.
//
// The shape below is the ordinary one: optional fields, a column that
// appears partway through an export.
// ═══════════════════════════════════════════════════════════════════════

const SPARSE: Record<string, CellValue>[] = [{ a: 1 }, { b: 2 }, { a: 3, c: 4 }]

describe("object writers take the union of every record's keys", () => {
  it("writeXlsxObjects keeps every column and every row", async () => {
    const bytes = await writeXlsxObjects(SPARSE)

    const { data, headers } = await readXlsxObjects(bytes)

    expect(headers).toEqual(["a", "b", "c"])
    expect(data).toEqual([
      { a: 1, b: null, c: null },
      { a: null, b: 2, c: null },
      { a: 3, b: null, c: 4 },
    ])
  })

  it("writeOdsObjects keeps every column and every row", async () => {
    const bytes = await writeOdsObjects(SPARSE)

    const { headers, data } = await readOdsObjects(bytes)

    expect(headers).toEqual(["a", "b", "c"])
    expect(data).toHaveLength(3)
    expect(data[1]!.b).toBe(2)
    expect(data[2]!.c).toBe(4)
  })

  it("writeCsvObjects keeps every column and every row", () => {
    const csv = writeCsvObjects(SPARSE)

    // Previously: "a\n1\n\n3" — one column, and the middle record blank.
    expect(csv.split(/\r?\n/)[0]).toBe("a,b,c")

    const { headers, data } = parseCsvObjects(csv, { header: true, typeInference: true })

    expect(headers).toEqual(["a", "b", "c"])
    expect(data).toHaveLength(3)
    expect(data[1]!.b).toBe(2)
    expect(data[2]!.c).toBe(4)
  })

  it("writeObjects keeps every column and every row", async () => {
    const bytes = await writeObjects(SPARSE)

    const { headers, data } = await readObjects(bytes)

    expect(headers).toEqual(["a", "b", "c"])
    expect(data).toHaveLength(3)
    expect(data[2]!.c).toBe(4)
  })

  it("still honours an explicit headers option, projecting away the rest", async () => {
    const bytes = await writeXlsxObjects(SPARSE, { headers: ["c", "a"] })

    const { headers, data } = await readXlsxObjects(bytes)

    // `b` was asked to be left out, so the record that held only `b` has
    // nothing left and reads back as the empty row it was written as.
    expect(headers).toEqual(["c", "a"])
    expect(data).toEqual([
      { c: null, a: 1 },
      { c: 4, a: 3 },
    ])
  })

  it("keeps first-seen order rather than sorting", async () => {
    const bytes = await writeXlsxObjects([{ z: 1 }, { a: 2 }, { m: 3 }])

    expect((await readXlsxObjects(bytes)).headers).toEqual(["z", "a", "m"])
  })

  it("still writes nothing for an empty list", async () => {
    expect(writeCsvObjects([])).toBe("")
    expect((await readXlsxObjects(await writeXlsxObjects([]))).headers).toEqual([])
  })
})
