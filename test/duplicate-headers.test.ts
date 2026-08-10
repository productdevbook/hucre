import { describe, expect, it } from "vitest"
import { disambiguate } from "../src/_objects"
import { sheetToObjects } from "../src/sheet-utils"
import { parseCsvObjects } from "../src/csv/reader"
import { readXlsxObjects } from "../src/xlsx/objects"
import { writeXlsx } from "../src/xlsx/writer"
import type { Sheet } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #439 §AG — two columns sharing a header collapsed into one key, and the
// later column's values overwrote the earlier's:
//
//   sheetToObjects({ rows: [["a","a","",null],[1,2,3,4]] })
//   // { data: [{ a: 2, "": 4 }], headers: ["a","a","",""] }
//
// Four columns in, two keys out, nothing warned. That is what a real
// spreadsheet looks like when someone repeated a label or left two spacer
// columns — and every objects-shaped API funnels through this projection.
//
// The blank key itself was never the problem: `""` for a single unnamed
// column is a settled contract across the readers and is pinned
// elsewhere. Only repeats are renamed.
// ═══════════════════════════════════════════════════════════════════════

describe("disambiguate", () => {
  it("leaves a list with no repeats exactly as it is", () => {
    expect(disambiguate(["a", "b", "c"])).toEqual(["a", "b", "c"])
  })

  it("leaves a single blank header blank", () => {
    // `""` as a key is the established behaviour of every *Objects reader.
    expect(disambiguate(["a", "", "b"])).toEqual(["a", "", "b"])
  })

  it("keeps the first of a repeated name and suffixes the rest", () => {
    expect(disambiguate(["a", "a", "a"])).toEqual(["a", "a_2", "a_3"])
  })

  it("names a repeated blank by its position", () => {
    // `_2` alone would read as nothing at all.
    expect(disambiguate(["", ""])).toEqual(["", "column2"])
    expect(disambiguate(["a", "", "b", ""])).toEqual(["a", "", "b", "column4"])
  })

  it("does not collide with a name that is already taken", () => {
    expect(disambiguate(["a", "a_2", "a"])).toEqual(["a", "a_2", "a_3"])
    expect(disambiguate(["a", "a", "a_2"])).toEqual(["a", "a_2", "a_2_2"])
  })
})

describe("no column is lost to a repeated header", () => {
  const ROWS = [
    ["a", "a", "", null],
    [1, 2, 3, 4],
  ]

  it("keeps all four values in sheetToObjects", () => {
    const sheet: Sheet = { name: "D", rows: ROWS }

    const { data, headers } = sheetToObjects(sheet)

    expect(headers).toEqual(["a", "a_2", "", "column4"])
    expect(data).toEqual([{ a: 1, a_2: 2, "": 3, column4: 4 }])
  })

  it("keeps all four in parseCsvObjects", () => {
    const { data, headers } = parseCsvObjects("a,a,,\n1,2,3,4", {
      header: true,
      typeInference: true,
    })

    expect(headers).toEqual(["a", "a_2", "", "column4"])
    expect(data).toEqual([{ a: 1, a_2: 2, "": 3, column4: 4 }])
  })

  it("keeps all four in readXlsxObjects", async () => {
    const bytes = await writeXlsx({ sheets: [{ name: "S", rows: ROWS }] })

    const { data, headers } = await readXlsxObjects(bytes)

    expect(headers).toEqual(["a", "a_2", "", "column4"])
    expect(data).toEqual([{ a: 1, a_2: 2, "": 3, column4: 4 }])
  })

  it("reports the names it actually keyed by", () => {
    // A caller who wants the original spelling has `sheet.rows[0]`; what
    // `headers` has to be is the keys of `data`, or the two disagree.
    const { data, headers } = sheetToObjects({ name: "D", rows: ROWS })

    expect(Object.keys(data[0]!)).toEqual(headers)
  })

  it("runs transformHeader first, so a transform can prevent the collision", () => {
    // `sheetToObjects` takes no transform hooks by design; the readers do.
    const { headers } = parseCsvObjects("a,a\n1,2", {
      header: true,
      transformHeader: (h, i) => `${h}${i}`,
    })

    expect(headers).toEqual(["a0", "a1"])
  })
})
