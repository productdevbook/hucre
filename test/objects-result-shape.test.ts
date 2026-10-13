import { describe, expect, it } from "vitest"
import { readObjects } from "../src/defter"
import { readXlsxObjects } from "../src/xlsx/objects"
import { readOdsObjects } from "../src/ods/objects"
import { writeXlsx } from "../src/xlsx/writer"
import { writeOds } from "../src/ods/writer"
import { parseCsvObjects } from "../src/csv/reader"
import { parseJson } from "../src/json/reader"
import { readXml } from "../src/xml/data-reader"
import { sheetToObjects } from "../src/sheet-utils"

// ── One result shape for every *Objects reader (#365 item 6) ─────────
// `readObjects` and `sheetToObjects` used to return a bare `T[]`, so the
// same "read a table as objects" idea had two incompatible shapes. v1
// freezes one: `{ data, headers }`.

const ROWS = [
  ["Name", "Age"],
  ["Alice", 30],
  ["Bob", 25],
]

const EXPECTED = [
  { Name: "Alice", Age: 30 },
  { Name: "Bob", Age: 25 },
]

describe("every *Objects reader returns { data, headers }", () => {
  it("readObjects", async () => {
    const result = await readObjects(await writeXlsx({ sheets: [{ name: "S", rows: ROWS }] }))
    expect(Object.keys(result).sort()).toEqual(["data", "headers"])
    expect(result.headers).toEqual(["Name", "Age"])
    expect(result.data).toEqual(EXPECTED)
  })

  it("readXlsxObjects", async () => {
    const result = await readXlsxObjects(await writeXlsx({ sheets: [{ name: "S", rows: ROWS }] }))
    expect(Object.keys(result).sort()).toEqual(["data", "headers"])
    expect(result.headers).toEqual(["Name", "Age"])
    expect(result.data).toEqual(EXPECTED)
  })

  it("readOdsObjects", async () => {
    const result = await readOdsObjects(await writeOds({ sheets: [{ name: "S", rows: ROWS }] }))
    expect(Object.keys(result).sort()).toEqual(["data", "headers"])
    expect(result.headers).toEqual(["Name", "Age"])
    expect(result.data).toEqual(EXPECTED)
  })

  it("parseCsvObjects", () => {
    const result = parseCsvObjects("Name,Age\nAlice,30\nBob,25", {
      hasHeaderRow: true,
      typeInference: true,
    })
    expect(Object.keys(result).sort()).toEqual(["data", "headers"])
    expect(result.headers).toEqual(["Name", "Age"])
    expect(result.data).toEqual(EXPECTED)
  })

  it("parseJson", () => {
    const result = parseJson(JSON.stringify(EXPECTED))
    expect(Object.keys(result).sort()).toEqual(["data", "headers"])
    expect(result.headers).toEqual(["Name", "Age"])
    expect(result.data).toEqual(EXPECTED)
  })

  it("readXml", () => {
    const result = readXml(
      "<rows><row><Name>Alice</Name><Age>30</Age></row><row><Name>Bob</Name><Age>25</Age></row></rows>",
      { rowTag: "row" },
    )
    // The one superset: readXml also reports the row tag it used.
    expect(Object.keys(result).sort()).toEqual(["data", "headers", "rowTag"])
    expect(result.headers).toEqual(["Name", "Age"])
    // readXml has no type inference, so values stay strings.
    expect(result.data).toEqual([
      { Name: "Alice", Age: "30" },
      { Name: "Bob", Age: "25" },
    ])
  })

  it("sheetToObjects", () => {
    const result = sheetToObjects({ name: "S", rows: ROWS })
    expect(Object.keys(result).sort()).toEqual(["data", "headers"])
    expect(result.headers).toEqual(["Name", "Age"])
    expect(result.data).toEqual(EXPECTED)
  })
})

describe("shared projection semantics", () => {
  it("readObjects, readXlsxObjects and readOdsObjects agree on the same table", async () => {
    const rows = [
      ["A", "B"],
      [null, null],
      [1, 2],
    ]
    const xlsx = await writeXlsx({ sheets: [{ name: "S", rows }] })
    const ods = await writeOds({ sheets: [{ name: "S", rows }] })

    const viaRead = await readObjects(xlsx)
    const viaXlsx = await readXlsxObjects(xlsx)
    const viaOds = await readOdsObjects(ods)

    expect(viaRead).toEqual(viaXlsx)
    expect(viaRead).toEqual(viaOds)
    // Empty rows skipped by default, in all three.
    expect(viaRead.data).toEqual([{ A: 1, B: 2 }])
  })

  it("keeps empty-string header keys everywhere", async () => {
    const rows = [
      ["A", "", "B"],
      [1, 2, 3],
    ]
    const xlsx = await writeXlsx({ sheets: [{ name: "S", rows }] })

    const viaRead = await readObjects(xlsx)
    const viaXlsx = await readXlsxObjects(xlsx)
    const viaCsv = parseCsvObjects("A,,B\n1,2,3", { hasHeaderRow: true, typeInference: true })

    expect(viaRead).toEqual(viaXlsx)
    expect(viaRead.data).toEqual([{ A: 1, "": 2, B: 3 }])
    expect(viaCsv.data).toEqual([{ A: 1, "": 2, B: 3 }])
  })
})
