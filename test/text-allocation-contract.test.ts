import { describe, expect, it } from "vitest"
import { read } from "../src/defter"
import { padToRectangle } from "../src/_grid"
import { parseJson, jsonToWorkbook, parseNdjson } from "../src/json/reader"
import { readXml } from "../src/xml/data-reader"
import { fromHtml } from "../src/export/html-import"
import { ParseError } from "../src/errors"

const enc = new TextEncoder()
const records = [{ a: 1, b: 2, c: 3 }, { a: 4 }]
const json = JSON.stringify(records)
const ndjson = records.map((r) => JSON.stringify(r)).join("\n")
const xml = "<rows><row><a>1</a><b>2</b><c>3</c></row><row><a>4</a></row></rows>"
const html =
  "<table><tr><td>1</td><td>2</td><td>3</td></tr><tr><td>4</td></tr><tr><td>4</td></tr></table>"

describe("dense text allocation limits", () => {
  it.each([
    ["CSV", "a,b,c\nx\nx", 9],
    ["JSON", json, 9],
    ["NDJSON", ndjson, 9],
    ["XML", xml, 9],
    ["HTML", html, 9],
  ] as const)("forwards the limit through read(%s)", async (_, text, slots) => {
    await expect(read(enc.encode(text), { maxTotalCells: slots - 1 })).rejects.toThrow(ParseError)
    const sheet = (await read(enc.encode(text), { maxTotalCells: slots })).sheets[0]
    expect(sheet.rows).toHaveLength(3)
    expect(sheet.rows.every((r) => r.length === 3)).toBe(true)
  })

  it.each([
    ["JSON", parseJson, json],
    ["NDJSON", parseNdjson, ndjson],
    ["XML", readXml, xml],
  ] as const)("bounds %s record expansion before invoking cell transforms", (_, parse, text) => {
    let calls = 0
    expect(() =>
      parse(text, {
        maxTotalCells: 5,
        transformValue(value) {
          calls++
          return value
        },
      }),
    ).toThrow(ParseError)
    expect(calls).toBe(0)
    expect(parse(text, { maxTotalCells: 6 }).data).toHaveLength(2)
  })

  it("counts JSON workbook headers before record expansion", () => {
    let calls = 0
    expect(() =>
      jsonToWorkbook(json, {
        maxTotalCells: 8,
        transformValue(value) {
          calls++
          return value
        },
      }),
    ).toThrow(ParseError)
    expect(calls).toBe(0)
    expect(jsonToWorkbook(json, { maxTotalCells: 9 }).sheets[0].rows).toEqual([
      ["a", "b", "c"],
      [1, 2, 3],
      [4, null, null],
    ])
  })

  it("counts earlier HTML width before accepting a short row", () => {
    expect(() => fromHtml(html, { maxTotalCells: 8 })).toThrow(ParseError)
    expect(fromHtml(html, { maxTotalCells: 9 }).rows).toEqual([
      [1, 2, 3],
      [4, null, null],
      [4, null, null],
    ])
  })

  it("rejects rectangular padding before modifying the short row", () => {
    const rows = [[1, 2, 3], [4]]
    expect(() => padToRectangle(rows, 5)).toThrow(ParseError)
    expect(rows).toEqual([[1, 2, 3], [4]])
    expect(padToRectangle(rows, 6)).toEqual([
      [1, 2, 3],
      [4, null, null],
    ])
  })
})
