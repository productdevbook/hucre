import { describe, expect, it } from "vitest"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"
import { readXlsx } from "../src/xlsx/reader"
import { writeXlsx } from "../src/xlsx/writer"
import type { AutoFilter } from "../src/_types"
import { InvalidArgumentError } from "../src/errors"
import { cloneSheet } from "../src/sheet-ops"

describe("auto filter columns: review follow-ups (#565)", () => {
  const rows = [
    ["a", "b", "c"],
    [1, 2, 3],
  ]
  const write = (af: AutoFilter) => writeXlsx({ sheets: [{ name: "S", rows, autoFilter: af }] })

  it("writes and reads a column that only hides its button", async () => {
    const af: AutoFilter = {
      range: "A1:C2",
      columns: [
        { colIndex: 1, hiddenButton: true },
        { colIndex: 2, showButton: false },
      ],
    }
    expect((await readXlsx(await write(af))).sheets[0].autoFilter).toEqual(af)
  })

  it("round-trips the attributes beside the criteria", async () => {
    const af = (await withAutoFilterXml(
      `<autoFilter ref="A1:D3">` +
        `<filterColumn colId="0"><filters calendarType="hijri"><filter val="1"/></filters></filterColumn>` +
        `<filterColumn colId="1"><top10 val="5" filterVal="42"/></filterColumn>` +
        `<filterColumn colId="2"><dynamicFilter type="between" val="1" maxVal="2" valIso="2024-01-01T00:00:00" maxValIso="2024-02-01T00:00:00"/></filterColumn>` +
        `</autoFilter>`,
    ))!
    expect(af.columns).toEqual([
      { colIndex: 0, filters: ["1"], calendarType: "hijri" },
      { colIndex: 1, top10: { val: 5, filterVal: 42 } },
      {
        colIndex: 2,
        dynamic: {
          type: "between",
          val: 1,
          maxVal: 2,
          valIso: "2024-01-01T00:00:00",
          maxValIso: "2024-02-01T00:00:00",
        },
      },
    ])
    expect((await readXlsx(await write(af))).sheets[0].autoFilter).toEqual(af)
  })

  it("refuses more than two custom conditions instead of truncating", async () => {
    const c = { operator: "equal" as const, value: "x" }
    await expect(
      write({ range: "A1:C2", columns: [{ colIndex: 0, custom: { conditions: [c, c, c] } }] }),
    ).rejects.toThrow(InvalidArgumentError)
  })

  it("refuses a column with more than one criterion kind", async () => {
    await expect(
      write({
        range: "A1:C2",
        columns: [{ colIndex: 0, filters: ["a"], top10: { val: 3 } }],
      }),
    ).rejects.toThrow(/more than one criterion/)
  })

  it("cloneSheet detaches the auto filter's columns", async () => {
    const af: AutoFilter = { range: "A1:C2", columns: [{ colIndex: 0, filters: ["1"] }] }
    const sheet = (await readXlsx(await write(af))).sheets[0]
    const clone = cloneSheet(sheet, "Copy")
    clone.autoFilter!.columns![0].filters!.push("2")
    clone.autoFilter!.columns!.push({ colIndex: 1, hiddenButton: true })
    expect(sheet.autoFilter!.columns).toEqual([{ colIndex: 0, filters: ["1"] }])
  })
})

// #565: an auto filter on two or more columns came back with one column.
// Only `<filters><filter val/>` was read, so every column filtered by a
// custom, top-10, dynamic or date-group criterion was dropped.

const ROWS = [
  ["a", "b", "c", "d"],
  [1, 2, 3, 4],
]

async function withAutoFilterXml(autoFilterXml: string): Promise<AutoFilter | undefined> {
  const base = await writeXlsx({ sheets: [{ name: "S", rows: ROWS }] })
  const files = await new ZipReader(base).extractAll()
  const path = "xl/worksheets/sheet1.xml"
  const xml = new TextDecoder()
    .decode(files.get(path)!)
    .replace("</sheetData>", `</sheetData>${autoFilterXml}`)
  const zip = new ZipWriter()
  for (const [name, data] of files) {
    zip.add(name, name === path ? new TextEncoder().encode(xml) : data)
  }
  return (await readXlsx(await zip.build())).sheets[0].autoFilter
}

const MIXED = `<autoFilter ref="A1:D3">
  <filterColumn colId="0"><filters blank="1"><filter val="1"/><filter val="x"/></filters></filterColumn>
  <filterColumn colId="1"><customFilters and="1"><customFilter operator="greaterThan" val="5"/><customFilter operator="lessThan" val="9"/></customFilters></filterColumn>
  <filterColumn colId="2"><top10 top="0" percent="1" val="10"/></filterColumn>
  <filterColumn colId="3" hiddenButton="1"><dynamicFilter type="aboveAverage" val="2.5"/></filterColumn>
</autoFilter>`

describe("auto filter columns (#565)", () => {
  it("reads every filterColumn, whatever its criterion kind", async () => {
    const af = await withAutoFilterXml(MIXED)
    expect(af).toEqual({
      range: "A1:D3",
      columns: [
        { colIndex: 0, filters: ["1", "x"], blank: true },
        {
          colIndex: 1,
          custom: {
            and: true,
            conditions: [
              { operator: "greaterThan", value: "5" },
              { operator: "lessThan", value: "9" },
            ],
          },
        },
        { colIndex: 2, top10: { bottom: true, percent: true, val: 10 } },
        { colIndex: 3, hiddenButton: true, dynamic: { type: "aboveAverage", val: 2.5 } },
      ],
    })
  })

  it("reads a customFilter with no operator as equal, and top10 as top", async () => {
    const af = await withAutoFilterXml(
      `<autoFilter ref="A1:D3"><filterColumn colId="0"><customFilters><customFilter val="a*"/></customFilters></filterColumn><filterColumn colId="1"><top10 val="3"/></filterColumn></autoFilter>`,
    )
    expect(af?.columns).toEqual([
      { colIndex: 0, custom: { conditions: [{ operator: "equal", value: "a*" }] } },
      { colIndex: 1, top10: { val: 3 } },
    ])
  })

  it("reads date groups of a value filter", async () => {
    const af = await withAutoFilterXml(
      `<autoFilter ref="A1:D3"><filterColumn colId="0"><filters><dateGroupItem year="2024" month="3" dateTimeGrouping="month"/></filters></filterColumn></autoFilter>`,
    )
    expect(af?.columns).toEqual([
      { colIndex: 0, dateGroups: [{ year: 2024, month: 3, grouping: "month" }] },
    ])
  })

  it("leaves out a column whose only criterion is not modelled (colour filter)", async () => {
    const af = await withAutoFilterXml(
      `<autoFilter ref="A1:D3"><filterColumn colId="0"><colorFilter dxfId="0"/></filterColumn><filterColumn colId="1"><filters><filter val="2"/></filters></filterColumn></autoFilter>`,
    )
    expect(af?.columns).toEqual([{ colIndex: 1, filters: ["2"] }])
  })

  it("writes what it read, so a second read is identical", async () => {
    const first = (await withAutoFilterXml(MIXED))!
    const out = await writeXlsx({ sheets: [{ name: "S", rows: ROWS, autoFilter: first }] })
    expect((await readXlsx(out)).sheets[0].autoFilter).toEqual(first)
  })

  it("writes date groups", async () => {
    const af: AutoFilter = {
      range: "A1:B2",
      columns: [{ colIndex: 0, dateGroups: [{ year: 2024, grouping: "year" }] }],
    }
    const out = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [
            ["a", "b"],
            [1, 2],
          ],
          autoFilter: af,
        },
      ],
    })
    expect((await readXlsx(out)).sheets[0].autoFilter).toEqual(af)
  })
})
