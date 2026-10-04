import { describe, expect, it } from "vitest"
import { cellEntries, getCell } from "../src/cell-store"
import { readOds } from "../src/ods/reader"
import { ParseError } from "../src/errors"
import { odsFromContent } from "./support/ods"

const number = (value: number, attrs = "") =>
  `<table:table-cell office:value-type="float" office:value="${value}" ${attrs}/>`
const row = (cells: string, attrs = "") => `<table:table-row ${attrs}>${cells}</table:table-row>`
const file = (rows: string, styles = "") =>
  odsFromContent(`<table:table table:name="S">${rows}</table:table>`, styles)

describe("ODS expansion contracts", () => {
  it.each(["populated", "empty"])("counts earlier width before %s row repeats", async (kind) => {
    const bytes = await file(
      row(number(1) + number(2) + number(3)) +
        (kind === "populated"
          ? row(number(4), 'table:number-rows-repeated="2"')
          : row("", 'table:number-rows-repeated="1"') + row(number(4))),
    )
    await expect(readOds(bytes, { maxTotalCells: 8 })).rejects.toThrow(ParseError)
    expect((await readOds(bytes, { maxTotalCells: 9 })).sheets[0].rows).toEqual([
      [1, 2, 3],
      kind === "populated" ? [4, null, null] : [null, null, null],
      [4, null, null],
    ])
  })

  it("copies metadata on every repeated row, with independent cell records", async () => {
    const bytes = await file(
      row(
        number(7, 'table:style-name="bold" table:formula="of:=1+6"') +
          '<table:table-cell office:value-type="string"><text:p><text:a xlink:href="https://example.com">link</text:a></text:p></table:table-cell>',
        'table:number-rows-repeated="3"',
      ),
      '<style:style style:name="bold" style:family="table-cell"><style:text-properties fo:font-weight="bold"/></style:style>',
    )
    const sheet = (await readOds(bytes, { readStyles: true })).sheets[0]
    for (let r = 0; r < 3; r++) {
      expect(getCell(sheet.cells, r, 0)).toMatchObject({
        value: 7,
        type: "formula",
        formula: "1+6",
        formulaResult: 7,
        style: { font: { bold: true } },
      })
      expect(getCell(sheet.cells, r, 1)?.hyperlink?.target).toBe("https://example.com")
    }
    expect(getCell(sheet.cells, 1, 0)).not.toBe(getCell(sheet.cells, 0, 0))
  })

  it("copies horizontal merges to each repeated row", async () => {
    const sheet = (
      await readOds(
        await file(
          row(
            number(1, 'table:number-columns-spanned="2"') +
              "<table:covered-table-cell/>" +
              number(2),
            'table:number-rows-repeated="2"',
          ),
        ),
      )
    ).sheets[0]
    expect(sheet.merges).toEqual([
      { startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
      { startRow: 1, startCol: 0, endRow: 1, endCol: 1 },
    ])
  })

  it("applies maxRows before repeated allocation and metadata collection", async () => {
    const sheet = (
      await readOds(
        await file(row(number(7, 'table:formula="of:=7"'), 'table:number-rows-repeated="100"')),
        { maxRows: 2, maxTotalCells: 2 },
      )
    ).sheets[0]
    expect(sheet.rows).toEqual([[7], [7]])
    expect([...cellEntries(sheet.cells)].map(([r, c]) => [r, c])).toEqual([
      [0, 0],
      [1, 0],
    ])
  })

  it("skips metadata past maxRows after an empty repeated run", async () => {
    const sheet = (
      await readOds(
        await file(
          row("", 'table:number-rows-repeated="100"') + row(number(7, 'table:formula="of:=7"')),
        ),
        { maxRows: 2, maxTotalCells: 2 },
      )
    ).sheets[0]
    expect(sheet.rows).toEqual([])
    expect(sheet.cells).toBeUndefined()
  })

  it("does not let a malformed repeat poison subsequent coordinates", async () => {
    const sheet = (
      await readOds(
        await file(
          row(number(1, 'table:formula="of:=1"'), 'table:number-rows-repeated="bad"') +
            row("", 'table:number-rows-repeated="bad"') +
            row(number(7, 'table:formula="of:=7"')),
        ),
      )
    ).sheets[0]
    expect(sheet.rows).toEqual([[7]])
    expect([...cellEntries(sheet.cells)].map(([r, c, cell]) => [r, c, cell.value])).toEqual([
      [0, 0, 7],
    ])
  })

  it("clamps the aggregate column count, including metadata", async () => {
    const sheet = (
      await readOds(
        await file(
          row(
            number(1, 'table:number-columns-repeated="16384"') + number(2, 'table:formula="of:=2"'),
          ),
        ),
      )
    ).sheets[0]
    expect(sheet.rows[0]).toHaveLength(16384)
    expect(sheet.rows[0][16383]).toBe(1)
    expect(sheet.cells).toBeUndefined()
  })
})
