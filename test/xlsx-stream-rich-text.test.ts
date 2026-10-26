import { describe, expect, it } from "vitest"
import type { CellValue, RichTextRun } from "../src/_types"
import { readXlsx } from "../src/xlsx/reader"
import {
  writeXlsxStream,
  writeXlsxStreamSheets,
  XlsxStreamWriter,
  type StreamStyledCell,
} from "../src/xlsx/stream-writer"
import { ZipReader } from "../src/zip/reader"

type Row = Array<CellValue | StreamStyledCell>
const modes = [
  "single-inline",
  "single-shared",
  "multi-inline",
  "multi-shared",
  "incremental",
] as const
type Mode = (typeof modes)[number]

async function write(mode: Mode, rows: Iterable<Row>, maxRowsPerSheet = 10): Promise<Uint8Array> {
  const options = { name: "Runs", maxRowsPerSheet }
  if (mode === "incremental") {
    const writer = new XlsxStreamWriter(options)
    for (const row of rows) writer.addRow(row)
    return writer.finish()
  }
  const inlineStrings = mode.endsWith("inline")
  const stream = mode.startsWith("single")
    ? writeXlsxStream(rows, { ...options, inlineStrings })
    : writeXlsxStreamSheets([{ ...options, rows }], { inlineStrings })
  return new Uint8Array(await new Response(stream).arrayBuffer())
}

// The buffered writer already carried richText into the shared serializer,
// but the streaming wrappers dropped it entirely. Exercise both paths and
// both string modes, including the header snapshot used for rollover (#637).
describe.each(modes)("streaming rich text — %s (#637)", (mode) => {
  it("preserves runs and cell styles, with or without a fallback value", async () => {
    const richText: RichTextRun[] = [
      { text: "=SUM(A1:A2)", font: { bold: true, color: { rgb: "0000FF" } } },
      { text: " <&> " },
    ]
    const bytes = await write(mode, [
      [
        { richText },
        { value: "fallback", richText, style: { alignment: { horizontal: "center" } } },
      ],
    ])
    const sheet = (await readXlsx(bytes, { readStyles: true })).sheets[0]
    for (const column of [0, 1]) {
      const cell = sheet.cells?.get(`0,${column}`)
      expect(cell?.value).toBe("=SUM(A1:A2) <&> ")
      expect(cell?.richText).toEqual(richText)
      expect(cell?.formula).toBeUndefined()
    }
    expect(sheet.cells?.get("0,1")?.style?.alignment?.horizontal).toBe("center")

    const zip = new ZipReader(bytes)
    const xml = new TextDecoder().decode(await zip.extract("xl/worksheets/sheet1.xml"))
    if (mode.endsWith("inline")) {
      expect(xml).toContain('t="inlineStr"')
      expect(zip.has("xl/sharedStrings.xml")).toBe(false)
    } else {
      expect(xml).toContain('t="s"')
      expect(xml).not.toContain("inlineStr")
      const strings = new TextDecoder().decode(await zip.extract("xl/sharedStrings.xml"))
      expect(strings).toContain("<rPr>")
      expect(strings).toContain("=SUM(A1:A2)")
    }
  })

  it("keeps the original rich-text header when callers mutate nested runs before rollover", async () => {
    const richText: RichTextRun[] = [
      { text: "Header", font: { bold: true, color: { rgb: "0000FF" } } },
    ]
    const original = structuredClone(richText)
    const cell: StreamStyledCell = { richText }
    function* rows(): Generator<Row> {
      yield [cell]
      richText[0].text = "Changed"
      richText[0].font!.color!.rgb = "FF0000"
      richText.push({ text: " tail" })
      yield [cell]
      yield ["last"]
    }
    const workbook = await readXlsx(await write(mode, rows(), 2))
    expect(workbook.sheets).toHaveLength(2)
    for (const sheet of workbook.sheets) {
      expect(sheet.rows[0][0]).toBe("Header")
      expect(sheet.cells?.get("0,0")?.richText).toEqual(original)
    }
    expect(workbook.sheets[0].rows[1][0]).toBe("Changed tail")
    expect(workbook.sheets[0].cells?.get("1,0")?.richText).toEqual(richText)
    expect(workbook.sheets[1].rows[1][0]).toBe("last")
  })

  it("falls back to values and formulas when the runs are empty", async () => {
    const bytes = await write(mode, [
      [
        { richText: [] },
        { value: "plain", richText: [] },
        { value: 3, formula: "1+2", richText: [] },
      ],
    ])
    const sheet = (await readXlsx(bytes)).sheets[0]
    expect(sheet.rows[0][0] ?? null).toBeNull()
    expect(sheet.rows[0][1]).toBe("plain")
    expect(sheet.cells?.get("0,1")?.richText).toBeUndefined()
    expect(sheet.cells?.get("0,2")?.value).toBe(3)
    expect(sheet.cells?.get("0,2")?.formula).toBe("1+2")
  })
})
