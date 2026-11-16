import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"

// Built-in id 14 is Excel's locale short date. The reader resolves it to
// "m/d/yyyy", and the writer used to register that code as custom format
// 164, so every date in a round-tripped workbook turned into a literal
// month-first date (3/15/2026 instead of 15.03.2026 in a Polish Excel).

async function part(bytes: Uint8Array, path: string): Promise<string> {
  return new TextDecoder().decode(await new ZipReader(bytes).extract(path))
}

async function rewrite(
  bytes: Uint8Array,
  path: string,
  edit: (xml: string) => string,
): Promise<Uint8Array> {
  const all = await new ZipReader(bytes).extractAll()
  const zw = new ZipWriter()
  for (const [name, data] of all) {
    zw.add(
      name,
      name === path ? new TextEncoder().encode(edit(new TextDecoder().decode(data))) : data,
    )
  }
  return zw.build()
}

const cellXfs = (styles: string) => styles.match(/<cellXfs[^>]*>[\s\S]*<\/cellXfs>/)![0]

/** A date cell styled with built-in id 14 and no <numFmt> entry, as Excel writes it. */
async function excelDate(): Promise<Uint8Array> {
  const base = await writeXlsx({
    sheets: [
      {
        name: "S",
        rows: [[46096]],
        cells: new Map([["0,0", { value: 46096, style: { numFmt: "yyyy-mm-dd" } }]]),
      },
    ],
  })
  return rewrite(base, "xl/styles.xml", (x) =>
    x.replace(/<numFmts[^>]*>[\s\S]*?<\/numFmts>/, "").replace(/numFmtId="164"/g, 'numFmtId="14"'),
  )
}

describe("built-in number format ids", () => {
  it("keeps id 14 through openXlsx -> saveXlsx", async () => {
    const out = await saveXlsx(await openXlsx(await excelDate(), { readStyles: true }))
    const styles = await part(out, "xl/styles.xml")
    expect(cellXfs(styles)).toContain('numFmtId="14"')
    expect(styles).not.toContain('formatCode="m/d/yyyy"')
  })

  it("writes a built-in format code under its id, with no <numFmt> entry", async () => {
    const styles = await part(
      await writeXlsx({
        sheets: [
          {
            name: "S",
            rows: [[1]],
            cells: new Map([["0,0", { value: 1, style: { numFmt: "0.00%" } }]]),
          },
        ],
      }),
      "xl/styles.xml",
    )
    expect(cellXfs(styles)).toContain('numFmtId="10"')
    expect(styles).not.toContain("<numFmts")
  })

  it("keeps a literal dollar format custom: ids 5-8 mean the locale's currency", async () => {
    const styles = await part(
      await writeXlsx({
        sheets: [
          {
            name: "S",
            rows: [[1]],
            cells: new Map([["0,0", { value: 1, style: { numFmt: "$#,##0_);\\($#,##0\\)" } }]]),
          },
        ],
      }),
      "xl/styles.xml",
    )
    expect(cellXfs(styles)).toContain('numFmtId="164"')
  })

  it("keeps the space-spelled accounting formats custom: Excel defines 37-40 with _);", async () => {
    const styles = await part(
      await writeXlsx({
        sheets: [
          {
            name: "S",
            rows: [[1]],
            cells: new Map([["0,0", { value: 1, style: { numFmt: "#,##0 ;(#,##0)" } }]]),
          },
        ],
      }),
      "xl/styles.xml",
    )
    expect(cellXfs(styles)).toContain('numFmtId="164"')
    expect(styles).toContain('formatCode="#,##0 ;(#,##0)"')
  })

  it("keeps the ECMA-spelled elapsed-minutes format custom: Excel defines 47 with a colon", async () => {
    const styles = await part(
      await writeXlsx({
        sheets: [
          {
            name: "S",
            rows: [[1]],
            cells: new Map([["0,0", { value: 1, style: { numFmt: "mmss.0" } }]]),
          },
        ],
      }),
      "xl/styles.xml",
    )
    expect(cellXfs(styles)).toContain('numFmtId="164"')
    expect(styles).toContain('formatCode="mmss.0"')
  })
})
