import { describe, expect, it } from "vitest"
import { createStylesCollector } from "../src/xlsx/styles-writer"
import { writeXlsxStream } from "../src/xlsx/stream-writer"
import { readXlsx } from "../src/xlsx/reader"
import type { CellStyle } from "../src/_types"

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = []
  let total = 0
  const reader = stream.getReader()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    chunks.push(value)
    total += value.length
  }
  const out = new Uint8Array(total)
  let offset = 0
  for (const chunk of chunks) {
    out.set(chunk, offset)
    offset += chunk.length
  }
  return out
}

describe("styles are snapshotted when registered", () => {
  it("keeps the font a cell was given after the caller mutates it", () => {
    const styles = createStylesCollector()
    const font = { name: "Arial", bold: false }

    styles.addStyle({ font })
    font.bold = true

    expect(styles.toXml()).toContain('<name val="Arial"/>')
    expect(styles.toXml().match(/<b\/>/g) ?? []).toHaveLength(0)
  })

  it("keeps the fill colour a cell was given", () => {
    const styles = createStylesCollector()
    const fgColor = { rgb: "FFFF0000" }

    styles.addStyle({ fill: { type: "pattern", pattern: "solid", fgColor } })
    fgColor.rgb = "FF00FF00"

    expect(styles.toXml()).toContain('rgb="FFFF0000"')
    expect(styles.toXml()).not.toContain('rgb="FF00FF00"')
  })

  it("keeps the border a cell was given", () => {
    const styles = createStylesCollector()
    const top = { style: "thin" as const }

    styles.addStyle({ border: { top } })
    top.style = "thick" as typeof top.style

    expect(styles.toXml()).toContain('style="thin"')
    expect(styles.toXml()).not.toContain('style="thick"')
  })

  it("keeps the alignment a cell was given", () => {
    const styles = createStylesCollector()
    const alignment = { horizontal: "left" as const }

    styles.addStyle({ alignment })
    alignment.horizontal = "right" as typeof alignment.horizontal

    expect(styles.toXml()).toContain('horizontal="left"')
    expect(styles.toXml()).not.toContain('horizontal="right"')
  })

  it("does not retroactively restyle earlier rows of a stream", async () => {
    const font = { name: "Arial", bold: false }
    const style: CellStyle = { font }

    function* rows() {
      yield [1]
      font.bold = true
      yield [2]
    }

    const bytes = await collect(writeXlsxStream(rows(), { name: "S", columns: [{ style }] }))
    const wb = await readXlsx(bytes, { readStyles: true })
    const cells = wb.sheets[0]!.cells!

    expect(cells.get("0,0")!.style!.font!.bold ?? false).toBe(false)
    expect(cells.get("1,0")!.style!.font!.bold).toBe(true)
  })
})
