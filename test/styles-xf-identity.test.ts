import { describe, expect, it } from "vitest"
import { createStylesCollector } from "../src/xlsx/styles-writer"
import { writeXlsx } from "../src/xlsx/writer"
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

describe("styles collector — xf reuse", () => {
  it("gives one xf to a style object reused across cells", () => {
    const styles = createStylesCollector(undefined, { reuseStyleIdentity: true })
    const shared: CellStyle = {
      font: { name: "Arial", size: 8 },
      border: { top: { style: "thin" }, bottom: { style: "thin" } },
    }

    const ids = Array.from({ length: 100 }, () => styles.addStyle(shared))

    expect(new Set(ids).size).toBe(1)
  })

  it("still collapses distinct objects that describe the same format", () => {
    const styles = createStylesCollector(undefined, { reuseStyleIdentity: true })

    const first = styles.addStyle({ font: { name: "Arial", size: 8 } })
    const second = styles.addStyle({ font: { name: "Arial", size: 8 } })

    expect(second).toBe(first)
  })

  it("keeps different formats apart", () => {
    const styles = createStylesCollector(undefined, { reuseStyleIdentity: true })

    const arial = styles.addStyle({ font: { name: "Arial", size: 8 } })
    const manrope = styles.addStyle({ font: { name: "Manrope", size: 20 } })

    expect(manrope).not.toBe(arial)
  })

  it("emits one cellXfs entry per distinct format, whatever the call pattern", () => {
    const styles = createStylesCollector(undefined, { reuseStyleIdentity: true })
    const shared: CellStyle = { font: { name: "Arial", size: 8 } }

    styles.addStyle(shared)
    styles.addStyle({ font: { name: "Arial", size: 8 } })
    styles.addStyle(shared)
    styles.addStyle({ font: { name: "Manrope", size: 20 } })

    const xml = styles.toXml()

    expect(xml.match(/<cellXfs count="(\d+)"/)?.[1]).toBe("3")
  })

  it("re-reads a mutated style when identity reuse is off", () => {
    const styles = createStylesCollector()
    const shared: CellStyle = { numFmt: "0.00" }

    const before = styles.addStyle(shared)
    shared.numFmt = "0.0000"
    const after = styles.addStyle(shared)

    expect(after).not.toBe(before)
  })
})

describe("style identity reuse is scoped to the writer that can rely on it", () => {
  it("writeXlsx reuses the format across rows of a column", async () => {
    const style: CellStyle = { numFmt: "0.00", font: { name: "Arial" } }
    const bytes = await writeXlsx({
      sheets: [{ name: "S", columns: [{ style }], rows: [[1], [2], [3]] }],
    })

    const wb = await readXlsx(bytes, { readStyles: true })
    const cells = wb.sheets[0]!.cells!

    for (const ref of ["0,0", "1,0", "2,0"]) {
      expect(cells.get(ref)!.style!.numFmt).toBe("0.00")
    }
    expect(bytes.length).toBeGreaterThan(0)
  })

  it("writeXlsxStream still sees a style mutated between rows", async () => {
    const style: CellStyle = { numFmt: "0.00" }

    function* rows() {
      yield [1]
      style.numFmt = "0.0000"
      yield [2]
    }

    const bytes = await collect(writeXlsxStream(rows(), { name: "S", columns: [{ style }] }))
    const wb = await readXlsx(bytes, { readStyles: true })
    const cells = wb.sheets[0]!.cells!

    expect(cells.get("0,0")!.style!.numFmt).toBe("0.00")
    expect(cells.get("1,0")!.style!.numFmt).toBe("0.0000")
  })
})
