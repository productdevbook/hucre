import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"

// #622: a <dataValidation> with no `type` (an input message and nothing
// else) was dropped on read. `type` defaults to "none" (§18.3.1.32).

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

async function template(): Promise<Uint8Array> {
  const base = await writeXlsx({
    sheets: [
      {
        name: "S",
        rows: [["x"]],
        dataValidations: [{ type: "whole", range: "A1", formula1: "1" }],
      },
    ],
  })
  return rewrite(base, "xl/worksheets/sheet1.xml", (x) =>
    x.replace(
      /<dataValidation [^>]*?(\/>|>[\s\S]*?<\/dataValidation>)/,
      '<dataValidation allowBlank="1" showInputMessage="1" promptTitle="Hi" prompt="Fill me" sqref="B2:B5"/>',
    ),
  )
}

describe("a data validation without a type (#622)", () => {
  it("is read as type none", async () => {
    const dv = (await readXlsx(await template())).sheets[0]!.dataValidations
    expect(dv).toEqual([
      {
        type: "none",
        range: "B2:B5",
        allowBlank: true,
        showInputMessage: true,
        inputTitle: "Hi",
        inputMessage: "Fill me",
      },
    ])
  })

  it("survives open -> save, and is written without a type attribute", async () => {
    const out = await saveXlsx(await openXlsx(await template()))
    const bytes = await new ZipReader(out).extract("xl/worksheets/sheet1.xml")
    const xml = new TextDecoder().decode(bytes)
    expect(xml).toMatch(/<dataValidation [^>]*prompt="Fill me"/)
    expect(xml).not.toMatch(/<dataValidation [^>]*type=/)
    expect((await readXlsx(out)).sheets[0]!.dataValidations?.[0]?.inputMessage).toBe("Fill me")
  })
})
