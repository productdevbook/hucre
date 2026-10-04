import { expect, it } from "vitest"
import { readXlsx, writeXlsx } from "../src/xlsx"
import { ParseError } from "../src/errors"
import type { ReadWarning } from "../src/_types"
import { patchXml } from "./support/xlsx"

const parts = ["xl/worksheets/sheet1.xml", "xl/comments1.xml"] as const

async function withReference(part: string, ref: string): Promise<Uint8Array> {
  const input = await writeXlsx({
    sheets: [
      {
        name: "Metadata",
        rows: [
          [
            {
              value: 1,
              comment: { text: "note" },
              hyperlink: { target: "https://example.test" },
            },
          ],
        ],
      },
    ],
  })
  return patchXml(input, part, (xml) => xml.replace('ref="A1"', `ref="${ref}"`))
}

it.each(parts)("reports out-of-grid metadata from %s as file damage", async (part) => {
  for (const ref of ["XFE1", "A1048577"]) {
    await expect(readXlsx(await withReference(part, ref))).rejects.toThrow(ParseError)
  }
})

it.each(parts)("drops malformed metadata from %s with a warning", async (part) => {
  const warnings: ReadWarning[] = []
  const workbook = await readXlsx(await withReference(part, "A0"), {
    onWarning: (warning) => warnings.push(warning),
  })
  expect(workbook.sheets[0].rows).toEqual([[1]])
  expect(warnings).toContainEqual(
    expect.objectContaining({ code: "malformed-cell-ref", sheet: "Metadata" }),
  )
})
