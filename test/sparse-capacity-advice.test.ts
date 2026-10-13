import { describe, expect, it } from "vitest"
import { oversizeSheetMessage } from "../src/xlsx/worksheet"

describe("dense-grid limit advice", () => {
  it.each([
    ["Data", 18_654, 16_384, 305_612_208, 82_000, "0.03%"],
    ["Log", 916_449, 33, 30_242_817, 28_413_226, "93.95%"],
  ] as const)(
    "offers sparse metadata for %s regardless of a flat Map's old ceiling",
    (name, rows, cols, total, count, density) => {
      const message = oversizeSheetMessage(name, rows, cols, total, count, 20_000_000)
      expect(message).toContain('Sheet "' + name + '"')
      expect(message).toContain(density + " of them filled")
      expect(message).toContain("readXlsx(input, { sparse: true })")
      expect(message).toContain("streamXlsxRows")
      expect(message).toContain("`range` or `maxRows`")
      expect(message).toContain("`maxTotalCells`")
      expect(message).not.toContain("cannot help")
    },
  )
})
