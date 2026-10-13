import { describe, expect, it, vi } from "vitest"
import { readXlsx, writeXlsx } from "../src/xlsx"

describe("cell storage capacity", () => {
  it("reads more cells than one metadata Map can hold", async () => {
    const rows = Array.from({ length: 135 * 128 }, (_, row) => (row % 128 === 0 ? [row + 1] : []))
    const input = await writeXlsx({ sheets: [{ name: "Log", rows }] })
    const originalSet = Map.prototype.set
    // Scale V8's per-Map ceiling down, without allocating millions of
    // cells. Limit only metadata maps so ZIP/style maps are unaffected.
    const guard = vi
      .spyOn(Map.prototype, "set")
      .mockImplementation(function (this: Map<unknown, unknown>, key, value) {
        if (
          value !== null &&
          typeof value === "object" &&
          "type" in value &&
          value.type === "number" &&
          this.size >= 128 &&
          !this.has(key)
        ) {
          throw new RangeError("Map maximum size exceeded")
        }
        return originalSet.call(this, key, value)
      })
    try {
      const workbook = await readXlsx(input, { sparse: true })
      expect(workbook.sheets[0].rows).toEqual([])
      expect(workbook.sheets[0].cells?.size).toBe(135)
    } finally {
      guard.mockRestore()
    }
  })
})
