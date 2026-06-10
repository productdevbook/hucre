import { describe, expect, it } from "vitest"
import { flattenValue } from "../src/json"

describe("flattenValue — __proto__ / constructor keys are preserved", () => {
  it("keeps a primitive value under a __proto__ key instead of dropping it", () => {
    // JSON.parse produces "__proto__" as an ordinary own property.
    const parsed = JSON.parse('{"__proto__": "x", "a": 1}')
    const flat = flattenValue(parsed)
    expect(flat["__proto__"]).toBe("x")
    expect(flat["a"]).toBe(1)
    expect(Object.keys(flat).sort()).toEqual(["__proto__", "a"])
  })

  it("does not pollute Object.prototype", () => {
    const parsed = JSON.parse('{"__proto__": {"polluted": true}}')
    flattenValue(parsed)
    expect(({} as Record<string, unknown>).polluted).toBeUndefined()
  })

  it("keeps a constructor key", () => {
    const parsed = JSON.parse('{"constructor": "boom"}')
    const flat = flattenValue(parsed)
    expect(flat["constructor"]).toBe("boom")
  })
})
