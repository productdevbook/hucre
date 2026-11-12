import { afterEach, describe, expect, it } from "vitest"
import { parseCsv } from "../src/csv/reader"
import { parseJson } from "../src/json/reader"

const originalTz = process.env.TZ

describe("typeInference of a timestamp without a zone", () => {
  afterEach(() => {
    process.env.TZ = originalTz
  })

  it("reads it as UTC in CSV and JSON whatever the machine's zone", () => {
    process.env.TZ = "America/New_York"

    const rows = parseCsv("event,start\nstandup,2026-03-05T10:00:00\n", { typeInference: true })
    expect((rows[1]![1] as Date).toISOString()).toBe("2026-03-05T10:00:00.000Z")

    const { data } = parseJson('[{"start":"2026-03-05T10:00:00"}]', { typeInference: true })
    expect((data[0]!.start as Date).toISOString()).toBe("2026-03-05T10:00:00.000Z")
  })
})
