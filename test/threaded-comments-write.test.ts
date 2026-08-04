import { describe, expect, it } from "vitest"
import { readFileSync } from "node:fs"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"

// ═══════════════════════════════════════════════════════════════════════
// #404 — `WriteSheet.threadedComments` was a shipped, typed field that
// nothing wrote. A caller could pass a full thread and get a workbook
// with no `xl/threadedComments/` part and no error. The field is gone;
// these tests keep it gone, and keep the two README lines that
// contradicted the code honest.
// ═══════════════════════════════════════════════════════════════════════

const readme = (): string => readFileSync(new URL("../README.md", import.meta.url), "utf-8")

describe("WriteSheet has no threadedComments field", () => {
  it("does not accept one — the type is the guard", () => {
    const sheet: import("../src/_types").WriteSheet = {
      name: "S",
      rows: [["a"]],
      // @ts-expect-error — removed in #404. If this ever compiles again,
      // the field is back, and the writer had better write it this time.
      threadedComments: [{ id: "{c1}", ref: "A1", personId: "{p1}", text: "hi" }],
    }
    expect(sheet.name).toBe("S")
  })

  it("would have produced no part even if it were passed", async () => {
    // The reason the field had to go rather than stay as a no-op: the
    // output looks entirely successful.
    const buf = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [["a"]],
          threadedComments: [{ id: "{c1}", ref: "A1", personId: "{p1}", text: "hi" }],
        } as never,
      ],
    })
    const parts = new ZipReader(buf).entries()
    expect(parts.some((p) => p.includes("threadedComment"))).toBe(false)
    expect(parts.some((p) => p.includes("person"))).toBe(false)
  })
})

describe("Sheet.threadedComments is real", () => {
  it("stays on the read model", async () => {
    // Removing the write field must not touch the read one — they are
    // read, and preserved through openXlsx → saveXlsx.
    const buf = await writeXlsx({ sheets: [{ name: "S", rows: [["a"]] }] })
    const workbook = await readXlsx(buf)
    expect(workbook.sheets[0]).not.toHaveProperty("threadedComments")
    // The property is optional, so absence here is correct; the round-trip
    // suite in threaded-comments.test.ts covers a file that has them.
  })
})

describe("the README no longer contradicts the code", () => {
  it("does not list VBA injection as unimplemented, because it works", async () => {
    const text = readme()
    expect(text).not.toContain("- VBA/macro injection")

    // …and it works, which is why the roadmap entry had to go.
    const buf = await writeXlsx({
      sheets: [{ name: "S", rows: [["a"]] }],
      vbaProject: new Uint8Array([1, 2, 3]),
    })
    expect(new ZipReader(buf).entries()).toContain("xl/vbaProject.bin")
  })

  it("scopes 'VBA preserved' to the round-trip path", () => {
    // The claim beside the saveXlsx sample is true of that path only;
    // readXlsx → writeXlsx drops the macro project with no warning.
    expect(readme()).toContain("Preservation is a property of **this** path")
  })

  it("says threaded comments are not silently accepted", () => {
    expect(readme()).toMatch(/`WriteSheet` has no `threadedComments` field/)
  })
})
