import { describe, expect, it } from "vitest"
import { readdirSync, readFileSync } from "node:fs"
import { replaceA1Refs, replaceA1Ranges } from "../src/cell-utils"

// #591: JavaScriptCore before Safari 16.4 throws a SyntaxError for a regex
// containing lookbehind, and a bundler cannot transpile a regex. A
// module-level one therefore blanks the page of every app that bundles hucre.

function sourceFiles(dir: URL): URL[] {
  const out: URL[] = []
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const child = new URL(entry.name + (entry.isDirectory() ? "/" : ""), dir)
    if (entry.isDirectory()) out.push(...sourceFiles(child))
    else if (/\.(ts|mts|js|mjs)$/.test(entry.name)) out.push(child)
  }
  return out
}

describe("no regex lookbehind in src (#591)", () => {
  it("no source file contains (?<= or (?<!", () => {
    const offenders: string[] = []
    for (const file of sourceFiles(new URL("../src/", import.meta.url))) {
      const lines = readFileSync(file, "utf-8").split("\n")
      lines.forEach((line, i) => {
        if (/\(\?<[=!]/.test(line)) offenders.push(`${file.pathname}:${i + 1}`)
      })
    }
    expect(offenders).toEqual([])
  })
})

describe("replaceA1Refs without lookbehind", () => {
  const wrap = (f: string) => replaceA1Refs(f, (r) => `[${r}]`)

  it("matches a ref at the start of the string", () => {
    expect(wrap("A1")).toBe("[A1]")
    expect(wrap("A1+B1")).toBe("[A1]+[B1]")
  })

  it("matches adjacent tokens separated by a single character", () => {
    expect(wrap("A1+B1+C1")).toBe("[A1]+[B1]+[C1]")
    expect(wrap("A1,B2")).toBe("[A1],[B2]")
    expect(wrap("A1 B2")).toBe("[A1] [B2]")
    expect(wrap("A1:B2,C3")).toBe("[A1]:[B2],[C3]")
  })

  it("keeps absolute refs whole", () => {
    expect(wrap("$A$1")).toBe("[$A$1]")
    expect(wrap("$A$1+$B2")).toBe("[$A$1]+[$B2]")
    expect(wrap("A$1*$B$2")).toBe("[A$1]*[$B$2]")
  })

  it("leaves identifiers, function names and string literals alone", () => {
    expect(wrap("LOG10(A1)")).toBe("LOG10([A1])")
    expect(wrap("ATAN2(A1,B1)")).toBe("ATAN2([A1],[B1])")
    expect(wrap('"AB1"&A1')).toBe('"AB1"&[A1]')
    expect(wrap("_A1+x.A1")).toBe("_A1+x.[A1]")
    expect(wrap("1A1")).toBe("1A1")
  })

  it("does not re-match across a consumed prefix", () => {
    // `$B2` directly after `A1` has `$` as its only possible prefix.
    expect(wrap("A1$B2")).toBe("[A1]$B2")
  })
})

describe("replaceA1Ranges without lookbehind", () => {
  const show = (f: string) =>
    replaceA1Ranges(f, (m) => `<${m.sheet1 ?? ""}|${m.ref1}|${m.sheet2 ?? ""}|${m.ref2 ?? ""}>`)

  it("matches at the start and for adjacent tokens", () => {
    expect(show("A1")).toBe("<|A1||>")
    expect(show("A1+B1")).toBe("<|A1||>+<|B1||>")
    expect(show("A1:B2,C3")).toBe("<|A1||B2>,<|C3||>")
    expect(show("SUM(A1:B2,C3)")).toBe("SUM(<|A1||B2>,<|C3||>)")
  })

  it("handles absolute refs", () => {
    expect(show("$A$1")).toBe("<|$A$1||>")
    expect(show("$A$1:$B$2")).toBe("<|$A$1||$B$2>")
  })

  it("handles sheet-qualified refs", () => {
    expect(show("Sheet2!A1")).toBe("<Sheet2|A1||>")
    expect(show("Sheet2!A1:B2+Sheet3!C3")).toBe("<Sheet2|A1||B2>+<Sheet3|C3||>")
    expect(show("'My Sheet'!A1")).toBe("<'My Sheet'|A1||>")
    expect(show("Sheet2.A1")).toBe("<Sheet2|A1||>")
    expect(show("Sheet1!A1:Sheet3!B2")).toBe("<Sheet1|A1|Sheet3|B2>")
  })

  it("rejects refs that are part of a longer name or qualifier", () => {
    expect(show("LOG10(1)")).toBe("LOG10(1)")
    expect(show("xA1")).toBe("xA1")
    expect(show('"A1"&B2')).toBe('"A1"&<|B2||>')
  })
})
