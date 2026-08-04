import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { consola } from "consola"
import {
  CliError,
  convertCommand,
  delimiterForExtension,
  detectFormatFromExtension,
  formatCellValue,
  inspectCommand,
  mainCommand,
  pkgVersion,
  readFile,
  validateCommand,
} from "../src/cli/commands"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { writeOds } from "../src/ods/writer"

// ═══════════════════════════════════════════════════════════════════════
// #399 — the CLI was the only module in the tree at 0% coverage, because
// src/cli.ts ended in a bare runMain(): importing it ran it. The commands
// now live in src/cli/commands.ts and are importable, so this suite runs
// them the way citty does — cmd.run({ args }) — with no child process.
// ═══════════════════════════════════════════════════════════════════════

let dir: string
let logs: string[]

/** Run a command's handler the way citty would, with defaults applied. */
function run(cmd: any, args: Record<string, unknown>): Promise<void> {
  const defaults: Record<string, unknown> = {}
  for (const [name, spec] of Object.entries(cmd.args as Record<string, any>)) {
    if (spec.default !== undefined) defaults[name] = spec.default
  }
  return cmd.run({ args: { ...defaults, ...args }, cmd, rawArgs: [] })
}

const path = (name: string) => join(dir, name)

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), "hucre-cli-"))
  logs = []
  // consola writes to stdout; capture instead of spraying the test output.
  for (const level of ["start", "success", "info", "log", "error"] as const) {
    vi.spyOn(consola, level).mockImplementation(((...a: unknown[]) => {
      logs.push(a.map(String).join(" "))
    }) as never)
  }
})

afterEach(() => {
  vi.restoreAllMocks()
  rmSync(dir, { recursive: true, force: true })
})

async function makeXlsx(name: string, rows: unknown[][], sheet = "Sheet1"): Promise<string> {
  const p = path(name)
  writeFileSync(p, await writeXlsx({ sheets: [{ name: sheet, rows: rows as never }] }))
  return p
}

// ── Format detection ────────────────────────────────────────────────

describe("detectFormatFromExtension", () => {
  it("maps the extensions it supports", () => {
    expect(detectFormatFromExtension("a.xlsx")).toBe("xlsx")
    expect(detectFormatFromExtension("a.ODS")).toBe("ods")
    expect(detectFormatFromExtension("a.csv")).toBe("csv")
    expect(detectFormatFromExtension("a.tsv")).toBe("csv")
  })

  it("rejects anything else, naming what is supported", () => {
    expect(() => detectFormatFromExtension("a.numbers")).toThrow(CliError)
    expect(() => detectFormatFromExtension("a.numbers")).toThrow(/\.xlsx, \.ods, \.csv, \.tsv/)
  })

  it("says so when there is no extension at all", () => {
    expect(() => detectFormatFromExtension("README")).toThrow(/\(none\)/)
  })
})

// ── The .tsv delimiter bug ──────────────────────────────────────────

describe("tab-separated files", () => {
  it("writes .tsv with tabs, not commas", async () => {
    // Both extensions map to the format "csv", and the writer then took
    // its default delimiter — so `convert x.xlsx out.tsv` produced a
    // comma-separated file whose name said otherwise.
    const input = await makeXlsx("in.xlsx", [
      ["h1", "h2"],
      ["b", 2],
    ])
    await run(convertCommand, { input, output: path("out.tsv") })

    const text = readFileSync(path("out.tsv"), "utf-8")
    expect(text).toContain("h1\th2")
    expect(text).not.toContain("h1,h2")
  })

  it("still writes .csv with commas", async () => {
    const input = await makeXlsx("in.xlsx", [["h1", "h2"]])
    await run(convertCommand, { input, output: path("out.csv") })
    expect(readFileSync(path("out.csv"), "utf-8")).toContain("h1,h2")
  })

  it("round-trips a tab-separated file without changing its separator", async () => {
    writeFileSync(path("in.tsv"), "a\tb\nc\td", "utf-8")
    await run(convertCommand, { input: path("in.tsv"), output: path("out.tsv") })
    const text = readFileSync(path("out.tsv"), "utf-8")
    expect(text).toContain("a\tb")
    expect(text).toContain("c\td")
  })

  it("reads a comma file named .tsv by its extension, not its content", () => {
    // Deliberate: the extension is the declaration. parseCsv's sniffing
    // is a fallback for input of unknown provenance, not a licence to
    // ignore what the caller named the file.
    expect(delimiterForExtension("x.tsv")).toBe("\t")
    expect(delimiterForExtension("x.csv")).toBe(",")
    expect(delimiterForExtension("x.xlsx")).toBe(",")
  })
})

// ── convert ─────────────────────────────────────────────────────────

describe("convert", () => {
  it("formats the first row like every other row", async () => {
    // Row 0 used to be pulled out as `headers` and stringified through
    // formatCellValue, while rows 1..n went through writeCsv — so the
    // same Date rendered two different ways depending on its row.
    const d = new Date(Date.UTC(2024, 0, 15))
    const input = await makeXlsx("in.xlsx", [[d], [d]])
    await run(convertCommand, { input, output: path("out.csv") })

    const [first, second] = readFileSync(path("out.csv"), "utf-8").trim().split(/\r?\n/)
    expect(first).toBe(second)
  })

  it("converts xlsx to ods", async () => {
    const input = await makeXlsx("in.xlsx", [["a", 1]], "Data")
    await run(convertCommand, { input, output: path("out.ods") })
    const back = await readFile(path("out.ods"))
    expect(back.sheets[0].name).toBe("Data")
    expect(back.sheets[0].rows[0]).toEqual(["a", 1])
  })

  it("converts ods to xlsx", async () => {
    writeFileSync(path("in.ods"), await writeOds({ sheets: [{ name: "S", rows: [["x", 2]] }] }))
    await run(convertCommand, { input: path("in.ods"), output: path("out.xlsx") })
    const back = await readXlsx(readFileSync(path("out.xlsx")))
    expect(back.sheets[0].rows[0]).toEqual(["x", 2])
  })

  it("converts csv to xlsx", async () => {
    writeFileSync(path("in.csv"), "a,b\n1,2", "utf-8")
    await run(convertCommand, { input: path("in.csv"), output: path("out.xlsx") })
    const back = await readXlsx(readFileSync(path("out.xlsx")))
    expect(back.sheets[0].rows).toEqual([
      ["a", "b"],
      ["1", "2"],
    ])
  })

  it("takes the first sheet only when the target is csv", async () => {
    const p = path("multi.xlsx")
    writeFileSync(
      p,
      await writeXlsx({
        sheets: [
          { name: "One", rows: [["first"]] },
          { name: "Two", rows: [["second"]] },
        ],
      }),
    )
    await run(convertCommand, { input: p, output: path("out.csv") })
    const text = readFileSync(path("out.csv"), "utf-8")
    expect(text).toContain("first")
    expect(text).not.toContain("second")
  })

  it("rejects an output extension it cannot write", async () => {
    const input = await makeXlsx("in.xlsx", [["a"]])
    await expect(run(convertCommand, { input, output: path("out.pdf") })).rejects.toThrow(CliError)
  })

  it("says in --help that it carries values only", () => {
    expect(convertCommand.meta).toMatchObject({
      description: expect.stringContaining("cell values only"),
    })
  })
})

// ── inspect ─────────────────────────────────────────────────────────

describe("inspect", () => {
  it("reports sheets, dimensions and a cell-type breakdown", async () => {
    const input = await makeXlsx(
      "in.xlsx",
      [
        ["a", 1, true],
        [new Date(0), null],
      ],
      "Mixed",
    )
    await run(inspectCommand, { file: input })

    const out = logs.join("\n")
    expect(out).toContain("Sheets: 1")
    expect(out).toContain('"Mixed"')
    expect(out).toContain("2 rows x 3 cols")
    expect(out).toMatch(/string: \d/)
    expect(out).toMatch(/number: \d/)
    expect(out).toMatch(/boolean: \d/)
  })

  it("prints document properties when the file carries them", async () => {
    const p = path("props.xlsx")
    writeFileSync(
      p,
      await writeXlsx({
        sheets: [{ name: "S", rows: [["a"]] }],
        properties: { title: "Quarterly", creator: "Ada", created: new Date(Date.UTC(2020, 0, 1)) },
      }),
    )
    await run(inspectCommand, { file: p })

    const out = logs.join("\n")
    expect(out).toContain("Title: Quarterly")
    expect(out).toContain("Creator: Ada")
    expect(out).toContain("2020-01-01")
  })

  it("previews a sheet with --sheet", async () => {
    const input = await makeXlsx("in.xlsx", [["header"], ["value"]])
    await run(inspectCommand, { file: input, sheet: "0" })
    expect(logs.join("\n")).toContain("header")
  })

  it("truncates a long cell in the preview", async () => {
    const input = await makeXlsx("in.xlsx", [["x".repeat(100)]])
    await run(inspectCommand, { file: input, sheet: "0" })
    expect(logs.join("\n")).toContain("...")
  })

  it("caps the preview at ten rows and says how many were left out", async () => {
    const rows = Array.from({ length: 25 }, (_, i) => [`r${i}`])
    const input = await makeXlsx("in.xlsx", rows)
    await run(inspectCommand, { file: input, sheet: "0" })

    const out = logs.join("\n")
    expect(out).toContain("and 15 more rows")
    expect(out).not.toContain("r10")
  })

  it("handles an empty sheet", async () => {
    writeFileSync(path("empty.csv"), "", "utf-8")
    await run(inspectCommand, { file: path("empty.csv"), sheet: "0" })
    expect(logs.join("\n")).toContain("(empty sheet)")
  })

  const badIndexes = ["9", "-1", "abc", "1.5"]
  for (const bad of badIndexes) {
    it(`rejects --sheet ${bad}`, async () => {
      const input = await makeXlsx("in.xlsx", [["a"]])
      await expect(run(inspectCommand, { file: input, sheet: bad })).rejects.toThrow(
        /Invalid sheet index/,
      )
    })
  }
})

// ── validate ────────────────────────────────────────────────────────

describe("validate", () => {
  const schema = {
    name: { type: "string", required: true },
    age: { type: "number" },
  }

  async function withSchema(rows: unknown[][]): Promise<{ file: string; schema: string }> {
    const file = await makeXlsx("data.xlsx", rows)
    const schemaPath = path("schema.json")
    writeFileSync(schemaPath, JSON.stringify(schema), "utf-8")
    return { file, schema: schemaPath }
  }

  it("passes a conforming sheet", async () => {
    const paths = await withSchema([
      ["name", "age"],
      ["Ada", 36],
    ])
    await run(validateCommand, paths)
    expect(logs.join("\n")).toContain("1 row(s) passed")
  })

  it("fails, lists the errors, and throws", async () => {
    const paths = await withSchema([
      ["name", "age"],
      [null, "not a number"],
    ])
    await expect(run(validateCommand, paths)).rejects.toThrow(/Validation failed/)
    expect(logs.join("\n")).toMatch(/Row \d+, Column/)
  })

  it("caps the listed errors at twenty", async () => {
    const rows: unknown[][] = [["name", "age"]]
    for (let i = 0; i < 30; i++) rows.push([null, 1])
    const paths = await withSchema(rows)

    await expect(run(validateCommand, paths)).rejects.toThrow(CliError)
    expect(logs.join("\n")).toContain("and 10 more errors")
  })

  it("rejects a schema file that is not JSON", async () => {
    const file = await makeXlsx("data.xlsx", [["name"], ["Ada"]])
    const schemaPath = path("broken.json")
    writeFileSync(schemaPath, "{ not json", "utf-8")
    await expect(run(validateCommand, { file, schema: schemaPath })).rejects.toThrow(
      /Invalid JSON schema/,
    )
  })

  it("rejects a sheet index the file does not have", async () => {
    const paths = await withSchema([["name"], ["Ada"]])
    await expect(run(validateCommand, { ...paths, sheet: "3" })).rejects.toThrow(
      /File has 1 sheet\(s\)/,
    )
  })

  it("rejects a sheet index that is not a number", async () => {
    const paths = await withSchema([["name"], ["Ada"]])
    await expect(run(validateCommand, { ...paths, sheet: "x" })).rejects.toThrow(
      /Invalid sheet index/,
    )
  })

  it("defaults to the first sheet", async () => {
    const paths = await withSchema([["name"], ["Ada"]])
    await run(validateCommand, paths)
    expect(logs.join("\n")).toContain("passed validation")
  })
})

// ── Wiring ──────────────────────────────────────────────────────────

describe("the command tree", () => {
  it("exposes all three subcommands", () => {
    expect(Object.keys(mainCommand.subCommands as object).sort()).toEqual([
      "convert",
      "inspect",
      "validate",
    ])
  })

  it("reports the package version, not the 0.0.0 fallback", () => {
    // The fallback exists so the CLI still runs if package.json cannot be
    // resolved; reaching it in a normal install means the bundle's
    // relative path to package.json is wrong, which is #357's failure.
    expect(pkgVersion).toMatch(/^\d+\.\d+\.\d+/)
    expect(pkgVersion).not.toBe("0.0.0")
  })
})

describe("the node-builtin boundary", () => {
  it("is crossed by the CLI and nothing else in src/", async () => {
    // The library is platform-neutral: Web APIs only, so it runs the same
    // in a browser, a worker and Deno. tsconfig.json enforces that by
    // declaring no ambient types, which makes `process` and `Buffer` type
    // errors — but an `import ... from "node:fs"` would still resolve at
    // run time in Node, so check the imports directly too.
    const { readdirSync, readFileSync: read } = await import("node:fs")
    const { join: j } = await import("node:path")

    const offenders: string[] = []
    const walk = (rel: string): void => {
      for (const entry of readdirSync(j("src", rel), { withFileTypes: true })) {
        const child = rel ? `${rel}/${entry.name}` : entry.name
        if (entry.isDirectory()) {
          walk(child)
        } else if (entry.name.endsWith(".ts")) {
          if (/from\s+["']node:/.test(read(j("src", child), "utf-8"))) offenders.push(child)
        }
      }
    }
    walk("")

    expect(offenders.sort()).toEqual(["cli/commands.ts"])
  })
})

describe("formatCellValue", () => {
  it("renders each cell type", () => {
    expect(formatCellValue(null)).toBe("")
    expect(formatCellValue(undefined as never)).toBe("")
    expect(formatCellValue(new Date(0))).toBe("1970-01-01T00:00:00.000Z")
    expect(formatCellValue(42)).toBe("42")
    expect(formatCellValue(true)).toBe("true")
    expect(formatCellValue("x")).toBe("x")
  })
})
