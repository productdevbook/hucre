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
  detectOutputFormat,
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
import { writeCfb } from "../src/xlsx/crypto/cfb"
import { ZipWriter } from "../src/zip/writer"

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

// ── Legacy read-only fixtures (#411) ────────────────────────────────
// There is no .xls or .xlsb writer to build these with, so both are
// assembled by hand — the smallest file each reader accepts — to exercise
// `convert legacy.xls out.xlsx`, which used to fail on the extension
// alone.

const u16 = (n: number): number[] => [n & 0xff, (n >> 8) & 0xff]
const u32 = (n: number): number[] => [
  n & 0xff,
  (n >> 8) & 0xff,
  (n >> 16) & 0xff,
  (n >>> 24) & 0xff,
]
const chars = (s: string): number[] => [...s].map((c) => c.charCodeAt(0))

/** A one-sheet BIFF8 .xls holding a single LABEL cell. */
function makeXls(name: string, text: string): string {
  const biff = (sid: number, data: number[]): number[] => [
    ...u16(sid),
    ...u16(data.length),
    ...data,
  ]
  const bof = (dt: number): number[] =>
    biff(0x0809, [...u16(0x0600), ...u16(dt), ...u16(0), ...u16(0), ...u32(0), ...u32(0)])
  const eof = (): number[] => biff(0x000a, [])
  const sheet = [
    ...bof(0x0010),
    ...biff(0x0204, [...u16(0), ...u16(0), ...u16(0), ...u16(text.length), 0, ...chars(text)]),
    ...eof(),
  ]
  // BOUNDSHEET carries the sheet substream's byte offset, so the globals
  // are laid out twice: once to measure, once for real.
  const globals = (sheetPos: number): number[] => [
    ...bof(0x0005),
    ...biff(0x0085, [...u32(sheetPos), 0, 0, "Legacy".length, 0, ...chars("Legacy")]),
    ...eof(),
  ]
  const stream = new Uint8Array([...globals(globals(0).length), ...sheet])
  const p = path(name)
  writeFileSync(p, writeCfb([{ name: "Workbook", data: stream }]))
  return p
}

/** A one-sheet .xlsb holding a single inline-string cell. */
async function makeXlsb(name: string, text: string): Promise<string> {
  const varint = (n: number): number[] => {
    const out: number[] = []
    let s = n
    do {
      let b = s & 0x7f
      s >>>= 7
      if (s) b |= 0x80
      out.push(b)
    } while (s)
    return out
  }
  const rec = (id: number, body: number[]): number[] => [
    ...(id < 0x80 ? [id] : [(id & 0x7f) | 0x80, (id >> 7) & 0x7f]),
    ...varint(body.length),
    ...body,
  ]
  /** XLWideString: u32 char count + UTF-16LE units. */
  const wstr = (s: string): number[] => [
    ...u32(s.length),
    ...[...s].flatMap((c) => u16(c.charCodeAt(0))),
  ]

  const ws = [
    ...rec(0, u32(0)), // BrtRowHdr
    ...rec(6, [...u32(0), ...u32(0), ...wstr(text)]), // BrtCellSt
  ]
  const wb = rec(156, [...u32(0), ...u32(0), ...wstr("rId1"), ...wstr("Legacy")]) // BrtBundleSh
  const ns = "http://schemas.openxmlformats.org/package/2006/relationships"
  const rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  const enc = new TextEncoder()
  const zw = new ZipWriter()
  zw.add(
    "_rels/.rels",
    enc.encode(
      `<Relationships xmlns="${ns}"><Relationship Id="r" Type="${rel}/officeDocument" Target="xl/workbook.bin"/></Relationships>`,
    ),
  )
  zw.add("xl/workbook.bin", new Uint8Array(wb))
  zw.add(
    "xl/_rels/workbook.bin.rels",
    enc.encode(
      `<Relationships xmlns="${ns}"><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.bin"/></Relationships>`,
    ),
  )
  zw.add("xl/worksheets/sheet1.bin", new Uint8Array(ws))

  const p = path(name)
  writeFileSync(p, await zw.build())
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

  it("maps the read-only legacy extensions too", () => {
    expect(detectFormatFromExtension("a.xls")).toBe("xls")
    expect(detectFormatFromExtension("a.XLSB")).toBe("xlsb")
  })

  it("rejects anything else, naming what is supported", () => {
    // The supported list now distinguishes the two directions, so the
    // message says which formats are input-only rather than leaving a
    // user to infer it from a failure two commands later.
    expect(() => detectFormatFromExtension("a.numbers")).toThrow(CliError)
    expect(() => detectFormatFromExtension("a.numbers")).toThrow(/\.xlsx, \.ods, \.csv, \.tsv/)
    expect(() => detectFormatFromExtension("a.numbers")).toThrow(/read-only: \.xls, \.xlsb/)
  })

  it("says so when there is no extension at all", () => {
    expect(() => detectFormatFromExtension("README")).toThrow(/\(none\)/)
  })
})

describe("detectOutputFormat", () => {
  it("accepts every writable format", () => {
    expect(detectOutputFormat("a.xlsx")).toBe("xlsx")
    expect(detectOutputFormat("a.ods")).toBe("ods")
    expect(detectOutputFormat("a.tsv")).toBe("csv")
  })

  for (const ext of [".xls", ".xlsb"]) {
    it(`refuses to write ${ext}, and says it is read-only rather than unsupported`, () => {
      expect(() => detectOutputFormat(`out${ext}`)).toThrow(CliError)
      expect(() => detectOutputFormat(`out${ext}`)).toThrow(/read-only format/)
      expect(() => detectOutputFormat(`out${ext}`)).not.toThrow(/Unsupported file extension/)
    })
  }
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

  it("writes a UTF-8 BOM when asked, and not otherwise", async () => {
    // Excel on a non-UTF-8 locale reads a UTF-8 CSV as the system code
    // page unless the file opens with EF BB BF, so `convert x.xlsx
    // out.csv` produced mojibake and there was no way to ask for better.
    // See #475 — which answered the read side and documented `bom: true`
    // as the write side's remedy, while the CLI could not pass it.
    const input = await makeXlsx("in.xlsx", [["Şehir", "Ürün"]])

    await run(convertCommand, { input, output: path("plain.csv") })
    const plain = readFileSync(path("plain.csv"))
    expect(plain[0]).not.toBe(0xef)

    await run(convertCommand, { input, output: path("bom.csv"), bom: true })
    const withBom = readFileSync(path("bom.csv"))
    expect([withBom[0], withBom[1], withBom[2]]).toEqual([0xef, 0xbb, 0xbf])
    // The mark is a prefix, not a replacement — the rows are still there.
    expect(new TextDecoder().decode(withBom)).toContain("Şehir")
  })

  it("writes the BOM on .tsv too", async () => {
    const input = await makeXlsx("in.xlsx", [["a", "b"]])
    await run(convertCommand, { input, output: path("bom.tsv"), bom: true })
    const bytes = readFileSync(path("bom.tsv"))
    expect([bytes[0], bytes[1], bytes[2]]).toEqual([0xef, 0xbb, 0xbf])
    expect(new TextDecoder().decode(bytes)).toContain("a\tb")
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

  // ── Legacy input formats (#411) ───────────────────────────────────
  // The readers shipped, but the CLI's extension switch did not know
  // about them, so the most obvious use of a read-only format —
  // rescuing an archive into .xlsx — failed before a byte was read.

  it("converts a legacy .xls to xlsx", async () => {
    const input = makeXls("legacy.xls", "from-xls")
    await run(convertCommand, { input, output: path("out.xlsx") })
    const back = await readXlsx(readFileSync(path("out.xlsx")))
    expect(back.sheets[0].name).toBe("Legacy")
    expect(back.sheets[0].rows[0]).toEqual(["from-xls"])
  })

  it("converts an .xlsb to csv", async () => {
    const input = await makeXlsb("legacy.xlsb", "from-xlsb")
    await run(convertCommand, { input, output: path("out.csv") })
    expect(readFileSync(path("out.csv"), "utf-8")).toContain("from-xlsb")
  })

  it("refuses .xls and .xlsb as an output target", async () => {
    const input = await makeXlsx("in.xlsx", [["a"]])
    for (const out of ["out.xls", "out.xlsb"]) {
      await expect(run(convertCommand, { input, output: path(out) })).rejects.toThrow(
        /read-only format/,
      )
    }
  })

  it("says in --help that it carries values only", () => {
    expect(convertCommand.meta).toMatchObject({
      description: expect.stringContaining("Convert between spreadsheet formats"),
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

  it("inspects a legacy .xls", async () => {
    await run(inspectCommand, { file: makeXls("legacy.xls", "cell"), sheet: "0" })
    const out = logs.join("\n")
    expect(out).toContain('"Legacy"')
    expect(out).toContain("cell")
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

// ═══════════════════════════════════════════════════════════════════════
// #469 — JSON, NDJSON, XML, HTML and Markdown all had readers and/or
// writers in the library and none was reachable from the terminal. Nor
// was stdin or stdout, which is most of what a CLI is for.
// ═══════════════════════════════════════════════════════════════════════

describe("convert reaches the formats the library has", () => {
  async function csvFile(name = "in.csv"): Promise<string> {
    const p = path(name)
    writeFileSync(p, "name,qty\nWidget,3\nGadget,7\n", "utf-8")
    return p
  }

  it("writes JSON", async () => {
    const out = path("out.json")
    await run(convertCommand, { input: await csvFile(), output: out })

    expect(JSON.parse(readFileSync(out, "utf-8"))).toEqual([
      { name: "Widget", qty: "3" },
      { name: "Gadget", qty: "7" },
    ])
  })

  it("writes NDJSON, one object per line", async () => {
    const out = path("out.ndjson")
    await run(convertCommand, { input: await csvFile(), output: out })

    const lines = readFileSync(out, "utf-8").trim().split("\n")
    expect(lines).toHaveLength(2)
    expect(JSON.parse(lines[0]!)).toEqual({ name: "Widget", qty: "3" })
  })

  it("takes .jsonl as a name for the same thing", () => {
    expect(detectFormatFromExtension("a.jsonl")).toBe("ndjson")
    expect(detectFormatFromExtension("a.ndjson")).toBe("ndjson")
  })

  it("writes XML, HTML and Markdown", async () => {
    for (const [ext, needle] of [
      ["xml", "<name>"],
      ["html", "<table"],
      ["md", "| name"],
    ] as const) {
      const out = path(`out.${ext}`)
      await run(convertCommand, { input: await csvFile(), output: out })

      expect(readFileSync(out, "utf-8"), ext).toContain(needle)
    }
  })

  it("reads JSON back into a spreadsheet", async () => {
    const src = path("in.json")
    writeFileSync(src, JSON.stringify([{ name: "Widget", qty: 3 }]), "utf-8")

    const out = path("out.xlsx")
    await run(convertCommand, { input: src, output: out })

    const wb = await readXlsx(new Uint8Array(readFileSync(out)))
    expect(wb.sheets[0]!.rows).toEqual([
      ["name", "qty"],
      ["Widget", 3],
    ])
  })

  it("refuses to read Markdown, because there is no reader", async () => {
    // Output only, and saying so is better than a confusing parse error.
    const src = path("in.md")
    writeFileSync(src, "| a |\n| - |\n| 1 |\n", "utf-8")

    await expect(run(convertCommand, { input: src, output: path("o.csv") })).rejects.toThrow(
      /output only/,
    )
  })
})

describe("convert names the formats it cannot write", () => {
  it("lists the writable ones when asked for a read-only format", () => {
    expect(() => detectOutputFormat("a.xls")).toThrow(/read-only/)
    expect(() => detectOutputFormat("a.xls")).toThrow(/\.json/)
  })

  it("still rejects an extension that is nothing", () => {
    expect(() => detectFormatFromExtension("a.pdf")).toThrow(CliError)
  })
})

describe("the stdin/stdout convention", () => {
  it("asks for --to when writing to a pipe, rather than guessing", async () => {
    // There is no extension to read, and defaulting to a binary format
    // would spray a ZIP into a terminal.
    const src = path("in.csv")
    writeFileSync(src, "a,b\n1,2\n", "utf-8")

    await expect(run(convertCommand, { input: src, output: "-" })).rejects.toThrow(/--to/)
  })

  it("resolves --to through the same table as an extension", async () => {
    const src = path("in.csv")
    writeFileSync(src, "a,b\n1,2\n", "utf-8")

    // A format that does not exist fails the same way a bad extension does.
    await expect(run(convertCommand, { input: src, output: "-", to: "pdf" })).rejects.toThrow(
      CliError,
    )

    // And a read-only one is named as read-only, not as unknown.
    await expect(run(convertCommand, { input: src, output: "-", to: "xls" })).rejects.toThrow(
      /read-only/,
    )
  })
})

describe("validate --header-row", () => {
  it("names the header row, and -1 says there is none", async () => {
    const dir = mkdtempSync(join(tmpdir(), "hucre-cli-header-"))
    try {
      const file = join(dir, "titled.xlsx")
      writeFileSync(
        file,
        await writeXlsx({
          sheets: [{ name: "S", rows: [["Report title"], ["qty"], [1], [2]] }],
        }),
      )
      const schemaPath = join(dir, "schema.json")
      writeFileSync(
        schemaPath,
        JSON.stringify({ qty: { column: "qty", type: "number", required: true } }),
      )

      // Row 0 is a title, so the default header row finds no "qty" column.
      await expect(run(validateCommand, { file, schema: schemaPath })).rejects.toThrow(CliError)
      // Naming row 1 as the header makes the file valid.
      await run(validateCommand, { file, schema: schemaPath, headerRow: "1" })
      await expect(
        run(validateCommand, { file, schema: schemaPath, headerRow: "x" }),
      ).rejects.toThrow(/Invalid header row/)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })
})
