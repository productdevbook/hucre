#!/usr/bin/env node
// ── Packaged-artifact verification ──────────────────────────────────
// Packs the library exactly as npm would publish it, installs the
// tarball into a throwaway project, and exercises it from there.
//
// This exists because `bun run test` only ever sees `src/`. The published
// CLI was dead across two releases — `dist/cli.mjs` imported `citty` and
// `consola` while `package.json` declared no runtime dependencies — and
// nothing caught it, because nothing ran the packaged binary (#357).
//
// Run: node scripts/verify-package.mjs

import { execFileSync } from "node:child_process"
import { existsSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { fileURLToPath } from "node:url"

const repoRoot = fileURLToPath(new URL("..", import.meta.url))

let failures = 0
let workDir

function check(label, fn) {
  try {
    fn()
    console.log(`  ok    ${label}`)
  } catch (error) {
    failures++
    console.log(`  FAIL  ${label}`)
    console.log(`        ${error.message.split("\n")[0]}`)
  }
}

function run(command, args, options = {}) {
  // The harness may be launched with an explicit Node 24 binary while
  // PATH points at another version; probes must use the requested engine.
  return execFileSync(command === "node" ? process.execPath : command, args, {
    encoding: "utf8",
    stdio: ["ignore", "pipe", "pipe"],
    ...options,
  })
}

function assert(condition, message) {
  if (!condition) throw new Error(message)
}

// ── Static check: no bare specifiers survive into the CLI ───────────
//
// Cheap and precise — this is the exact shape of the #357 regression,
// and it fails before the slower install-based checks below.

check("dist/cli.mjs imports nothing outside node: and its own dist", () => {
  const source = readFileSync(join(repoRoot, "dist/cli.mjs"), "utf8")
  const specifiers = [
    ...source.matchAll(/(?:^|[\s;])(?:import|export)[^'"]*?from\s*["']([^"']+)["']/g),
  ]
    .map((match) => match[1])
    .concat([...source.matchAll(/\bimport\s*\(\s*["']([^"']+)["']\s*\)/g)].map((m) => m[1]))

  const external = specifiers.filter(
    (specifier) => !specifier.startsWith(".") && !specifier.startsWith("node:"),
  )

  assert(
    external.length === 0,
    `bare imports would not resolve for an installed user: ${[...new Set(external)].join(", ")}`,
  )
})

// ── Static check: every shipped module keeps its types ─────────────
//
// A declaration oxc cannot infer under --isolatedDeclarations only warns
// during the build: the .mjs is written, the .d.mts is not, and the
// package ships an entry whose types resolve to nothing. Only the bundled
// CLI (built with `dts: false`) and its inlined chunks are exempt.

check("every shipped module has a matching .d.mts", () => {
  const distDir = join(repoRoot, "dist")
  const orphans = readdirSync(distDir, { recursive: true })
    .map((entry) => String(entry).replaceAll("\\", "/"))
    .filter(
      (entry) => entry.endsWith(".mjs") && entry !== "cli.mjs" && !entry.startsWith("_chunks/"),
    )
    .filter((entry) => !existsSync(join(distDir, `${entry.slice(0, -4)}.d.mts`)))

  assert(orphans.length === 0, `declaration files missing for: ${orphans.join(", ")}`)
})

check("package.json declares no runtime dependencies", () => {
  const pkg = JSON.parse(readFileSync(join(repoRoot, "package.json"), "utf8"))
  const deps = Object.keys(pkg.dependencies ?? {})
  assert(deps.length === 0, `unexpected runtime dependencies: ${deps.join(", ")}`)
})

// ── Install the tarball and use it like a consumer ─────────────────

try {
  workDir = mkdtempSync(join(tmpdir(), "hucre-pkg-"))

  const packOutput = run("npm", ["pack", "--silent", "--pack-destination", workDir], {
    cwd: repoRoot,
  })
  const tarball = packOutput.trim().split("\n").pop()

  writeFileSync(join(workDir, "package.json"), JSON.stringify({ name: "consumer", private: true }))
  run("npm", ["install", "--silent", "--no-audit", "--no-fund", join(workDir, tarball)], {
    cwd: workDir,
  })

  const cli = join(workDir, "node_modules/.bin/hucre")

  check("installing the tarball pulls in no transitive packages", () => {
    const installed = run("ls", [join(workDir, "node_modules")])
      .split("\n")
      .filter((name) => name && !name.startsWith("."))
    assert(
      installed.length === 1 && installed[0] === "hucre",
      `expected only hucre, got: ${installed.join(", ")}`,
    )
  })

  check("hucre --help runs", () => {
    const output = run(cli, ["--help"])
    assert(output.includes("convert"), "help output does not list the convert command")
  })

  check("hucre convert produces a readable workbook", () => {
    writeFileSync(join(workDir, "in.csv"), "name,qty\nfoo,1\nbar,2\n")
    run(cli, ["convert", join(workDir, "in.csv"), join(workDir, "out.xlsx")])
    const output = run(cli, ["inspect", join(workDir, "out.xlsx")])
    assert(output.includes("3 rows"), `inspect did not report 3 rows:\n${output}`)
  })

  check("the library entry points import cleanly from an install", () => {
    const probe = join(workDir, "probe.mjs")
    writeFileSync(
      probe,
      [
        `import { read, ParseError, readXlsx, writeXlsx, insertRows, replaceCells, readObjects, sheetToObjects, toJson, write, writeXlsxStream, XlsxStreamWriter, writeOds, readOds as readOdsRoot, writeOdsStream, OdsStreamWriter, link } from "hucre"`,
        `import { writeXlsx as x } from "hucre/xlsx"`,
        `import { parseCsv } from "hucre/csv"`,
        `import { readOds } from "hucre/ods"`,
        `import { parseJson } from "hucre/json"`,
        `import { parseChart } from "hucre/ooxml"`,
        `import { readXml } from "hucre/xml"`,
        `import { createCellStore, getCell, setCell, hasCell, deleteCell, cellEntries } from "hucre/cell"`,
        `import { formatValue } from "hucre/format"`,
        `import { audit } from "hucre/a11y"`,
        `if (!readXlsx || !writeXlsx || !x || !parseCsv || !readOds || !parseJson || !parseChart || !readXml || !formatValue || !audit) {`,
        `  throw new Error("an entry point resolved to undefined")`,
        `}`,
        `const cells = createCellStore([[128, 3, { value: 42, type: "number" }]])`,
        `setCell(cells, 129, 4, { value: 7, type: "number" })`,
        `const copy = structuredClone(cells)`,
        `deleteCell(copy, 128, 3)`,
        `if (cells.size !== 2 || copy.size !== 1 || !hasCell(copy, 129, 4) || getCell(copy, 129, 4)?.value !== 7 || [...cellEntries(copy)].length !== 1) {`,
        `  throw new Error("installed cell helpers lost metadata")`,
        `}`,
        `const wb = await readXlsx(await writeXlsx({ sheets: [{ name: "Metadata", rows: [[1]], cells }] }), { sparse: true })`,
        `if (getCell(wb.sheets[0].cells, 128, 3)?.value !== 42) throw new Error("installed metadata round trip failed")`,
        `const sheet = { name: "Edit", rows: [[1], [2]], cells: createCellStore([[1, 0, { value: 2, type: "formula", formula: "Edit!A2" }]]) }`,
        `insertRows(sheet, 0, 1)`,
        `if (sheet.rows[2][0] !== 2 || getCell(sheet.cells, 2, 0)?.formula !== "Edit!A3") throw new Error("installed editing lost references")`,
        `if (replaceCells(sheet, 2, false) !== 1 || getCell(sheet.cells, 2, 0)?.formulaResult !== false) throw new Error("installed value edit lost its cache")`,
        `const cleared = await readXlsx(await writeXlsx({ sheets: [{ name: "Null", rows: [[42]], cells: createCellStore([[0, 0, { value: null, style: { font: { bold: true } } }]]) }] }), { readStyles: true })`,
        `if (cleared.sheets[0].rows[0][0] !== null) throw new Error("installed writer discarded a null override")`,
        `const table = await writeXlsx({ sheets: [{ name: "S", rows: [["Name"], ["Ada"]] }] })`,
        `const sparse = (await readXlsx(table, { sparse: true })).sheets[0]`,
        `if ((await readObjects(table, { sparse: true })).data[0]?.Name !== "Ada" || sheetToObjects(sparse).data[0]?.Name !== "Ada" || JSON.parse(toJson(sparse))[0]?.Name !== "Ada") throw new Error("installed sparse projection lost values")`,
        `if (!(new TextDecoder().decode(await write({ sheets: [sparse] }, { format: "csv" }))).includes("Ada")) throw new Error("installed text write lost sparse values")`,
        `let bounded = false`,
        `const authored = { sheets: [{ name: "Data", data: [{ name: "Ada" }, { score: 42 }] }] }`,
        `if (new TextDecoder().decode(await write(authored, { format: "csv" })) !== "name,score\\r\\nAda,\\r\\n,42") throw new Error("installed object-data authoring lost fields")`,
        `if ((await readXlsx(await x(authored))).sheets[0].rows[2][1] !== 42) throw new Error("installed XLSX data normalization failed")`,
        `try { await read(new TextEncoder().encode("a,b,c\\nx\\nx"), { maxTotalCells: 8 }) } catch (error) { bounded = error instanceof ParseError }`,
        `if (!bounded) throw new Error("installed read discarded maxTotalCells")`,
        `bounded = false`,
        `try { parseJson('[{"a":1,"b":2},{"a":3}]', { maxTotalCells: 3 }) } catch (error) { bounded = error instanceof ParseError }`,
        `if (!bounded) throw new Error("installed JSON expansion ignored its bound")`,
        `const inline = [[{ value: 1, formula: "6*7", formulaResult: 42 }, link("Open", "https://example.com/a"), { value: true, checkbox: true, comment: { text: "Note" } }]]`,
        `const inc = new XlsxStreamWriter({ name: "S" }); inc.addRow(inline[0])`,
        `for (const out of [await inc.finish(), new Uint8Array(await new Response(writeXlsxStream(inline, { name: "S" })).arrayBuffer())]) {`,
        `  const s = (await readXlsx(out)).sheets[0]`,
        `  if (s.rows[0][0] !== 42 || getCell(s.cells, 0, 1)?.hyperlink?.target !== "https://example.com/a" || getCell(s.cells, 0, 2)?.comment?.text !== "Note" || !getCell(s.cells, 0, 2)?.checkbox) throw new Error("installed XLSX stream discarded inline metadata")`,
        `}`,
        `const oi = new OdsStreamWriter({ name: "S" }); oi.addRow(inline[0])`,
        `for (const out of [await writeOds({ sheets: [{ name: "S", rows: inline }] }), await oi.finish(), new Uint8Array(await new Response(writeOdsStream(inline, { name: "S" })).arrayBuffer())]) {`,
        `  if ((await readOdsRoot(out)).sheets[0].rows[0][0] !== 42) throw new Error("installed ODS writer discarded the explicit cache")`,
        `}`,
        `console.log("entry points ok")`,
      ].join("\n"),
    )
    const output = run("node", [probe], { cwd: workDir })
    assert(output.includes("entry points ok"), output)
  })
} catch (error) {
  failures++
  console.log(`  FAIL  packaging harness`)
  console.log(`        ${error.message.split("\n")[0]}`)
} finally {
  if (workDir) rmSync(workDir, { recursive: true, force: true })
}

if (failures > 0) {
  console.error(`\n${failures} packaged-artifact check(s) failed.`)
  process.exit(1)
}
console.log("\nPackaged artifact verified.")
