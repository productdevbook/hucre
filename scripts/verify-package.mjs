#!/usr/bin/env node
// ── Packaged-artifact verification ──────────────────────────────────
// Packs the library exactly as npm would publish it, installs the
// tarball into a throwaway project, and exercises it from there.
//
// This exists because `pnpm test` only ever sees `src/`. The published
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
  return execFileSync(command, args, {
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
        `import { readXlsx, writeXlsx } from "hucre"`,
        `import { writeXlsx as x } from "hucre/xlsx"`,
        `import { parseCsv } from "hucre/csv"`,
        `import { readOds } from "hucre/ods"`,
        `import { parseJson } from "hucre/json"`,
        `import { parseChart } from "hucre/ooxml"`,
        `if (!readXlsx || !writeXlsx || !x || !parseCsv || !readOds || !parseJson || !parseChart) {`,
        `  throw new Error("an entry point resolved to undefined")`,
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
