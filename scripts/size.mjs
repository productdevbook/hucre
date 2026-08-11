#!/usr/bin/env node
// ── Bundle-size budget ──────────────────────────────────────────────
//
// The README compares hucre against a competitor's ~300 KB without ever
// publishing hucre's own number, and nothing anywhere pinned it — so a
// regression that doubled what a caller pulls in would ship silently.
// See #474.
//
// What this measures is what a caller actually gets: a synthetic entry
// importing one API from one entry point, bundled and minified, so the
// tree shaking counts. That is a different — and much smaller — number
// than the size of `dist/`, which holds every format.
//
// Run: node scripts/size.mjs          check against the budgets below
//      node scripts/size.mjs --update rewrite the budgets from measurement

import { gzipSync } from "node:zlib"
import { mkdtempSync, rmSync, writeFileSync, readFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { fileURLToPath } from "node:url"
import { rolldown } from "rolldown"

const repoRoot = fileURLToPath(new URL("..", import.meta.url))
const budgetPath = join(repoRoot, "scripts", "size-budget.json")

/**
 * One entry per thing worth pinning. `from` is the published subpath so
 * the numbers describe what a caller imports, not what `src/` contains.
 */
const SCENARIOS = [
  // The four the README's bundle-size footnote publishes. Pinning
  // exactly these is the point: the footnote had already drifted — it
  // claimed 114 KB for the whole library, which measures 127.
  { name: "csv (hucre/csv)", from: "hucre/csv", named: ["parseCsv", "writeCsv"] },
  { name: "readXlsx (hucre/xlsx)", from: "hucre/xlsx", named: ["readXlsx"] },
  { name: "read+write xlsx (hucre/xlsx)", from: "hucre/xlsx", named: ["readXlsx", "writeXlsx"] },
  { name: "everything (hucre)", from: "hucre", named: null },
  // Worth watching besides.
  { name: "writeXlsx (hucre)", from: "hucre", named: ["writeXlsx"] },
  { name: "readOds (hucre/ods)", from: "hucre/ods", named: ["readOds"] },
]

/** Map a published subpath to the built file it resolves to. */
function resolveEntry(subpath) {
  const pkg = JSON.parse(readFileSync(join(repoRoot, "package.json"), "utf8"))
  const key = subpath === "hucre" ? "." : `./${subpath.slice("hucre/".length)}`
  const target = pkg.exports[key]?.default
  if (!target) throw new Error(`package.json#exports has no entry for ${subpath}`)
  return join(repoRoot, target)
}

async function measure(scenario, workDir) {
  const entryFile = resolveEntry(scenario.from)
  const source = scenario.named
    ? `import { ${scenario.named.join(", ")} } from ${JSON.stringify(entryFile)}\n` +
      `console.log(${scenario.named.join(", ")})\n`
    : `import * as all from ${JSON.stringify(entryFile)}\nconsole.log(all)\n`

  const input = join(workDir, `${scenario.name.replace(/\W+/g, "-")}.mjs`)
  writeFileSync(input, source)

  const bundle = await rolldown({ input, logLevel: "silent" })
  const { output } = await bundle.write({
    dir: join(workDir, "out"),
    format: "esm",
    minify: true,
  })
  await bundle.close()

  const code = output
    .filter((c) => c.type === "chunk")
    .map((c) => c.code)
    .join("")
  const bytes = Buffer.from(code, "utf8")
  return { min: bytes.length, gzip: gzipSync(bytes, { level: 9 }).length }
}

const update = process.argv.includes("--update")
const workDir = mkdtempSync(join(tmpdir(), "hucre-size-"))
let failures = 0

try {
  const budgets = update ? {} : JSON.parse(readFileSync(budgetPath, "utf8"))
  const measured = {}

  console.log("\n  scenario                        minified      gzipped")
  console.log("  " + "-".repeat(58))

  for (const scenario of SCENARIOS) {
    const { min, gzip } = await measure(scenario, workDir)
    measured[scenario.name] = { min, gzip }

    const budget = budgets[scenario.name]
    const over = budget && (min > budget.min || gzip > budget.gzip)
    if (over) failures++

    const mark = update ? " " : over ? "!" : " "
    console.log(
      `${mark} ${scenario.name.padEnd(30)} ${kb(min).padStart(9)} ${kb(gzip).padStart(12)}` +
        (over ? `   over ${kb(budget.min)} / ${kb(budget.gzip)}` : ""),
    )
  }

  if (update) {
    writeFileSync(budgetPath, `${JSON.stringify(withHeadroom(measured), null, 2)}\n`)
    console.log(`\n  budgets written to scripts/size-budget.json (+5% headroom)\n`)
  } else if (failures > 0) {
    console.log(
      `\n  ${failures} scenario(s) over budget. If the growth is intended, ` +
        `run \`node scripts/size.mjs --update\` and say why in the commit.\n`,
    )
    process.exitCode = 1
  } else {
    console.log("\n  all within budget\n")
  }
} finally {
  rmSync(workDir, { recursive: true, force: true })
}

/**
 * Budgets carry 5% headroom over the measurement that set them.
 *
 * A budget pinned to the exact byte fails on a comment, which teaches
 * people to run `--update` reflexively and stops the check meaning
 * anything. 5% is loose enough to absorb ordinary edits and tight enough
 * that pulling in a new subsystem is still caught.
 */
function withHeadroom(measured) {
  const out = {}
  for (const [name, { min, gzip }] of Object.entries(measured)) {
    out[name] = { min: Math.ceil(min * 1.05), gzip: Math.ceil(gzip * 1.05) }
  }
  return out
}

function kb(bytes) {
  return `${(bytes / 1024).toFixed(1)} KB`
}
