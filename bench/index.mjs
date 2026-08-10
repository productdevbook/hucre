// ── pnpm bench ───────────────────────────────────────────────────────
//
// Runs every scenario, one child process each, so `maxRSS` measures that
// scenario and nothing else. Takes a couple of minutes.
//
//   pnpm bench            # 100k rows
//   pnpm bench 300000     # or however many

import { spawnSync } from "node:child_process"
import { existsSync } from "node:fs"
import { fileURLToPath } from "node:url"

const rows = process.argv[2] ?? "100000"
const here = (name) => fileURLToPath(new URL(name, import.meta.url))

if (!existsSync(fileURLToPath(new URL("../dist/index.mjs", import.meta.url)))) {
  console.error("dist/ not found — run `pnpm build` first.")
  process.exit(1)
}

const runs = [
  ["write.mjs", "writeXlsx", rows],
  ["write.mjs", "writeXlsxStream", rows],
  ["write.mjs", "XlsxStreamWriter", rows],
  ["read.mjs", "readXlsx", "high-cardinality", rows],
  ["read.mjs", "readXlsxMaxRows", "high-cardinality", rows],
  ["read.mjs", "streamXlsxRows", "high-cardinality", rows],
  ["read.mjs", "readXlsx", "low-cardinality", rows],
  ["read.mjs", "readXlsxMaxRows", "low-cardinality", rows],
  ["read.mjs", "streamXlsxRows", "low-cardinality", rows],
]

let failed = 0
for (const [script, ...args] of runs) {
  const result = spawnSync(process.execPath, [here(script), ...args], { stdio: "inherit" })
  if (result.status !== 0) failed++
}

if (failed > 0) {
  console.error(`\n${failed} scenario(s) failed`)
  process.exit(1)
}
