// ── XLSX read paths ──────────────────────────────────────────────────
//
//   node bench/read.mjs [scenario] [fixture] [rows]
//
// The fixture is written by a child process and read here, so the read's
// peak RSS is not the write's. The two fixtures differ in one line — the
// number of *distinct* strings — which is what the streaming reader's
// memory actually tracks. See bench/README.md.

import { mkdirSync, existsSync, writeFileSync, readFileSync } from "node:fs"
import { fileURLToPath } from "node:url"
import { readXlsx, streamXlsxRows, writeXlsx } from "../dist/index.mjs"

const SCENARIOS = ["readXlsx", "readXlsxStyles", "readXlsxMaxRows", "streamXlsxRows"]
const FIXTURES = ["high-cardinality", "low-cardinality"]

const scenario = process.argv[2] ?? "readXlsx"
const fixture = process.argv[3] ?? "high-cardinality"
const rowCount = Number(process.argv[4] ?? 100000)
const COLS = 12
const WORDS = [
  "alpha",
  "beta",
  "gamma",
  "delta",
  "epsilon",
  "zeta",
  "eta",
  "theta",
  "iota",
  "kappa",
]

const peakMb = () => Math.round(process.resourceUsage().maxRSS / 1024)
const nowMs = () => Number(process.hrtime.bigint() / 1000000n)

/**
 * The one line that matters. High cardinality gives every text cell its
 * own string, so `xl/sharedStrings.xml` holds ~400k entries; low
 * cardinality reuses ten.
 */
function makeRow(i, distinct) {
  const row = []
  for (let c = 0; c < COLS; c++) {
    row.push(
      c % 3 === 0
        ? distinct
          ? `text ${i}-${c}`
          : WORDS[(i + c) % WORDS.length]
        : c % 3 === 1
          ? i * c
          : new Date(Date.UTC(2024, 0, 1 + (i % 28))),
    )
  }
  return row
}

async function fixturePath() {
  const dir = fileURLToPath(new URL("./.fixtures/", import.meta.url))
  mkdirSync(dir, { recursive: true })
  const path = `${dir}${fixture}-${rowCount}.xlsx`
  if (!existsSync(path)) {
    const distinct = fixture === "high-cardinality"
    const rows = Array.from({ length: rowCount }, (_, i) => makeRow(i, distinct))
    writeFileSync(path, await writeXlsx({ sheets: [{ name: "S", rows }] }))
  }
  return path
}

async function run() {
  const path = await fixturePath()

  // Building the fixture in this process would put its peak into our
  // measurement, so the parent builds and a child measures.
  if (process.env.HUCRE_BENCH_READY !== "1") {
    const { spawnSync } = await import("node:child_process")
    const result = spawnSync(
      process.execPath,
      [fileURLToPath(import.meta.url), scenario, fixture, String(rowCount)],
      { stdio: "inherit", env: { ...process.env, HUCRE_BENCH_READY: "1" } },
    )
    process.exitCode = result.status ?? 0
    return
  }

  const buf = new Uint8Array(readFileSync(path))
  const started = nowMs()
  let rows
  if (scenario === "readXlsx") {
    rows = (await readXlsx(buf)).sheets[0].rows.length
  } else if (scenario === "readXlsxStyles") {
    rows = (await readXlsx(buf, { readStyles: true })).sheets[0].rows.length
  } else if (scenario === "readXlsxMaxRows") {
    rows = (await readXlsx(buf, { maxRows: 1000 })).sheets[0].rows.length
  } else {
    rows = 0
    for await (const _row of streamXlsxRows(buf)) rows++
  }

  const ms = nowMs() - started
  console.log(
    `${scenario.padEnd(18)} ${fixture.padEnd(17)} ${String(ms).padStart(6)} ms  ` +
      `peakRSS ${String(peakMb()).padStart(5)} MB  ${rows} rows  (${buf.length} bytes in)`,
  )
}

if (!SCENARIOS.includes(scenario) || !FIXTURES.includes(fixture)) {
  console.error(
    `usage: node bench/read.mjs <${SCENARIOS.join("|")}> <${FIXTURES.join("|")}> [rows]`,
  )
  process.exitCode = 1
} else {
  await run()
}
