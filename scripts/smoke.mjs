// ── Cross-runtime smoke test ─────────────────────────────────────────
//
// The package advertises Deno, Bun, browsers, Workers and Edge, and the
// core is written to Web APIs only for exactly that reason — but nothing
// ever ran it anywhere but Node on Linux. This is the smallest thing that
// would have caught a runtime-specific break: load the built ESM and run
// one round trip through each format that touches the platform.
//
// It is deliberately not the test suite. Vitest is a Node harness; the
// question here is whether `dist/` works where the README says it does.

import {
  parseCsv,
  readXlsx,
  writeCsv,
  writeXlsx,
  writeXlsxStream,
  readOds,
  writeOds,
  createCellStore,
  getCell,
  setCell,
  hasCell,
  deleteCell,
  cellEntries,
} from "../dist/index.mjs"

let failures = 0

function check(what, condition) {
  if (condition) {
    console.log(`  ok   ${what}`)
  } else {
    console.log(`  FAIL ${what}`)
    failures++
  }
}

async function drain(stream) {
  const chunks = []
  let total = 0
  const reader = stream.getReader()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    chunks.push(value)
    total += value.length
  }
  const out = new Uint8Array(total)
  let offset = 0
  for (const chunk of chunks) {
    out.set(chunk, offset)
    offset += chunk.length
  }
  return out
}

const ROWS = [
  ["Name", "Amount", "When"],
  ["Ada", 1234.5, new Date(Date.UTC(2024, 0, 15))],
]

console.log("cell metadata")
{
  const cells = createCellStore([[127, 1, { value: 42, type: "number" }]])
  setCell(cells, 128, 2, { value: 7, type: "number" })
  const cloned = structuredClone(cells)
  deleteCell(cloned, 127, 1)
  check(
    "numeric blocks survive structured clone",
    cells.size === 2 &&
      cloned.size === 1 &&
      hasCell(cloned, 128, 2) &&
      getCell(cloned, 128, 2)?.value === 7 &&
      [...cellEntries(cloned)].length === 1,
  )
  const bytes = await writeXlsx({ sheets: [{ name: "Metadata", cells }] })
  const wb = await readXlsx(bytes, { sparse: true })
  check(
    "sparse metadata round trip",
    wb.sheets[0].rows.length === 0 &&
      getCell(wb.sheets[0].cells, 127, 1)?.value === 42 &&
      getCell(wb.sheets[0].cells, 128, 2)?.value === 7,
  )
}

console.log("csv")
{
  const csv = writeCsv(ROWS)
  const back = parseCsv(csv, { typeInference: true })
  check("round trip", back[1][0] === "Ada" && back[1][1] === 1234.5)
}

console.log("xlsx")
{
  // Exercises DEFLATE via CompressionStream and the ZIP writer.
  const bytes = await writeXlsx({ sheets: [{ name: "S", rows: ROWS }] })
  const wb = await readXlsx(bytes)
  check("round trip", wb.sheets[0].rows[1][0] === "Ada")
  check("numbers survive", wb.sheets[0].rows[1][1] === 1234.5)
  check("dates survive", wb.sheets[0].rows[1][2] instanceof Date)
}

console.log("xlsx streaming")
{
  // Exercises the streaming ZIP writer and backpressure across ReadableStream.
  const bytes = await drain(writeXlsxStream(ROWS, { name: "S" }))
  const wb = await readXlsx(bytes)
  check("round trip", wb.sheets[0].rows[1][0] === "Ada")
}

console.log("ods")
{
  const bytes = await writeOds({ sheets: [{ name: "S", rows: ROWS }] })
  const wb = await readOds(bytes)
  check("round trip", wb.sheets[0].rows[1][0] === "Ada")
}

if (failures > 0) {
  console.log(`\n${failures} check(s) failed`)
  throw new Error(`smoke test failed: ${failures} check(s)`)
}
console.log("\nall checks passed")
