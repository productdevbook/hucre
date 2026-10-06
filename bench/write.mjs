// ── XLSX write paths ─────────────────────────────────────────────────
//
//   node bench/write.mjs [scenario] [rows]
//
// One scenario per process, deliberately: maxRSS is a high-water mark for
// the whole process, so a second measurement in the same run inherits the
// first one's peak. See bench/README.md.

const { writeXlsx, writeXlsxStream, XlsxStreamWriter } = await import(
  process.env.HUCRE_BENCH_ENTRY ?? "../dist/index.mjs"
)

const SCENARIOS = [
  "writeXlsx",
  "writeXlsxData",
  "writeXlsxDataLinks",
  "writeXlsxStream",
  "XlsxStreamWriter",
]

const scenario = process.argv[2] ?? "writeXlsxStream"
const rowCount = Number(process.argv[3] ?? 100000)
const COLS = 12

function makeRow(i) {
  const row = []
  for (let c = 0; c < COLS; c++) {
    row.push(
      c % 3 === 0
        ? `text ${i}-${c}`
        : c % 3 === 1
          ? i * c
          : new Date(Date.UTC(2024, 0, 1 + (i % 28))),
    )
  }
  return row
}

const peakMb = () => Math.round(process.resourceUsage().maxRSS / 1024)
const nowMs = () => Number(process.hrtime.bigint() / 1000000n)

async function drain(stream) {
  let total = 0
  const reader = stream.getReader()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    total += value.length
  }
  return total
}

async function run() {
  let started
  let bytes

  if (scenario === "writeXlsxData" || scenario === "writeXlsxDataLinks") {
    const columns = Array.from({ length: COLS }, (_, c) => ({
      key: `c${c}`,
      header: `Column ${c}`,
      numFmt: c % 3 === 1 ? "0.00" : undefined,
      style: { font: { bold: true } },
    }))
    const data = Array.from({ length: rowCount }, (_, i) => {
      const values = makeRow(i)
      if (scenario === "writeXlsxDataLinks") {
        values[0] = { text: `text ${i}-0`, hyperlink: `https://example.com/items/${i}` }
      }
      return Object.fromEntries(values.map((value, c) => [`c${c}`, value]))
    })
    started = nowMs()
    bytes = (await writeXlsx({ sheets: [{ name: "S", columns, data }] })).length
  } else if (scenario === "writeXlsx") {
    // The whole model is built first, which is the cost being measured —
    // so it is built before the clock starts.
    const rows = Array.from({ length: rowCount }, (_, i) => makeRow(i))
    started = nowMs()
    bytes = (await writeXlsx({ sheets: [{ name: "S", rows }] })).length
  } else if (scenario === "writeXlsxStream") {
    function* rows() {
      for (let i = 0; i < rowCount; i++) yield makeRow(i)
    }
    started = nowMs()
    bytes = await drain(writeXlsxStream(rows(), { name: "S" }))
  } else {
    started = nowMs()
    const writer = new XlsxStreamWriter({ name: "S" })
    for (let i = 0; i < rowCount; i++) writer.addRow(makeRow(i))
    bytes = (await writer.finish()).length
  }

  const ms = nowMs() - started
  console.log(
    `${scenario.padEnd(18)} rows=${String(rowCount).padStart(7)}  ` +
      `${String(ms).padStart(6)} ms  peakRSS ${String(peakMb()).padStart(5)} MB  ` +
      `${Math.round((rowCount * COLS) / (ms / 1000)).toLocaleString("en-US")} cells/s  ` +
      `${bytes} bytes`,
  )
}

if (!SCENARIOS.includes(scenario)) {
  console.error(`unknown scenario "${scenario}" — one of: ${SCENARIOS.join(", ")}`)
  process.exitCode = 1
} else {
  await run()
}
