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
  read,
  ParseError,
  InvalidArgumentError,
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
  insertRows,
  deleteRows,
  insertColumns,
  deleteColumns,
  findCells,
  replaceCells,
  fillTemplate,
  readObjects,
  sheetToObjects,
  toJson,
  write,
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

console.log("model projection")
{
  const authored = { sheets: [{ name: "Data", data: [{ Name: "Ada" }, { Score: 42 }] }] }
  const dataRows = (await readXlsx(await writeXlsx(authored))).sheets[0].rows
  check("object data infers every field", dataRows[0][1] === "Score" && dataRows[2][1] === 42)
  check(
    "object data uses the same text rows",
    new TextDecoder().decode(await write(authored, { format: "csv" })) ===
      "Name,Score\r\nAda,\r\n,42",
  )
  let ambiguous = false
  try {
    await write({ sheets: [{ name: "S", data: [], rows: [] }] }, { format: "csv" })
  } catch (error) {
    ambiguous = error instanceof InvalidArgumentError
  }
  check("competing row sources reject", ambiguous)
  const bytes = await writeXlsx({
    sheets: [
      {
        name: "S",
        rows: [
          ["Name", "Score"],
          ["Ada", 42],
        ],
      },
    ],
  })
  const sparse = (await readXlsx(bytes, { sparse: true })).sheets[0]
  const result = await readObjects(bytes, { sparse: true })
  check("sparse object projection", result.data[0]?.Name === "Ada" && result.data[0]?.Score === 42)
  check("sparse JSON projection", JSON.parse(toJson(sparse))[0]?.Score === 42)
  check(
    "sparse CSV projection",
    new TextDecoder()
      .decode(await write({ sheets: [sparse] }, { format: "csv" }))
      .includes("Ada,42"),
  )
  const far = {
    name: "Far",
    rows: [],
    cells: createCellStore([
      [0, 0, { value: "Name", type: "string" }],
      [1048575, 0, { value: "last", type: "string" }],
    ]),
  }
  check(
    "far-away values without dense gaps",
    sheetToObjects(far).data[0]?.Name === "last" && far.rows.length === 0,
  )
}

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

console.log("sheet editing")
{
  const sheet = {
    name: "Edit",
    rows: [
      [1, 2],
      [3, 4],
    ],
    cells: createCellStore([[1, 1, { value: 4, type: "formula", formula: "Edit!A2" }]]),
  }
  insertRows(sheet, 0, 1)
  insertColumns(sheet, 0, 1)
  check(
    "values and own-sheet references move together",
    sheet.rows.length === 3 &&
      sheet.rows[2][2] === 4 &&
      getCell(sheet.cells, 2, 2)?.formula === "Edit!B3",
  )
  deleteRows(sheet, 0, 1)
  deleteColumns(sheet, 0, 1)
  check(
    "deletion uses the same coordinate rules",
    sheet.rows.length === 2 &&
      sheet.rows[1][1] === 4 &&
      getCell(sheet.cells, 1, 1)?.formula === "Edit!A2",
  )
}

console.log("csv")
{
  const csv = writeCsv(ROWS)
  const back = parseCsv(csv, { typeInference: true })
  check("round trip", back[1][0] === "Ada" && back[1][1] === 1234.5)
}

console.log("value editing")
{
  const cell = {
    value: "{{value}}",
    type: "formula",
    formula: '"unchanged"',
    formulaResult: "{{value}}",
  }
  const sheet = { name: "Values", rows: [["{{value}}"]], cells: createCellStore([[0, 0, cell]]) }
  fillTemplate({ sheets: [sheet] }, { value: 7 })
  check(
    "template keeps the formula cache synchronized",
    sheet.rows[0][0] === 7 && cell.formulaResult === 7 && cell.type === "formula",
  )
  const sparse = {
    name: "Sparse",
    rows: [],
    cells: createCellStore([[1048575, 16383, { value: "old", type: "string" }]]),
  }
  check(
    "find and replace includes sparse values",
    findCells(sparse, "old").length === 1 && replaceCells(sparse, "old", false) === 1,
  )
  check(
    "sparse edits keep the grid empty",
    sparse.rows.length === 0 && getCell(sparse.cells, 1048575, 16383).type === "boolean",
  )
}

console.log("dense text bounds")
{
  const input = new TextEncoder().encode("a,b,c\nx\nx")
  let bounded = false
  try {
    await read(input, { maxTotalCells: 8 })
  } catch (error) {
    bounded = error instanceof ParseError
  }
  check("read forwards the dense cell bound", bounded)
  const sheet = (await read(input, { maxTotalCells: 9 })).sheets[0]
  check(
    "the exact boundary remains rectangular",
    sheet.rows.length === 3 && sheet.rows[2].length === 3,
  )
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
