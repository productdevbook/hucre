// One fresh process per sample: node --expose-gc bench/cell-store.mjs blocks wide
// Compares actual shipped helpers with v1's flat string-keyed Map.
// Measures metadata only, not XLSX parsing or total workbook memory.
import { createCellStore, getCell, setCell } from "../dist/cell.mjs"

const [mode, shape] = process.argv.slice(2)
if (!["flat", "blocks"].includes(mode) || !["wide", "narrow"].includes(shape) || !globalThis.gc) {
  throw new Error("Use node --expose-gc bench/cell-store.mjs <flat|blocks> <wide|narrow>")
}
const [rows, cols] = shape === "wide" ? [20_000, 32] : [240_000, 1]
const cells = mode === "flat" ? new Map() : createCellStore()
const set =
  mode === "flat"
    ? (row, col, cell) => cells.set(`${row},${col}`, cell)
    : (row, col, cell) => setCell(cells, row, col, cell)
const get =
  mode === "flat"
    ? (row, col) => cells.get(`${row},${col}`)
    : (row, col) => getCell(cells, row, col)

globalThis.gc()
const before = process.memoryUsage().heapUsed
const buildStart = performance.now()
for (let row = 0; row < rows; row++) {
  for (let col = 0; col < cols; col++) set(row, col, { value: row + col, type: "number" })
}
const buildMs = performance.now() - buildStart
const lookupStart = performance.now()
let sum = 0
for (let row = 0; row < rows; row++) {
  for (let col = 0; col < cols; col++) sum += get(row, col).value
}
const lookupMs = performance.now() - lookupStart
const expected = (cols * rows * (rows - 1)) / 2 + (rows * cols * (cols - 1)) / 2
if (sum !== expected || cells.size !== rows * cols) throw new Error("Incorrect cell storage result")
globalThis.gc()
console.log(
  JSON.stringify({
    node: process.version,
    mode,
    shape,
    cells: cells.size,
    buildMs: +buildMs.toFixed(1),
    lookupMs: +lookupMs.toFixed(1),
    retainedMB: +((process.memoryUsage().heapUsed - before) / 1024 ** 2).toFixed(1),
    peakMB: +(process.resourceUsage().maxRSS / 1024).toFixed(1),
    sum,
    last: get(rows - 1, cols - 1).value,
  }),
)
