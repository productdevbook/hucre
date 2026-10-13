// Numeric keys avoid one coordinate string per cell. Blocks also avoid
// one Map per row: that costs more memory than the old flat Map on tall,
// narrow sheets. At Excel's coordinate bounds, the outer Map has at most
// 8,192 entries and an inner Map at most 2,097,152, below V8's 2^24 cap.
import type { Cell, CellStore } from "./_types"
import { InvalidArgumentError } from "./errors"
import { MAX_COL_INDEX } from "./limits"
import { validCoordinates } from "./_validate"

const BLOCK_BITS = 7
const BLOCK_ROWS = 1 << BLOCK_BITS
const COLUMNS = MAX_COL_INDEX + 1

/** Build sparse cell metadata from zero-based [row, column, cell] entries. */
export function createCellStore<T extends Partial<Cell> = Cell>(
  entries?: Iterable<readonly [number, number, T]>,
): CellStore<T> {
  const cells: CellStore<T> = { blocks: new Map(), size: 0 }
  if (entries) for (const [row, col, cell] of entries) setCell(cells, row, col, cell)
  return cells
}

/** Read metadata at zero-based coordinates; absent cells return undefined. */
export function getCell<T extends Partial<Cell>>(
  cells: CellStore<T> | undefined,
  row: number,
  col: number,
): T | undefined {
  // Otherwise an invalid column can alias a valid cell on the next row.
  if (!validCoordinates(row, col)) return undefined
  return cells?.blocks.get(row >>> BLOCK_BITS)?.get((row % BLOCK_ROWS) * COLUMNS + col)
}

/**
 * Set metadata without changing the separate dense rows grid.
 * Coordinates must be integers within Excel's row and column bounds.
 */
export function setCell<T extends Partial<Cell>>(
  cells: CellStore<T>,
  row: number,
  col: number,
  cell: T,
): void {
  if (!validCoordinates(row, col)) {
    throw new InvalidArgumentError("Invalid Excel cell coordinates.")
  }
  const blockIndex = row >>> BLOCK_BITS
  let block = cells.blocks.get(blockIndex)
  if (!block) cells.blocks.set(blockIndex, (block = new Map()))
  const index = (row % BLOCK_ROWS) * COLUMNS + col
  if (!block.has(index)) cells.size++
  block.set(index, cell)
}

/** Whether metadata exists, including an empty override object. */
export function hasCell<T extends Partial<Cell>>(
  cells: CellStore<T> | undefined,
  row: number,
  col: number,
): boolean {
  return getCell(cells, row, col) !== undefined
}

/** Remove metadata and release a block when its final cell is removed. */
export function deleteCell<T extends Partial<Cell>>(
  cells: CellStore<T>,
  row: number,
  col: number,
): boolean {
  if (!validCoordinates(row, col)) return false
  const blockIndex = row >>> BLOCK_BITS
  const block = cells.blocks.get(blockIndex)
  if (!block?.delete((row % BLOCK_ROWS) * COLUMNS + col)) return false
  cells.size--
  if (block.size === 0) cells.blocks.delete(blockIndex)
  return true
}

/** Walk metadata without reconstructing or parsing coordinate strings. */
export function* cellEntries<T extends Partial<Cell>>(
  cells: CellStore<T> | undefined,
): Generator<[number, number, T]> {
  if (!cells) return
  for (const [blockIndex, block] of cells.blocks) {
    const firstRow = blockIndex * BLOCK_ROWS
    for (const [index, cell] of block) {
      yield [firstRow + Math.floor(index / COLUMNS), index % COLUMNS, cell]
    }
  }
}
