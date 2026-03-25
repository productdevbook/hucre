// ── Cell & Range Utility Functions ──────────────────────────────────
// Re-exports existing utilities and adds new ones.

import { parseCellRef } from "./xlsx/worksheet";

// Re-export existing functions
export { parseCellRef } from "./xlsx/worksheet";
export { colToLetter, cellRef, rangeRef } from "./xlsx/worksheet-writer";

// ── New Utilities ──────────────────────────────────────────────────

/**
 * Convert a column letter (e.g. "A", "Z", "AA") to a 0-based column index.
 * This is the inverse of `colToLetter`.
 *
 *   "A" → 0, "Z" → 25, "AA" → 26, "ZZ" → 701
 */
export function letterToCol(letter: string): number {
  let col = 0;
  for (let i = 0; i < letter.length; i++) {
    const code = letter.charCodeAt(i);
    // Support both uppercase and lowercase
    let value: number;
    if (code >= 65 && code <= 90) {
      value = code - 64; // A=1, B=2, ...
    } else if (code >= 97 && code <= 122) {
      value = code - 96; // a=1, b=2, ...
    } else {
      break;
    }
    col = col * 26 + value;
  }
  return col - 1; // Convert to 0-based
}

/**
 * Parse a range string like "A1:D10" into 0-based coordinates.
 */
export function parseRange(range: string): {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
} {
  const parts = range.split(":");
  const start = parseCellRef(parts[0]!);
  const end = parts.length > 1 ? parseCellRef(parts[1]!) : start;
  return {
    startRow: start.row,
    startCol: start.col,
    endRow: end.row,
    endCol: end.col,
  };
}

/**
 * Check if a cell (0-based row and col) falls within a range.
 */
export function isInRange(
  cellRow: number,
  cellCol: number,
  range: { startRow: number; startCol: number; endRow: number; endCol: number },
): boolean {
  return (
    cellRow >= range.startRow &&
    cellRow <= range.endRow &&
    cellCol >= range.startCol &&
    cellCol <= range.endCol
  );
}
