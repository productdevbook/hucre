// ── hucre/cell entry point ───────────────────────────────────────────
//
// Cell and range reference helpers, on their own so a string helper does
// not need a whole format's entry point. v1 had four of these on
// `hucre/xlsx` and nine on the root (#474); every one is here.
export {
  parseCellRef,
  colToLetter,
  letterToCol,
  cellRef,
  rangeRef,
  parseRange,
  isInRange,
  r1c1ToA1,
  a1ToR1C1,
  toRange,
  toRanges,
} from "./cell-utils"
export type { RangeLike } from "./cell-utils"
export type { MergeRange } from "./_types"
