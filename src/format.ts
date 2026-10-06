// ── hucre/format entry point ─────────────────────────────────────────
//
// Dates, serials and number formats — the arithmetic behind every reader
// and writer, usable without one.
export {
  serialToDate,
  dateToSerial,
  isDateFormat,
  formatDate,
  parseDate,
  serialToTime,
  timeToSerial,
} from "./_date"
export { formatValue } from "./_format"
export type { FormatOptions, LocaleFormat } from "./_format"
export { cloneCellStyle } from "./_style"
