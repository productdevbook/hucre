// ── Built-in Number Formats ──────────────────────────────────────────
// Ids 0-49 that a styles.xml may reference without a <numFmt> entry.
// Shared by the reader (id -> format code) and the writer (format code -> id).

export const BUILTIN_NUM_FMTS: Record<number, string> = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  5: "$#,##0_);\\($#,##0\\)",
  6: "$#,##0_);[Red]\\($#,##0\\)",
  7: "$#,##0.00_);\\($#,##0.00\\)",
  8: "$#,##0.00_);[Red]\\($#,##0.00\\)",
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "m/d/yyyy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yyyy h:mm",
  37: "#,##0 ;(#,##0)",
  38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mmss.0",
  48: "##0.0E+0",
  49: "@",
}

/**
 * Format code -> built-in id, for the writer. The reader resolves id 14 to
 * "m/d/yyyy"; writing that back as custom format 164 turned Excel's locale
 * short date (dd.mm.yyyy in a Polish Excel) into a literal month-first
 * date on every cell of a round-tripped workbook. Mapping the code back to
 * its id keeps the cell's meaning, and is what ExcelJS and openpyxl do.
 *
 * Ids 5-8 are left out: their codes spell a literal "$", while the ids
 * mean the locale's currency. An author who writes "$#,##0…" wants
 * dollars, and the read side cannot tell the two apart.
 *
 * Ids 37-40 are left out for the same reason in reverse: this table spells
 * them with a space ("#,##0 ;(#,##0)") while Excel defines them with a
 * paren-width gap ("#,##0_);(#,##0)"). Mapping the code back to the id
 * would change the rendering; keeping the exact code custom preserves it.
 */
export const BUILTIN_NUM_FMT_IDS: ReadonlyMap<string, number> = new Map(
  Object.entries(BUILTIN_NUM_FMTS)
    .filter(([id]) => {
      const n = Number(id)
      if (n <= 0) return false
      if (n >= 5 && n <= 8) return false
      if (n >= 37 && n <= 40) return false
      return true
    })
    .map(([id, code]) => [code, Number(id)]),
)
