// ── The legacy indexed colour palette ────────────────────────────────
//
// A colour can name a palette index instead of an RGB — `indexed="2"` —
// and hucre carried the index and nothing else, because it had no palette
// to resolve it against. A caller got `{ indexed: 2 }` and no way to know
// that means red, and a file that *overrode* the palette was dropped
// entirely: `<indexedColors>` was not read at all, so even a caller with
// its own copy of the defaults got the wrong answer for those files.
//
// Found by `scripts/spec-coverage.mjs`, which crosses the schema with the
// fixture corpus: `indexedColors` and `rgbColor` are in ECMA-376, and in
// the corpus, and were nowhere in `src/`.
//
// ECMA-376 Part 1 §18.8.27 defines the defaults and says why they are odd:
// "0-7 are redundant of 8-15 to preserve backwards compatibility". Indices
// 64 and 65 are the system foreground and background — they have no ARGB
// in the table and are deliberately absent here, so a colour naming one
// stays unresolved rather than being given a colour the file did not
// choose.
//
// The values were lifted from the specification text rather than typed
// from memory, and cross-checked against the `<indexedColors>` block
// openpyxl writes into `test/fixtures/openpyxl-basic.xlsx` — an
// independent implementation, agreeing on all 64.

/**
 * Indices 0-63 of the default palette, as 6-digit RGB.
 *
 * The spec writes them ARGB with a `00` alpha, which is not transparency
 * — `ColorSpec.rgb` is 6 hex digits, so the prefix is dropped here as it
 * is everywhere else in the reader.
 */
export const DEFAULT_INDEXED_PALETTE: readonly string[] = [
  "000000", // 0
  "FFFFFF", // 1
  "FF0000", // 2
  "00FF00", // 3
  "0000FF", // 4
  "FFFF00", // 5
  "FF00FF", // 6
  "00FFFF", // 7
  "000000", // 8
  "FFFFFF", // 9
  "FF0000", // 10
  "00FF00", // 11
  "0000FF", // 12
  "FFFF00", // 13
  "FF00FF", // 14
  "00FFFF", // 15
  "800000", // 16
  "008000", // 17
  "000080", // 18
  "808000", // 19
  "800080", // 20
  "008080", // 21
  "C0C0C0", // 22
  "808080", // 23
  "9999FF", // 24
  "993366", // 25
  "FFFFCC", // 26
  "CCFFFF", // 27
  "660066", // 28
  "FF8080", // 29
  "0066CC", // 30
  "CCCCFF", // 31
  "000080", // 32
  "FF00FF", // 33
  "FFFF00", // 34
  "00FFFF", // 35
  "800080", // 36
  "800000", // 37
  "008080", // 38
  "0000FF", // 39
  "00CCFF", // 40
  "CCFFFF", // 41
  "CCFFCC", // 42
  "FFFF99", // 43
  "99CCFF", // 44
  "FF99CC", // 45
  "CC99FF", // 46
  "FFCC99", // 47
  "3366FF", // 48
  "33CCCC", // 49
  "99CC00", // 50
  "FFCC00", // 51
  "FF9900", // 52
  "FF6600", // 53
  "666699", // 54
  "969696", // 55
  "003366", // 56
  "339966", // 57
  "003300", // 58
  "333300", // 59
  "993300", // 60
  "993366", // 61
  "333399", // 62
  "333333", // 63
]
