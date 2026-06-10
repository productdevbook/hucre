// ── Resource limits (DoS hardening) ─────────────────────────────────
// Central place for the bounds used to defend against malicious /
// malformed input (zip bombs, billion-laughs cell refs, etc.).

/**
 * Absolute hard cap on the number of bytes any single entry may
 * decompress to. Defends against zip bombs that claim a small
 * compressed size but expand to gigabytes. Default ~2 GiB.
 */
export const MAX_DECOMPRESSED_BYTES = 2 * 1024 * 1024 * 1024

/** Maximum row index (0-based) — Excel supports 1,048,576 rows. */
export const MAX_ROW_INDEX = 1_048_575

/** Maximum column index (0-based) — Excel supports 16,384 columns. */
export const MAX_COL_INDEX = 16_383

/**
 * Upper bound on the password-derivation spin count accepted from an
 * encrypted workbook. Office uses 100,000; we allow a generous ceiling
 * so a hostile file cannot pin a CPU for minutes.
 */
export const MAX_SPIN_COUNT = 10_000_000
