// ── Limits as the ZIP readers apply them ─────────────────────────────
// The buffered, local-header and random-access readers each take the
// same caller options and the same sizes from the archive. These are the
// one reading of them, so a ceiling cannot be spelled one way in one
// reader and drift in another.

/**
 * `value` when it is a positive finite number, else `fallback`.
 *
 * A caller passing 0, a negative number, `NaN` or `Infinity` has not
 * asked for "no limit"; the documented default applies instead, as it
 * does in `bufferReadableStream`.
 */
export function positiveOr(value: number | undefined, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) && value > 0 ? value : fallback
}
