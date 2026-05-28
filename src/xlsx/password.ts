// ── Sheet Protection Password Hash ──────────────────────────────────
// Legacy Excel-compatible password hash for sheet protection.
// This is NOT cryptographically secure — it's a simple 16-bit hash
// that Excel uses for backward-compatible sheet protection.

/**
 * Compute the legacy Excel sheet protection password hash.
 *
 * The result is a 4-character uppercase hex string (e.g. "CC3D")
 * suitable for the `password` attribute on `<sheetProtection>`.
 */
export function hashSheetPassword(password: string): string {
  let hash = 0

  for (let i = password.length - 1; i >= 0; i--) {
    const charCode = password.charCodeAt(i)
    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff)
    hash ^= charCode
  }

  hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff)
  hash ^= password.length
  hash ^= 0xce4b

  return hash.toString(16).toUpperCase().padStart(4, "0")
}
