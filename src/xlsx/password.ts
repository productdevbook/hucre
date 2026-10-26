import type { PasswordDigest } from "../_types"

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

/**
 * Attribute names of a digest. Sheets use bare names (`password`,
 * `algorithmName`); the workbook element prefixes them (`workbookPassword`,
 * `revisionsAlgorithmName`).
 */
function digestNames(prefix: string) {
  const c = (n: string) => (prefix ? n[0].toUpperCase() + n.slice(1) : n)
  return {
    hash: prefix ? `${prefix}Password` : "password",
    algorithmName: prefix + c("algorithmName"),
    hashValue: prefix + c("hashValue"),
    saltValue: prefix + c("saltValue"),
    spinCount: prefix + c("spinCount"),
    characterSet: prefix + "PasswordCharacterSet",
  }
}

/** Read a digest off an element's attributes; undefined when it has none. */
export function readPasswordDigest(
  attrs: Record<string, string>,
  prefix: "" | "workbook" | "revisions",
): PasswordDigest | undefined {
  const n = digestNames(prefix)
  const d: PasswordDigest = {}
  if (attrs[n.hash]) d.hash = attrs[n.hash]
  if (attrs[n.algorithmName]) d.algorithmName = attrs[n.algorithmName]
  if (attrs[n.hashValue]) d.hashValue = attrs[n.hashValue]
  if (attrs[n.saltValue]) d.saltValue = attrs[n.saltValue]
  // <sheetProtection> has no character-set attribute.
  if (prefix && attrs[n.characterSet]) d.characterSet = attrs[n.characterSet]
  const spin = Number(attrs[n.spinCount])
  if (attrs[n.spinCount] && Number.isInteger(spin) && spin >= 0) d.spinCount = spin
  return Object.keys(d).length > 0 ? d : undefined
}

/**
 * Put a digest onto an element's attributes. A plaintext password, when
 * given, wins: its legacy hash is written and any carried digest is dropped,
 * because a strong digest of the old password beside a hash of the new one
 * would leave the file with two answers.
 */
export function writePasswordDigest(
  attrs: Record<string, string | number>,
  prefix: "" | "workbook" | "revisions",
  plaintext: string | undefined,
  digest: PasswordDigest | undefined,
): void {
  const n = digestNames(prefix)
  if (plaintext) {
    attrs[n.hash] = hashSheetPassword(plaintext)
    return
  }
  if (!digest) return
  if (digest.hash) attrs[n.hash] = digest.hash
  if (digest.algorithmName) attrs[n.algorithmName] = digest.algorithmName
  if (digest.hashValue) attrs[n.hashValue] = digest.hashValue
  if (digest.saltValue) attrs[n.saltValue] = digest.saltValue
  if (digest.spinCount !== undefined) attrs[n.spinCount] = digest.spinCount
  if (prefix && digest.characterSet) attrs[n.characterSet] = digest.characterSet
}
