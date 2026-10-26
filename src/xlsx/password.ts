import type { ProtectionPasswordHash } from "../_types"

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

// ── Protection Password Digest ──────────────────────────────────────
// A protected sheet or workbook stores only a digest of its password:
// the legacy 16-bit hash above, or the hashValue/saltValue/spinCount of
// ECMA-376 §18.2.29 / §18.3.1.85. The password cannot be recovered from
// it, so the digest is carried as-is. Dropping it on read meant a read →
// write round trip kept the protection and lost the password, leaving a
// "protected" sheet anyone could unprotect.

/** Field of `ProtectionPasswordHash` → attribute name on `<sheetProtection>`. */
const HASH_ATTRS: Array<[keyof ProtectionPasswordHash, string]> = [
  ["legacy", "password"],
  ["algorithmName", "algorithmName"],
  ["hashValue", "hashValue"],
  ["saltValue", "saltValue"],
  ["spinCount", "spinCount"],
]

/** `<workbookProtection>` uses the same names with a `workbook` prefix. */
function hashAttrName(attr: string, prefix: "" | "workbook"): string {
  return prefix ? prefix + attr[0].toUpperCase() + attr.slice(1) : attr
}

/** Read the password digest from protection element attributes. */
export function readPasswordHash(
  attrs: Record<string, string>,
  prefix: "" | "workbook",
): ProtectionPasswordHash | undefined {
  const hash: ProtectionPasswordHash = {}
  let found = false
  for (const [field, attr] of HASH_ATTRS) {
    const val = attrs[hashAttrName(attr, prefix)]
    if (val === undefined) continue
    if (field === "spinCount") {
      const n = Number(val)
      if (!Number.isInteger(n) || n < 0) continue
      hash.spinCount = n
    } else {
      hash[field] = val
    }
    found = true
  }
  return found ? hash : undefined
}

/** Write a password digest onto protection element attributes. */
export function writePasswordHash(
  hash: ProtectionPasswordHash,
  prefix: "" | "workbook",
  attrs: Record<string, string | number>,
): void {
  for (const [field, attr] of HASH_ATTRS) {
    const val = hash[field]
    if (val !== undefined) attrs[hashAttrName(attr, prefix)] = val
  }
}
