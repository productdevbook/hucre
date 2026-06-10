// ── ECMA-376 Agile Encryption ────────────────────────────────────────
// Decrypt and encrypt password-protected OOXML workbooks using the Agile
// encryption scheme (Excel 2010+ default), per [MS-OFFCRYPTO].
//
// The cipher is AES-CBC. WebCrypto's AES-CBC always applies PKCS#7
// padding, but OOXML segments are raw (unpadded, block-aligned), so we use
// two well-known tricks:
//   • decrypt: append one ciphertext block crafted so its plaintext is a
//     full 0x10×16 PKCS#7 pad block, which WebCrypto strips — leaving the
//     true plaintext.
//   • encrypt: encrypt normally, then drop the trailing padding block;
//     the leading blocks are the raw CBC ciphertext.
// SHA hashing uses WebCrypto's `subtle.digest` directly.

import { readCfb, writeCfb } from "./cfb"
import { DecryptionError } from "../../errors"
import { MAX_SPIN_COUNT } from "../../limits"

// Block keys (MS-OFFCRYPTO §2.3.4.x).
const BLOCK_VERIFIER_INPUT = new Uint8Array([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79])
const BLOCK_VERIFIER_VALUE = new Uint8Array([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e])
const BLOCK_KEY_VALUE = new Uint8Array([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6])
const BLOCK_HMAC_KEY = new Uint8Array([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6])
const BLOCK_HMAC_VALUE = new Uint8Array([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33])

const SEGMENT = 4096
const ZERO_IV = new Uint8Array(16)

interface AgileKeyInfo {
  saltValue: Uint8Array
  hashAlgorithm: string // WebCrypto name, e.g. "SHA-512"
  keyBytes: number
  blockSize: number
  spinCount: number
  encryptedVerifierHashInput: Uint8Array
  encryptedVerifierHashValue: Uint8Array
  encryptedKeyValue: Uint8Array
}

interface AgileKeyData {
  saltValue: Uint8Array
  hashAlgorithm: string
  keyBytes: number
  blockSize: number
}

// ── Public: decrypt ──────────────────────────────────────────────────

/**
 * Decrypt an Agile-encrypted OOXML container into the inner ZIP package.
 * Throws {@link DecryptionError} for a wrong password or unsupported
 * encryption variant.
 */
export async function decryptAgile(data: Uint8Array, password: string): Promise<Uint8Array> {
  let streams: Map<string, Uint8Array>
  try {
    streams = readCfb(data)
  } catch (err) {
    throw new DecryptionError("Not a valid encrypted workbook container.", { cause: err })
  }
  const info = streams.get("EncryptionInfo")
  const pkg = streams.get("EncryptedPackage")
  if (!info || !pkg)
    throw new DecryptionError("Encrypted workbook missing EncryptionInfo/EncryptedPackage.")

  const major = info[0] | (info[1] << 8)
  const minor = info[2] | (info[3] << 8)
  if (major !== 4 || minor !== 4) {
    throw new DecryptionError(
      `Unsupported encryption (version ${major}.${minor}); only ECMA-376 Agile is supported.`,
    )
  }
  const xml = new TextDecoder("utf-8").decode(info.subarray(8))
  const key = parseKeyEncryptor(xml)
  const keyData = parseKeyData(xml)

  const secretKey = await deriveSecretKey(password, key)
  if (!secretKey) throw new DecryptionError("Incorrect password.")

  // Data integrity (MS-OFFCRYPTO §2.3.4.14): verify the HMAC over the
  // EncryptedPackage stream before trusting the decrypted bytes. Tamper or
  // truncation of the ciphertext is rejected here rather than surfacing as
  // silently-corrupt output.
  await verifyDataIntegrity(xml, pkg, secretKey, keyData)

  return decryptPackage(pkg, secretKey, keyData)
}

/**
 * Verify the `<dataIntegrity>` HMAC (MS-OFFCRYPTO §2.3.4.14). The HMAC key
 * and value are AES-CBC encrypted with the package secret key; once
 * recovered, the HMAC is computed over the entire EncryptedPackage stream
 * (8-byte size prefix included). A mismatch means the ciphertext was
 * tampered with or corrupted. No-op when the element is absent (older
 * writers) or the hash isn't SHA-512.
 */
async function verifyDataIntegrity(
  xml: string,
  pkg: Uint8Array,
  secretKey: Uint8Array,
  keyData: AgileKeyData,
): Promise<void> {
  const tag = getElementTag(xml, "dataIntegrity")
  if (!tag) return // not present — nothing to verify
  if (keyData.hashAlgorithm !== "SHA-512") return // only SHA-512 HMAC supported

  const encHmacKey = base64Decode(getAttr(tag, "encryptedHmacKey"))
  const encHmacValue = base64Decode(getAttr(tag, "encryptedHmacValue"))
  if (encHmacKey.length === 0 || encHmacValue.length === 0) return

  const ivKey = await iv(
    keyData.saltValue,
    BLOCK_HMAC_KEY,
    keyData.hashAlgorithm,
    keyData.blockSize,
  )
  const ivVal = await iv(
    keyData.saltValue,
    BLOCK_HMAC_VALUE,
    keyData.hashAlgorithm,
    keyData.blockSize,
  )
  const hmacKey = await aesDecrypt(secretKey, ivKey, encHmacKey)
  const expected = await aesDecrypt(secretKey, ivVal, encHmacValue)

  const actual = await hmacSha512(hmacKey, pkg)
  // Compare the leading hashSize (64) bytes — the decrypted value is
  // block-padded to 16-byte alignment.
  const n = Math.min(actual.length, expected.length)
  let diff = actual.length === 0 ? 1 : 0
  for (let i = 0; i < n; i++) diff |= actual[i] ^ expected[i]
  if (diff !== 0) {
    throw new DecryptionError("Encrypted package failed integrity check (HMAC mismatch).")
  }
}

// ── Public: encrypt ──────────────────────────────────────────────────

/**
 * Encrypt an OOXML ZIP package into an Agile-encrypted OLE2/CFB container.
 * `spinCount` defaults to Excel's 100000; tests may lower it for speed.
 */
export async function encryptAgile(
  zip: Uint8Array,
  password: string,
  opts?: { spinCount?: number },
): Promise<Uint8Array> {
  const spinCount = opts?.spinCount ?? 100000
  const hashAlgorithm = "SHA-512"
  const keyBytes = 32
  const blockSize = 16
  const hashSize = 64

  const keyDataSalt = randomBytes(16)
  const pwSalt = randomBytes(16)
  const secretKey = randomBytes(keyBytes)

  // Password-derived key chain.
  const chain = await passwordChain(password, pwSalt, spinCount, hashAlgorithm)

  // Verifier: random input, its hash, both encrypted under password keys.
  const verifierInput = randomBytes(blockSize)
  const verifierHash = await digest(hashAlgorithm, verifierInput)
  const kVin = await deriveBlockKey(chain, BLOCK_VERIFIER_INPUT, keyBytes, hashAlgorithm)
  const kVval = await deriveBlockKey(chain, BLOCK_VERIFIER_VALUE, keyBytes, hashAlgorithm)
  const kVkey = await deriveBlockKey(chain, BLOCK_KEY_VALUE, keyBytes, hashAlgorithm)
  const encryptedVerifierHashInput = await aesEncrypt(kVin, pwSalt, verifierInput)
  const encryptedVerifierHashValue = await aesEncrypt(
    kVval,
    pwSalt,
    padBlock(verifierHash, blockSize),
  )
  const encryptedKeyValue = await aesEncrypt(kVkey, pwSalt, secretKey)

  // Encrypted package (8-byte size prefix + segmented AES-CBC).
  const encryptedPackage = await encryptPackage(zip, secretKey, {
    saltValue: keyDataSalt,
    hashAlgorithm,
    keyBytes,
    blockSize,
  })

  // Data integrity: HMAC-SHA512 over the encrypted package.
  const hmacKey = randomBytes(hashSize)
  const ivHmacKey = await iv(keyDataSalt, BLOCK_HMAC_KEY, hashAlgorithm, blockSize)
  const encryptedHmacKey = await aesEncrypt(secretKey, ivHmacKey, padBlock(hmacKey, blockSize))
  const hmacValue = await hmacSha512(hmacKey, encryptedPackage)
  const ivHmacValue = await iv(keyDataSalt, BLOCK_HMAC_VALUE, hashAlgorithm, blockSize)
  const encryptedHmacValue = await aesEncrypt(
    secretKey,
    ivHmacValue,
    padBlock(hmacValue, blockSize),
  )

  const xml = buildEncryptionInfoXml({
    keyDataSalt,
    pwSalt,
    spinCount,
    keyBits: keyBytes * 8,
    hashSize,
    blockSize,
    encryptedVerifierHashInput,
    encryptedVerifierHashValue,
    encryptedKeyValue,
    encryptedHmacKey,
    encryptedHmacValue,
  })
  const header = new Uint8Array(8)
  const dv = new DataView(header.buffer)
  dv.setUint16(0, 4, true)
  dv.setUint16(2, 4, true)
  dv.setUint32(4, 0x40, true)
  const encryptionInfo = concat([header, new TextEncoder().encode(xml)])

  return writeCfb([
    { name: "EncryptionInfo", data: encryptionInfo },
    { name: "EncryptedPackage", data: encryptedPackage },
  ])
}

// ── Key derivation ───────────────────────────────────────────────────

function utf16le(s: string): Uint8Array {
  const out = new Uint8Array(s.length * 2)
  const dv = new DataView(out.buffer)
  for (let i = 0; i < s.length; i++) dv.setUint16(i * 2, s.charCodeAt(i), true)
  return out
}

async function passwordChain(
  password: string,
  salt: Uint8Array,
  spinCount: number,
  algo: string,
): Promise<Uint8Array> {
  let h = await digest(algo, concat([salt, utf16le(password)]))
  const counter = new Uint8Array(4)
  const cv = new DataView(counter.buffer)
  for (let i = 0; i < spinCount; i++) {
    cv.setUint32(0, i, true)
    h = await digest(algo, concat([counter, h]))
  }
  return h
}

async function deriveBlockKey(
  chain: Uint8Array,
  blockKey: Uint8Array,
  keyBytes: number,
  algo: string,
): Promise<Uint8Array> {
  const h = await digest(algo, concat([chain, blockKey]))
  return fitKey(h, keyBytes)
}

/** Truncate or 0x36-pad a hash to the required key length (MS-OFFCRYPTO §2.3.4.11). */
function fitKey(hashBytes: Uint8Array, keyBytes: number): Uint8Array {
  if (hashBytes.length >= keyBytes) return hashBytes.subarray(0, keyBytes)
  const out = new Uint8Array(keyBytes).fill(0x36)
  out.set(hashBytes, 0)
  return out
}

/** Derive the package secret key, verifying the password. Returns null on mismatch. */
async function deriveSecretKey(password: string, key: AgileKeyInfo): Promise<Uint8Array | null> {
  const chain = await passwordChain(password, key.saltValue, key.spinCount, key.hashAlgorithm)
  const kVin = await deriveBlockKey(chain, BLOCK_VERIFIER_INPUT, key.keyBytes, key.hashAlgorithm)
  const kVval = await deriveBlockKey(chain, BLOCK_VERIFIER_VALUE, key.keyBytes, key.hashAlgorithm)
  const kVkey = await deriveBlockKey(chain, BLOCK_KEY_VALUE, key.keyBytes, key.hashAlgorithm)

  const verifierInput = await aesDecrypt(kVin, key.saltValue, key.encryptedVerifierHashInput)
  const expectedHash = await aesDecrypt(kVval, key.saltValue, key.encryptedVerifierHashValue)
  const actualHash = await digest(key.hashAlgorithm, verifierInput)
  const hashLen = Math.min(actualHash.length, expectedHash.length)
  if (!bytesEqual(actualHash.subarray(0, hashLen), expectedHash.subarray(0, hashLen))) {
    return null
  }
  const secret = await aesDecrypt(kVkey, key.saltValue, key.encryptedKeyValue)
  return secret.subarray(0, key.keyBytes)
}

// ── Package (de/en)cryption ──────────────────────────────────────────

async function iv(
  salt: Uint8Array,
  blockKey: Uint8Array,
  algo: string,
  blockSize: number,
): Promise<Uint8Array> {
  const h = await digest(algo, concat([salt, blockKey]))
  return fitIv(h, blockSize)
}

function fitIv(h: Uint8Array, blockSize: number): Uint8Array {
  if (h.length >= blockSize) return h.subarray(0, blockSize)
  const out = new Uint8Array(blockSize).fill(0x36)
  out.set(h, 0)
  return out
}

async function decryptPackage(
  pkg: Uint8Array,
  secretKey: Uint8Array,
  keyData: AgileKeyData,
): Promise<Uint8Array> {
  const totalSize = Number(new DataView(pkg.buffer, pkg.byteOffset, 8).getBigUint64(0, true))
  const cipher = pkg.subarray(8)
  const out = new Uint8Array(Math.ceil(cipher.length / SEGMENT) * SEGMENT)
  const counter = new Uint8Array(4)
  const cv = new DataView(counter.buffer)
  let off = 0
  let segIndex = 0
  while (off < cipher.length) {
    const chunk = cipher.subarray(off, Math.min(off + SEGMENT, cipher.length))
    cv.setUint32(0, segIndex, true)
    const segIv = fitIv(
      await digest(keyData.hashAlgorithm, concat([keyData.saltValue, counter])),
      keyData.blockSize,
    )
    const plain = await aesDecrypt(secretKey, segIv, chunk)
    out.set(plain, off)
    off += SEGMENT
    segIndex++
  }
  return out.subarray(0, totalSize)
}

async function encryptPackage(
  zip: Uint8Array,
  secretKey: Uint8Array,
  keyData: AgileKeyData,
): Promise<Uint8Array> {
  const header = new Uint8Array(8)
  new DataView(header.buffer).setBigUint64(0, BigInt(zip.length), true)
  const padded = padBlock(zip, SEGMENT) // segment-align (also block-aligned)
  const out = new Uint8Array(padded.length)
  const counter = new Uint8Array(4)
  const cv = new DataView(counter.buffer)
  let off = 0
  let segIndex = 0
  while (off < padded.length) {
    const chunk = padded.subarray(off, off + SEGMENT)
    cv.setUint32(0, segIndex, true)
    const segIv = fitIv(
      await digest(keyData.hashAlgorithm, concat([keyData.saltValue, counter])),
      keyData.blockSize,
    )
    const enc = await aesEncrypt(secretKey, segIv, chunk)
    out.set(enc, off)
    off += SEGMENT
    segIndex++
  }
  return concat([header, out])
}

// ── WebCrypto primitives ─────────────────────────────────────────────

// WebCrypto's lib types want `BufferSource` backed by a plain ArrayBuffer;
// our byte views are `Uint8Array<ArrayBufferLike>`. They're always
// ArrayBuffer-backed at runtime, so cast at the boundary.
const bs = (u: Uint8Array): BufferSource => u as unknown as BufferSource

async function digest(algo: string, data: Uint8Array): Promise<Uint8Array> {
  return new Uint8Array(await crypto.subtle.digest(algo, bs(data)))
}

async function importAesKey(raw: Uint8Array): Promise<CryptoKey> {
  return crypto.subtle.importKey("raw", bs(raw), { name: "AES-CBC" }, false, ["encrypt", "decrypt"])
}

/** AES-CBC encrypt of one block via zero IV → raw ECB block. */
async function ecbBlock(key: CryptoKey, block: Uint8Array): Promise<Uint8Array> {
  const out = new Uint8Array(
    await crypto.subtle.encrypt({ name: "AES-CBC", iv: bs(ZERO_IV) }, key, bs(block)),
  )
  return out.subarray(0, 16)
}

/** Raw (unpadded) AES-CBC decrypt of block-aligned data. */
async function aesDecryptRaw(
  key: CryptoKey,
  ivBytes: Uint8Array,
  data: Uint8Array,
): Promise<Uint8Array> {
  if (data.length === 0) return new Uint8Array(0)
  const last = data.subarray(data.length - 16)
  const x = new Uint8Array(16)
  for (let i = 0; i < 16; i++) x[i] = last[i] ^ 0x10
  const pad = await ecbBlock(key, x)
  const ext = concat([data, pad])
  const dec = new Uint8Array(
    await crypto.subtle.decrypt({ name: "AES-CBC", iv: bs(ivBytes) }, key, bs(ext)),
  )
  return dec.subarray(0, data.length)
}

/** Raw (unpadded) AES-CBC encrypt of block-aligned data. */
async function aesEncryptRawKey(
  key: CryptoKey,
  ivBytes: Uint8Array,
  data: Uint8Array,
): Promise<Uint8Array> {
  if (data.length === 0) return new Uint8Array(0)
  const enc = new Uint8Array(
    await crypto.subtle.encrypt({ name: "AES-CBC", iv: bs(ivBytes) }, key, bs(data)),
  )
  return enc.subarray(0, data.length)
}

async function aesDecrypt(
  rawKey: Uint8Array,
  ivBytes: Uint8Array,
  data: Uint8Array,
): Promise<Uint8Array> {
  return aesDecryptRaw(await importAesKey(rawKey), ivBytes, data)
}

async function aesEncrypt(
  rawKey: Uint8Array,
  ivBytes: Uint8Array,
  data: Uint8Array,
): Promise<Uint8Array> {
  return aesEncryptRawKey(await importAesKey(rawKey), ivBytes, data)
}

async function hmacSha512(rawKey: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
  const key = await crypto.subtle.importKey(
    "raw",
    bs(rawKey),
    { name: "HMAC", hash: "SHA-512" },
    false,
    ["sign"],
  )
  return new Uint8Array(await crypto.subtle.sign("HMAC", key, bs(data)))
}

// ── EncryptionInfo XML ───────────────────────────────────────────────

function getElementTag(xml: string, local: string): string | null {
  const re = new RegExp(`<(?:\\w+:)?${local}\\b[^>]*?/?>`)
  const m = re.exec(xml)
  return m ? m[0] : null
}

function getAttr(tag: string, name: string): string {
  const m = new RegExp(`\\b${name}="([^"]*)"`).exec(tag)
  return m ? m[1] : ""
}

const HASH_NAME_MAP: Record<string, string> = {
  SHA512: "SHA-512",
  SHA384: "SHA-384",
  SHA256: "SHA-256",
  SHA1: "SHA-1",
}

function mapHash(name: string): string {
  return HASH_NAME_MAP[name] ?? "SHA-512"
}

function parseKeyData(xml: string): AgileKeyData {
  const tag = getElementTag(xml, "keyData")
  if (!tag) throw new DecryptionError("EncryptionInfo missing keyData.")
  return {
    saltValue: base64Decode(getAttr(tag, "saltValue")),
    hashAlgorithm: mapHash(getAttr(tag, "hashAlgorithm")),
    keyBytes: parseInt(getAttr(tag, "keyBits"), 10) / 8,
    blockSize: parseInt(getAttr(tag, "blockSize"), 10),
  }
}

function parseKeyEncryptor(xml: string): AgileKeyInfo {
  const tag = getElementTag(xml, "encryptedKey")
  if (!tag) throw new DecryptionError("EncryptionInfo missing encryptedKey.")
  const rawSpinCount = parseInt(getAttr(tag, "spinCount"), 10)
  if (!Number.isFinite(rawSpinCount) || rawSpinCount < 0) {
    throw new DecryptionError("EncryptionInfo has an invalid spinCount.")
  }
  // The spinCount drives the password-derivation loop. Cap the untrusted
  // value so a hostile file can't pin a CPU for minutes (Office uses
  // 100,000; the ceiling is deliberately generous).
  if (rawSpinCount > MAX_SPIN_COUNT) {
    throw new DecryptionError(
      `EncryptionInfo spinCount ${rawSpinCount} exceeds the maximum of ${MAX_SPIN_COUNT}.`,
    )
  }
  return {
    saltValue: base64Decode(getAttr(tag, "saltValue")),
    hashAlgorithm: mapHash(getAttr(tag, "hashAlgorithm")),
    keyBytes: parseInt(getAttr(tag, "keyBits"), 10) / 8,
    blockSize: parseInt(getAttr(tag, "blockSize"), 10),
    spinCount: rawSpinCount,
    encryptedVerifierHashInput: base64Decode(getAttr(tag, "encryptedVerifierHashInput")),
    encryptedVerifierHashValue: base64Decode(getAttr(tag, "encryptedVerifierHashValue")),
    encryptedKeyValue: base64Decode(getAttr(tag, "encryptedKeyValue")),
  }
}

function buildEncryptionInfoXml(p: {
  keyDataSalt: Uint8Array
  pwSalt: Uint8Array
  spinCount: number
  keyBits: number
  hashSize: number
  blockSize: number
  encryptedVerifierHashInput: Uint8Array
  encryptedVerifierHashValue: Uint8Array
  encryptedKeyValue: Uint8Array
  encryptedHmacKey: Uint8Array
  encryptedHmacValue: Uint8Array
}): string {
  const b = base64Encode
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" ` +
    `xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">` +
    `<keyData saltSize="${p.keyDataSalt.length}" blockSize="${p.blockSize}" keyBits="${p.keyBits}" ` +
    `hashSize="${p.hashSize}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" ` +
    `hashAlgorithm="SHA512" saltValue="${b(p.keyDataSalt)}"/>` +
    `<dataIntegrity encryptedHmacKey="${b(p.encryptedHmacKey)}" ` +
    `encryptedHmacValue="${b(p.encryptedHmacValue)}"/>` +
    `<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">` +
    `<p:encryptedKey spinCount="${p.spinCount}" saltSize="${p.pwSalt.length}" blockSize="${p.blockSize}" ` +
    `keyBits="${p.keyBits}" hashSize="${p.hashSize}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" ` +
    `hashAlgorithm="SHA512" saltValue="${b(p.pwSalt)}" ` +
    `encryptedVerifierHashInput="${b(p.encryptedVerifierHashInput)}" ` +
    `encryptedVerifierHashValue="${b(p.encryptedVerifierHashValue)}" ` +
    `encryptedKeyValue="${b(p.encryptedKeyValue)}"/>` +
    `</keyEncryptor></keyEncryptors></encryption>`
  )
}

// ── Helpers ──────────────────────────────────────────────────────────

function randomBytes(n: number): Uint8Array {
  const b = new Uint8Array(n)
  crypto.getRandomValues(b)
  return b
}

function padBlock(data: Uint8Array, multiple: number): Uint8Array {
  if (data.length % multiple === 0) return data
  const out = new Uint8Array(Math.ceil(data.length / multiple) * multiple)
  out.set(data, 0)
  return out
}

function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false
  let diff = 0
  for (let i = 0; i < a.length; i++) diff |= a[i] ^ b[i]
  return diff === 0
}

function concat(chunks: Uint8Array[]): Uint8Array {
  let total = 0
  for (const c of chunks) total += c.length
  const out = new Uint8Array(total)
  let off = 0
  for (const c of chunks) {
    out.set(c, off)
    off += c.length
  }
  return out
}

const B64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

function base64Encode(bytes: Uint8Array): string {
  let out = ""
  let i = 0
  for (; i + 2 < bytes.length; i += 3) {
    const n = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2]
    out += B64[(n >> 18) & 63] + B64[(n >> 12) & 63] + B64[(n >> 6) & 63] + B64[n & 63]
  }
  if (i < bytes.length) {
    const rem = bytes.length - i
    if (rem === 1) {
      const n = bytes[i] << 16
      out += B64[(n >> 18) & 63] + B64[(n >> 12) & 63] + "=="
    } else {
      const n = (bytes[i] << 16) | (bytes[i + 1] << 8)
      out += B64[(n >> 18) & 63] + B64[(n >> 12) & 63] + B64[(n >> 6) & 63] + "="
    }
  }
  return out
}

function base64Decode(s: string): Uint8Array {
  const clean = s.replace(/[^A-Za-z0-9+/]/g, "")
  const len = Math.floor((clean.length * 3) / 4)
  const out = new Uint8Array(len)
  let bits = 0
  let acc = 0
  let oi = 0
  for (let i = 0; i < clean.length; i++) {
    acc = (acc << 6) | B64.indexOf(clean[i])
    bits += 6
    if (bits >= 8) {
      bits -= 8
      out[oi++] = (acc >> bits) & 0xff
    }
  }
  return out.subarray(0, oi)
}
