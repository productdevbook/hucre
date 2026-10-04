import { afterEach, describe, expect, it, vi } from "vitest"
import { decryptAgile, encryptAgile } from "../src/xlsx/crypto/agile"
import { sha512, sha512Spin } from "../src/xlsx/crypto/sha512"

const hex = (b: Uint8Array) => [...b].map((x) => x.toString(16).padStart(2, "0")).join("")
const subtle512 = async (b: Uint8Array) =>
  new Uint8Array(await crypto.subtle.digest("SHA-512", b as Uint8Array<ArrayBuffer>))

describe("sha512", () => {
  it("matches the FIPS 180-4 'abc' vector", () => {
    expect(hex(sha512(new TextEncoder().encode("abc")))).toBe(
      "ddaf35a193617abacc417349ae20413112e6fa4e89a97ea20a9eeee64b55d39a" +
        "2192992a274fc1a836ba3c23a3feebbd454d4423643ce80e2a9ac94fa54ca49f",
    )
  })

  // Every length from empty through three blocks, so each padding branch
  // (tail fits in one block, tail spills into a second) is crossed.
  it("matches WebCrypto for every length 0–300", async () => {
    for (let n = 0; n <= 300; n++) {
      const data = new Uint8Array(n).map((_, i) => (i * 131 + n) & 0xff)
      expect(hex(sha512(data)), `length ${n}`).toBe(hex(await subtle512(data)))
    }
  })
})

describe("sha512Spin", () => {
  // The loop as MS-OFFCRYPTO §2.3.4.11 writes it, on WebCrypto.
  async function reference(h: Uint8Array, spinCount: number) {
    for (let i = 0; i < spinCount; i++) {
      const input = new Uint8Array(68)
      new DataView(input.buffer).setUint32(0, i, true)
      input.set(h, 4)
      h = await subtle512(input)
    }
    return h
  }

  it("matches the WebCrypto spin, including counters past one byte", async () => {
    for (const spins of [0, 1, 2, 255, 256, 1000]) {
      const seed = await subtle512(new TextEncoder().encode(`seed ${spins}`))
      expect(hex(sha512Spin(seed, spins)), `${spins} spins`).toBe(hex(await reference(seed, spins)))
    }
  })
})

describe("sha512Spin in chunks", () => {
  // `passwordChain` spins in chunks (see agile.ts); the counter has to carry
  // across them or every chunk after the first hashes the wrong rounds.
  it("a then b rounds from counter a equals a + b rounds at once", async () => {
    const seed = await subtle512(new TextEncoder().encode("chunks"))
    const whole = sha512Spin(seed, 25_000)
    let h: Uint8Array = seed
    for (const [from, n] of [
      [0, 10_000],
      [10_000, 10_000],
      [20_000, 5_000],
    ] as const) {
      h = sha512Spin(h, n, from)
    }
    expect(hex(h)).toBe(hex(whole))
    // Boundaries that are not multiples of 256 move the counter's low byte too.
    expect(hex(sha512Spin(sha512Spin(seed, 300), 24_700, 300))).toBe(hex(whole))
  })
})

describe("agile key derivation", () => {
  afterEach(() => vi.restoreAllMocks())

  // The spin used to await `crypto.subtle.digest` once per round. On Node
  // each call is a thread-pool round trip, so Excel's 100,000 rounds took
  // ~1 s on a laptop and timed out a 256 MB Lambda at 30 s. Counting calls
  // pins the fix without a timing assertion: before it, this file made one
  // digest call per spin (2,000+ here); now a handful per workbook.
  it("does not call subtle.digest once per spin", async () => {
    const digest = vi.spyOn(crypto.subtle, "digest")
    const payload = new TextEncoder().encode("PK" + "x".repeat(5000))
    const enc = await encryptAgile(payload, "pw", { spinCount: 2000 })
    expect([...(await decryptAgile(enc, "pw"))]).toEqual([...payload])
    expect(digest.mock.calls.length).toBeLessThan(50)
  })

  // A file being decrypted picks its own spin count, up to MAX_SPIN_COUNT
  // (10,000,000 — ~9 s of hashing). The spin must give the event loop a
  // turn between chunks rather than run as one unbroken synchronous loop.
  it("yields to the event loop between spin chunks", async () => {
    const timer = vi.spyOn(globalThis, "setTimeout")
    const enc = await encryptAgile(new TextEncoder().encode("PK yield"), "pw", {
      spinCount: 35_000,
    })
    const yields = timer.mock.calls.filter((c) => c[1] === 0).length
    expect(yields).toBeGreaterThanOrEqual(3) // 35,000 rounds = 4 chunks of ≤ 10,000
    expect(new TextDecoder().decode(await decryptAgile(enc, "pw"))).toBe("PK yield")
  })
})
