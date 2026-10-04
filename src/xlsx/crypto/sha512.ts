// ── Synchronous SHA-512 ──────────────────────────────────────────────
// Exists for one loop: Agile key derivation hashes the password chain
// `spinCount` times (Excel writes 100,000). With `crypto.subtle.digest`
// every round is an awaited call — on Node each one is a job handed to the
// libuv thread pool and back — so the loop cost ~1 s on a laptop, and past
// a 30 s Lambda timeout on a 256 MB function's CPU share. A block hashed
// in place here costs about a microsecond and no scheduling at all.
//
// Plain FIPS 180-4 on 32-bit halves (JS has no fast 64-bit integers). Only
// SHA-512 — what Excel writes and hucre encrypts with; a file naming
// another hash keeps going through WebCrypto (see `passwordChain`).

// Round constants and initial hash, as [hi, lo] pairs. FIPS 180-4 defines
// them as the first 64 fractional bits of the cube roots of the first 80
// primes (K) and the square roots of the first 8 (IV). They are computed
// once, on first use, rather than shipped: as a 160-entry hex table they
// added ~2 KB gzipped to every bundle that reaches the Agile code.
let K: Int32Array | undefined
let IV: Int32Array | undefined

function constants(): void {
  const primes: bigint[] = []
  for (let n = 2n; primes.length < 80; n++) {
    if (primes.every((p) => n % p !== 0n)) primes.push(n)
  }
  // Integer n-th root by Newton's method: the largest r with r**n <= x.
  const root = (x: bigint, n: bigint): bigint => {
    let r = 1n << (BigInt(x.toString(2).length) / n + 1n)
    for (;;) {
      const next = ((n - 1n) * r + x / r ** (n - 1n)) / n
      if (next >= r) return r
      r = next
    }
  }
  // The fractional bits are the low 64 bits of root(p << 64n).
  const table = (count: number, n: bigint) => {
    const out = new Int32Array(count * 2)
    for (let i = 0; i < count; i++) {
      const v = root(primes[i]! << (64n * n), n) & 0xffffffffffffffffn
      out[i * 2] = Number(v >> 32n) | 0
      out[i * 2 + 1] = Number(v & 0xffffffffn) | 0
    }
    return out
  }
  K = table(80, 3n)
  IV = table(8, 2n)
}

/** Compress one 128-byte block at `block[off..]` into `state` (16 ints, hi/lo pairs). */
function compress(state: Int32Array, w: Int32Array, block: Uint8Array, off: number): void {
  const k = K!
  for (let i = 0; i < 32; i++) {
    const j = off + i * 4
    w[i] = (block[j]! << 24) | (block[j + 1]! << 16) | (block[j + 2]! << 8) | block[j + 3]!
  }
  for (let i = 16; i < 80; i++) {
    // σ0(w[i-15])
    let xh = w[(i - 15) * 2]!
    let xl = w[(i - 15) * 2 + 1]!
    const s0h = ((xh >>> 1) | (xl << 31)) ^ ((xh >>> 8) | (xl << 24)) ^ (xh >>> 7)
    const s0l = ((xl >>> 1) | (xh << 31)) ^ ((xl >>> 8) | (xh << 24)) ^ ((xl >>> 7) | (xh << 25))
    // σ1(w[i-2])
    xh = w[(i - 2) * 2]!
    xl = w[(i - 2) * 2 + 1]!
    const s1h = ((xh >>> 19) | (xl << 13)) ^ ((xl >>> 29) | (xh << 3)) ^ (xh >>> 6)
    const s1l = ((xl >>> 19) | (xh << 13)) ^ ((xh >>> 29) | (xl << 3)) ^ ((xl >>> 6) | (xh << 26))
    // w[i] = σ1 + w[i-7] + σ0 + w[i-16]
    let lo = (s1l >>> 0) + (w[(i - 7) * 2 + 1]! >>> 0) + (s0l >>> 0) + (w[(i - 16) * 2 + 1]! >>> 0)
    const hi = (s1h + w[(i - 7) * 2]! + s0h + w[(i - 16) * 2]! + ((lo / 0x100000000) | 0)) | 0
    lo |= 0
    w[i * 2] = hi
    w[i * 2 + 1] = lo
  }

  let ah = state[0]!,
    al = state[1]!,
    bh = state[2]!,
    bl = state[3]!
  let ch = state[4]!,
    cl = state[5]!,
    dh = state[6]!,
    dl = state[7]!
  let eh = state[8]!,
    el = state[9]!,
    fh = state[10]!,
    fl = state[11]!
  let gh = state[12]!,
    gl = state[13]!,
    hh = state[14]!,
    hl = state[15]!

  for (let i = 0; i < 80; i++) {
    // Σ1(e)
    const S1h = ((eh >>> 14) | (el << 18)) ^ ((eh >>> 18) | (el << 14)) ^ ((el >>> 9) | (eh << 23))
    const S1l = ((el >>> 14) | (eh << 18)) ^ ((el >>> 18) | (eh << 14)) ^ ((eh >>> 9) | (el << 23))
    // Ch(e, f, g)
    const chh = (eh & fh) ^ (~eh & gh)
    const chl = (el & fl) ^ (~el & gl)
    // T1 = h + Σ1 + Ch + K[i] + W[i]
    let t1l = (hl >>> 0) + (S1l >>> 0) + (chl >>> 0) + (k[i * 2 + 1]! >>> 0) + (w[i * 2 + 1]! >>> 0)
    const t1h = (hh + S1h + chh + k[i * 2]! + w[i * 2]! + ((t1l / 0x100000000) | 0)) | 0
    t1l |= 0
    // Σ0(a)
    const S0h = ((ah >>> 28) | (al << 4)) ^ ((al >>> 2) | (ah << 30)) ^ ((al >>> 7) | (ah << 25))
    const S0l = ((al >>> 28) | (ah << 4)) ^ ((ah >>> 2) | (al << 30)) ^ ((ah >>> 7) | (al << 25))
    // Maj(a, b, c)
    const mjh = (ah & bh) ^ (ah & ch) ^ (bh & ch)
    const mjl = (al & bl) ^ (al & cl) ^ (bl & cl)
    // T2 = Σ0 + Maj
    let t2l = (S0l >>> 0) + (mjl >>> 0)
    const t2h = (S0h + mjh + ((t2l / 0x100000000) | 0)) | 0
    t2l |= 0

    hh = gh
    hl = gl
    gh = fh
    gl = fl
    fh = eh
    fl = el
    // e = d + T1
    let sl = (dl >>> 0) + (t1l >>> 0)
    eh = (dh + t1h + ((sl / 0x100000000) | 0)) | 0
    el = sl | 0
    dh = ch
    dl = cl
    ch = bh
    cl = bl
    bh = ah
    bl = al
    // a = T1 + T2
    sl = (t1l >>> 0) + (t2l >>> 0)
    ah = (t1h + t2h + ((sl / 0x100000000) | 0)) | 0
    al = sl | 0
  }

  add(state, 0, ah, al)
  add(state, 2, bh, bl)
  add(state, 4, ch, cl)
  add(state, 6, dh, dl)
  add(state, 8, eh, el)
  add(state, 10, fh, fl)
  add(state, 12, gh, gl)
  add(state, 14, hh, hl)
}

function add(state: Int32Array, i: number, h: number, l: number): void {
  const lo = (state[i + 1]! >>> 0) + (l >>> 0)
  state[i] = (state[i]! + h + ((lo / 0x100000000) | 0)) | 0
  state[i + 1] = lo | 0
}

function writeState(state: Int32Array, out: Uint8Array, off: number): void {
  for (let i = 0; i < 16; i++) {
    const v = state[i]!
    const j = off + i * 4
    out[j] = v >>> 24
    out[j + 1] = v >>> 16
    out[j + 2] = v >>> 8
    out[j + 3] = v
  }
}

/** SHA-512 of `data`. */
export function sha512(data: Uint8Array): Uint8Array {
  if (!IV) constants()
  const state = IV!.slice()
  const w = new Int32Array(160)
  const full = data.length - (data.length % 128)
  for (let off = 0; off < full; off += 128) compress(state, w, data, off)

  // Final block(s): remaining bytes, 0x80, zeros, then the bit length in
  // the last 16 bytes (only the low 53 bits can be non-zero here).
  const rest = data.length - full
  const tail = new Uint8Array(rest < 112 ? 128 : 256)
  tail.set(data.subarray(full))
  tail[rest] = 0x80
  const bits = data.length * 8
  const dv = new DataView(tail.buffer)
  dv.setUint32(tail.length - 8, Math.floor(bits / 0x100000000))
  dv.setUint32(tail.length - 4, bits >>> 0)
  for (let off = 0; off < tail.length; off += 128) compress(state, w, tail, off)

  const out = new Uint8Array(64)
  writeState(state, out, 0)
  return out
}

/**
 * The Agile spin: `spinCount` times, h = SHA-512(LE32(i) ‖ h), starting
 * from the 64-byte `h`. Each round's input is 68 bytes — one padded
 * block — so the block is built once and only its counter and hash bytes
 * change between rounds. `start` is the first round's counter, so a long
 * spin can run in chunks: spinning a rounds then b rounds from `start` a
 * gives the same hash as a + b rounds at once.
 */
export function sha512Spin(h: Uint8Array, spinCount: number, start = 0): Uint8Array {
  if (h.length !== 64) throw new RangeError("sha512Spin expects a 64-byte SHA-512 hash")
  if (!IV) constants()
  const block = new Uint8Array(128)
  block.set(h, 4)
  block[68] = 0x80
  block[127] = 68 * 8 // message length in bits (544), big-endian in the last bytes
  block[126] = (68 * 8) >>> 8
  const state = new Int32Array(16)
  const w = new Int32Array(160)
  for (let i = start; i < start + spinCount; i++) {
    block[0] = i
    block[1] = i >>> 8
    block[2] = i >>> 16
    block[3] = i >>> 24
    state.set(IV!)
    compress(state, w, block, 0)
    writeState(state, block, 4) // the new hash is the next round's input
  }
  return block.slice(4, 68)
}
