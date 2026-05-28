// ── Compound File Binary (CFB / OLE2) ────────────────────────────────
// Minimal reader + writer for the OLE2 container Office uses to wrap a
// password-protected workbook. We only need two named streams:
// `EncryptionInfo` (small) and `EncryptedPackage` (the encrypted OOXML
// ZIP). The writer lays out a compliant container generically, routing
// each stream to the mini-stream or the regular FAT by size, so it works
// for tiny and large packages alike.
//
// Reference: [MS-CFB] Compound File Binary File Format.

import { ParseError } from "../../errors"

const SIG = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]
const ENDOFCHAIN = 0xfffffffe
const FREESECT = 0xffffffff
const FATSECT = 0xfffffffd
const MINI_CUTOFF = 4096
const MINI_SECTOR = 64

// ── Reader ───────────────────────────────────────────────────────────

interface DirEntry {
  name: string
  type: number // 0 unknown, 1 storage, 2 stream, 5 root
  startSector: number
  size: number
}

/**
 * Parse a CFB container and return its named streams. Storage hierarchy
 * is flattened — we only deal with top-level streams, which is all an
 * encrypted workbook has.
 */
export function readCfb(data: Uint8Array): Map<string, Uint8Array> {
  if (data.length < 512) throw new ParseError("CFB: file too small")
  for (let i = 0; i < SIG.length; i++) {
    if (data[i] !== SIG[i]) throw new ParseError("CFB: bad signature")
  }
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength)
  const sectorShift = view.getUint16(30, true)
  const sectorSize = 1 << sectorShift
  const miniSectorShift = view.getUint16(32, true)
  const miniSectorSize = 1 << miniSectorShift
  const numFatSectors = view.getUint32(44, true)
  const firstDirSector = view.getUint32(48, true)
  const miniCutoff = view.getUint32(56, true)
  const firstMiniFatSector = view.getUint32(60, true)
  const numMiniFatSectors = view.getUint32(64, true)
  const firstDifatSector = view.getUint32(68, true)
  const numDifatSectors = view.getUint32(72, true)

  const sectorOffset = (sector: number): number => (sector + 1) * sectorSize

  const readSector = (sector: number): Uint8Array => {
    const off = sectorOffset(sector)
    if (off + sectorSize > data.length) throw new ParseError("CFB: sector out of range")
    return data.subarray(off, off + sectorSize)
  }

  // Collect FAT sector locations: first 109 from the header DIFAT, then
  // any chained DIFAT sectors.
  const fatSectorList: number[] = []
  for (let i = 0; i < 109 && fatSectorList.length < numFatSectors; i++) {
    const s = view.getUint32(76 + i * 4, true)
    if (s === FREESECT || s === ENDOFCHAIN) break
    fatSectorList.push(s)
  }
  let difatSector = firstDifatSector
  let difatGuard = 0
  while (
    difatSector !== ENDOFCHAIN &&
    difatSector !== FREESECT &&
    numDifatSectors > 0 &&
    difatGuard++ < numDifatSectors + 1
  ) {
    const sec = readSector(difatSector)
    const dv = new DataView(sec.buffer, sec.byteOffset, sec.byteLength)
    const entries = sectorSize / 4 - 1
    for (let i = 0; i < entries && fatSectorList.length < numFatSectors; i++) {
      const s = dv.getUint32(i * 4, true)
      if (s !== FREESECT && s !== ENDOFCHAIN) fatSectorList.push(s)
    }
    difatSector = dv.getUint32(sectorSize - 4, true)
  }

  // Build the FAT.
  const fat: number[] = []
  for (const fs of fatSectorList) {
    const sec = readSector(fs)
    const dv = new DataView(sec.buffer, sec.byteOffset, sec.byteLength)
    for (let i = 0; i < sectorSize / 4; i++) fat.push(dv.getUint32(i * 4, true))
  }

  const chain = (start: number): number[] => {
    const out: number[] = []
    let s = start
    let guard = 0
    while (s !== ENDOFCHAIN && s !== FREESECT && guard++ < fat.length + 1) {
      out.push(s)
      s = fat[s] ?? ENDOFCHAIN
    }
    return out
  }

  const readChain = (start: number, size: number): Uint8Array => {
    const sectors = chain(start)
    const buf = new Uint8Array(sectors.length * sectorSize)
    let off = 0
    for (const s of sectors) {
      buf.set(readSector(s), off)
      off += sectorSize
    }
    return size >= 0 ? buf.subarray(0, size) : buf
  }

  // Directory entries (each 128 bytes).
  const dirBytes = readChain(firstDirSector, -1)
  const dirView = new DataView(dirBytes.buffer, dirBytes.byteOffset, dirBytes.byteLength)
  const entries: DirEntry[] = []
  for (let off = 0; off + 128 <= dirBytes.length; off += 128) {
    const type = dirBytes[off + 66]
    if (type === 0) continue // unused entry
    const nameLen = dirView.getUint16(off + 64, true)
    let name = ""
    for (let i = 0; i + 1 < nameLen; i += 2) {
      const code = dirView.getUint16(off + i, true)
      if (code === 0) break
      name += String.fromCharCode(code)
    }
    const startSector = dirView.getUint32(off + 116, true)
    const size = dirView.getUint32(off + 120, true) // low 32 bits (streams here are < 4 GB)
    entries.push({ name, type, startSector, size })
  }

  const root = entries.find((e) => e.type === 5)
  if (!root) throw new ParseError("CFB: missing root entry")

  // Mini-stream container lives in the regular FAT, anchored on the root.
  const miniStream = readChain(root.startSector, root.size)

  // Mini-FAT.
  const miniFat: number[] = []
  if (numMiniFatSectors > 0 && firstMiniFatSector !== ENDOFCHAIN) {
    const mfBytes = readChain(firstMiniFatSector, numMiniFatSectors * sectorSize)
    const mfView = new DataView(mfBytes.buffer, mfBytes.byteOffset, mfBytes.byteLength)
    for (let i = 0; i < mfBytes.length / 4; i++) miniFat.push(mfView.getUint32(i * 4, true))
  }

  const readMini = (start: number, size: number): Uint8Array => {
    const out = new Uint8Array(size)
    let s = start
    let off = 0
    let guard = 0
    while (s !== ENDOFCHAIN && s !== FREESECT && off < size && guard++ < miniFat.length + 1) {
      const from = s * miniSectorSize
      const take = Math.min(miniSectorSize, size - off)
      out.set(miniStream.subarray(from, from + take), off)
      off += take
      s = miniFat[s] ?? ENDOFCHAIN
    }
    return out
  }

  const streams = new Map<string, Uint8Array>()
  for (const e of entries) {
    if (e.type !== 2) continue
    const cutoff = miniCutoff || MINI_CUTOFF
    streams.set(
      e.name,
      e.size < cutoff ? readMini(e.startSector, e.size) : readChain(e.startSector, e.size),
    )
  }
  return streams
}

// ── Writer ───────────────────────────────────────────────────────────

interface StreamInput {
  name: string
  data: Uint8Array
}

/**
 * Assemble a CFB container holding the given top-level streams (plus the
 * implicit Root Entry). Uses 512-byte sectors and 64-byte mini sectors;
 * streams smaller than 4096 bytes go in the mini-stream, the rest in the
 * regular FAT.
 */
export function writeCfb(inputs: StreamInput[]): Uint8Array {
  const sectorSize = 512
  const perSector = sectorSize / 4 // FAT entries per sector

  // 1. Partition streams into mini vs regular and build the mini-stream.
  const miniChunks: Uint8Array[] = []
  const miniStartMini: number[] = [] // mini-sector index where each mini stream starts
  const regular: Array<{ name: string; data: Uint8Array }> = []
  const placement = new Map<string, { mini: boolean }>()

  let miniSectorCount = 0
  for (const s of inputs) {
    if (s.data.length < MINI_CUTOFF) {
      placement.set(s.name, { mini: true })
      miniStartMini.push(miniSectorCount)
      const padded = padTo(s.data, MINI_SECTOR)
      miniChunks.push(padded)
      miniSectorCount += padded.length / MINI_SECTOR
    } else {
      placement.set(s.name, { mini: false })
      regular.push({ name: s.name, data: s.data })
    }
  }
  const miniStream = concat(miniChunks)

  // 2. Lay out regular-FAT payload sectors: each regular stream, then the
  // mini-stream container, then the mini-FAT, then the directory. The FAT
  // sectors themselves come last (their count depends on the total).
  interface Region {
    startSector: number
    sectorCount: number
  }
  let cursor = 0
  const alloc = (byteLen: number): Region => {
    const sectorCount = Math.max(1, Math.ceil(byteLen / sectorSize))
    const startSector = cursor
    cursor += sectorCount
    return { startSector, sectorCount }
  }

  const regularRegions = regular.map((r) => ({
    name: r.name,
    data: r.data,
    region: alloc(r.data.length),
  }))
  const miniStreamRegion = miniStream.length > 0 ? alloc(miniStream.length) : null

  // Mini-FAT: one entry per mini sector.
  const miniFat = new Uint8Array(Math.ceil(Math.max(miniSectorCount, 0) / perSector) * sectorSize)
  {
    const dv = new DataView(miniFat.buffer)
    for (let i = 0; i < miniFat.length / 4; i++) dv.setUint32(i * 4, FREESECT, true)
    // Chain each mini stream's mini sectors.
    let idx = 0
    let mi = 0
    for (const s of inputs) {
      if (s.data.length >= MINI_CUTOFF) continue
      const sectors = padTo(s.data, MINI_SECTOR).length / MINI_SECTOR
      for (let k = 0; k < sectors; k++) {
        const cur = miniStartMini[mi] + k
        dv.setUint32(cur * 4, k === sectors - 1 ? ENDOFCHAIN : cur + 1, true)
      }
      mi++
      idx += sectors
    }
    void idx
  }
  const miniFatRegion = miniSectorCount > 0 ? alloc(miniFat.length) : null

  // Directory: Root + one entry per stream. 4 entries per 512 sector.
  const dirEntryCount = 1 + inputs.length
  const dirBytesLen = Math.ceil(dirEntryCount / (sectorSize / 128)) * sectorSize
  const dirRegion = alloc(dirBytesLen)

  // 3. FAT sizing: total payload sectors so far + the FAT sectors we are
  // about to add. Solve for FAT sector count (self-referential).
  const payloadSectors = cursor
  let numFatSectors = 1
  for (;;) {
    const totalSectors = payloadSectors + numFatSectors
    const need = Math.ceil(totalSectors / perSector)
    if (need <= numFatSectors) break
    numFatSectors = need
  }
  const fatRegion = { startSector: cursor, sectorCount: numFatSectors }
  cursor += numFatSectors
  const totalSectors = cursor

  // 4. Build the FAT.
  const fat = new Uint32Array(numFatSectors * perSector)
  fat.fill(FREESECT)
  const chainRegion = (r: Region): void => {
    for (let i = 0; i < r.sectorCount; i++) {
      fat[r.startSector + i] = i === r.sectorCount - 1 ? ENDOFCHAIN : r.startSector + i + 1
    }
  }
  for (const r of regularRegions) chainRegion(r.region)
  if (miniStreamRegion) chainRegion(miniStreamRegion)
  if (miniFatRegion) chainRegion(miniFatRegion)
  chainRegion(dirRegion)
  for (let i = 0; i < numFatSectors; i++) fat[fatRegion.startSector + i] = FATSECT

  // 5. Directory bytes.
  const dir = new Uint8Array(dirBytesLen)
  const dirView = new DataView(dir.buffer)
  const writeDirEntry = (
    idx: number,
    name: string,
    type: number,
    color: number,
    left: number,
    right: number,
    child: number,
    start: number,
    size: number,
  ): void => {
    const o = idx * 128
    for (let i = 0; i < name.length; i++) dirView.setUint16(o + i * 2, name.charCodeAt(i), true)
    dirView.setUint16(o + 64, (name.length + 1) * 2, true)
    dir[o + 66] = type
    dir[o + 67] = color
    dirView.setUint32(o + 68, left, true)
    dirView.setUint32(o + 72, right, true)
    dirView.setUint32(o + 76, child, true)
    dirView.setUint32(o + 116, start, true)
    dirView.setUint32(o + 120, size, true)
  }
  // Mark all slots free first.
  for (let i = 0; i < dirBytesLen / 128; i++) {
    dirView.setUint32(i * 128 + 68, FREESECT, true)
    dirView.setUint32(i * 128 + 72, FREESECT, true)
    dirView.setUint32(i * 128 + 76, FREESECT, true)
  }
  // Root entry (index 0): its child is the first stream; start = mini-stream container.
  writeDirEntry(
    0,
    "Root Entry",
    5,
    1,
    FREESECT,
    FREESECT,
    inputs.length > 0 ? 1 : FREESECT,
    miniStreamRegion ? miniStreamRegion.startSector : ENDOFCHAIN,
    miniStream.length,
  )
  // Stream entries 1..n laid out as a degenerate right-leaning tree:
  // entry i's right sibling is i+1. (Excel's reader accepts this.)
  for (let i = 0; i < inputs.length; i++) {
    const s = inputs[i]
    const place = placement.get(s.name)!
    const entryIdx = i + 1
    const right = i + 1 < inputs.length ? entryIdx + 1 : FREESECT
    let start: number
    if (place.mini) {
      // mini sector index where this stream starts
      let miOff = 0
      for (let j = 0; j < i; j++) {
        const sj = inputs[j]
        if (sj.data.length < MINI_CUTOFF) miOff += padTo(sj.data, MINI_SECTOR).length / MINI_SECTOR
      }
      start = miOff
    } else {
      const reg = regularRegions.find((r) => r.name === s.name)!
      start = reg.region.startSector
    }
    writeDirEntry(entryIdx, s.name, 2, 1, FREESECT, right, FREESECT, start, s.data.length)
  }

  // 6. Assemble the file: header sector + payload sectors.
  const file = new Uint8Array((totalSectors + 1) * sectorSize)
  const fileView = new DataView(file.buffer)
  // Header.
  for (let i = 0; i < SIG.length; i++) file[i] = SIG[i]
  fileView.setUint16(24, 0x003e, true) // minor version
  fileView.setUint16(26, 0x0003, true) // major version (3 → 512-byte sectors)
  fileView.setUint16(28, 0xfffe, true) // byte order
  fileView.setUint16(30, 9, true) // sector shift (512)
  fileView.setUint16(32, 6, true) // mini sector shift (64)
  fileView.setUint32(44, numFatSectors, true)
  fileView.setUint32(48, dirRegion.startSector, true)
  fileView.setUint32(56, MINI_CUTOFF, true)
  fileView.setUint32(60, miniFatRegion ? miniFatRegion.startSector : ENDOFCHAIN, true)
  fileView.setUint32(64, miniFatRegion ? miniFatRegion.sectorCount : 0, true)
  fileView.setUint32(68, ENDOFCHAIN, true) // first DIFAT sector
  fileView.setUint32(72, 0, true) // num DIFAT sectors
  for (let i = 0; i < 109; i++) {
    fileView.setUint32(76 + i * 4, i < numFatSectors ? fatRegion.startSector + i : FREESECT, true)
  }

  const sectorAt = (sector: number): number => (sector + 1) * sectorSize
  const putBytes = (region: Region, bytes: Uint8Array): void => {
    file.set(bytes, sectorAt(region.startSector))
  }
  for (const r of regularRegions) putBytes(r.region, r.data)
  if (miniStreamRegion) putBytes(miniStreamRegion, miniStream)
  if (miniFatRegion) putBytes(miniFatRegion, miniFat)
  putBytes(dirRegion, dir)
  // FAT.
  {
    const off = sectorAt(fatRegion.startSector)
    const dv = new DataView(file.buffer, off, numFatSectors * sectorSize)
    for (let i = 0; i < fat.length; i++) dv.setUint32(i * 4, fat[i], true)
  }
  return file
}

function padTo(data: Uint8Array, multiple: number): Uint8Array {
  if (data.length % multiple === 0) return data
  const out = new Uint8Array(Math.ceil(data.length / multiple) * multiple)
  out.set(data, 0)
  return out
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
