// ── node bench/stream-layout.mjs [rows] ─────────────────────────────────
//
// What streamXlsxRows holds for a large workbook in Excel's part order —
// the sheet before styles.xml and sharedStrings.xml, sizes in the local
// headers, no data descriptors — read as a ReadableStream (which has to
// buffer it) and as a RandomAccessInput (which does not).
//
// The workbook is generated, not committed: 4,000,000 rows (the default)
// is ~300 MB zipped and ~1.5 GB of sheet XML, the shape of a large
// battery-cycler or instrument export. It is written to the OS temp dir
// and removed afterwards. One child process per measurement, because
// maxRSS is a high-water mark for the whole process.
//
//   bun run build && node bench/stream-layout.mjs            # 4M rows
//   node bench/stream-layout.mjs 500000                     # smaller
//   node bench/stream-layout.mjs generate <rows> <file>    # keep the file
//   node bench/stream-layout.mjs measure <dist> <stream|random> <file>
//
// `measure` takes any build's dist/index.mjs, so the same file can be read
// by an older release for a before/after (its `stream` mode needs nothing
// this change added).

import { spawnSync } from "node:child_process"
import { openAsBlob } from "node:fs"
import { mkdtemp, open, rm } from "node:fs/promises"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { fileURLToPath, pathToFileURL } from "node:url"
import { constants, crc32, createDeflateRaw } from "node:zlib"

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const PKG = "http://schemas.openxmlformats.org/package/2006"

if (process.argv[2] === "measure") {
  await measure(...process.argv.slice(3))
} else if (process.argv[2] === "generate") {
  await generate(process.argv[4], Number(process.argv[3]))
} else {
  await main(Number(process.argv[2] ?? 4_000_000))
}

async function main(rows) {
  const dist = fileURLToPath(new URL("../dist/index.mjs", import.meta.url))
  const dir = await mkdtemp(join(tmpdir(), "hucre-bench-"))
  const file = join(dir, "excel-order.xlsx")
  try {
    console.log(`generating ${rows.toLocaleString()} rows in Excel's part order…`)
    await generate(file, rows)
    for (const mode of ["stream", "random"]) {
      const r = spawnSync(
        process.execPath,
        [fileURLToPath(import.meta.url), "measure", dist, mode, file],
        {
          stdio: "inherit",
        },
      )
      if (r.status !== 0) process.exitCode = 1
    }
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
}

// ── Measurement ─────────────────────────────────────────────────────────

async function measure(dist, mode, file) {
  const hucre = await import(pathToFileURL(dist).href)
  const blob = await openAsBlob(file)
  let read = 0
  let input
  if (mode === "stream") {
    const inner = blob.stream().getReader()
    input = new ReadableStream(
      {
        async pull(c) {
          const { value, done } = await inner.read()
          if (done) return c.close()
          read += value.length
          c.enqueue(value)
        },
        cancel: (reason) => inner.cancel(reason),
      },
      { highWaterMark: 0 },
    )
  } else {
    const source = hucre.randomAccessFromBlob(blob)
    input = {
      size: source.size,
      read(offset, length) {
        read += length
        return source.read(offset, length)
      },
    }
  }

  const t0 = performance.now()
  let rows = 0
  let readBeforeFirstRow = 0
  for await (const _ of hucre.streamXlsxRows(input)) {
    if (rows === 0) readBeforeFirstRow = read
    rows++
  }
  const label = mode === "stream" ? "ReadableStream   " : "RandomAccessInput"
  console.log(
    `  ${label}  ${(blob.size / 1e6).toFixed(1)} MB file  ` +
      `${rows.toLocaleString()} rows  ` +
      `${((100 * readBeforeFirstRow) / blob.size).toFixed(1)}% read before first row  ` +
      `${((performance.now() - t0) / 1000).toFixed(1)} s  ` +
      `peak RSS ${Math.round(process.resourceUsage().maxRSS / 1024)} MB`,
  )
}

// ── Generation ──────────────────────────────────────────────────────────
// A minimal seekable ZIP writer: each local header is written with zero
// sizes, the body streamed after it, then the header patched — the way
// Excel's writer leaves sizes in the header with no data descriptor.

async function generate(file, rows) {
  const fh = await open(file, "w")
  const records = []
  let offset = 0
  const write = async (bytes) => {
    await fh.write(bytes, 0, bytes.length, offset)
    offset += bytes.length
  }
  const enc = new TextEncoder()

  async function entry(name, chunks) {
    const nameBytes = enc.encode(name)
    const headerAt = offset
    await write(new Uint8Array(30 + nameBytes.length))
    let crc = 0
    let usize = 0
    let csize = 0
    // One raw-DEFLATE stream across all chunks, flushed per chunk.
    const z = createDeflateRaw({ level: 6 })
    const out = []
    z.on("data", (d) => out.push(d))
    const flush = () => new Promise((res) => z.flush(constants.Z_SYNC_FLUSH, res))
    for await (const text of chunks) {
      const bytes = enc.encode(text)
      crc = crc32(bytes, crc)
      usize += bytes.length
      z.write(bytes)
      await flush()
      for (const d of out.splice(0)) {
        csize += d.length
        await write(d)
      }
    }
    await new Promise((res) => {
      z.on("end", res)
      z.end()
    })
    for (const d of out.splice(0)) {
      csize += d.length
      await write(d)
    }
    if (usize >= 2 ** 32 || csize >= 2 ** 32) throw new Error("bench generator does not do ZIP64")
    const h = new DataView(new ArrayBuffer(30))
    h.setUint32(0, 0x04034b50, true)
    h.setUint16(4, 20, true)
    h.setUint16(8, 8, true)
    h.setUint32(14, crc, true)
    h.setUint32(18, csize, true)
    h.setUint32(22, usize, true)
    h.setUint16(26, nameBytes.length, true)
    await fh.write(new Uint8Array(h.buffer), 0, 30, headerAt)
    await fh.write(nameBytes, 0, nameBytes.length, headerAt + 30)
    records.push({ nameBytes, crc, csize, usize, headerAt })
  }

  async function* once(text) {
    yield text
  }
  async function* sheet() {
    let x = 42
    const next = () => {
      x ^= x << 13
      x ^= x >>> 17
      x ^= x << 5
      return ((x >>> 0) / 2 ** 32) * 4.2
    }
    yield `<?xml version="1.0"?><worksheet ${NS} ${R}><sheetData>`
    let buf = ""
    for (let r = 1; r <= rows; r++) {
      let cells = ""
      for (const c of "CDEFGHIJ") cells += `<c r="${c}${r}"><v>${next().toFixed(6)}</v></c>`
      buf +=
        `<row r="${r}"><c r="A${r}" t="s"><v>${r % 3}</v></c>` +
        `<c r="B${r}" s="1"><v>${(45000 + r / 86400).toFixed(8)}</v></c>${cells}</row>`
      if (r % 5000 === 0) {
        yield buf
        buf = ""
      }
    }
    yield `${buf}</sheetData></worksheet>`
  }

  await entry(
    "[Content_Types].xml",
    once(
      `<?xml version="1.0"?><Types xmlns="${PKG}/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/></Types>`,
    ),
  )
  await entry(
    "_rels/.rels",
    once(
      `<?xml version="1.0"?><Relationships xmlns="${PKG}/relationships">` +
        `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    ),
  )
  await entry(
    "xl/workbook.xml",
    once(
      `<?xml version="1.0"?><workbook ${NS} ${R}><sheets>` +
        `<sheet name="Record" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
  )
  await entry(
    "xl/_rels/workbook.xml.rels",
    once(
      `<?xml version="1.0"?><Relationships xmlns="${PKG}/relationships">` +
        `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>` +
        `<Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/>` +
        `<Relationship Id="rId3" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>` +
        `</Relationships>`,
    ),
  )
  // Excel's order: the sheet, then styles and shared strings after it.
  await entry("xl/worksheets/sheet1.xml", sheet())
  await entry(
    "xl/styles.xml",
    once(
      `<?xml version="1.0"?><styleSheet ${NS}><cellXfs count="2">` +
        `<xf numFmtId="0"/><xf numFmtId="22" applyNumberFormat="1"/></cellXfs></styleSheet>`,
    ),
  )
  await entry(
    "xl/sharedStrings.xml",
    once(
      `<?xml version="1.0"?><sst ${NS} count="3" uniqueCount="3">` +
        `<si><t>charge</t></si><si><t>rest</t></si><si><t>discharge</t></si></sst>`,
    ),
  )

  const cdStart = offset
  for (const rec of records) {
    const c = new DataView(new ArrayBuffer(46))
    c.setUint32(0, 0x02014b50, true)
    c.setUint16(4, 20, true)
    c.setUint16(6, 20, true)
    c.setUint16(10, 8, true)
    c.setUint32(16, rec.crc, true)
    c.setUint32(20, rec.csize, true)
    c.setUint32(24, rec.usize, true)
    c.setUint16(28, rec.nameBytes.length, true)
    c.setUint32(42, rec.headerAt, true)
    await write(new Uint8Array(c.buffer))
    await write(rec.nameBytes)
  }
  const e = new DataView(new ArrayBuffer(22))
  e.setUint32(0, 0x06054b50, true)
  e.setUint16(8, records.length, true)
  e.setUint16(10, records.length, true)
  e.setUint32(12, offset - cdStart, true)
  e.setUint32(16, cdStart, true)
  await write(new Uint8Array(e.buffer))
  await fh.close()
}
