// One adapter inventory for cell regressions and independent workbook saves.
import type { CellInput } from "../../src/_types"
import { writeXlsx } from "../../src/xlsx/writer"
import { readXlsx } from "../../src/xlsx/reader"
import {
  XlsxStreamWriter,
  writeXlsxStream,
  writeXlsxStreamSheets,
} from "../../src/xlsx/stream-writer"
import { writeOds } from "../../src/ods/writer"
import { readOds } from "../../src/ods/reader"
import { OdsStreamWriter } from "../../src/ods/incremental-writer"
import { writeOdsStream } from "../../src/ods/stream-writer"

export const drain = async (stream: ReadableStream<Uint8Array>): Promise<Uint8Array> =>
  new Uint8Array(await new Response(stream).arrayBuffer())

export const writerPaths = [
  {
    name: "buffered XLSX",
    write: (rows: CellInput[][]) => writeXlsx({ sheets: [{ name: "S", rows }] }),
    read: readXlsx,
  },
  {
    name: "incremental XLSX",
    write: async (rows: CellInput[][]) => {
      const writer = new XlsxStreamWriter({ name: "S" })
      for (const row of rows) writer.addRow(row)
      return writer.finish()
    },
    read: readXlsx,
  },
  {
    name: "streamed XLSX",
    write: (rows: CellInput[][]) => drain(writeXlsxStream(rows, { name: "S" })),
    read: readXlsx,
  },
  {
    name: "multi-sheet XLSX",
    write: (rows: CellInput[][]) => drain(writeXlsxStreamSheets([{ name: "S", rows }])),
    read: readXlsx,
  },
  {
    name: "buffered ODS",
    write: (rows: CellInput[][]) => writeOds({ sheets: [{ name: "S", rows }] }),
    read: readOds,
  },
  {
    name: "incremental ODS",
    write: async (rows: CellInput[][]) => {
      const writer = new OdsStreamWriter({ name: "S" })
      for (const row of rows) writer.addRow(row)
      return writer.finish()
    },
    read: readOds,
  },
  {
    name: "streamed ODS",
    write: (rows: CellInput[][]) => drain(writeOdsStream(rows, { name: "S" })),
    read: readOds,
  },
]
