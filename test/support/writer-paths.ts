// One adapter inventory for cell regressions and independent workbook saves.
import type { CellInput, ColumnDef } from "../../src/_types"
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
    write: (rows: CellInput[][], columns?: ColumnDef[]) =>
      writeXlsx({ sheets: [{ name: "S", rows, columns }] }),
    read: readXlsx,
  },
  {
    name: "incremental XLSX",
    write: async (rows: CellInput[][], columns?: ColumnDef[]) => {
      const writer = new XlsxStreamWriter({ name: "S", columns })
      for (const row of rows) writer.addRow(row)
      return writer.finish()
    },
    read: readXlsx,
  },
  {
    name: "streamed XLSX",
    write: (rows: CellInput[][], columns?: ColumnDef[]) =>
      drain(writeXlsxStream(rows, { name: "S", columns })),
    read: readXlsx,
  },
  {
    name: "multi-sheet XLSX",
    write: (rows: CellInput[][], columns?: ColumnDef[]) =>
      drain(writeXlsxStreamSheets([{ name: "S", rows, columns }])),
    read: readXlsx,
  },
  {
    name: "buffered ODS",
    write: (rows: CellInput[][], columns?: ColumnDef[]) =>
      writeOds({ sheets: [{ name: "S", rows, columns }] }),
    read: readOds,
  },
  {
    name: "incremental ODS",
    write: async (rows: CellInput[][], columns?: ColumnDef[]) => {
      const writer = new OdsStreamWriter({ name: "S", columns })
      for (const row of rows) writer.addRow(row)
      return writer.finish()
    },
    read: readOds,
  },
  {
    name: "streamed ODS",
    write: (rows: CellInput[][], columns?: ColumnDef[]) =>
      drain(writeOdsStream(rows, { name: "S", columns })),
    read: readOds,
  },
]

export const objectWriterPaths = [
  {
    name: "buffered XLSX",
    write: (data: Record<string, CellInput>[], columns: ColumnDef[]) =>
      writeXlsx({ sheets: [{ name: "S", data, columns }] }),
    read: readXlsx,
  },
  {
    name: "buffered ODS",
    write: (data: Record<string, CellInput>[], columns: ColumnDef[]) =>
      writeOds({ sheets: [{ name: "S", data, columns }] }),
    read: readOds,
  },
  {
    name: "incremental XLSX",
    write: async (data: Record<string, CellInput>[], columns: ColumnDef[]) => {
      const w = new XlsxStreamWriter({ name: "S", columns })
      for (const row of data) w.addObject(row)
      return w.finish()
    },
    read: readXlsx,
  },
  {
    name: "incremental ODS",
    write: async (data: Record<string, CellInput>[], columns: ColumnDef[]) => {
      const w = new OdsStreamWriter({ name: "S", columns })
      for (const row of data) w.addObject(row)
      return w.finish()
    },
    read: readOds,
  },
  {
    name: "streamed XLSX",
    write: (data: Record<string, CellInput>[], columns: ColumnDef[]) =>
      drain(writeXlsxStream(data, { name: "S", columns })),
    read: readXlsx,
  },
  {
    name: "multi-sheet XLSX",
    write: (data: Record<string, CellInput>[], columns: ColumnDef[]) =>
      drain(writeXlsxStreamSheets([{ name: "S", rows: data, columns }])),
    read: readXlsx,
  },
]
