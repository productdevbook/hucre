import { ZipReader } from "../../src/zip/reader"
import { ZipWriter } from "../../src/zip/writer"

const encoder = new TextEncoder()
const decoder = new TextDecoder()

/** Mutate one XML part while preserving all other bytes. For unit regressions;
 * integration tests use committed files from independent producers. */
export async function patchXml(
  bytes: Uint8Array,
  path: string,
  edit: (xml: string) => string,
): Promise<Uint8Array> {
  const zip = new ZipWriter()
  for (const [name, data] of await new ZipReader(bytes).extractAll()) {
    zip.add(name, name === path ? encoder.encode(edit(decoder.decode(data))) : data)
  }
  return zip.build()
}

/** Minimal independent OOXML scaffold. Raw cell XML stays visible in tests. */
export async function xlsxWithCells(cellsXml: string, sharedString?: string): Promise<Uint8Array> {
  const zip = new ZipWriter()
  const parts: Record<string, string> = {
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>${sharedString === undefined ? "" : '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'}</Relationships>`,
    "xl/worksheets/sheet1.xml": `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">${cellsXml}</row></sheetData></worksheet>`,
  }
  if (sharedString !== undefined) {
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
      "</Types>",
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
    )
    parts["xl/sharedStrings.xml"] =
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>${sharedString}</t></si></sst>`
  }
  for (const [name, xml] of Object.entries(parts)) zip.add(name, encoder.encode(xml))
  return zip.build()
}
