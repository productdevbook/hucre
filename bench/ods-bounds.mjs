// Run one fresh process per sample. An optional module path compares a
// historical reader built with the same bundler, without changing the checkout.
import { resolve } from "node:path"
import { pathToFileURL } from "node:url"
import { ZipWriter } from "../dist/zip/writer.mjs"

const { readOds } = await import(
  process.argv[2] ? pathToFileURL(resolve(process.argv[2])).href : "../dist/ods.mjs"
)
const zip = new ZipWriter()
const enc = new TextEncoder()
zip.add("mimetype", enc.encode("application/vnd.oasis.opendocument.spreadsheet"), {
  compress: false,
})
zip.add(
  "content.xml",
  enc.encode(`<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body><office:spreadsheet><table:table table:name="S">
    <table:table-row><table:table-cell office:value-type="float" office:value="1" table:number-columns-repeated="2048"/></table:table-row>
    <table:table-row table:number-rows-repeated="10000"><table:table-cell office:value-type="float" office:value="2"/></table:table-row>
  </table:table></office:spreadsheet></office:body></office:document-content>`),
)
const bytes = await zip.build()
const start = performance.now()
let result
try {
  const workbook = await readOds(bytes)
  result = {
    outcome: "accepted",
    cells: workbook.sheets[0].rows.length * workbook.sheets[0].rows[0].length,
  }
} catch (error) {
  if (error.name !== "ParseError") throw error
  result = { outcome: "rejected", message: error.message }
}
console.log(
  JSON.stringify({
    node: process.version,
    inputBytes: bytes.length,
    limit: 20_000_000,
    ...result,
    elapsedMs: +(performance.now() - start).toFixed(1),
    peakMB: +(process.resourceUsage().maxRSS / 1024).toFixed(1),
  }),
)
