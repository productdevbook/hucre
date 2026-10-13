const original = console.log.bind(console)
console.log = (...args) => {
  postMessage({ log: args.join(" ") })
  original(...args)
}
try {
  await import("/scripts/smoke.mjs")
  const h = await import("/dist/index.mjs")
  self.onmessage = async ({ data }) => {
    try {
      const value = h.getCell(data.workbook.sheets[0].cells, 127, 2)?.value
      if (!(value instanceof Date) || value.toISOString() !== "2026-10-05T00:00:00.000Z")
        throw new Error("incoming workbook clone failed")
      h.insertRows(data.workbook.sheets[0], 0, 1)
      const result = await h.readXlsx(data.encrypted, { password: "local-test-only" })
      if (result.sheets[0].rows[1][0] !== "Ada") throw new Error("worker crypto failed")
      postMessage({ done: true, pass: true, workbook: data.workbook })
    } catch (error) {
      postMessage({ error: error.stack })
    }
  }
  postMessage({ ready: true })
} catch (error) {
  postMessage({ error: error.stack })
}
