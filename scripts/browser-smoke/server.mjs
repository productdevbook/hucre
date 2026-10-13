// Serve only the browser harness and built public modules on loopback.
import { createServer } from "node:http"
import { readFile } from "node:fs/promises"
import { resolve, sep } from "node:path"
import { fileURLToPath } from "node:url"
const root = fileURLToPath(new URL("../../", import.meta.url))
const harness = fileURLToPath(new URL(".", import.meta.url))
const server = createServer(async (request, response) => {
  const path = new URL(request.url, "http://127.0.0.1").pathname
  let file
  if (path === "/") file = resolve(harness, "index.html")
  else if (path === "/worker.mjs") file = resolve(harness, "worker.mjs")
  else if (path === "/scripts/smoke.mjs") file = resolve(root, "scripts/smoke.mjs")
  else if (path.startsWith("/dist/") && path.endsWith(".mjs")) {
    const candidate = resolve(root, "." + path)
    if (candidate.startsWith(resolve(root, "dist") + sep)) file = candidate
  }
  if (!file) {
    response.writeHead(404)
    response.end("Not found")
    return
  }
  try {
    const bytes = await readFile(file)
    response.writeHead(200, {
      "Content-Type": file.endsWith(".html")
        ? "text/html; charset=utf-8"
        : "text/javascript; charset=utf-8",
      "Cache-Control": "no-store",
    })
    response.end(bytes)
  } catch (error) {
    response.writeHead(500)
    response.end(String(error))
  }
})
server.listen(0, "127.0.0.1", () => console.log("http://127.0.0.1:" + server.address().port + "/"))
