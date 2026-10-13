import { readFileSync } from "node:fs"

const root = new URL("../fixtures/", import.meta.url)

/** Committed external-producer inputs, shared by integration suites. */
export function fixtureBytes(name: string): Uint8Array {
  return new Uint8Array(readFileSync(new URL(name, root)))
}

export function fixtureJson<T>(name: string): T {
  return JSON.parse(readFileSync(new URL(name, root), "utf8")) as T
}
