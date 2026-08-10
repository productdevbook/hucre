# Security

hucre parses files it did not write. That is the whole point of it, and it
is also where the risk is: a spreadsheet is an archive of XML, and both
layers have well-known ways to make a parser allocate forever.

## Reporting a vulnerability

Please **do not open a public issue** for a security problem.

Use GitHub's private reporting — [Security → Report a
vulnerability](https://github.com/productdevbook/hucre/security/advisories/new)
— or email <mehmet.k.hob@gmail.com>.

Include the input that triggers it if you can. A file, or the few lines of
markup or XML that reproduce it, is worth more than a description.

You should get an acknowledgement within a few days. Fixes go out as a
patch release, and the advisory is published once one is available.

## What is in scope

- Reading any supported format — XLSX, XLSM, XLSB, XLS, ODS, CSV, JSON,
  NDJSON, XML, HTML — from untrusted input.
- The ZIP and OLE2/CFB container readers.
- The XML parser and the HTML importer.
- Encryption and decryption (ECMA-376 Agile).
- Anything that makes hucre allocate without bound, hang, or read outside
  the input it was given.

## What is not

- Writing a file with content a caller supplied. `writeXlsx` escapes what
  it writes, but hucre is a serialiser: if you put a formula in a cell, it
  writes a formula. Deciding what is safe to put in a spreadsheet is the
  caller's, which is why `escapeFormulae` exists on the CSV writer and is
  opt-in.
- Denial of service through input the caller chose to accept above the
  documented bounds — see below.
- Anything requiring the attacker to control your own source code or
  dependencies.

## The bounds hucre already enforces

`src/limits.ts` is the list, and each entry says what it is defending
against:

| bound                    | default                                 | what it stops                                                                         |
| ------------------------ | --------------------------------------- | ------------------------------------------------------------------------------------- |
| `MAX_DECOMPRESSED_BYTES` | 2 GiB per entry                         | ZIP bombs — a small entry claiming a small compressed size and expanding to gigabytes |
| `MAX_INPUT_BYTES`        | 1 GiB, overridable with `maxInputBytes` | a `ReadableStream` that never ends                                                    |
| `MAX_TOTAL_CELLS`        | 20,000,000                              | two legal cells at opposite corners describing a 1.7e10-slot rectangle                |
| `MAX_REPEAT_COUNT`       | 100,000                                 | ODS `text:c="900000000"`, a gigabyte of spaces in one cell                            |
| `MAX_SPAN_CELLS`         | 1,000,000                               | HTML `rowspan` × `colspan` bombs                                                      |
| `MAX_SPIN_COUNT`         | 10,000,000                              | a hostile encrypted file pinning a CPU in key derivation                              |

Two structural properties are worth naming because they remove whole
classes of attack rather than bounding them:

- **The XML parser expands no DTD.** It handles the five predefined
  entities and numeric character references, and nothing else, so
  entity-expansion attacks — billion laughs, quadratic blowup — have
  nothing to work with. There is no external-entity resolution either, so
  XXE is not reachable.
- **The library reads no filesystem and opens no network.** There are no
  `node:` imports in `src/` outside the CLI, and `tsconfig.json` sets
  `"types": []` so reaching for one is a compile error. A malicious file
  cannot make hucre fetch anything.

If you find input that gets past a bound, or a path where one is not
applied, that is a vulnerability and worth reporting.

## Supported versions

The latest minor of the current major receives security fixes.
