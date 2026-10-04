# bench

Reproducible measurements for the claims in the README, and for anyone
changing a hot path.

```bash
bun run build     # bench runs against dist/, which is what users get
bun run bench     # every scenario, one process each
```

Or one at a time:

```bash
node bench/write.mjs writeXlsxStream 300000
node bench/read.mjs streamXlsxRows high-cardinality
```

## Why one process per measurement

`process.resourceUsage().maxRSS` is a high-water mark for the **whole
process**. Measure two things in one run and the second inherits the
first's peak — which is how a streaming writer can appear to use 800 MB.
`bun run bench` forks a child per scenario for that reason, and the numbers
below were taken that way.

## The scenarios

`write.mjs` — 12 columns of mixed text, number and date, through each of
the three XLSX write paths.

`read.mjs` — the same sheet read back four ways, against **two** fixtures
that differ in one line:

```js
c % 3 === 0 ? `text ${i}-${c}`        // high cardinality: ~400k distinct strings
c % 3 === 0 ? WORDS[(i + c) % 10]     // low cardinality: 10 distinct strings
```

That one line is the difference between 90 MB and 556 MB in
`streamXlsxRows`, which is the thing worth knowing about the streaming
reader: its peak tracks the number of **distinct strings**, not the number
of rows, because `xl/sharedStrings.xml` has to be read up front.

## Numbers to compare against

Taken on Linux, Node 24, 100,000 rows × 12 columns. Yours will differ;
the ratios are the point.

| write              |     time |  peak RSS |
| ------------------ | -------: | --------: |
| `writeXlsxStream`  | 2,230 ms | **92 MB** |
| `XlsxStreamWriter` | 2,877 ms |    604 MB |
| `writeXlsx`        | 3,186 ms |    758 MB |

At 300,000 rows `writeXlsxStream` grows to 126 MB and `XlsxStreamWriter`
to 1,515 MB — flat against linear, which is the promise the README makes.

| read                         |      high cardinality |      low cardinality |
| ---------------------------- | --------------------: | -------------------: |
| `readXlsx`                   |     2,721 ms / 662 MB |    1,896 ms / 392 MB |
| `readXlsx` + `maxRows: 1000` |     1,953 ms / 656 MB |    1,150 ms / 221 MB |
| `streamXlsxRows`             | 2,446 ms / **556 MB** | 1,592 ms / **90 MB** |

Two things fall out of that table and are worth keeping in view when
changing the readers:

- `streamXlsxRows` is constant-memory in rows and **linear in distinct
  strings**.
- `maxRows` bounds the output, not the work: on the high-cardinality
  fixture it saved 28% of the time and 1% of the memory.

## Numeric cell storage

```bash
bun run build
node --expose-gc bench/cell-store.mjs flat wide
node --expose-gc bench/cell-store.mjs blocks wide
node --expose-gc bench/cell-store.mjs flat narrow
node --expose-gc bench/cell-store.mjs blocks narrow
```

`blocks` calls the shipped `hucre/cell` helpers, including coordinate
validation and cell counts. `flat` reproduces the old string-keyed Map.
Both create identical cell objects and verify their counts and checksums.
This measures metadata storage, not XLSX throughput or total workbook memory.

macOS arm64, Node 24.21.0, three fresh processes per mode and shape:

| Shape                    | Store  |  Build (ms) | Lookup (ms) | Retained heap (MB) | Peak RSS (MB) |
| ------------------------ | ------ | ----------: | ----------: | -----------------: | ------------: |
| 20,000 × 32 (640k cells) | flat   | 168.1–185.2 |  87.6–106.4 |          66.6–67.2 |   174.9–175.6 |
| 20,000 × 32              | blocks |   31.6–33.3 |   11.8–12.3 |          41.1–41.3 |   128.3–129.2 |
| 240,000 × 1 (240k cells) | flat   |   42.9–45.5 |   26.6–28.6 |          21.2–21.4 |   107.4–107.6 |
| 240,000 × 1              | blocks |   15.0–16.9 |     6.2–7.2 |          15.3–15.7 |     82.7–83.2 |

Retained heap is measured after GC while the store is still reachable;
peak RSS includes runtime and module loading. Results vary by engine,
machine and data shape. These samples do not allocate 2^24 cells: the
capacity regression scales the per-Map threshold down, and the numeric
block bounds establish why no metadata Map reaches V8's real threshold.
