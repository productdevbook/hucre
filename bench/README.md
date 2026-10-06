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

## ODS bounding-box enforcement

```bash
bun run build
task_dir=$(mktemp -d)
git archive 0c0fc04 src | tar -x -C "$task_dir"
bunx rolldown "$task_dir/src/ods/reader.ts" --file "$task_dir/before.mjs" --format esm
bunx rolldown src/ods/reader.ts --file "$task_dir/after.mjs" --format esm
node bench/ods-bounds.mjs "$task_dir/before.mjs"
node bench/ods-bounds.mjs "$task_dir/after.mjs"
```

Run each command in a fresh process for every sample. Both readers are
bundled the same way; the input is raw ODF XML packaged with the ZIP
writer, independent of the spreadsheet writer. `node bench/ods-bounds.mjs`
without a module argument tests the built package.

The 475-byte archive contains a 2,048-column first row followed by a
one-column row repeated 10,000 times: 20,482,048 slots after padding.
The default bound is 20,000,000. macOS arm64, Node 24.21.0, three fresh
processes per reader:

| Reader             | Result                               | Read/rejection time (ms) | Peak RSS (MB) |
| ------------------ | ------------------------------------ | -----------------------: | ------------: |
| V2 `0c0fc04`       | Incorrectly accepts 20,482,048 cells |              203.3–209.0 |   407.0–408.3 |
| Bounding-box check | Rejects before row expansion         |                      1.4 |     49.5–49.8 |

This measures one rejected amplification input, not general ODS throughput.
Peak RSS includes runtime, module loading and ZIP preparation. Results
vary with engine, machine and input shape.

## Shared authoring row resolution

`writeXlsxData` and `writeXlsxDataLinks` in `bench/write.mjs` exercise
50,000 object records × 12 columns. Four columns each contain strings,
numbers and dates; all columns are bold and numeric columns use `0.00`.
The links case makes the first value in every record a rich hyperlink.

```sh
bun run build
node bench/write.mjs writeXlsxData 50000
node bench/write.mjs writeXlsxDataLinks 50000
HUCRE_BENCH_ENTRY=/absolute/path/to/baseline/dist/index.mjs node bench/write.mjs writeXlsxData 50000
```

Build the baseline with the same dependency versions; `HUCRE_BENCH_ENTRY`
selects its built entry point while retaining the same scenario generator.
Run each sample as a fresh process. macOS arm64, Node 24.21.0, three samples
per scenario/tree, alternating baseline and current:

| Scenario               | V2 `380a03c` time / peak RSS | Shared boundary time / peak RSS |
| ---------------------- | ---------------------------: | ------------------------------: |
| Object data            |      596–610 ms / 505–506 MB |         512–526 ms / 488–491 MB |
| Object data with links |      724–746 ms / 581–582 MB |         696–737 ms / 529–546 MB |

The clock covers writing after input construction; peak RSS includes
module loading, input and output allocation. These are two formatted
object-data scenarios on one machine, not a claim about every writer or
workbook. The independent invoice and regression suites verify values,
headers, caches and styles separately from the measurements.

## Shared streaming cell resolution

The existing `writeXlsxStream` and `XlsxStreamWriter` scenarios emit
100,000 rows × 12 columns of strings, numbers and dates. They contain no
links/comments, so the comparison checks the ordinary-cell path after
metadata support was added; it does not measure metadata-heavy output.

```sh
bun run build
node bench/write.mjs writeXlsxStream 100000
node bench/write.mjs XlsxStreamWriter 100000
HUCRE_BENCH_ENTRY=/absolute/path/to/baseline/dist/index.mjs node bench/write.mjs writeXlsxStream 100000
HUCRE_BENCH_ENTRY=/absolute/path/to/baseline/dist/index.mjs node bench/write.mjs XlsxStreamWriter 100000
```

Build V2 `adcf5e5` with the same dependencies for the baseline. Use one
scenario per fresh process, with three samples per scenario/tree and
alternate baseline/current. macOS arm64, Node 24.21.0:

| Scenario           | V2 `adcf5e5` time / peak RSS | Shared cell model time / peak RSS |
| ------------------ | ---------------------------: | --------------------------------: |
| `writeXlsxStream`  |        755–807 ms / 98–99 MB |            755–771 ms / 97–101 MB |
| `XlsxStreamWriter` |  1,018–1,064 ms / 602–603 MB |       1,004–1,043 ms / 601–603 MB |

These intervals overlap; they do not establish a throughput improvement.
Output byte counts match per scenario (5,330,764 streamed, 6,230,593
incremental). The streaming sink discards bytes; the class retains its
fragments/output until finish. Peak RSS includes runtime and module loading.
Distinct styles and shared strings are retained, and links/comments now
retain metadata until a physical sheet's related parts are emitted.
