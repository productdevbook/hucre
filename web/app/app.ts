import {
  readXlsx,
  writeXlsx,
  parseCsv,
  parseCsvObjects,
  detectDelimiter,
  writeCsv,
  validateWithSchema,
  streamXlsxRows,
  XlsxStreamWriter,
  readOds,
  writeOds,
  toHtml,
  toMarkdown,
  toJson,
  formatValue,
  parseJson,
  parseNdjson,
  writeJson,
  writeNdjson,
  readXml,
  writeXml,
  writeXlsxObjects,
  getCharts,
  cloneChart,
  chartKindToWriteKind,
  addChart,
} from "hucre";
import type {
  CellValue,
  WriteSheet,
  SchemaDefinition,
  Sheet,
  Chart,
  ChartKind,
  WriteChartKind,
  CloneChartOptions,
  CloneChartSeriesOverride,
  SheetChart,
  Workbook,
} from "hucre";

// ── Toast ─────────────────────────────────────────────────────────

let toastTimer: ReturnType<typeof setTimeout>;
function toast(msg: string) {
  const el = document.getElementById("toast")!;
  el.textContent = msg;
  el.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove("show"), 2000);
}

// ── Tabs ──────────────────────────────────────────────────────────

function setupTabs() {
  const tabs = document.querySelectorAll<HTMLButtonElement>(".tab");
  const panels = document.querySelectorAll<HTMLElement>(".panel");
  tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      tabs.forEach((t) => t.classList.remove("active"));
      panels.forEach((p) => p.classList.remove("active"));
      tab.classList.add("active");
      const target = tab.dataset["tab"];
      document.querySelector(`[data-panel="${target}"]`)?.classList.add("active");
    });
  });
}

// ── Helpers ───────────────────────────────────────────────────────

function $(id: string) {
  return document.getElementById(id)!;
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function cellClass(v: CellValue): string {
  if (v === null || v === undefined) return "null";
  if (typeof v === "number") return "num";
  if (typeof v === "boolean") return "bool";
  if (v instanceof Date) return "date";
  return "";
}

function cellDisplay(v: CellValue): string {
  if (v === null || v === undefined) return "null";
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  return String(v);
}

function renderTable(headers: string[], rows: CellValue[][]): string {
  let html = "<table><thead><tr>";
  for (const h of headers) html += `<th>${escapeHtml(h)}</th>`;
  html += "</tr></thead><tbody>";
  for (const row of rows.slice(0, 200)) {
    html += "<tr>";
    for (const cell of row) {
      html += `<td class="${cellClass(cell)}">${escapeHtml(cellDisplay(cell))}</td>`;
    }
    html += "</tr>";
  }
  html += "</tbody></table>";
  if (rows.length > 200) {
    html += `<div class="meta">Showing 200 of ${rows.length} rows</div>`;
  }
  return html;
}

// ── Read XLSX ─────────────────────────────────────────────────────

let lastReadResult: { headers: string[]; rows: CellValue[][] } | null = null;

async function handleReadFile(file: File) {
  const output = $("read-output");
  const stats = $("read-stats");

  try {
    output.innerHTML = '<p style="color:var(--text-dim);text-align:center">Parsing...</p>';
    const buffer = await file.arrayBuffer();
    const data = new Uint8Array(buffer);

    const headerRow = parseInt(($("read-header") as HTMLInputElement).value) || 0;
    const skipHidden = ($("read-skip-hidden") as HTMLInputElement).checked;
    const wb = await readXlsx(data, {
      readStyles: ($("read-styles") as HTMLInputElement).checked,
      ...(skipHidden
        ? {
            sheets: (info: { hidden?: boolean; veryHidden?: boolean }) =>
              !info.hidden && !info.veryHidden,
          }
        : {}),
    });

    if (wb.sheets.length === 0) {
      output.innerHTML = '<p class="error">No sheets found</p>';
      return;
    }

    const sheet = wb.sheets[0];
    const rows = sheet.rows;

    // Stats
    stats.hidden = false;
    stats.innerHTML = `
      <div class="stat"><div class="value">${wb.sheets.length}</div><div class="label">Sheets</div></div>
      <div class="stat"><div class="value">${rows.length}</div><div class="label">Rows</div></div>
      <div class="stat"><div class="value">${rows[0]?.length || 0}</div><div class="label">Columns</div></div>
      <div class="stat"><div class="value">${(data.byteLength / 1024).toFixed(1)} KB</div><div class="label">File Size</div></div>
    `;

    // Build headers and data
    let headers: string[];
    let dataRows: CellValue[][];
    if (headerRow > 0 && rows.length > 0) {
      headers = (rows[headerRow - 1] || []).map((v, i) => (v != null ? String(v) : `Col ${i + 1}`));
      dataRows = rows.slice(headerRow);
    } else {
      headers = (rows[0] || []).map((_, i) => `Col ${i + 1}`);
      dataRows = rows;
    }

    lastReadResult = { headers, rows: dataRows };
    output.innerHTML = renderTable(headers, dataRows);

    ($("read-copy") as HTMLButtonElement).disabled = false;
    ($("read-download") as HTMLButtonElement).disabled = false;
  } catch (e: unknown) {
    output.innerHTML = `<p class="error">${escapeHtml(String(e))}</p>`;
    stats.hidden = true;
  }
}

function setupRead() {
  const drop = $("read-drop");
  const fileInput = $("read-file") as HTMLInputElement;

  drop.addEventListener("click", () => fileInput.click());
  drop.addEventListener("dragover", (e) => {
    e.preventDefault();
    drop.classList.add("drag-over");
  });
  drop.addEventListener("dragleave", () => drop.classList.remove("drag-over"));
  drop.addEventListener("drop", (e) => {
    e.preventDefault();
    drop.classList.remove("drag-over");
    const file = (e as DragEvent).dataTransfer?.files[0];
    if (file) handleReadFile(file);
  });
  fileInput.addEventListener("change", () => {
    if (fileInput.files?.[0]) handleReadFile(fileInput.files[0]);
  });

  $("read-copy").addEventListener("click", () => {
    if (!lastReadResult) return;
    const json = lastReadResult.rows.map((row) => {
      const obj: Record<string, CellValue> = {};
      lastReadResult!.headers.forEach((h, i) => {
        obj[h] = row[i] ?? null;
      });
      return obj;
    });
    navigator.clipboard.writeText(JSON.stringify(json, null, 2));
    toast("JSON copied to clipboard");
  });

  $("read-download").addEventListener("click", () => {
    if (!lastReadResult) return;
    const csv = writeCsv([lastReadResult.headers, ...lastReadResult.rows], { bom: true });
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "export.csv";
    a.click();
    URL.revokeObjectURL(url);
    toast("CSV downloaded");
  });
}

// ── Write XLSX ────────────────────────────────────────────────────

let lastXlsxBlob: Blob | null = null;

function setupWrite() {
  $("write-generate").addEventListener("click", async () => {
    const output = $("write-output");
    try {
      const rawData = JSON.parse(($("write-data") as HTMLTextAreaElement).value);
      const rawCols = JSON.parse(($("write-cols") as HTMLTextAreaElement).value);
      const sheetName = ($("write-name") as HTMLInputElement).value || "Sheet1";
      const freezeRows = parseInt(($("write-freeze") as HTMLInputElement).value) || 0;
      const autoFilter = ($("write-autofilter") as HTMLInputElement).checked;
      const autoWidth = ($("write-autowidth") as HTMLInputElement).checked;
      const checkboxCols = ($("write-checkbox-cols") as HTMLInputElement).value
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean);

      const columns: Record<string, { header?: string; width?: number; numFmt?: string }> = rawCols;
      const columnKeys = Object.keys(columns);

      // Build a cells Map flagging boolean cells in checkboxCols as Excel 2024 checkboxes.
      let cellsMap: Map<string, { value: CellValue; type: "boolean"; checkbox: true }> | undefined;
      if (checkboxCols.length > 0) {
        cellsMap = new Map();
        // header is row 0; data rows start at row 1
        for (let r = 0; r < rawData.length; r++) {
          for (const cbKey of checkboxCols) {
            const colIdx = columnKeys.indexOf(cbKey);
            if (colIdx === -1) continue;
            const v = (rawData[r] as Record<string, CellValue>)[cbKey];
            if (typeof v !== "boolean") continue;
            cellsMap.set(`${r + 1},${colIdx}`, {
              value: v,
              type: "boolean",
              checkbox: true,
            });
          }
        }
      }

      const sheet: WriteSheet = {
        name: sheetName,
        data: rawData,
        columns: Object.entries(columns).map(([key, col]) => ({
          key,
          header: col.header || key,
          width: col.width,
          numFmt: col.numFmt,
          autoWidth: autoWidth && !col.width,
        })),
        cells: cellsMap,
        freezePane: freezeRows > 0 ? { rows: freezeRows } : undefined,
        autoFilter: autoFilter
          ? {
              range: `A1:${String.fromCharCode(64 + columnKeys.length)}${rawData.length + 1}`,
            }
          : undefined,
      };

      const result = await writeXlsx({ sheets: [sheet] });
      lastXlsxBlob = new Blob([result], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      // Show preview as table
      const headers = sheet.columns?.map((c) => c.header || c.key || "") || [];
      const rows: CellValue[][] = rawData.map((obj: Record<string, CellValue>) =>
        Object.keys(columns).map((k) => obj[k] ?? null),
      );

      output.innerHTML = renderTable(headers, rows);
      output.innerHTML += `<div class="meta">Generated: ${(result.byteLength / 1024).toFixed(1)} KB XLSX</div>`;

      ($("write-download") as HTMLButtonElement).disabled = false;
    } catch (e: unknown) {
      output.innerHTML = `<p class="error">${escapeHtml(String(e))}</p>`;
    }
  });

  $("write-download").addEventListener("click", () => {
    if (!lastXlsxBlob) return;
    const url = URL.createObjectURL(lastXlsxBlob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${($("write-name") as HTMLInputElement).value || "sheet"}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
    toast("XLSX downloaded");
  });
}

// ── CSV ───────────────────────────────────────────────────────────

let lastCsvParsed: { headers: string[]; rows: CellValue[][] } | null = null;

function setupCsv() {
  $("csv-parse").addEventListener("click", () => {
    const output = $("csv-output");
    try {
      const input = ($("csv-input") as HTMLTextAreaElement).value;
      const delimSel = ($("csv-delim") as HTMLSelectElement).value;
      const hasHeader = ($("csv-header") as HTMLInputElement).checked;
      const typeInference = ($("csv-types") as HTMLInputElement).checked;
      const skipEmptyRows = ($("csv-skip-empty") as HTMLInputElement).checked;

      const delimiter = delimSel === "auto" ? undefined : delimSel;
      const detected = detectDelimiter(input);

      if (hasHeader) {
        const result = parseCsvObjects(input, {
          header: true,
          delimiter,
          typeInference,
          skipEmptyRows,
        });
        const headers = result.headers;
        const rows = result.data.map((obj) =>
          headers.map((h) => (obj as Record<string, CellValue>)[h] ?? null),
        );
        lastCsvParsed = { headers, rows };
        output.innerHTML = renderTable(headers, rows);
      } else {
        const rows = parseCsv(input, { delimiter, typeInference, skipEmptyRows });
        const headers = rows[0]?.map((_, i) => `Col ${i + 1}`) || [];
        lastCsvParsed = { headers, rows };
        output.innerHTML = renderTable(headers, rows);
      }

      output.innerHTML += `<div class="meta">Detected delimiter: "${escapeHtml(detected)}" &middot; ${lastCsvParsed.rows.length} rows</div>`;

      ($("csv-copy") as HTMLButtonElement).disabled = false;
      ($("csv-to-xlsx") as HTMLButtonElement).disabled = false;
    } catch (e: unknown) {
      output.innerHTML = `<p class="error">${escapeHtml(String(e))}</p>`;
    }
  });

  $("csv-copy").addEventListener("click", () => {
    if (!lastCsvParsed) return;
    const json = lastCsvParsed.rows.map((row) => {
      const obj: Record<string, CellValue> = {};
      lastCsvParsed!.headers.forEach((h, i) => {
        obj[h] = row[i] ?? null;
      });
      return obj;
    });
    navigator.clipboard.writeText(JSON.stringify(json, null, 2));
    toast("JSON copied to clipboard");
  });

  $("csv-to-xlsx").addEventListener("click", async () => {
    if (!lastCsvParsed) return;
    try {
      const result = await writeXlsx({
        sheets: [
          {
            name: "CSV Import",
            columns: lastCsvParsed.headers.map((h) => ({ header: h, key: h })),
            rows: [lastCsvParsed.headers, ...lastCsvParsed.rows],
          },
        ],
      });
      const blob = new Blob([result], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "csv-export.xlsx";
      a.click();
      URL.revokeObjectURL(url);
      toast("XLSX downloaded");
    } catch (e: unknown) {
      toast(`Error: ${String(e)}`);
    }
  });
}

// ── Schema Validation ─────────────────────────────────────────────

function setupSchema() {
  $("schema-validate").addEventListener("click", () => {
    const output = $("schema-output");
    try {
      const csvInput = ($("schema-data") as HTMLTextAreaElement).value;
      const schemaDef = JSON.parse(($("schema-def") as HTMLTextAreaElement).value);

      // Convert pattern strings to RegExp
      const schema: SchemaDefinition = {};
      for (const [key, field] of Object.entries(schemaDef) as Array<
        [string, Record<string, unknown>]
      >) {
        schema[key] = { ...field } as SchemaDefinition[string];
        if (typeof field["pattern"] === "string") {
          schema[key].pattern = new RegExp(field["pattern"] as string);
        }
      }

      // Parse CSV first
      const rows = parseCsv(csvInput, { typeInference: true });

      // Validate
      const result = validateWithSchema(rows, schema, {
        headerRow: 1,
        skipEmptyRows: false,
        errorMode: "collect",
      });

      let html = "";

      // Valid data table
      if (result.data.length > 0) {
        const headers = Object.keys(schema);
        html +=
          '<div style="margin-bottom:0.5rem;color:var(--accent);font-weight:600;font-size:0.8rem">VALID ROWS</div>';
        html += "<table><thead><tr>";
        for (const h of headers) html += `<th>${escapeHtml(h)}</th>`;
        html += "</tr></thead><tbody>";
        for (const row of result.data) {
          html += "<tr>";
          for (const h of headers) {
            const v = (row as Record<string, CellValue>)[h];
            html += `<td class="${cellClass(v)}">${escapeHtml(cellDisplay(v))}</td>`;
          }
          html += "</tr>";
        }
        html += "</tbody></table>";
      }

      // Errors
      if (result.errors.length > 0) {
        html +=
          '<div style="margin-top:1rem;margin-bottom:0.5rem;color:var(--error);font-weight:600;font-size:0.8rem">VALIDATION ERRORS</div>';
        html +=
          "<table><thead><tr><th>Row</th><th>Field</th><th>Message</th><th>Value</th></tr></thead><tbody>";
        for (const err of result.errors) {
          html += `<tr>
            <td class="num">${err.row}</td>
            <td>${escapeHtml(err.field)}</td>
            <td style="color:var(--error)">${escapeHtml(err.message)}</td>
            <td class="null">${escapeHtml(String(err.value ?? "null"))}</td>
          </tr>`;
        }
        html += "</tbody></table>";
      }

      html += `<div class="meta">${result.data.length} valid rows, ${result.errors.length} errors</div>`;
      output.innerHTML = html;
    } catch (e: unknown) {
      output.innerHTML = `<p class="error">${escapeHtml(String(e))}</p>`;
    }
  });
}

// ── Streaming ─────────────────────────────────────────────────────

let lastStreamBlob: Blob | null = null;

function setupStreaming() {
  $("stream-generate").addEventListener("click", async () => {
    const output = $("stream-output");
    const rowCount = parseInt(($("stream-rows") as HTMLInputElement).value) || 10000;
    const colCount = parseInt(($("stream-cols") as HTMLInputElement).value) || 5;

    try {
      output.innerHTML = '<p style="color:var(--text-dim);text-align:center">Generating...</p>';

      const t0 = performance.now();

      // Write with streaming writer
      const headers = Array.from({ length: colCount }, (_, i) => `Col ${i + 1}`);
      const writer = new XlsxStreamWriter({
        name: "StreamData",
        columns: headers.map((h) => ({ header: h, key: h })),
        freezePane: { rows: 1 },
      });

      for (let r = 0; r < rowCount; r++) {
        const row: CellValue[] = [];
        for (let c = 0; c < colCount; c++) {
          row.push(c === 0 ? `Row ${r + 1}` : Math.round(Math.random() * 10000) / 100);
        }
        writer.addRow(row);
      }

      const xlsxBuffer = await writer.finish();
      const writeTime = performance.now() - t0;

      // Read back with streaming reader
      const t1 = performance.now();
      let streamedRows = 0;
      let firstRow: CellValue[] | null = null;
      let lastRow: CellValue[] | null = null;

      for await (const row of streamXlsxRows(xlsxBuffer)) {
        streamedRows++;
        if (streamedRows === 1) firstRow = row.values;
        lastRow = row.values;
      }

      const readTime = performance.now() - t1;
      const fileSize = (xlsxBuffer.byteLength / 1024).toFixed(1);

      let html = '<div class="stats" style="margin-bottom:1rem">';
      html += `<div class="stat"><div class="value">${rowCount.toLocaleString()}</div><div class="label">Rows Written</div></div>`;
      html += `<div class="stat"><div class="value">${streamedRows.toLocaleString()}</div><div class="label">Rows Read</div></div>`;
      html += `<div class="stat"><div class="value">${writeTime.toFixed(0)}ms</div><div class="label">Write Time</div></div>`;
      html += `<div class="stat"><div class="value">${readTime.toFixed(0)}ms</div><div class="label">Read Time</div></div>`;
      html += `<div class="stat"><div class="value">${fileSize} KB</div><div class="label">File Size</div></div>`;
      html += "</div>";

      if (firstRow) {
        html +=
          '<div style="color:var(--accent);font-weight:600;font-size:0.8rem;margin-bottom:0.5rem">FIRST ROW</div>';
        html += `<div style="font-family:monospace;font-size:0.8rem;color:var(--text-muted);margin-bottom:1rem">${firstRow.map((v) => escapeHtml(String(v))).join(" | ")}</div>`;
      }
      if (lastRow) {
        html +=
          '<div style="color:var(--accent);font-weight:600;font-size:0.8rem;margin-bottom:0.5rem">LAST ROW</div>';
        html += `<div style="font-family:monospace;font-size:0.8rem;color:var(--text-muted)">${lastRow.map((v) => escapeHtml(String(v))).join(" | ")}</div>`;
      }

      output.innerHTML = html;

      lastStreamBlob = new Blob([xlsxBuffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      ($("stream-download") as HTMLButtonElement).disabled = false;
    } catch (e: unknown) {
      output.innerHTML = `<p class="error">${escapeHtml(String(e))}</p>`;
    }
  });

  $("stream-download").addEventListener("click", () => {
    if (!lastStreamBlob) return;
    const url = URL.createObjectURL(lastStreamBlob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "stream-data.xlsx";
    a.click();
    URL.revokeObjectURL(url);
    toast("XLSX downloaded");
  });

  // ── Stream from File (ReadableStream input) ───────────────────
  const fileDrop = $("stream-file-drop");
  const fileInput = $("stream-file-input") as HTMLInputElement;

  fileDrop.addEventListener("click", () => fileInput.click());
  fileDrop.addEventListener("dragover", (e) => {
    e.preventDefault();
    fileDrop.classList.add("drag-over");
  });
  fileDrop.addEventListener("dragleave", () => fileDrop.classList.remove("drag-over"));
  fileDrop.addEventListener("drop", (e) => {
    e.preventDefault();
    fileDrop.classList.remove("drag-over");
    const file = (e as DragEvent).dataTransfer?.files[0];
    if (file) handleStreamFile(file);
  });
  fileInput.addEventListener("change", () => {
    if (fileInput.files?.[0]) handleStreamFile(fileInput.files[0]);
  });

  // ── Auto-split past Excel's row limit ─────────────────────────
  let lastAutosplitBlob: Blob | null = null;

  $("autosplit-generate").addEventListener("click", async () => {
    const output = $("autosplit-output");
    try {
      const totalRows = parseInt(($("autosplit-rows") as HTMLInputElement).value) || 2500;
      const limit = parseInt(($("autosplit-limit") as HTMLInputElement).value) || 1000;
      const repeat = ($("autosplit-repeat") as HTMLInputElement).checked;

      output.innerHTML = '<p style="color:var(--text-dim);text-align:center">Generating...</p>';
      const start = performance.now();

      const writer = new XlsxStreamWriter({
        name: "BigData",
        columns: [
          { key: "id", header: "ID" },
          { key: "value", header: "Value" },
        ],
        maxRowsPerSheet: limit,
        repeatHeaders: repeat,
      });
      for (let i = 0; i < totalRows; i++) {
        writer.addRow([i + 1, Math.random().toFixed(4)]);
      }
      const xlsx = await writer.finish();
      const writeTime = performance.now() - start;

      const wb = await readXlsx(xlsx);
      const expectedSheets = Math.ceil(repeat ? totalRows / (limit - 1) : totalRows / limit);

      lastAutosplitBlob = new Blob([new Uint8Array(xlsx)], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      let html = `<div class="meta">${totalRows} rows → ${wb.sheets.length} sheet${wb.sheets.length !== 1 ? "s" : ""} (${writeTime.toFixed(0)} ms, ${(xlsx.byteLength / 1024).toFixed(1)} KB)</div>`;
      html += `<div class="meta">Expected ~${expectedSheets} sheets at maxRowsPerSheet=${limit}${repeat ? " (with repeated header)" : ""}</div>`;
      html +=
        "<table><thead><tr><th>Sheet</th><th>Rows</th><th>First row</th><th>Last row</th></tr></thead><tbody>";
      for (const sheet of wb.sheets) {
        const first = sheet.rows[0];
        const last = sheet.rows[sheet.rows.length - 1];
        html += `<tr><td><strong>${escapeHtml(sheet.name)}</strong></td><td class="num">${sheet.rows.length}</td><td>${escapeHtml(JSON.stringify(first))}</td><td>${escapeHtml(JSON.stringify(last))}</td></tr>`;
      }
      html += "</tbody></table>";
      output.innerHTML = html;

      ($("autosplit-download") as HTMLButtonElement).disabled = false;
    } catch (e: unknown) {
      output.innerHTML = `<p class="error">${escapeHtml(String(e))}</p>`;
    }
  });

  $("autosplit-download").addEventListener("click", () => {
    if (!lastAutosplitBlob) return;
    const url = URL.createObjectURL(lastAutosplitBlob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "autosplit.xlsx";
    a.click();
    URL.revokeObjectURL(url);
    toast("XLSX downloaded");
  });
}

async function handleStreamFile(file: File) {
  const output = $("stream-file-output");
  output.innerHTML = '<p style="color:var(--text-dim);text-align:center">Streaming...</p>';

  try {
    const buffer = await file.arrayBuffer();
    const data = new Uint8Array(buffer);
    const fileSize = (data.byteLength / 1024).toFixed(1);

    // ── Buffered read (Uint8Array) ──
    const t0 = performance.now();
    let bufferedCount = 0;
    for await (const _row of streamXlsxRows(data)) {
      bufferedCount++;
    }
    const bufferedTime = performance.now() - t0;

    // ── Streaming read (ReadableStream) ──
    const readableStream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(data);
        controller.close();
      },
    });

    const t1 = performance.now();
    let streamCount = 0;
    let firstRowStream: CellValue[] | null = null;
    let lastRowStream: CellValue[] | null = null;
    for await (const row of streamXlsxRows(readableStream)) {
      streamCount++;
      if (streamCount === 1) firstRowStream = row.values;
      lastRowStream = row.values;
    }
    const streamTime = performance.now() - t1;

    const rowsMatch = bufferedCount === streamCount;

    let html = `<div style="font-size:0.8rem;color:var(--text-muted);margin-bottom:0.75rem"><strong>${escapeHtml(file.name)}</strong> &mdash; ${fileSize} KB</div>`;

    html += '<div class="stats" style="margin-bottom:1rem">';
    html += `<div class="stat"><div class="value">${streamCount.toLocaleString()}</div><div class="label">Rows</div></div>`;
    html += `<div class="stat"><div class="value">${bufferedTime.toFixed(0)}ms</div><div class="label">Uint8Array</div></div>`;
    html += `<div class="stat"><div class="value">${streamTime.toFixed(0)}ms</div><div class="label">ReadableStream</div></div>`;
    html += `<div class="stat"><div class="value">${rowsMatch ? "✓" : "✗"}</div><div class="label">Match</div></div>`;
    html += "</div>";

    if (firstRowStream) {
      html +=
        '<div style="color:var(--accent);font-weight:600;font-size:0.8rem;margin-bottom:0.5rem">FIRST ROW</div>';
      html += `<div style="font-family:monospace;font-size:0.75rem;color:var(--text-muted);margin-bottom:1rem;word-break:break-all">${firstRowStream.map((v) => escapeHtml(String(v))).join(" | ")}</div>`;
    }
    if (lastRowStream) {
      html +=
        '<div style="color:var(--accent);font-weight:600;font-size:0.8rem;margin-bottom:0.5rem">LAST ROW</div>';
      html += `<div style="font-family:monospace;font-size:0.75rem;color:var(--text-muted);word-break:break-all">${lastRowStream.map((v) => escapeHtml(String(v))).join(" | ")}</div>`;
    }

    output.innerHTML = html;
    toast(`Streamed ${streamCount.toLocaleString()} rows`);
  } catch (e: unknown) {
    output.innerHTML = `<p class="error">${escapeHtml(String(e))}</p>`;
  }
}

// ── ODS ───────────────────────────────────────────────────────────

let lastOdsBlob: Blob | null = null;

async function handleOdsFile(file: File) {
  const output = $("ods-output");
  try {
    output.innerHTML = '<p style="color:var(--text-dim);text-align:center">Parsing...</p>';
    const buffer = await file.arrayBuffer();
    const data = new Uint8Array(buffer);
    const wb = await readOds(data);

    if (wb.sheets.length === 0) {
      output.innerHTML = '<p class="error">No sheets found</p>';
      return;
    }

    const sheet = wb.sheets[0];
    const rows = sheet.rows;
    const headers = rows[0]?.map((v, i) => (v != null ? String(v) : `Col ${i + 1}`)) || [];
    const dataRows = rows.slice(1);

    output.innerHTML = renderTable(headers, dataRows);
    output.innerHTML += `<div class="meta">${wb.sheets.length} sheets, ${rows.length} rows, ${(data.byteLength / 1024).toFixed(1)} KB</div>`;
  } catch (e: unknown) {
    output.innerHTML = `<p class="error">${escapeHtml(String(e))}</p>`;
  }
}

function setupOds() {
  const drop = $("ods-drop");
  const fileInput = $("ods-file") as HTMLInputElement;

  drop.addEventListener("click", () => fileInput.click());
  drop.addEventListener("dragover", (e) => {
    e.preventDefault();
    drop.classList.add("drag-over");
  });
  drop.addEventListener("dragleave", () => drop.classList.remove("drag-over"));
  drop.addEventListener("drop", (e) => {
    e.preventDefault();
    drop.classList.remove("drag-over");
    const file = (e as DragEvent).dataTransfer?.files[0];
    if (file) handleOdsFile(file);
  });
  fileInput.addEventListener("change", () => {
    if (fileInput.files?.[0]) handleOdsFile(fileInput.files[0]);
  });

  $("ods-generate").addEventListener("click", async () => {
    const output = $("ods-output");
    try {
      const rawData = JSON.parse(($("ods-data") as HTMLTextAreaElement).value);
      const keys = Object.keys(rawData[0] || {});

      const result = await writeOds({
        sheets: [
          {
            name: "Sheet1",
            columns: keys.map((k) => ({ header: k, key: k })),
            data: rawData,
          },
        ],
      });

      lastOdsBlob = new Blob([result], {
        type: "application/vnd.oasis.opendocument.spreadsheet",
      });

      // Read it back to show preview
      const wb = await readOds(result);
      const sheet = wb.sheets[0];
      const rows = sheet.rows;
      const headers = rows[0]?.map((v, i) => (v != null ? String(v) : `Col ${i + 1}`)) || [];
      output.innerHTML = renderTable(headers, rows.slice(1));
      output.innerHTML += `<div class="meta">Generated: ${(result.byteLength / 1024).toFixed(1)} KB ODS</div>`;

      ($("ods-download") as HTMLButtonElement).disabled = false;
    } catch (e: unknown) {
      output.innerHTML = `<p class="error">${escapeHtml(String(e))}</p>`;
    }
  });

  $("ods-download").addEventListener("click", () => {
    if (!lastOdsBlob) return;
    const url = URL.createObjectURL(lastOdsBlob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "export.ods";
    a.click();
    URL.revokeObjectURL(url);
    toast("ODS downloaded");
  });
}

// ── Export (HTML / Markdown) ───────────────────────────────────────

let lastExportText = "";
let exportSheet: Sheet | null = null;

async function loadExportSheet(): Promise<Sheet> {
  // If a file was loaded, use that
  if (exportSheet) return exportSheet;

  // Otherwise parse the CSV textarea
  const csvText = ($("export-csv") as HTMLTextAreaElement).value;
  if (!csvText.trim()) throw new Error("No data — drop a file or paste CSV");
  const rows = parseCsv(csvText, { typeInference: true });
  return { name: "Export", rows };
}

async function handleExportFile(file: File) {
  const buffer = new Uint8Array(await file.arrayBuffer());
  const ext = file.name.split(".").pop()?.toLowerCase();

  if (ext === "xlsx") {
    const wb = await readXlsx(buffer);
    exportSheet = wb.sheets[0];
  } else if (ext === "ods") {
    const wb = await readOds(buffer);
    exportSheet = wb.sheets[0];
  } else if (ext === "csv" || ext === "tsv" || ext === "txt") {
    const text = new TextDecoder().decode(buffer);
    ($("export-csv") as HTMLTextAreaElement).value = text;
    exportSheet = null; // use CSV path
  } else {
    throw new Error(`Unsupported format: .${ext}`);
  }

  toast(`Loaded ${file.name}`);
}

function setupExport() {
  const drop = $("export-drop");
  const fileInput = $("export-file") as HTMLInputElement;

  drop.addEventListener("click", () => fileInput.click());
  drop.addEventListener("dragover", (e) => {
    e.preventDefault();
    drop.classList.add("drag-over");
  });
  drop.addEventListener("dragleave", () => drop.classList.remove("drag-over"));
  drop.addEventListener("drop", async (e) => {
    e.preventDefault();
    drop.classList.remove("drag-over");
    const file = (e as DragEvent).dataTransfer?.files[0];
    if (file) {
      try {
        await handleExportFile(file);
      } catch (err: unknown) {
        $("export-output").innerHTML = `<p class="error">${escapeHtml(String(err))}</p>`;
      }
    }
  });
  fileInput.addEventListener("change", async () => {
    if (fileInput.files?.[0]) {
      try {
        await handleExportFile(fileInput.files[0]);
      } catch (err: unknown) {
        $("export-output").innerHTML = `<p class="error">${escapeHtml(String(err))}</p>`;
      }
    }
  });

  // Reset file sheet when CSV is edited
  ($("export-csv") as HTMLTextAreaElement).addEventListener("input", () => {
    exportSheet = null;
  });

  $("export-html").addEventListener("click", async () => {
    const output = $("export-output");
    try {
      const sheet = await loadExportSheet();
      const headerRow = ($("export-header") as HTMLInputElement).checked;
      const styles = ($("export-styles") as HTMLInputElement).checked;
      const html = toHtml(sheet, { headerRow, styles, classes: true, includeStyleTag: true });
      lastExportText = html;
      output.innerHTML = `<div>${html}</div>`;
      output.innerHTML += `<details style="margin-top:0.75rem"><summary style="cursor:pointer;color:var(--text-dim);font-size:0.75rem">View source (${html.length} chars)</summary><pre style="font-size:0.7rem;overflow-x:auto;margin-top:0.5rem;color:var(--text-muted)">${escapeHtml(html)}</pre></details>`;
      ($("export-copy") as HTMLButtonElement).disabled = false;
    } catch (e: unknown) {
      output.innerHTML = `<p class="error">${escapeHtml(String(e))}</p>`;
    }
  });

  $("export-md").addEventListener("click", async () => {
    const output = $("export-output");
    try {
      const sheet = await loadExportSheet();
      const headerRow = ($("export-header") as HTMLInputElement).checked;
      const md = toMarkdown(sheet, { headerRow });
      lastExportText = md;
      output.innerHTML = `<pre style="font-size:0.8rem;line-height:1.6;color:var(--text-muted)">${escapeHtml(md)}</pre>`;
      ($("export-copy") as HTMLButtonElement).disabled = false;
    } catch (e: unknown) {
      output.innerHTML = `<p class="error">${escapeHtml(String(e))}</p>`;
    }
  });

  $("export-json").addEventListener("click", async () => {
    const output = $("export-output");
    try {
      const sheet = await loadExportSheet();
      const headerRow = ($("export-header") as HTMLInputElement).checked;
      const json = toJson(sheet, { headerRow, format: "objects", pretty: true });
      lastExportText = json;
      output.innerHTML = `<pre style="font-size:0.8rem;line-height:1.5;color:var(--text-muted)">${escapeHtml(json)}</pre>`;
      ($("export-copy") as HTMLButtonElement).disabled = false;
    } catch (e: unknown) {
      output.innerHTML = `<p class="error">${escapeHtml(String(e))}</p>`;
    }
  });

  $("export-ndjson").addEventListener("click", async () => {
    const output = $("export-output");
    try {
      const sheet = await loadExportSheet();
      const data = sheetToObjectsLocal(sheet);
      const out = writeNdjson(data);
      lastExportText = out;
      output.innerHTML = `<pre style="font-size:0.8rem;line-height:1.5;color:var(--text-muted)">${escapeHtml(out)}</pre>`;
      ($("export-copy") as HTMLButtonElement).disabled = false;
    } catch (e: unknown) {
      output.innerHTML = `<p class="error">${escapeHtml(String(e))}</p>`;
    }
  });

  $("export-xml").addEventListener("click", async () => {
    const output = $("export-output");
    try {
      const sheet = await loadExportSheet();
      const data = sheetToObjectsLocal(sheet);
      const out = writeXml(data, { rootTag: "rows", rowTag: "row", pretty: true });
      lastExportText = out;
      output.innerHTML = `<pre style="font-size:0.8rem;line-height:1.5;color:var(--text-muted)">${escapeHtml(out)}</pre>`;
      ($("export-copy") as HTMLButtonElement).disabled = false;
    } catch (e: unknown) {
      output.innerHTML = `<p class="error">${escapeHtml(String(e))}</p>`;
    }
  });

  $("export-copy").addEventListener("click", () => {
    if (!lastExportText) return;
    navigator.clipboard.writeText(lastExportText);
    toast("Copied to clipboard");
  });
}

// Local helper: build objects from a Sheet's rows (first row = headers).
function sheetToObjectsLocal(sheet: Sheet): Record<string, CellValue>[] {
  if (sheet.rows.length === 0) return [];
  const headers = sheet.rows[0]!.map((h) =>
    h === null || h === undefined ? "" : String(h).trim(),
  );
  const out: Record<string, CellValue>[] = [];
  for (let i = 1; i < sheet.rows.length; i++) {
    const row = sheet.rows[i]!;
    const obj: Record<string, CellValue> = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]!] = j < row.length ? (row[j] ?? null) : null;
    }
    out.push(obj);
  }
  return out;
}

// ── JSON / NDJSON ─────────────────────────────────────────────────

function setupJson() {
  let lastResult: { data: Record<string, CellValue>[]; headers: string[] } | null = null;

  function getMode(): "json" | "ndjson" {
    const checked = document.querySelector<HTMLInputElement>('input[name="json-mode"]:checked');
    return (checked?.value as "json" | "ndjson") ?? "json";
  }

  $("json-parse").addEventListener("click", () => {
    const output = $("json-output");
    try {
      const input = ($("json-input") as HTMLTextAreaElement).value;
      const flatten = ($("json-flatten") as HTMLInputElement).checked;
      const arrayJoin = ($("json-array-join") as HTMLInputElement).value || ", ";
      const rowsAt = ($("json-rows-at") as HTMLInputElement).value.trim() || undefined;
      const mode = getMode();

      const result =
        mode === "ndjson"
          ? parseNdjson(input, { flatten, arrayJoin })
          : parseJson(input, { flatten, arrayJoin, rowsAt });

      lastResult = result;
      const tableRows = result.data.map((row) => result.headers.map((h) => row[h] ?? null));

      const stats = `<div class="meta">${result.data.length} rows × ${result.headers.length} columns · mode: ${mode}</div>`;
      output.innerHTML = stats + renderTable(result.headers, tableRows);

      ($("json-to-ndjson") as HTMLButtonElement).disabled = false;
      ($("json-to-xlsx") as HTMLButtonElement).disabled = false;
      ($("json-copy") as HTMLButtonElement).disabled = false;
    } catch (e: unknown) {
      output.innerHTML = `<p class="error">${escapeHtml(String(e))}</p>`;
    }
  });

  $("json-to-ndjson").addEventListener("click", () => {
    if (!lastResult) return;
    const output = $("json-output");
    const ndjson = writeNdjson(lastResult.data);
    output.innerHTML = `<pre style="font-size:0.8rem;line-height:1.5;color:var(--text-muted)">${escapeHtml(ndjson)}</pre>`;
  });

  $("json-to-xlsx").addEventListener("click", async () => {
    if (!lastResult) return;
    try {
      const xlsx = await writeXlsxObjects(lastResult.data, {
        sheetName: "Imported",
        headers: lastResult.headers,
      });
      const blob = new Blob([new Uint8Array(xlsx)], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "imported.xlsx";
      a.click();
      URL.revokeObjectURL(url);
      toast("Downloaded imported.xlsx");
    } catch (e: unknown) {
      toast(`Error: ${String(e)}`);
    }
  });

  $("json-copy").addEventListener("click", () => {
    if (!lastResult) return;
    navigator.clipboard.writeText(writeJson(lastResult.data, { pretty: true }));
    toast("Copied JSON to clipboard");
  });
}

// ── XML ──────────────────────────────────────────────────────────

function setupXml() {
  let lastResult: {
    data: Record<string, CellValue>[];
    headers: string[];
    rowTag: string;
  } | null = null;

  $("xml-parse").addEventListener("click", () => {
    const output = $("xml-output");
    try {
      const input = ($("xml-input") as HTMLTextAreaElement).value;
      const rowTag = ($("xml-row-tag") as HTMLInputElement).value.trim() || undefined;
      const attrPrefix = ($("xml-attr-prefix") as HTMLInputElement).value || "@";
      const flatten = ($("xml-flatten") as HTMLInputElement).checked;
      const stripNamespaces = ($("xml-strip-ns") as HTMLInputElement).checked;

      const result = readXml(input, { rowTag, attrPrefix, flatten, stripNamespaces });
      lastResult = result;

      const tableRows = result.data.map((row) => result.headers.map((h) => row[h] ?? null));

      const stats = `<div class="meta">rowTag: <code>${escapeHtml(result.rowTag)}</code> · ${result.data.length} rows × ${result.headers.length} columns</div>`;
      output.innerHTML = stats + renderTable(result.headers, tableRows);

      ($("xml-roundtrip") as HTMLButtonElement).disabled = false;
      ($("xml-to-xlsx") as HTMLButtonElement).disabled = false;
      ($("xml-copy") as HTMLButtonElement).disabled = false;
    } catch (e: unknown) {
      output.innerHTML = `<p class="error">${escapeHtml(String(e))}</p>`;
    }
  });

  $("xml-roundtrip").addEventListener("click", () => {
    if (!lastResult) return;
    const output = $("xml-output");
    const xml = writeXml(lastResult.data, {
      rootTag: "Catalog",
      rowTag: lastResult.rowTag || "row",
      pretty: true,
    });
    output.innerHTML = `<pre style="font-size:0.8rem;line-height:1.5;color:var(--text-muted)">${escapeHtml(xml)}</pre>`;
  });

  $("xml-to-xlsx").addEventListener("click", async () => {
    if (!lastResult) return;
    try {
      const xlsx = await writeXlsxObjects(lastResult.data, {
        sheetName: lastResult.rowTag || "Imported",
        headers: lastResult.headers,
      });
      const blob = new Blob([new Uint8Array(xlsx)], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "xml-import.xlsx";
      a.click();
      URL.revokeObjectURL(url);
      toast("Downloaded xml-import.xlsx");
    } catch (e: unknown) {
      toast(`Error: ${String(e)}`);
    }
  });

  $("xml-copy").addEventListener("click", () => {
    if (!lastResult) return;
    navigator.clipboard.writeText(writeJson(lastResult.data, { pretty: true }));
    toast("Copied JSON to clipboard");
  });
}

// ── Format Value ──────────────────────────────────────────────────

function setupFormat() {
  function doFormat() {
    const output = $("fmt-output");
    try {
      const rawVal = ($("fmt-value") as HTMLInputElement).value;
      const fmt = ($("fmt-format") as HTMLInputElement).value;

      // Try to parse as number
      let value: unknown = rawVal;
      const num = Number(rawVal);
      if (!Number.isNaN(num) && rawVal.trim() !== "") value = num;
      if (rawVal.toLowerCase() === "true") value = true;
      if (rawVal.toLowerCase() === "false") value = false;

      const result = formatValue(value, fmt);

      let html = `<div style="text-align:center;padding:2rem 0">`;
      html += `<div style="font-size:2.5rem;font-weight:700;color:var(--accent);font-family:'SF Mono','Fira Code',monospace">${escapeHtml(result)}</div>`;
      html += `<div style="margin-top:1rem;color:var(--text-dim);font-size:0.8rem">`;
      html += `<span style="color:var(--text-muted)">formatValue(</span>`;
      html += `<span style="color:#60a5fa">${escapeHtml(String(value))}</span>`;
      html += `<span style="color:var(--text-muted)">, </span>`;
      html += `<span style="color:#fbbf24">"${escapeHtml(fmt)}"</span>`;
      html += `<span style="color:var(--text-muted)">)</span>`;
      html += `</div></div>`;
      output.innerHTML = html;
    } catch (e: unknown) {
      output.innerHTML = `<p class="error">${escapeHtml(String(e))}</p>`;
    }
  }

  $("fmt-apply").addEventListener("click", doFormat);

  // Quick example buttons
  document.querySelectorAll<HTMLButtonElement>("[data-fmt-val]").forEach((btn) => {
    btn.addEventListener("click", () => {
      ($("fmt-value") as HTMLInputElement).value = btn.dataset["fmtVal"] || "";
      ($("fmt-format") as HTMLInputElement).value = btn.dataset["fmtStr"] || "";
      doFormat();
    });
  });
}

// ── Chart Clone ───────────────────────────────────────────────────
//
// Demo for the chart cloning / dashboard composition flow tracked in
// issue #136. The user drops a template workbook, the panel surfaces
// every chart via `getCharts(workbook)`, and per-chart override knobs
// let them re-anchor / retitle / recolor / coerce a chart kind before
// re-emitting through `cloneChart` + `addChart` + `writeXlsx`.

interface ChartCloneState {
  workbook: Workbook;
  rawBytes: Uint8Array;
  charts: ReturnType<typeof getCharts>;
  fileName: string;
}

let chartCloneState: ChartCloneState | null = null;

const WRITE_CHART_KINDS: WriteChartKind[] = [
  "bar",
  "column",
  "line",
  "pie",
  "doughnut",
  "scatter",
  "area",
];

function colIndexToLetter(col: number): string {
  let n = col;
  let s = "";
  while (n >= 0) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

function parseAnchorCell(input: string): { row: number; col: number } | null {
  const m = /^\s*([A-Z]+)\s*([0-9]+)\s*$/i.exec(input);
  if (!m) return null;
  const letters = m[1].toUpperCase();
  const row1 = parseInt(m[2], 10);
  if (!Number.isFinite(row1) || row1 < 1) return null;
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  col -= 1;
  return { row: row1 - 1, col };
}

function describeKinds(kinds: ChartKind[]): string {
  if (!kinds || kinds.length === 0) return "(unknown)";
  return kinds.join(", ");
}

function defaultWriteKind(source: Chart): WriteChartKind | "" {
  for (const k of source.kinds) {
    const w = chartKindToWriteKind(k);
    if (w) return w;
  }
  return "";
}

function defaultAnchorFor(source: Chart, fallbackIndex: number): string {
  const a = source.anchor?.from;
  if (a) return `${colIndexToLetter(a.col)}${a.row + 1}`;
  // Fallback: stagger 18 rows per chart starting at B2 so multi-chart
  // dashboards do not stack on top of each other.
  return `B${2 + fallbackIndex * 18}`;
}

function renderChartList(state: ChartCloneState): string {
  if (state.charts.length === 0) {
    return `<p class="error">No charts found in this workbook. Try a file containing at least one Excel chart.</p>`;
  }

  let html = `<div class="meta" style="border-top:none;padding-top:0;margin-top:0">`;
  html += `<strong>${state.charts.length}</strong> chart${state.charts.length === 1 ? "" : "s"} found in <em>${escapeHtml(state.fileName)}</em>`;
  html += `</div>`;

  state.charts.forEach((loc, i) => {
    const c = loc.chart;
    const writeKind = defaultWriteKind(c);
    const writable = writeKind !== "";
    const anchor = defaultAnchorFor(c, i);
    const seriesCount = c.seriesCount ?? c.series?.length ?? 0;
    const title = c.title ?? "(untitled)";

    html += `<details class="section" data-chart-row="${i}" ${i === 0 ? "open" : ""} style="margin-top:0.5rem">`;
    html += `<summary>`;
    html += `<input type="checkbox" data-chart-pick="${i}" ${writable ? "checked" : ""} ${writable ? "" : "disabled"} style="margin-right:0.4rem" onclick="event.stopPropagation()" />`;
    html += `<span style="text-transform:none;letter-spacing:0;color:var(--text);font-weight:600">#${i + 1} · ${escapeHtml(loc.sheetName)}</span>`;
    html += ` <span style="color:var(--text-muted);font-weight:400;margin-left:0.4rem">${escapeHtml(describeKinds(c.kinds))} · ${seriesCount} ser · ${escapeHtml(title)}</span>`;
    html += `</summary>`;

    html += `<div class="field-grid">`;
    html += `<div class="field"><label>Title override</label><input type="text" data-chart-title="${i}" placeholder="(keep ${escapeHtml(title)})" /></div>`;

    html += `<div class="field"><label>Type coercion</label><select data-chart-type="${i}">`;
    html += `<option value="">(auto: ${escapeHtml(writeKind || "n/a")})</option>`;
    for (const k of WRITE_CHART_KINDS) {
      html += `<option value="${k}">${k}</option>`;
    }
    html += `</select></div>`;

    html += `<div class="field"><label>Anchor cell</label><input type="text" data-chart-anchor="${i}" value="${escapeHtml(anchor)}" /></div>`;

    html += `<div class="field"><label>Series #1 color (RRGGBB)</label><input type="text" data-chart-color="${i}" placeholder="e.g. 1F77B4" maxlength="6" /></div>`;

    if (!writable) {
      html += `<div class="field full"><span class="error" style="display:block">This chart kind ("${escapeHtml(describeKinds(c.kinds))}") cannot be cloned via the writer yet — pick a target type above to coerce it.</span></div>`;
    }

    if (c.series && c.series.length > 0) {
      html += `<div class="field full"><label>Series</label><div style="font-size:0.75rem;color:var(--text-muted);font-family:'SF Mono','Fira Code',monospace">`;
      for (const s of c.series) {
        html += `<div>${escapeHtml(String(s.index ?? 0))}: ${escapeHtml(s.name ?? "(unnamed)")} ← ${escapeHtml(s.valuesRef ?? "(literal)")}</div>`;
      }
      html += `</div></div>`;
    }

    html += `</div></details>`;
  });

  return html;
}

function readChartUiOverride(
  i: number,
  source: Chart,
): { picked: boolean; options: CloneChartOptions } | { error: string } {
  const pickEl = document.querySelector<HTMLInputElement>(`[data-chart-pick="${i}"]`);
  if (!pickEl) return { error: `Chart #${i + 1}: missing picker` };
  const picked = !pickEl.disabled && pickEl.checked;

  const titleEl = document.querySelector<HTMLInputElement>(`[data-chart-title="${i}"]`);
  const typeEl = document.querySelector<HTMLSelectElement>(`[data-chart-type="${i}"]`);
  const anchorEl = document.querySelector<HTMLInputElement>(`[data-chart-anchor="${i}"]`);
  const colorEl = document.querySelector<HTMLInputElement>(`[data-chart-color="${i}"]`);

  const anchorRaw = (anchorEl?.value || "").trim() || defaultAnchorFor(source, i);
  const from = parseAnchorCell(anchorRaw);
  if (!from) {
    return {
      error: `Chart #${i + 1}: invalid anchor cell "${anchorRaw}". Use A1-style notation (e.g. B2).`,
    };
  }

  const opts: CloneChartOptions = { anchor: { from } };

  const typeChoice = (typeEl?.value || "").trim();
  if (typeChoice) opts.type = typeChoice as WriteChartKind;

  const titleChoice = (titleEl?.value || "").trim();
  if (titleChoice) opts.title = titleChoice;

  const colorChoice = (colorEl?.value || "").trim().replace(/^#/, "").toUpperCase();
  if (colorChoice) {
    if (!/^[0-9A-F]{6}$/.test(colorChoice)) {
      return { error: `Chart #${i + 1}: color must be 6 hex digits (e.g. 1F77B4).` };
    }
    const overrides: CloneChartSeriesOverride[] = [];
    overrides[0] = { color: colorChoice };
    opts.seriesOverrides = overrides;
  }

  return { picked, options: opts };
}

function cleanWorkbookForRewrite(wb: Workbook): WriteSheet[] {
  // The reader produces a `Workbook` with `Sheet[]` shaped objects.
  // For the purpose of round-tripping data to the writer we keep just
  // the rows/name and drop the read-side metadata the writer doesn't
  // accept verbatim. We deliberately strip pre-existing charts on each
  // sheet — the demo re-emits a curated chart selection, so leaking
  // every original chart back into the output would double up.
  return wb.sheets.map<WriteSheet>((s) => ({
    name: s.name,
    rows: (s.rows ?? []) as CellValue[][],
  }));
}

function setupChartClone() {
  const drop = $("chart-clone-drop");
  const fileInput = $("chart-clone-file") as HTMLInputElement;
  const output = $("chart-clone-output");
  const stats = $("chart-clone-stats");
  const runBtn = $("chart-clone-run") as HTMLButtonElement;
  const destSel = $("chart-clone-dest") as HTMLSelectElement;
  const modeSel = $("chart-clone-mode") as HTMLSelectElement;

  drop.addEventListener("click", () => fileInput.click());
  drop.addEventListener("dragover", (e) => {
    e.preventDefault();
    drop.classList.add("drag-over");
  });
  drop.addEventListener("dragleave", () => drop.classList.remove("drag-over"));
  drop.addEventListener("drop", (e) => {
    e.preventDefault();
    drop.classList.remove("drag-over");
    const file = (e as DragEvent).dataTransfer?.files[0];
    if (file) handleChartCloneFile(file);
  });
  fileInput.addEventListener("change", () => {
    if (fileInput.files?.[0]) handleChartCloneFile(fileInput.files[0]);
  });

  async function handleChartCloneFile(file: File) {
    try {
      output.innerHTML = '<p style="color:var(--text-dim);text-align:center">Parsing charts...</p>';
      const buffer = await file.arrayBuffer();
      const data = new Uint8Array(buffer);
      const wb = await readXlsx(data);
      const charts = getCharts(wb);

      chartCloneState = { workbook: wb, rawBytes: data, charts, fileName: file.name };

      // Stats
      stats.hidden = false;
      stats.innerHTML = `
        <div class="stat"><div class="value">${wb.sheets.length}</div><div class="label">Sheets</div></div>
        <div class="stat"><div class="value">${charts.length}</div><div class="label">Charts</div></div>
        <div class="stat"><div class="value">${(data.byteLength / 1024).toFixed(1)} KB</div><div class="label">File Size</div></div>
      `;

      // Refresh destination list with sheets from the workbook + a "new sheet" option.
      destSel.innerHTML = "";
      const newOpt = document.createElement("option");
      newOpt.value = "__new__";
      newOpt.textContent = '[ New sheet "Charts" ]';
      destSel.appendChild(newOpt);
      for (let si = 0; si < wb.sheets.length; si++) {
        const o = document.createElement("option");
        o.value = String(si);
        o.textContent = wb.sheets[si].name;
        destSel.appendChild(o);
      }

      output.innerHTML = renderChartList(chartCloneState);
      runBtn.disabled = charts.length === 0;
    } catch (e: unknown) {
      output.innerHTML = `<p class="error">${escapeHtml(String(e))}</p>`;
      stats.hidden = true;
      runBtn.disabled = true;
      chartCloneState = null;
    }
  }

  runBtn.addEventListener("click", async () => {
    if (!chartCloneState) return;
    try {
      const state = chartCloneState;
      const writeSheets = cleanWorkbookForRewrite(state.workbook);

      // Resolve destination sheet
      const destChoice = destSel.value;
      const mode = modeSel.value;
      let destSheet: WriteSheet;
      let destSheetLabel: string;
      if (destChoice === "__new__") {
        const newName = ($("chart-clone-dest-name") as HTMLInputElement).value.trim() || "Charts";
        // Avoid colliding with an existing sheet by suffixing a number if needed.
        let unique = newName;
        let n = 2;
        while (writeSheets.some((s) => s.name === unique)) {
          unique = `${newName} ${n++}`;
        }
        destSheet = { name: unique, rows: [["Generated by hucre cloneChart() demo"]] };
        writeSheets.push(destSheet);
        destSheetLabel = unique;
      } else {
        const idx = parseInt(destChoice, 10);
        destSheet = writeSheets[idx];
        destSheetLabel = destSheet.name;
      }

      const composedAnchors: Array<{ from: { row: number; col: number } }> = [];
      const addedTitles: string[] = [];
      let addedCount = 0;

      for (let i = 0; i < state.charts.length; i++) {
        const source = state.charts[i].chart;
        const result = readChartUiOverride(i, source);
        if ("error" in result) {
          throw new Error(result.error);
        }
        if (!result.picked) continue;

        // Compose mode lays charts out in a 2-column grid; clone mode
        // honors each chart's own anchor input as-is.
        let opts = result.options;
        if (mode === "compose") {
          const col = (addedCount % 2) * 9; // ~9 columns wide per chart
          const row = Math.floor(addedCount / 2) * 18 + 1; // ~18 rows tall
          opts = { ...opts, anchor: { from: { row, col } } };
        }

        let chart: SheetChart;
        try {
          chart = cloneChart(source, opts);
        } catch (e: unknown) {
          throw new Error(`Chart #${i + 1}: ${String(e)}`);
        }

        addChart(destSheet, chart);
        composedAnchors.push({ from: opts.anchor.from });
        addedTitles.push(chart.title ?? "(untitled)");
        addedCount++;
      }

      if (addedCount === 0) {
        toast("Pick at least one chart to clone");
        return;
      }

      const bytes = await writeXlsx({ sheets: writeSheets });
      const blob = new Blob([bytes], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = state.fileName.replace(/\.xlsx$/i, "") + ".clone.xlsx";
      a.click();
      URL.revokeObjectURL(url);

      // Append a small log to the output panel.
      const logHtml: string[] = [];
      logHtml.push(`<div class="meta" style="margin-top:0.75rem">`);
      logHtml.push(
        `Cloned <strong>${addedCount}</strong> chart${addedCount === 1 ? "" : "s"} onto sheet <strong>${escapeHtml(destSheetLabel)}</strong> (${(bytes.byteLength / 1024).toFixed(1)} KB).`,
      );
      logHtml.push(`<ul style="margin:0.4rem 0 0 1.2rem;padding:0;font-size:0.75rem">`);
      addedTitles.forEach((t, k) => {
        const a = composedAnchors[k].from;
        logHtml.push(
          `<li>${escapeHtml(t)} → ${escapeHtml(colIndexToLetter(a.col))}${a.row + 1}</li>`,
        );
      });
      logHtml.push(`</ul></div>`);
      output.insertAdjacentHTML("beforeend", logHtml.join(""));

      toast(`Downloaded ${addedCount} cloned chart${addedCount === 1 ? "" : "s"}`);
    } catch (e: unknown) {
      output.insertAdjacentHTML(
        "beforeend",
        `<p class="error" style="margin-top:0.75rem">${escapeHtml(String(e))}</p>`,
      );
    }
  });
}

// ── Init ──────────────────────────────────────────────────────────

export function setupApp() {
  setupTabs();
  setupRead();
  setupWrite();
  setupCsv();
  setupJson();
  setupXml();
  setupSchema();
  setupStreaming();
  setupOds();
  setupExport();
  setupFormat();
  setupChartClone();
}
