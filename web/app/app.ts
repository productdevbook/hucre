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
} from "defter";
import type { CellValue, WriteSheet, SchemaDefinition } from "defter";

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
    const wb = await readXlsx(data, { readStyles: ($("read-styles") as HTMLInputElement).checked });

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

      const columns: Record<string, { header?: string; width?: number; numFmt?: string }> = rawCols;
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
        freezePane: freezeRows > 0 ? { rows: freezeRows } : undefined,
        autoFilter: autoFilter
          ? {
              range: `A1:${String.fromCharCode(64 + Object.keys(columns).length)}${rawData.length + 1}`,
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

// ── Init ──────────────────────────────────────────────────────────

export function setupApp() {
  setupTabs();
  setupRead();
  setupWrite();
  setupCsv();
  setupSchema();
  setupStreaming();
  setupOds();
}
