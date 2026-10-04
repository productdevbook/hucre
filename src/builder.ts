import { createCellStore, setCell } from "./cell-store"
// ── Builder Pattern / Fluent API ─────────────────────────────────────
// Provides a method-chaining API for constructing workbooks.

import type {
  WorkbookInput,
  XlsxWriteOptions,
  SheetInput,
  CellInput,
  ColumnDef,
  DataValidation,
  Cell,
  WorkbookProperties,
  FontStyle,
} from "./_types"
import { writeXlsx } from "./xlsx/writer"

/**
 * Fluent builder for constructing XLSX workbooks.
 *
 * @example
 * ```ts
 * const data = await WorkbookBuilder.create()
 *   .addSheet("Sales")
 *     .columns([{ header: "Product", width: 20 }, { header: "Amount", width: 12 }])
 *     .row(["Widget", 100])
 *     .row(["Gadget", 250])
 *     .freeze(1)
 *   .done()
 *   .build();
 * ```
 */
export class WorkbookBuilder {
  private sheets: SheetBuilder[] = []
  private _writeOptions: XlsxWriteOptions = {}
  /** The rest of `WorkbookInput`, set through {@link set}. */
  private _rest: Partial<Omit<WorkbookInput, "sheets">> = {}

  static create(): WorkbookBuilder {
    return new WorkbookBuilder()
  }

  /**
   * Add a new sheet and return its builder.
   * Use `.done()` on the SheetBuilder to return to this WorkbookBuilder.
   */
  addSheet(name: string): SheetBuilder {
    const sb = new SheetBuilder(name, this)
    this.sheets.push(sb)
    return sb
  }

  /** Set workbook properties (title, creator, etc.) */
  properties(props: WorkbookProperties): this {
    this._rest.properties = props
    return this
  }

  /** Set the default font for the workbook */
  defaultFont(font: FontStyle): this {
    this._rest.defaultFont = font
    return this
  }

  /** Set the date system (1900 or 1904) */
  dateSystem(system: "1900" | "1904"): this {
    this._rest.dateSystem = system
    return this
  }

  /** Set the active sheet index (0-based) */
  activeSheet(index: number): this {
    this._rest.activeSheet = index
    return this
  }

  /** Define workbook-level named ranges. */
  namedRanges(ranges: NonNullable<WorkbookInput["namedRanges"]>): this {
    this._rest.namedRanges = ranges
    return this
  }

  /** Lock the workbook's structure and/or windows. */
  protect(protection: NonNullable<WorkbookInput["workbookProtection"]>): this {
    this._rest.workbookProtection = protection
    return this
  }

  /** Store strings in a shared table (default) or inline per cell. */
  stringMode(mode: NonNullable<XlsxWriteOptions["stringMode"]>): this {
    this._writeOptions.stringMode = mode
    return this
  }

  /** Encrypt the output (ECMA-376 Agile). */
  encrypt(encryption: NonNullable<XlsxWriteOptions["encryption"]>): this {
    this._writeOptions.encryption = encryption
    return this
  }

  /** Embed a VBA project, making the output macro-enabled. */
  vbaProject(project: NonNullable<XlsxWriteOptions["vbaProject"]>): this {
    this._writeOptions.vbaProject = project
    return this
  }

  /**
   * Set any other `WorkbookInput` field. The escape hatch, so the builder
   * cannot fall behind the type. `sheets` comes from `addSheet`.
   */
  set(fields: Partial<Omit<WorkbookInput, "sheets">>): this {
    Object.assign(this._rest, fields)
    return this
  }

  /** Build XLSX bytes, with optional encoding and loss reporting. */
  async build(options?: XlsxWriteOptions): Promise<Uint8Array> {
    return writeXlsx(
      {
        ...this._rest,
        sheets: this.sheets.map((s) => s._toWriteSheet()),
      },
      { ...this._writeOptions, ...options },
    )
  }
}

/**
 * Fluent builder for constructing a single worksheet.
 */
export class SheetBuilder {
  // Named methods and set() share one authoring state. Reassembling
  // separate fields at build time used to overwrite set() silently.
  // The model, rather than a second field list, defines the surface (#439).
  private _rest: Partial<SheetInput> = {}

  constructor(
    private _name: string,
    private _wb: WorkbookBuilder,
  ) {}

  /** Add a single column definition. */
  column(col: ColumnDef): this {
    ;(this._rest.columns ??= []).push(col)
    return this
  }

  /** Add multiple column definitions at once. */
  columns(cols: ColumnDef[]): this {
    ;(this._rest.columns ??= []).push(...cols)
    return this
  }

  /** Add a single row of values. */
  row(values: CellInput[]): this {
    ;(this._rest.rows ??= []).push(values)
    return this
  }

  /** Add multiple rows of values at once. */
  rows(data: CellInput[][]): this {
    ;(this._rest.rows ??= []).push(...data)
    return this
  }

  /** Add a merge range (0-based, inclusive). */
  merge(startRow: number, startCol: number, endRow: number, endCol: number): this {
    ;(this._rest.merges ??= []).push({ startRow, startCol, endRow, endCol })
    return this
  }

  /** Freeze rows and/or columns. */
  freeze(rows?: number, columns?: number): this {
    this._rest.freezePane = { rows, columns }
    return this
  }

  /** Add a data validation rule. */
  validation(v: DataValidation): this {
    ;(this._rest.dataValidations ??= []).push(v)
    return this
  }

  /** Set a cell-level override at zero-based row and column coordinates. */
  cell(row: number, col: number, cell: Partial<Cell>): this {
    setCell((this._rest.cells ??= createCellStore()), row, col, cell)
    return this
  }

  /** Mark the sheet as hidden. */
  hidden(value = true): this {
    this._rest.hidden = value
    return this
  }

  /** Mark the sheet as very hidden (only unhideable via VBA). */
  veryHidden(value = true): this {
    this._rest.veryHidden = value
    return this
  }

  /** Add a conditional formatting rule. */
  conditionalRule(rule: NonNullable<SheetInput["conditionalRules"]>[number]): this {
    ;(this._rest.conditionalRules ??= []).push(rule)
    return this
  }

  /** Set the auto-filter range (and optional per-column value filters). */
  autoFilter(filter: NonNullable<SheetInput["autoFilter"]>): this {
    this._rest.autoFilter = filter
    return this
  }

  /** Split the sheet into panes, in twips. */
  split(xSplit?: number, ySplit?: number): this {
    this._rest.splitPane = { xSplit, ySplit }
    return this
  }

  /** Set row-level properties — height, hidden, outline level, collapsed. */
  rowDef(
    row: number,
    def: NonNullable<SheetInput["rowDefs"]> extends Map<number, infer T> ? T : never,
  ): this {
    ;(this._rest.rowDefs ??= new Map()).set(row, def)
    return this
  }

  /** Page setup: orientation, scale, margins, print area, paper size. */
  pageSetup(setup: NonNullable<SheetInput["pageSetup"]>): this {
    this._rest.pageSetup = setup
    return this
  }

  /** Headers and footers. */
  headerFooter(hf: NonNullable<SheetInput["headerFooter"]>): this {
    this._rest.headerFooter = hf
    return this
  }

  /** Sheet view: grid lines, zoom, tab colour, right-to-left. */
  view(view: NonNullable<SheetInput["view"]>): this {
    this._rest.view = view
    return this
  }

  /** Protect the sheet. */
  protect(protection: NonNullable<SheetInput["protection"]>): this {
    this._rest.protection = protection
    return this
  }

  /** Define an Excel table (ListObject) over a range. */
  table(table: NonNullable<SheetInput["tables"]>[number]): this {
    ;(this._rest.tables ??= []).push(table)
    return this
  }

  /** Place an image. */
  image(image: NonNullable<SheetInput["images"]>[number]): this {
    ;(this._rest.images ??= []).push(image)
    return this
  }

  /** Add a chart. */
  chart(chart: NonNullable<SheetInput["charts"]>[number]): this {
    ;(this._rest.charts ??= []).push(chart)
    return this
  }

  /**
   * Set any other `SheetInput` field — sparklines, text boxes, page
   * breaks, outline properties, a background image, pivot tables, a11y
   * metadata.
   *
   * The escape hatch, so the builder cannot fall behind the type. `name`
   * is fixed by `addSheet` and is rejected here.
   */
  set(fields: Partial<Omit<SheetInput, "name">>): this {
    Object.assign(this._rest, fields)
    return this
  }

  /** Go back to the workbook builder to add another sheet or finish. */
  done(): WorkbookBuilder {
    return this._wb
  }

  /** Build directly, forwarding encoding and loss reporting options. */
  async build(options?: XlsxWriteOptions): Promise<Uint8Array> {
    return this._wb.build(options)
  }

  /** @internal Assemble this builder's state into a SheetInput. */
  _toWriteSheet(): SheetInput {
    return {
      ...this._rest,
      name: this._name,
    }
  }
}
