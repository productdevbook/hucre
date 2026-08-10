// ── Builder Pattern / Fluent API ─────────────────────────────────────
// Provides a method-chaining API for constructing workbooks.

import type {
  WriteOptions,
  WriteSheet,
  CellValue,
  ColumnDef,
  DataValidation,
  MergeRange,
  FreezePane,
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
  private _properties?: WriteOptions["properties"]
  private _defaultFont?: WriteOptions["defaultFont"]
  private _dateSystem?: WriteOptions["dateSystem"]
  private _activeSheet?: WriteOptions["activeSheet"]
  /** The rest of `WriteOptions`, set through {@link set}. */
  private _rest: Partial<Omit<WriteOptions, "sheets">> = {}

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
    this._properties = props
    return this
  }

  /** Set the default font for the workbook */
  defaultFont(font: FontStyle): this {
    this._defaultFont = font
    return this
  }

  /** Set the date system (1900 or 1904) */
  dateSystem(system: "1900" | "1904"): this {
    this._dateSystem = system
    return this
  }

  /** Set the active sheet index (0-based) */
  activeSheet(index: number): this {
    this._activeSheet = index
    return this
  }

  /** Build the workbook and return the XLSX as a Uint8Array. */

  /** Define workbook-level named ranges. */
  namedRanges(ranges: NonNullable<WriteOptions["namedRanges"]>): this {
    this._rest.namedRanges = ranges
    return this
  }

  /** Lock the workbook's structure and/or windows. */
  protect(protection: NonNullable<WriteOptions["workbookProtection"]>): this {
    this._rest.workbookProtection = protection
    return this
  }

  /** Store strings in a shared table (default) or inline per cell. */
  stringMode(mode: NonNullable<WriteOptions["stringMode"]>): this {
    this._rest.stringMode = mode
    return this
  }

  /** Encrypt the output (ECMA-376 Agile). */
  encrypt(encryption: NonNullable<WriteOptions["encryption"]>): this {
    this._rest.encryption = encryption
    return this
  }

  /** Embed a VBA project, making the output macro-enabled. */
  vbaProject(project: NonNullable<WriteOptions["vbaProject"]>): this {
    this._rest.vbaProject = project
    return this
  }

  /**
   * Set any other `WriteOptions` field. The escape hatch, so the builder
   * cannot fall behind the type. `sheets` comes from `addSheet`.
   */
  set(fields: Partial<Omit<WriteOptions, "sheets">>): this {
    Object.assign(this._rest, fields)
    return this
  }

  async build(): Promise<Uint8Array> {
    return writeXlsx({
      ...this._rest,
      sheets: this.sheets.map((s) => s._toWriteSheet()),
      properties: this._properties,
      defaultFont: this._defaultFont,
      dateSystem: this._dateSystem,
      activeSheet: this._activeSheet,
    })
  }
}

/**
 * Fluent builder for constructing a single worksheet.
 */
export class SheetBuilder {
  private _columns: ColumnDef[] = []
  private _rows: CellValue[][] = []
  private _merges: MergeRange[] = []
  private _freezePane?: FreezePane
  private _validations: DataValidation[] = []
  private _cells?: Map<string, Partial<Cell>>
  private _hidden?: boolean
  private _veryHidden?: boolean
  /**
   * Everything else on `WriteSheet`, set through {@link set}.
   *
   * The builder used to reach eight of the type's fields, so the first
   * sheet needing a page setup or a conditional rule had to abandon it
   * entirely. The named methods below cover what a builder is for; `set`
   * covers the rest without this class having to grow a method per field
   * and drift behind the type. See #439 §AJ.
   */
  private _rest: Partial<WriteSheet> = {}

  constructor(
    private _name: string,
    private _wb: WorkbookBuilder,
  ) {}

  /** Add a single column definition. */
  column(col: ColumnDef): this {
    this._columns.push(col)
    return this
  }

  /** Add multiple column definitions at once. */
  columns(cols: ColumnDef[]): this {
    this._columns.push(...cols)
    return this
  }

  /** Add a single row of values. */
  row(values: CellValue[]): this {
    this._rows.push(values)
    return this
  }

  /** Add multiple rows of values at once. */
  rows(data: CellValue[][]): this {
    this._rows.push(...data)
    return this
  }

  /** Add a merge range (0-based, inclusive). */
  merge(startRow: number, startCol: number, endRow: number, endCol: number): this {
    this._merges.push({ startRow, startCol, endRow, endCol })
    return this
  }

  /** Freeze rows and/or columns. */
  freeze(rows?: number, columns?: number): this {
    this._freezePane = { rows, columns }
    return this
  }

  /** Add a data validation rule. */
  validation(v: DataValidation): this {
    this._validations.push(v)
    return this
  }

  /** Set a cell-level override (keyed by "row,col", e.g. "0,2"). */
  cell(row: number, col: number, cell: Partial<Cell>): this {
    if (!this._cells) {
      this._cells = new Map()
    }
    this._cells.set(`${row},${col}`, cell)
    return this
  }

  /** Mark the sheet as hidden. */
  hidden(value = true): this {
    this._hidden = value
    return this
  }

  /** Mark the sheet as very hidden (only unhideable via VBA). */
  veryHidden(value = true): this {
    this._veryHidden = value
    return this
  }

  /** Add a conditional formatting rule. */
  conditionalRule(rule: NonNullable<WriteSheet["conditionalRules"]>[number]): this {
    ;(this._rest.conditionalRules ??= []).push(rule)
    return this
  }

  /** Set the auto-filter range (and optional per-column value filters). */
  autoFilter(filter: NonNullable<WriteSheet["autoFilter"]>): this {
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
    def: NonNullable<WriteSheet["rowDefs"]> extends Map<number, infer T> ? T : never,
  ): this {
    ;(this._rest.rowDefs ??= new Map()).set(row, def)
    return this
  }

  /** Page setup: orientation, scale, margins, print area, paper size. */
  pageSetup(setup: NonNullable<WriteSheet["pageSetup"]>): this {
    this._rest.pageSetup = setup
    return this
  }

  /** Headers and footers. */
  headerFooter(hf: NonNullable<WriteSheet["headerFooter"]>): this {
    this._rest.headerFooter = hf
    return this
  }

  /** Sheet view: grid lines, zoom, tab colour, right-to-left. */
  view(view: NonNullable<WriteSheet["view"]>): this {
    this._rest.view = view
    return this
  }

  /** Protect the sheet. */
  protect(protection: NonNullable<WriteSheet["protection"]>): this {
    this._rest.protection = protection
    return this
  }

  /** Define an Excel table (ListObject) over a range. */
  table(table: NonNullable<WriteSheet["tables"]>[number]): this {
    ;(this._rest.tables ??= []).push(table)
    return this
  }

  /** Place an image. */
  image(image: NonNullable<WriteSheet["images"]>[number]): this {
    ;(this._rest.images ??= []).push(image)
    return this
  }

  /** Add a chart. */
  chart(chart: NonNullable<WriteSheet["charts"]>[number]): this {
    ;(this._rest.charts ??= []).push(chart)
    return this
  }

  /**
   * Set any other `WriteSheet` field — sparklines, text boxes, page
   * breaks, outline properties, a background image, pivot tables, a11y
   * metadata.
   *
   * The escape hatch, so the builder cannot fall behind the type. `name`
   * is fixed by `addSheet` and is rejected here.
   */
  set(fields: Partial<Omit<WriteSheet, "name">>): this {
    Object.assign(this._rest, fields)
    return this
  }

  /** Go back to the workbook builder to add another sheet or finish. */
  done(): WorkbookBuilder {
    return this._wb
  }

  /** Build the workbook directly (shortcut that skips `.done().build()`). */
  async build(): Promise<Uint8Array> {
    return this._wb.build()
  }

  /** @internal Assemble this builder's state into a WriteSheet. */
  _toWriteSheet(): WriteSheet {
    return {
      ...this._rest,
      name: this._name,
      columns: this._columns.length > 0 ? this._columns : undefined,
      rows: this._rows,
      cells: this._cells,
      merges: this._merges.length > 0 ? this._merges : undefined,
      freezePane: this._freezePane,
      dataValidations: this._validations.length > 0 ? this._validations : undefined,
      hidden: this._hidden,
      veryHidden: this._veryHidden,
    }
  }
}
