// ── Builder Pattern / Fluent API ─────────────────────────────────────
// Provides a method-chaining API for constructing workbooks.

import type {
  WriteOptions,
  WriteSheet,
  CellStyle,
  CellValue,
  ColumnDef,
  DataValidation,
  MergeRange,
  FreezePane,
  Cell,
  WorkbookProperties,
  FontStyle,
} from "./_types";
import { writeXlsx } from "./xlsx/writer";

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
  private sheets: SheetBuilder[] = [];
  private _properties?: WriteOptions["properties"];
  private _defaultFont?: WriteOptions["defaultFont"];
  private _dateSystem?: WriteOptions["dateSystem"];
  private _activeSheet?: WriteOptions["activeSheet"];

  static create(): WorkbookBuilder {
    return new WorkbookBuilder();
  }

  /**
   * Add a new sheet and return its builder.
   * Use `.done()` on the SheetBuilder to return to this WorkbookBuilder.
   */
  addSheet(name: string): SheetBuilder {
    const sb = new SheetBuilder(name, this);
    this.sheets.push(sb);
    return sb;
  }

  /** Set workbook properties (title, creator, etc.) */
  properties(props: WorkbookProperties): this {
    this._properties = props;
    return this;
  }

  /** Set the default font for the workbook */
  defaultFont(font: FontStyle): this {
    this._defaultFont = font;
    return this;
  }

  /** Set the date system (1900 or 1904) */
  dateSystem(system: "1900" | "1904"): this {
    this._dateSystem = system;
    return this;
  }

  /** Set the active sheet index (0-based) */
  activeSheet(index: number): this {
    this._activeSheet = index;
    return this;
  }

  /** Build the workbook and return the XLSX as a Uint8Array. */
  async build(): Promise<Uint8Array> {
    return writeXlsx({
      sheets: this.sheets.map((s) => s._toWriteSheet()),
      properties: this._properties,
      defaultFont: this._defaultFont,
      dateSystem: this._dateSystem,
      activeSheet: this._activeSheet,
    });
  }
}

/**
 * Fluent builder for constructing a single worksheet.
 */
export class SheetBuilder {
  private _columns: ColumnDef[] = [];
  private _rows: CellValue[][] = [];
  private _data?: Array<Record<string, unknown>>;
  private _merges: MergeRange[] = [];
  private _freezePane?: FreezePane;
  private _validations: DataValidation[] = [];
  private _cells?: Map<string, Partial<Cell>>;
  private _hidden?: boolean;
  private _veryHidden?: boolean;

  constructor(
    private _name: string,
    private _wb: WorkbookBuilder,
  ) {}

  /** Add a single column definition. */
  column(col: ColumnDef): this {
    this._columns.push(col);
    return this;
  }

  /** Add multiple column definitions at once. */
  columns(cols: ColumnDef[]): this {
    this._columns.push(...cols);
    return this;
  }

  /** Add a single row of values. */
  row(values: CellValue[]): this {
    this._rows.push(values);
    return this;
  }

  /** Add multiple rows of values at once. */
  rows(data: CellValue[][]): this {
    this._rows.push(...data);
    return this;
  }

  /** Set object data with column definitions.
   *  Uses the data+columns path for full ColumnDef support (value accessors, transforms, formulas, etc.). */
  objectRows<T extends Record<string, unknown>>(data: T[], cols: ColumnDef<T>[]): this {
    this._data = data as Array<Record<string, unknown>>;
    this._columns = cols as ColumnDef[];
    return this;
  }

  /** Add a merge range (0-based, inclusive). */
  merge(startRow: number, startCol: number, endRow: number, endCol: number): this {
    this._merges.push({ startRow, startCol, endRow, endCol });
    return this;
  }

  /** Freeze rows and/or columns. */
  freeze(rows?: number, columns?: number): this {
    this._freezePane = { rows, columns };
    return this;
  }

  /** Add a data validation rule. */
  validation(v: DataValidation): this {
    this._validations.push(v);
    return this;
  }

  /** Set a cell-level override (keyed by "row,col", e.g. "0,2"). */
  cell(row: number, col: number, cell: Partial<Cell>): this {
    if (!this._cells) {
      this._cells = new Map();
    }
    this._cells.set(`${row},${col}`, cell);
    return this;
  }

  /** Mark the sheet as hidden. */
  hidden(value = true): this {
    this._hidden = value;
    return this;
  }

  /** Mark the sheet as very hidden (only unhideable via VBA). */
  veryHidden(value = true): this {
    this._veryHidden = value;
    return this;
  }

  /** Go back to the workbook builder to add another sheet or finish. */
  done(): WorkbookBuilder {
    return this._wb;
  }

  /** Build the workbook directly (shortcut that skips `.done().build()`). */
  async build(): Promise<Uint8Array> {
    return this._wb.build();
  }

  /** @internal Assemble this builder's state into a WriteSheet. */
  _toWriteSheet(): WriteSheet {
    return {
      name: this._name,
      columns: this._columns.length > 0 ? this._columns : undefined,
      ...(this._data ? { data: this._data } : { rows: this._rows }),
      cells: this._cells,
      merges: this._merges.length > 0 ? this._merges : undefined,
      freezePane: this._freezePane,
      dataValidations: this._validations.length > 0 ? this._validations : undefined,
      hidden: this._hidden,
      veryHidden: this._veryHidden,
    };
  }
}
