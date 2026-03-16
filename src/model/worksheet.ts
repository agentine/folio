import { parseAddress, parseRange } from './address.js';
import { Cell } from './cell.js';
import { Column } from './column.js';
import { Row } from './row.js';
import type { CellValue, ColumnDefinition, PageSetup, SheetState, SheetView } from './types.js';

export interface WorksheetOptions {
  state?: SheetState;
  properties?: Record<string, unknown>;
}

interface MergedRange {
  start: { col: number; row: number };
  end: { col: number; row: number };
  range: string;
}

export class Worksheet {
  readonly id: number;
  name: string;
  state: SheetState = 'visible';
  views: SheetView[] = [];
  pageSetup: PageSetup = {};

  private _rows: Map<number, Row> = new Map();
  private _columns: Column[] = [];
  private _columnKeys: Map<string, number> = new Map();
  private _mergedCells: MergedRange[] = [];

  /** @internal */
  constructor(id: number, name: string, options?: WorksheetOptions) {
    this.id = id;
    this.name = name;
    if (options?.state) this.state = options.state;
  }

  // --- Columns ---

  get columns(): ColumnDefinition[] {
    return this._columns.map((c) => ({
      header: c.header,
      key: c.key,
      width: c.width,
      style: c.style,
      hidden: c.hidden,
      outlineLevel: c.outlineLevel,
    }));
  }

  set columns(defs: ColumnDefinition[]) {
    this._columns = [];
    this._columnKeys.clear();
    for (let i = 0; i < defs.length; i++) {
      const col = new Column(i + 1, defs[i]);
      this._columns.push(col);
      if (col.key) {
        this._columnKeys.set(col.key, i + 1);
      }
    }

    // Set header row if headers are provided
    const hasHeaders = defs.some((d) => d.header !== undefined);
    if (hasHeaders) {
      const row = this.getRow(1);
      for (let i = 0; i < defs.length; i++) {
        if (defs[i].header !== undefined) {
          row.getCell(i + 1).value = defs[i].header!;
        }
      }
    }
  }

  getColumn(col: number): Column {
    while (this._columns.length < col) {
      this._columns.push(new Column(this._columns.length + 1));
    }
    return this._columns[col - 1];
  }

  // --- Rows ---

  getRow(rowNumber: number): Row {
    let row = this._rows.get(rowNumber);
    if (!row) {
      row = new Row(rowNumber);
      this._rows.set(rowNumber, row);
    }
    return row;
  }

  addRow(data: unknown[] | Record<string, unknown>): Row {
    const rowNumber = this.rowCount + 1;
    const row = this.getRow(rowNumber);

    if (Array.isArray(data)) {
      for (let i = 0; i < data.length; i++) {
        if (data[i] !== undefined) {
          row.getCell(i + 1).value = data[i] as CellValue;
        }
      }
    } else {
      // Object with column keys
      for (const [key, value] of Object.entries(data)) {
        const colNum = this._columnKeys.get(key);
        if (colNum !== undefined) {
          row.getCell(colNum).value = value as CellValue;
        }
      }
    }

    return row;
  }

  addRows(rows: (unknown[] | Record<string, unknown>)[]): void {
    for (const data of rows) {
      this.addRow(data);
    }
  }

  eachRow(callback: (row: Row, rowNumber: number) => void): void;
  eachRow(options: { includeEmpty?: boolean }, callback: (row: Row, rowNumber: number) => void): void;
  eachRow(
    optionsOrCallback: { includeEmpty?: boolean } | ((row: Row, rowNumber: number) => void),
    maybeCallback?: (row: Row, rowNumber: number) => void,
  ): void {
    let includeEmpty = false;
    let cb: (row: Row, rowNumber: number) => void;

    if (typeof optionsOrCallback === 'function') {
      cb = optionsOrCallback;
    } else {
      includeEmpty = optionsOrCallback.includeEmpty ?? false;
      cb = maybeCallback!;
    }

    if (includeEmpty) {
      const max = this.rowCount;
      for (let i = 1; i <= max; i++) {
        cb(this.getRow(i), i);
      }
    } else {
      const sorted = [...this._rows.entries()].sort((a, b) => a[0] - b[0]);
      for (const [num, row] of sorted) {
        cb(row, num);
      }
    }
  }

  // --- Cells ---

  getCell(address: string): Cell {
    const { col, row } = parseAddress(address);
    return this.getRow(row).getCell(col);
  }

  // --- Merge ---

  mergeCells(range: string): void {
    const { start, end } = parseRange(range);
    const masterCell = this.getRow(start.row).getCell(start.col);

    this._mergedCells.push({ start, end, range });

    for (let r = start.row; r <= end.row; r++) {
      for (let c = start.col; c <= end.col; c++) {
        const cell = this.getRow(r).getCell(c);
        if (r === start.row && c === start.col) {
          cell._setMerged(null); // master
        } else {
          cell._setMerged(masterCell);
        }
      }
    }
  }

  unMergeCells(range: string): void {
    const { start, end } = parseRange(range);

    this._mergedCells = this._mergedCells.filter((m) => m.range !== range);

    for (let r = start.row; r <= end.row; r++) {
      for (let c = start.col; c <= end.col; c++) {
        const cell = this.getRow(r)._getCellIfExists(c);
        if (cell) {
          cell._setMerged(null);
        }
      }
    }
  }

  get mergedCells(): string[] {
    return this._mergedCells.map((m) => m.range);
  }

  // --- Counts ---

  get rowCount(): number {
    if (this._rows.size === 0) return 0;
    return Math.max(...this._rows.keys());
  }

  get columnCount(): number {
    let max = this._columns.length;
    for (const row of this._rows.values()) {
      row.eachCell((_, col) => {
        if (col > max) max = col;
      });
    }
    return max;
  }

  get actualRowCount(): number {
    return this._rows.size;
  }

  get actualColumnCount(): number {
    const cols = new Set<number>();
    for (const row of this._rows.values()) {
      row.eachCell((_, col) => cols.add(col));
    }
    return cols.size;
  }
}
