import { Cell } from './cell.js';
import type { Style } from './types.js';

export class Row {
  readonly number: number;
  private _cells: Map<number, Cell> = new Map();
  height: number | undefined;
  hidden = false;
  style: Style = {};

  constructor(number: number) {
    this.number = number;
  }

  getCell(col: number): Cell {
    let cell = this._cells.get(col);
    if (!cell) {
      cell = new Cell(this.number, col);
      this._cells.set(col, cell);
    }
    return cell;
  }

  /** @internal */
  _getCellIfExists(col: number): Cell | undefined {
    return this._cells.get(col);
  }

  get values(): (unknown)[] {
    if (this._cells.size === 0) return [];
    const maxCol = Math.max(...this._cells.keys());
    const result: (unknown)[] = new Array(maxCol + 1);
    for (const [col, cell] of this._cells) {
      result[col] = cell.value;
    }
    return result;
  }

  set values(data: unknown[]) {
    // values is 1-based: index 0 is ignored, index 1 -> col 1
    for (let i = 1; i < data.length; i++) {
      if (data[i] !== undefined) {
        this.getCell(i).value = data[i] as import('./types.js').CellValue;
      }
    }
  }

  eachCell(callback: (cell: Cell, colNumber: number) => void): void;
  eachCell(options: { includeEmpty?: boolean }, callback: (cell: Cell, colNumber: number) => void): void;
  eachCell(
    optionsOrCallback: { includeEmpty?: boolean } | ((cell: Cell, colNumber: number) => void),
    maybeCallback?: (cell: Cell, colNumber: number) => void,
  ): void {
    let includeEmpty = false;
    let cb: (cell: Cell, colNumber: number) => void;

    if (typeof optionsOrCallback === 'function') {
      cb = optionsOrCallback;
    } else {
      includeEmpty = optionsOrCallback.includeEmpty ?? false;
      cb = maybeCallback!;
    }

    if (includeEmpty && this._cells.size > 0) {
      const maxCol = Math.max(...this._cells.keys());
      for (let col = 1; col <= maxCol; col++) {
        cb(this.getCell(col), col);
      }
    } else {
      for (const [col, cell] of this._cells) {
        cb(cell, col);
      }
    }
  }

  commit(): void {
    // Placeholder for streaming mode — currently a no-op.
  }
}
