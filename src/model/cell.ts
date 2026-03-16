import { formatAddress } from './address.js';
import type { CellValue, DataValidation, Style } from './types.js';
import { ValueType } from './types.js';

export class Cell {
  readonly row: number;
  readonly col: number;
  private _value: CellValue = null;
  private _style: Style = {};
  private _isMerged = false;
  private _master: Cell | null = null;
  dataValidation: DataValidation | undefined;

  constructor(row: number, col: number) {
    this.row = row;
    this.col = col;
  }

  get address(): string {
    return formatAddress(this.col, this.row);
  }

  get value(): CellValue {
    return this._value;
  }

  set value(v: CellValue) {
    this._value = v;
  }

  get type(): ValueType {
    const v = this._value;
    if (v === null || v === undefined) return ValueType.Null;
    if (typeof v === 'string') return ValueType.String;
    if (typeof v === 'number') return ValueType.Number;
    if (typeof v === 'boolean') return ValueType.Boolean;
    if (v instanceof Date) return ValueType.Date;
    if (typeof v === 'object') {
      if ('formula' in v) return ValueType.Formula;
      if ('richText' in v) return ValueType.RichText;
      if ('hyperlink' in v) return ValueType.Hyperlink;
      if ('error' in v) return ValueType.Error;
    }
    return ValueType.Null;
  }

  get style(): Style {
    return this._style;
  }

  set style(s: Style) {
    this._style = s;
  }

  get isMerged(): boolean {
    return this._isMerged;
  }

  get master(): Cell {
    return this._master ?? this;
  }

  /** @internal */
  _setMerged(master: Cell | null): void {
    this._isMerged = master !== null;
    this._master = master;
  }
}
