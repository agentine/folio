import type { Style } from './types.js';

export class Column {
  readonly number: number;
  key: string | undefined;
  header: string | undefined;
  width: number | undefined;
  style: Style = {};
  hidden = false;
  outlineLevel = 0;

  constructor(number: number, opts?: { key?: string; header?: string; width?: number; style?: Style; hidden?: boolean; outlineLevel?: number }) {
    this.number = number;
    if (opts) {
      this.key = opts.key;
      this.header = opts.header;
      this.width = opts.width;
      if (opts.style) this.style = opts.style;
      if (opts.hidden !== undefined) this.hidden = opts.hidden;
      if (opts.outlineLevel !== undefined) this.outlineLevel = opts.outlineLevel;
    }
  }
}
