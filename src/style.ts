import type { Font, Fill, PatternFill, Border, BorderStyle, Alignment, Color, Style } from './model/types.js';

// OOXML built-in number format IDs (subset — Excel defines 0-163)
const BUILTIN_NUMFMTS: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'mm-dd-yy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
};

// numFmt IDs that represent dates/times
const DATE_NUMFMT_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);

export function isDateNumFmt(numFmtId: number, customCode?: string): boolean {
  if (DATE_NUMFMT_IDS.has(numFmtId)) return true;
  if (customCode) {
    // Heuristic: contains date/time tokens but not just text
    const stripped = customCode.replace(/"[^"]*"/g, '').replace(/\\./, '');
    return /[ymdhsYMDHS]/.test(stripped) && !/^[#0.,E+\-;@ %()]*$/.test(stripped);
  }
  return false;
}

function colorKey(c: Color | undefined): string {
  if (!c) return '';
  return `${c.argb ?? ''}|${c.theme ?? ''}|${c.tint ?? ''}`;
}

function fontKey(f: Font): string {
  return [
    f.name ?? '',
    f.size ?? '',
    f.bold ? '1' : '0',
    f.italic ? '1' : '0',
    String(f.underline ?? '0'),
    f.strike ? '1' : '0',
    colorKey(f.color),
    f.family ?? '',
    f.vertAlign ?? '',
  ].join('|');
}

function borderStyleKey(bs: BorderStyle | undefined): string {
  if (!bs) return '';
  return `${bs.style ?? ''}:${colorKey(bs.color)}`;
}

function borderKey(b: Border): string {
  const d = b.diagonal;
  return [
    borderStyleKey(b.top),
    borderStyleKey(b.right),
    borderStyleKey(b.bottom),
    borderStyleKey(b.left),
    d ? `${borderStyleKey(d)}:${d.up ? '1' : '0'}:${d.down ? '1' : '0'}` : '',
  ].join('|');
}

function fillKey(f: Fill): string {
  if (f.type === 'pattern') {
    return `p|${f.pattern}|${colorKey(f.fgColor)}|${colorKey(f.bgColor)}`;
  }
  // gradient
  const stops = f.stops.map((s) => `${s.position}:${colorKey(s.color)}`).join(',');
  return `g|${f.gradient}|${f.degree ?? ''}|${f.center?.left ?? ''}:${f.center?.top ?? ''}|${stops}`;
}

function alignmentKey(a: Alignment): string {
  return [
    a.horizontal ?? '',
    a.vertical ?? '',
    a.wrapText ? '1' : '0',
    a.shrinkToFit ? '1' : '0',
    a.indent ?? '',
    a.textRotation ?? '',
    a.readingOrder ?? '',
  ].join('|');
}

interface CellXf {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  alignment?: Alignment;
  protection?: { locked?: boolean; hidden?: boolean };
}

export class StyleRegistry {
  private _fonts: Font[] = [];
  private _fontIndex: Map<string, number> = new Map();

  private _fills: Fill[] = [];
  private _fillIndex: Map<string, number> = new Map();

  private _borders: Border[] = [];
  private _borderIndex: Map<string, number> = new Map();

  private _numFmts: Map<string, number> = new Map(); // formatCode -> numFmtId
  private _numFmtCodes: Map<number, string> = new Map(); // numFmtId -> formatCode
  private _nextNumFmtId = 164; // custom formats start at 164

  private _xfs: CellXf[] = [];
  private _xfIndex: Map<string, number> = new Map();

  constructor() {
    // OOXML requires these defaults
    // Default font (index 0): Calibri 11pt
    this.registerFont({ name: 'Calibri', size: 11 });
    // Two required fills: index 0 = none, index 1 = gray125
    this.registerFill({ type: 'pattern', pattern: 'none' });
    this.registerFill({ type: 'pattern', pattern: 'gray125' });
    // Default border (index 0): empty
    this.registerBorder({});
    // Default cellXf (index 0)
    this._xfs.push({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 });
    this._xfIndex.set('0|0|0|0||', 0);
  }

  // --- Fonts ---

  registerFont(font: Font): number {
    const key = fontKey(font);
    const existing = this._fontIndex.get(key);
    if (existing !== undefined) return existing;
    const id = this._fonts.length;
    this._fonts.push({ ...font });
    this._fontIndex.set(key, id);
    return id;
  }

  getFont(id: number): Font | undefined {
    return this._fonts[id];
  }

  get fonts(): ReadonlyArray<Font> {
    return this._fonts;
  }

  // --- Fills ---

  registerFill(fill: Fill): number {
    const key = fillKey(fill);
    const existing = this._fillIndex.get(key);
    if (existing !== undefined) return existing;
    const id = this._fills.length;
    this._fills.push({ ...fill } as Fill);
    this._fillIndex.set(key, id);
    return id;
  }

  getFill(id: number): Fill | undefined {
    return this._fills[id];
  }

  get fills(): ReadonlyArray<Fill> {
    return this._fills;
  }

  // --- Borders ---

  registerBorder(border: Border): number {
    const key = borderKey(border);
    const existing = this._borderIndex.get(key);
    if (existing !== undefined) return existing;
    const id = this._borders.length;
    this._borders.push({ ...border });
    this._borderIndex.set(key, id);
    return id;
  }

  getBorder(id: number): Border | undefined {
    return this._borders[id];
  }

  get borders(): ReadonlyArray<Border> {
    return this._borders;
  }

  // --- Number Formats ---

  registerNumFmt(formatCode: string): number {
    // Check built-in formats first
    for (const [id, code] of Object.entries(BUILTIN_NUMFMTS)) {
      if (code === formatCode) return Number(id);
    }
    const existing = this._numFmts.get(formatCode);
    if (existing !== undefined) return existing;
    const id = this._nextNumFmtId++;
    this._numFmts.set(formatCode, id);
    this._numFmtCodes.set(id, formatCode);
    return id;
  }

  getNumFmtCode(id: number): string | undefined {
    return BUILTIN_NUMFMTS[id] ?? this._numFmtCodes.get(id);
  }

  /** Custom number formats (excludes built-ins). */
  get customNumFmts(): ReadonlyMap<number, string> {
    return this._numFmtCodes;
  }

  // --- Cell Formats (XF records) ---

  registerStyle(style: Style): number {
    const fontId = style.font ? this.registerFont(style.font) : 0;
    const fillId = style.fill ? this.registerFill(style.fill) : 0;
    const borderId = style.border ? this.registerBorder(style.border) : 0;
    const numFmtId = style.numFmt ? this.registerNumFmt(style.numFmt) : 0;

    const protKey = style.protection
      ? `${style.protection.locked ?? ''}:${style.protection.hidden ?? ''}`
      : '';
    const alnKey = style.alignment ? alignmentKey(style.alignment) : '';
    const key = `${numFmtId}|${fontId}|${fillId}|${borderId}|${alnKey}|${protKey}`;

    const existing = this._xfIndex.get(key);
    if (existing !== undefined) return existing;

    const id = this._xfs.length;
    const xf: CellXf = { numFmtId, fontId, fillId, borderId };
    if (style.alignment) xf.alignment = { ...style.alignment };
    if (style.protection) xf.protection = { ...style.protection };
    this._xfs.push(xf);
    this._xfIndex.set(key, id);
    return id;
  }

  getXf(id: number): CellXf | undefined {
    return this._xfs[id];
  }

  get xfs(): ReadonlyArray<CellXf> {
    return this._xfs;
  }

  /** Reconstruct a Style object from an xf index. */
  resolveStyle(xfId: number): Style {
    const xf = this._xfs[xfId];
    if (!xf) return {};
    const style: Style = {};
    if (xf.fontId !== 0) {
      const f = this._fonts[xf.fontId];
      if (f) style.font = { ...f };
    }
    if (xf.fillId > 1) {
      // 0=none, 1=gray125 are defaults — only include user fills
      const f = this._fills[xf.fillId];
      if (f) style.fill = { ...f } as Fill;
    }
    if (xf.borderId !== 0) {
      const b = this._borders[xf.borderId];
      if (b) style.border = { ...b };
    }
    if (xf.numFmtId !== 0) {
      const code = this.getNumFmtCode(xf.numFmtId);
      if (code) style.numFmt = code;
    }
    if (xf.alignment) style.alignment = { ...xf.alignment };
    if (xf.protection) style.protection = { ...xf.protection };
    return style;
  }
}

export { BUILTIN_NUMFMTS };
export type { CellXf };
