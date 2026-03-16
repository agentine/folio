import { readFile } from 'node:fs/promises';
import type { Readable } from 'node:stream';

import { Workbook } from './model/workbook.js';
import type { CellValue, Font, Fill, Border, BorderStyle, Alignment, Color, Style } from './model/types.js';
import { parseSax } from './xml/parser.js';
import type { XmlAttribute } from './xml/parser.js';
import { ZipReader } from './zip/reader.js';
import { StyleRegistry, isDateNumFmt, BUILTIN_NUMFMTS } from './style.js';
import { SharedStrings } from './shared-strings.js';

// Excel epoch: 1900-01-01, with Lotus 1-2-3 leap year bug
const EXCEL_EPOCH = Date.UTC(1899, 11, 30); // Dec 30, 1899
const MS_PER_DAY = 86400000;

function serialToDate(serial: number): Date {
  // Lotus bug: serials >= 60 are off by 1 because of phantom Feb 29, 1900
  const adjusted = serial >= 60 ? serial - 1 : serial;
  return new Date(EXCEL_EPOCH + adjusted * MS_PER_DAY);
}

function attrMap(attrs: XmlAttribute[]): Record<string, string> {
  const m: Record<string, string> = {};
  for (const a of attrs) m[a.name] = a.value;
  return m;
}

// ── Styles XML Parser ──

function parseStylesXml(xml: string, reg: StyleRegistry): void {
  // We need to parse numFmts, fonts, fills, borders, cellXfs in order
  // and register them so that indices match the file's indices.
  // Since StyleRegistry auto-assigns indices, and the file's defaults
  // match ours, we clear and re-add to ensure alignment.

  const numFmts: Map<number, string> = new Map();
  const fonts: Font[] = [];
  const fills: Fill[] = [];
  const borders: Border[] = [];
  const xfs: { numFmtId: number; fontId: number; fillId: number; borderId: number; alignment?: Alignment; protection?: { locked?: boolean; hidden?: boolean } }[] = [];

  let section = '';
  let currentFont: Font | null = null;
  let currentFill: { type?: string; pattern?: string; fgColor?: Color; bgColor?: Color; gradient?: string; degree?: number; center?: { left: number; top: number }; stops?: { position: number; color: Color }[] } | null = null;
  let currentBorder: { top?: BorderStyle; right?: BorderStyle; bottom?: BorderStyle; left?: BorderStyle; diagonal?: BorderStyle & { up?: boolean; down?: boolean }; diagonalUp?: boolean; diagonalDown?: boolean } | null = null;
  let currentBorderSide = '';
  let currentXf: typeof xfs[0] | null = null;

  parseSax(xml, {
    onOpenTag(name: string, attrs: XmlAttribute[]) {
      const a = attrMap(attrs);

      if (name === 'numFmt') {
        const id = parseInt(a.numFmtId, 10);
        numFmts.set(id, a.formatCode);
      } else if (name === 'fonts') {
        section = 'fonts';
      } else if (name === 'fills') {
        section = 'fills';
      } else if (name === 'borders') {
        section = 'borders';
      } else if (name === 'cellXfs') {
        section = 'cellXfs';
      } else if (section === 'fonts' && name === 'font') {
        currentFont = {};
      } else if (currentFont) {
        applyFontTag(currentFont, name, a);
      } else if (section === 'fills' && name === 'fill') {
        currentFill = {};
      } else if (currentFill && name === 'patternFill') {
        currentFill.type = 'pattern';
        currentFill.pattern = a.patternType || 'none';
      } else if (currentFill && name === 'fgColor') {
        currentFill.fgColor = parseColor(a);
      } else if (currentFill && name === 'bgColor') {
        currentFill.bgColor = parseColor(a);
      } else if (currentFill && name === 'gradientFill') {
        currentFill.type = 'gradient';
        currentFill.gradient = a.type || 'angle';
        if (a.degree) currentFill.degree = parseFloat(a.degree);
        if (a.left && a.top) currentFill.center = { left: parseFloat(a.left), top: parseFloat(a.top) };
        currentFill.stops = [];
      } else if (currentFill && currentFill.stops && name === 'stop') {
        currentFill.stops.push({ position: parseFloat(a.position || '0'), color: {} });
      } else if (currentFill && currentFill.stops && currentFill.stops.length > 0 && name === 'color') {
        const last = currentFill.stops[currentFill.stops.length - 1];
        last.color = parseColor(a);
      } else if (section === 'borders' && name === 'border') {
        currentBorder = {};
        if (a.diagonalUp === '1' || a.diagonalUp === 'true') currentBorder.diagonalUp = true;
        if (a.diagonalDown === '1' || a.diagonalDown === 'true') currentBorder.diagonalDown = true;
      } else if (currentBorder && (name === 'left' || name === 'right' || name === 'top' || name === 'bottom' || name === 'diagonal')) {
        currentBorderSide = name;
        const bs: BorderStyle = {};
        if (a.style) bs.style = a.style as BorderStyle['style'];
        (currentBorder as Record<string, unknown>)[name] = bs;
      } else if (currentBorder && currentBorderSide && name === 'color') {
        const bs = (currentBorder as Record<string, BorderStyle>)[currentBorderSide];
        if (bs) bs.color = parseColor(a);
      } else if (section === 'cellXfs' && name === 'xf') {
        currentXf = {
          numFmtId: parseInt(a.numFmtId || '0', 10),
          fontId: parseInt(a.fontId || '0', 10),
          fillId: parseInt(a.fillId || '0', 10),
          borderId: parseInt(a.borderId || '0', 10),
        };
      } else if (currentXf && name === 'alignment') {
        const aln: Alignment = {};
        if (a.horizontal) aln.horizontal = a.horizontal as Alignment['horizontal'];
        if (a.vertical) aln.vertical = a.vertical as Alignment['vertical'];
        if (a.wrapText === '1' || a.wrapText === 'true') aln.wrapText = true;
        if (a.shrinkToFit === '1' || a.shrinkToFit === 'true') aln.shrinkToFit = true;
        if (a.indent) aln.indent = parseInt(a.indent, 10);
        if (a.textRotation) aln.textRotation = parseInt(a.textRotation, 10);
        currentXf.alignment = aln;
      } else if (currentXf && name === 'protection') {
        const prot: { locked?: boolean; hidden?: boolean } = {};
        if (a.locked !== undefined) prot.locked = a.locked !== '0';
        if (a.hidden !== undefined) prot.hidden = a.hidden !== '0';
        currentXf.protection = prot;
      }
    },
    onCloseTag(name: string) {
      if (section === 'fonts' && name === 'font' && currentFont) {
        fonts.push(currentFont);
        currentFont = null;
      } else if (section === 'fills' && name === 'fill' && currentFill) {
        if (currentFill.type === 'gradient' && currentFill.stops) {
          fills.push({
            type: 'gradient',
            gradient: (currentFill.gradient as 'angle' | 'path') || 'angle',
            degree: currentFill.degree,
            center: currentFill.center,
            stops: currentFill.stops,
          } as Fill);
        } else {
          fills.push({
            type: 'pattern',
            pattern: currentFill.pattern || 'none',
            fgColor: currentFill.fgColor,
            bgColor: currentFill.bgColor,
          } as Fill);
        }
        currentFill = null;
      } else if (section === 'borders' && name === 'border' && currentBorder) {
        const b: Border = {};
        if (currentBorder.top) b.top = currentBorder.top;
        if (currentBorder.right) b.right = currentBorder.right;
        if (currentBorder.bottom) b.bottom = currentBorder.bottom;
        if (currentBorder.left) b.left = currentBorder.left;
        if (currentBorder.diagonal) {
          b.diagonal = { ...currentBorder.diagonal };
          if (currentBorder.diagonalUp) b.diagonal.up = true;
          if (currentBorder.diagonalDown) b.diagonal.down = true;
        }
        borders.push(b);
        currentBorder = null;
      } else if (name === 'left' || name === 'right' || name === 'top' || name === 'bottom' || name === 'diagonal') {
        currentBorderSide = '';
      } else if (section === 'cellXfs' && name === 'xf' && currentXf) {
        xfs.push(currentXf);
        currentXf = null;
      } else if (name === 'fonts' || name === 'fills' || name === 'borders' || name === 'cellXfs') {
        section = '';
      }
    },
  });

  // Store parsed data on the registry using a side-channel for the reader
  (reg as any)._parsedFonts = fonts;
  (reg as any)._parsedFills = fills;
  (reg as any)._parsedBorders = borders;
  (reg as any)._parsedXfs = xfs;
  (reg as any)._parsedNumFmts = numFmts;
}

/** Resolve a style from parsed xf data. */
function resolveStyleFromParsed(reg: StyleRegistry, xfId: number): Style {
  const xfs = (reg as any)._parsedXfs as { numFmtId: number; fontId: number; fillId: number; borderId: number; alignment?: Alignment; protection?: { locked?: boolean; hidden?: boolean } }[] | undefined;
  if (!xfs || !xfs[xfId]) return {};

  const xf = xfs[xfId];
  const fonts = (reg as any)._parsedFonts as Font[];
  const fills = (reg as any)._parsedFills as Fill[];
  const borders = (reg as any)._parsedBorders as Border[];
  const numFmts = (reg as any)._parsedNumFmts as Map<number, string>;

  const style: Style = {};
  if (xf.fontId > 0 && fonts[xf.fontId]) style.font = { ...fonts[xf.fontId] };
  if (xf.fillId > 1 && fills[xf.fillId]) style.fill = { ...fills[xf.fillId] } as Fill;
  if (xf.borderId > 0 && borders[xf.borderId]) style.border = { ...borders[xf.borderId] };
  if (xf.numFmtId !== 0) {
    const code = BUILTIN_NUMFMTS[xf.numFmtId] ?? numFmts.get(xf.numFmtId);
    if (code) style.numFmt = code;
  }
  if (xf.alignment) style.alignment = { ...xf.alignment };
  if (xf.protection) style.protection = { ...xf.protection };
  return style;
}

function isDateXf(reg: StyleRegistry, xfId: number): boolean {
  const xfs = (reg as any)._parsedXfs as { numFmtId: number }[] | undefined;
  if (!xfs || !xfs[xfId]) return false;
  const numFmtId = xfs[xfId].numFmtId;
  const numFmts = (reg as any)._parsedNumFmts as Map<number, string>;
  return isDateNumFmt(numFmtId, numFmts.get(numFmtId));
}

function parseColor(a: Record<string, string>): Color {
  const c: Color = {};
  if (a.rgb) c.argb = a.rgb;
  if (a.theme) c.theme = parseInt(a.theme, 10);
  if (a.tint) c.tint = parseFloat(a.tint);
  return c;
}

function applyFontTag(font: Font, name: string, a: Record<string, string>): void {
  switch (name) {
    case 'b': font.bold = a.val !== '0' && a.val !== 'false'; break;
    case 'i': font.italic = a.val !== '0' && a.val !== 'false'; break;
    case 'strike': font.strike = a.val !== '0' && a.val !== 'false'; break;
    case 'u': font.underline = (a.val as Font['underline']) || true; break;
    case 'vertAlign': if (a.val === 'superscript' || a.val === 'subscript') font.vertAlign = a.val; break;
    case 'sz': if (a.val) font.size = parseFloat(a.val); break;
    case 'color': font.color = parseColor(a); break;
    case 'name': case 'rFont': if (a.val) font.name = a.val; break;
    case 'family': if (a.val) font.family = parseInt(a.val, 10); break;
  }
}

// ── Workbook XML Parser ──

interface SheetInfo {
  name: string;
  sheetId: number;
  rId: string;
}

function parseWorkbookXml(xml: string): SheetInfo[] {
  const sheets: SheetInfo[] = [];
  parseSax(xml, {
    onOpenTag(name: string, attrs: XmlAttribute[]) {
      if (name === 'sheet') {
        const a = attrMap(attrs);
        sheets.push({
          name: a.name,
          sheetId: parseInt(a.sheetId, 10),
          rId: a['r:id'],
        });
      }
    },
  });
  return sheets;
}

// ── Relationships Parser ──

function parseRels(xml: string): Map<string, string> {
  const rels = new Map<string, string>();
  parseSax(xml, {
    onOpenTag(name: string, attrs: XmlAttribute[]) {
      if (name === 'Relationship') {
        const a = attrMap(attrs);
        rels.set(a.Id, a.Target);
      }
    },
  });
  return rels;
}

// ── Worksheet XML Parser ──

function parseWorksheetXml(
  xml: string,
  sst: SharedStrings,
  reg: StyleRegistry,
  workbook: Workbook,
  sheetName: string,
): void {
  const sheet = workbook.addWorksheet(sheetName);

  let inRow = false;
  let currentRow = 0;
  let inCell = false;
  let cellRef = '';
  let cellType = '';
  let cellStyleId = 0;
  let inValue = false;
  let inFormula = false;
  let valueBuf = '';
  let formulaBuf = '';
  const mergedRanges: string[] = [];

  parseSax(xml, {
    onOpenTag(name: string, attrs: XmlAttribute[]) {
      const a = attrMap(attrs);
      if (name === 'row') {
        inRow = true;
        currentRow = parseInt(a.r, 10);
        const row = sheet.getRow(currentRow);
        if (a.ht) {
          row.height = parseFloat(a.ht);
        }
        if (a.hidden === '1' || a.hidden === 'true') {
          row.hidden = true;
        }
      } else if (name === 'c' && inRow) {
        inCell = true;
        cellRef = a.r || '';
        cellType = a.t || '';
        cellStyleId = parseInt(a.s || '0', 10);
        valueBuf = '';
        formulaBuf = '';
      } else if (name === 'v' && inCell) {
        inValue = true;
        valueBuf = '';
      } else if (name === 'f' && inCell) {
        inFormula = true;
        formulaBuf = '';
      } else if (name === 'mergeCell') {
        if (a.ref) mergedRanges.push(a.ref);
      }
    },
    onText(text: string) {
      if (inValue) valueBuf += text;
      if (inFormula) formulaBuf += text;
    },
    onCloseTag(name: string) {
      if (name === 'v') {
        inValue = false;
      } else if (name === 'f') {
        inFormula = false;
      } else if (name === 'c' && inCell) {
        inCell = false;
        if (!cellRef) return;

        const cell = sheet.getCell(cellRef);
        let value: CellValue = null;

        if (formulaBuf) {
          // Formula cell
          const fv: { formula: string; result?: string | number | boolean } = { formula: formulaBuf };
          if (valueBuf) {
            if (cellType === 'b') {
              fv.result = valueBuf === '1';
            } else if (cellType === 'str' || cellType === 's') {
              fv.result = valueBuf;
            } else {
              const n = parseFloat(valueBuf);
              fv.result = isNaN(n) ? valueBuf : n;
            }
          }
          value = fv;
        } else if (cellType === 's') {
          // Shared string
          const idx = parseInt(valueBuf, 10);
          const item = sst.get(idx);
          if (item?.richText) {
            value = { richText: item.richText };
          } else {
            value = item?.text ?? '';
          }
        } else if (cellType === 'b') {
          value = valueBuf === '1';
        } else if (cellType === 'e') {
          value = { error: valueBuf as any };
        } else if (cellType === 'str' || cellType === 'inlineStr') {
          value = valueBuf;
        } else {
          // Number or date
          if (valueBuf) {
            const num = parseFloat(valueBuf);
            if (!isNaN(num) && isDateXf(reg, cellStyleId)) {
              value = serialToDate(num);
            } else {
              value = isNaN(num) ? valueBuf as any : num;
            }
          }
        }

        cell.value = value;

        // Apply style
        if (cellStyleId > 0) {
          const style = resolveStyleFromParsed(reg, cellStyleId);
          if (Object.keys(style).length > 0) {
            cell.style = style;
          }
        }
      } else if (name === 'row') {
        inRow = false;
      }
    },
  });

  // Apply merged cells
  for (const range of mergedRanges) {
    sheet.mergeCells(range);
  }
}

// ── Main Reader ──

export function readXlsx(buffer: Buffer): Workbook {
  const zip = new ZipReader(buffer);
  const workbook = new Workbook();

  // Parse workbook
  const wbXml = zip.extract('xl/workbook.xml').toString('utf8');
  const sheets = parseWorkbookXml(wbXml);

  // Parse workbook relationships
  const wbRelsXml = zip.extract('xl/_rels/workbook.xml.rels').toString('utf8');
  const wbRels = parseRels(wbRelsXml);

  // Parse styles
  const reg = new StyleRegistry();
  try {
    const stylesXml = zip.extract('xl/styles.xml').toString('utf8');
    parseStylesXml(stylesXml, reg);
  } catch {
    // styles.xml may not exist in minimal files
  }

  // Parse shared strings
  let sst = new SharedStrings();
  try {
    const sstXml = zip.extract('xl/sharedStrings.xml').toString('utf8');
    sst = SharedStrings.fromXml(sstXml);
  } catch {
    // sharedStrings.xml may not exist
  }

  // Parse each worksheet
  for (const sheetInfo of sheets) {
    const target = wbRels.get(sheetInfo.rId);
    if (!target) continue;
    const path = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
    try {
      const sheetXml = zip.extract(path).toString('utf8');
      parseWorksheetXml(sheetXml, sst, reg, workbook, sheetInfo.name);
    } catch {
      // Skip sheets that can't be parsed
    }
  }

  return workbook;
}

export async function readXlsxFile(filename: string): Promise<Workbook> {
  const buffer = await readFile(filename);
  return readXlsx(buffer);
}

export async function readXlsxStream(stream: Readable): Promise<Workbook> {
  const chunks: Buffer[] = [];
  for await (const chunk of stream) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  }
  return readXlsx(Buffer.concat(chunks));
}
