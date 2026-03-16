import { writeFile } from 'node:fs/promises';
import type { Writable } from 'node:stream';

import type { Workbook } from './model/workbook.js';
import type { Worksheet } from './model/worksheet.js';
import type { Cell } from './model/cell.js';
import { ValueType } from './model/index.js';
import type { Color, Font, Fill, Border, BorderStyle } from './model/types.js';
import { el, xmlDeclaration } from './xml/builder.js';
import { ZipWriter } from './zip/writer.js';
import { StyleRegistry } from './style.js';
import { SharedStrings } from './shared-strings.js';

// Excel epoch: 1900-01-01, with the Lotus 1-2-3 leap year bug (day 60 = Feb 29, 1900 which doesn't exist)
const EXCEL_EPOCH = Date.UTC(1899, 11, 30); // Dec 30, 1899
const MS_PER_DAY = 86400000;

function dateToSerial(d: Date): number {
  const ms = d.getTime() - EXCEL_EPOCH;
  const serial = ms / MS_PER_DAY;
  // Lotus bug: if serial >= 60, add 1 to skip the phantom Feb 29, 1900
  return serial >= 60 ? serial + 1 : serial;
}

function colorToAttr(c: Color): (parent: ReturnType<typeof el>) => void {
  return (parent) => {
    if (c.argb) parent.attr('rgb', c.argb);
    if (c.theme !== undefined) parent.attr('theme', c.theme);
    if (c.tint !== undefined) parent.attr('tint', c.tint);
  };
}

function writeFontXml(font: Font): ReturnType<typeof el> {
  const f = el('font');
  if (font.bold) f.child(el('b'));
  if (font.italic) f.child(el('i'));
  if (font.strike) f.child(el('strike'));
  if (font.underline) {
    const u = el('u');
    if (typeof font.underline === 'string') u.attr('val', font.underline);
    f.child(u);
  }
  if (font.vertAlign) f.child(el('vertAlign').attr('val', font.vertAlign));
  if (font.size !== undefined) f.child(el('sz').attr('val', font.size));
  if (font.color) {
    const c = el('color');
    colorToAttr(font.color)(c);
    f.child(c);
  }
  if (font.name) f.child(el('name').attr('val', font.name));
  if (font.family !== undefined) f.child(el('family').attr('val', font.family));
  return f;
}

function writeFillXml(fill: Fill): ReturnType<typeof el> {
  const f = el('fill');
  if (fill.type === 'pattern') {
    const pf = el('patternFill').attr('patternType', fill.pattern);
    if (fill.fgColor) {
      const c = el('fgColor');
      colorToAttr(fill.fgColor)(c);
      pf.child(c);
    }
    if (fill.bgColor) {
      const c = el('bgColor');
      colorToAttr(fill.bgColor)(c);
      pf.child(c);
    }
    f.child(pf);
  } else {
    // gradient fill
    const gf = el('gradientFill').attr('type', fill.gradient);
    if (fill.degree !== undefined) gf.attr('degree', fill.degree);
    if (fill.center) {
      gf.attr('left', fill.center.left);
      gf.attr('top', fill.center.top);
    }
    for (const stop of fill.stops) {
      const s = el('stop').attr('position', stop.position);
      const c = el('color');
      colorToAttr(stop.color)(c);
      s.child(c);
      gf.child(s);
    }
    f.child(gf);
  }
  return f;
}

function writeBorderStyleXml(name: string, bs: BorderStyle | undefined): ReturnType<typeof el> | null {
  if (!bs) return null;
  const e = el(name);
  if (bs.style) e.attr('style', bs.style);
  if (bs.color) {
    const c = el('color');
    colorToAttr(bs.color)(c);
    e.child(c);
  }
  return e;
}

function writeBorderXml(border: Border): ReturnType<typeof el> {
  const b = el('border');
  if (border.diagonal) {
    if (border.diagonal.up) b.attr('diagonalUp', true);
    if (border.diagonal.down) b.attr('diagonalDown', true);
  }
  const left = writeBorderStyleXml('left', border.left);
  const right = writeBorderStyleXml('right', border.right);
  const top = writeBorderStyleXml('top', border.top);
  const bottom = writeBorderStyleXml('bottom', border.bottom);
  const diag = writeBorderStyleXml('diagonal', border.diagonal);
  if (left) b.child(left);
  if (right) b.child(right);
  if (top) b.child(top);
  if (bottom) b.child(bottom);
  if (diag) b.child(diag);
  return b;
}

export function writeStylesXml(reg: StyleRegistry): string {
  const ss = el('styleSheet')
    .attr('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

  // numFmts
  if (reg.customNumFmts.size > 0) {
    const nfs = el('numFmts').attr('count', reg.customNumFmts.size);
    for (const [id, code] of reg.customNumFmts) {
      nfs.child(el('numFmt').attr('numFmtId', id).attr('formatCode', code));
    }
    ss.child(nfs);
  }

  // fonts
  const fonts = el('fonts').attr('count', reg.fonts.length);
  for (const font of reg.fonts) {
    fonts.child(writeFontXml(font));
  }
  ss.child(fonts);

  // fills
  const fills = el('fills').attr('count', reg.fills.length);
  for (const fill of reg.fills) {
    fills.child(writeFillXml(fill));
  }
  ss.child(fills);

  // borders
  const borders = el('borders').attr('count', reg.borders.length);
  for (const border of reg.borders) {
    borders.child(writeBorderXml(border));
  }
  ss.child(borders);

  // cellXfs
  const xfs = el('cellXfs').attr('count', reg.xfs.length);
  for (const xf of reg.xfs) {
    const x = el('xf')
      .attr('numFmtId', xf.numFmtId)
      .attr('fontId', xf.fontId)
      .attr('fillId', xf.fillId)
      .attr('borderId', xf.borderId);
    if (xf.fontId !== 0) x.attr('applyFont', 1);
    if (xf.fillId !== 0) x.attr('applyFill', 1);
    if (xf.borderId !== 0) x.attr('applyBorder', 1);
    if (xf.numFmtId !== 0) x.attr('applyNumberFormat', 1);
    if (xf.alignment) {
      x.attr('applyAlignment', 1);
      const a = el('alignment');
      if (xf.alignment.horizontal) a.attr('horizontal', xf.alignment.horizontal);
      if (xf.alignment.vertical) a.attr('vertical', xf.alignment.vertical);
      if (xf.alignment.wrapText) a.attr('wrapText', 1);
      if (xf.alignment.shrinkToFit) a.attr('shrinkToFit', 1);
      if (xf.alignment.indent) a.attr('indent', xf.alignment.indent);
      if (xf.alignment.textRotation !== undefined) a.attr('textRotation', xf.alignment.textRotation);
      x.child(a);
    }
    if (xf.protection) {
      x.attr('applyProtection', 1);
      const p = el('protection');
      if (xf.protection.locked !== undefined) p.attr('locked', xf.protection.locked ? 1 : 0);
      if (xf.protection.hidden !== undefined) p.attr('hidden', xf.protection.hidden ? 1 : 0);
      x.child(p);
    }
    xfs.child(x);
  }
  ss.child(xfs);

  return xmlDeclaration() + ss.toString();
}

function writeContentTypes(sheetCount: number, hasSharedStrings: boolean): string {
  const types = el('Types')
    .attr('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');

  types.child(el('Default').attr('Extension', 'rels').attr('ContentType', 'application/vnd.openxmlformats-package.relationships+xml'));
  types.child(el('Default').attr('Extension', 'xml').attr('ContentType', 'application/xml'));
  types.child(el('Override').attr('PartName', '/xl/workbook.xml').attr('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'));
  types.child(el('Override').attr('PartName', '/xl/styles.xml').attr('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'));

  for (let i = 1; i <= sheetCount; i++) {
    types.child(el('Override').attr('PartName', `/xl/worksheets/sheet${i}.xml`).attr('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'));
  }

  if (hasSharedStrings) {
    types.child(el('Override').attr('PartName', '/xl/sharedStrings.xml').attr('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'));
  }

  return xmlDeclaration() + types.toString();
}

function writeRootRels(): string {
  const rels = el('Relationships')
    .attr('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

  rels.child(el('Relationship')
    .attr('Id', 'rId1')
    .attr('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument')
    .attr('Target', 'xl/workbook.xml'));

  return xmlDeclaration() + rels.toString();
}

function writeWorkbookRels(sheetCount: number, hasSharedStrings: boolean): string {
  const rels = el('Relationships')
    .attr('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

  for (let i = 1; i <= sheetCount; i++) {
    rels.child(el('Relationship')
      .attr('Id', `rId${i}`)
      .attr('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet')
      .attr('Target', `worksheets/sheet${i}.xml`));
  }

  let nextRid = sheetCount + 1;
  rels.child(el('Relationship')
    .attr('Id', `rId${nextRid}`)
    .attr('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles')
    .attr('Target', 'styles.xml'));

  if (hasSharedStrings) {
    nextRid++;
    rels.child(el('Relationship')
      .attr('Id', `rId${nextRid}`)
      .attr('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings')
      .attr('Target', 'sharedStrings.xml'));
  }

  return xmlDeclaration() + rels.toString();
}

function writeWorkbookXml(sheets: { name: string; id: number }[]): string {
  const wb = el('workbook')
    .attr('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    .attr('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

  const sheetsEl = el('sheets');
  for (let i = 0; i < sheets.length; i++) {
    sheetsEl.child(el('sheet')
      .attr('name', sheets[i].name)
      .attr('sheetId', sheets[i].id)
      .attr('r:id', `rId${i + 1}`));
  }
  wb.child(sheetsEl);

  return xmlDeclaration() + wb.toString();
}

function writeCellXml(
  cell: Cell,
  sst: SharedStrings,
  reg: StyleRegistry,
): ReturnType<typeof el> | null {
  const v = cell.value;
  const type = cell.type;

  if (type === ValueType.Null) return null;

  const c = el('c').attr('r', cell.address);

  // Register style if present
  const style = cell.style;
  let styleId = 0;
  if (style && (style.font || style.fill || style.border || style.alignment || style.numFmt || style.protection)) {
    styleId = reg.registerStyle(style);
  }

  // For dates, auto-apply a date numFmt if not already set
  if (type === ValueType.Date && !style.numFmt) {
    styleId = reg.registerStyle({ ...style, numFmt: 'mm-dd-yy' });
  }

  if (styleId !== 0) {
    c.attr('s', styleId);
  }

  switch (type) {
    case ValueType.String: {
      const idx = sst.add(v as string);
      c.attr('t', 's');
      c.child(el('v').text(String(idx)));
      break;
    }
    case ValueType.Number:
      c.child(el('v').text(String(v)));
      break;
    case ValueType.Boolean:
      c.attr('t', 'b');
      c.child(el('v').text(v ? '1' : '0'));
      break;
    case ValueType.Date: {
      const serial = dateToSerial(v as Date);
      c.child(el('v').text(String(serial)));
      break;
    }
    case ValueType.Formula: {
      const fv = v as { formula: string; result?: string | number | boolean };
      c.child(el('f').text(fv.formula));
      if (fv.result !== undefined) {
        if (typeof fv.result === 'string') {
          c.attr('t', 'str');
          c.child(el('v').text(fv.result));
        } else if (typeof fv.result === 'boolean') {
          c.attr('t', 'b');
          c.child(el('v').text(fv.result ? '1' : '0'));
        } else {
          c.child(el('v').text(String(fv.result)));
        }
      }
      break;
    }
    case ValueType.RichText: {
      const rv = v as { richText: { font?: Font; text: string }[] };
      const idx = sst.addRichText(rv.richText);
      c.attr('t', 's');
      c.child(el('v').text(String(idx)));
      break;
    }
    case ValueType.Hyperlink: {
      const hv = v as { text: string; hyperlink: string };
      const idx = sst.add(hv.text);
      c.attr('t', 's');
      c.child(el('v').text(String(idx)));
      break;
    }
    case ValueType.Error: {
      const ev = v as { error: string };
      c.attr('t', 'e');
      c.child(el('v').text(ev.error));
      break;
    }
  }

  return c;
}

function writeWorksheetXml(
  sheet: Worksheet,
  sst: SharedStrings,
  reg: StyleRegistry,
): string {
  const ws = el('worksheet')
    .attr('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    .attr('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

  // Sheet data
  const sheetData = el('sheetData');

  sheet.eachRow((row, rowNum) => {
    const r = el('row').attr('r', rowNum);
    if (row.height !== undefined) {
      r.attr('ht', row.height);
      r.attr('customHeight', 1);
    }
    if (row.hidden) r.attr('hidden', 1);

    row.eachCell((cell, _col) => {
      const cellEl = writeCellXml(cell, sst, reg);
      if (cellEl) r.child(cellEl);
    });

    sheetData.child(r);
  });

  ws.child(sheetData);

  // Merged cells
  const merged = sheet.mergedCells;
  if (merged.length > 0) {
    const mc = el('mergeCells').attr('count', merged.length);
    for (const range of merged) {
      mc.child(el('mergeCell').attr('ref', range));
    }
    ws.child(mc);
  }

  return xmlDeclaration() + ws.toString();
}

export function writeXlsx(workbook: Workbook): Buffer {
  const zip = new ZipWriter();
  const reg = new StyleRegistry();
  const sst = new SharedStrings();

  const sheets = workbook.worksheets;
  const sheetInfos = sheets.map((s) => ({ name: s.name, id: s.id }));

  // Generate worksheet XML (populates sst and reg as side effects)
  const sheetXmls: string[] = [];
  for (const sheet of sheets) {
    sheetXmls.push(writeWorksheetXml(sheet, sst, reg));
  }

  const hasSharedStrings = sst.uniqueCount > 0;

  // Add files to ZIP
  zip.addFile('[Content_Types].xml', writeContentTypes(sheets.length, hasSharedStrings));
  zip.addFile('_rels/.rels', writeRootRels());
  zip.addFile('xl/_rels/workbook.xml.rels', writeWorkbookRels(sheets.length, hasSharedStrings));
  zip.addFile('xl/workbook.xml', writeWorkbookXml(sheetInfos));
  zip.addFile('xl/styles.xml', writeStylesXml(reg));

  for (let i = 0; i < sheetXmls.length; i++) {
    zip.addFile(`xl/worksheets/sheet${i + 1}.xml`, sheetXmls[i]);
  }

  if (hasSharedStrings) {
    zip.addFile('xl/sharedStrings.xml', sst.toXml());
  }

  return zip.toBuffer();
}

export async function writeXlsxFile(workbook: Workbook, filename: string): Promise<void> {
  const buffer = writeXlsx(workbook);
  await writeFile(filename, buffer);
}

export async function writeXlsxStream(workbook: Workbook, stream: Writable): Promise<void> {
  const buffer = writeXlsx(workbook);
  return new Promise((resolve, reject) => {
    stream.write(buffer, (err) => {
      if (err) reject(err);
      else resolve();
    });
  });
}

export async function writeXlsxBuffer(workbook: Workbook): Promise<Buffer> {
  return writeXlsx(workbook);
}
