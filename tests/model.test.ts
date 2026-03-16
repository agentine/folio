import { describe, it, expect } from 'vitest';
import { Workbook, Worksheet, Row, Cell, ValueType, parseAddress, formatAddress, colToLetter, letterToCol, parseRange } from '../src/model/index.js';

describe('Address utilities', () => {
  it('converts column numbers to letters', () => {
    expect(colToLetter(1)).toBe('A');
    expect(colToLetter(26)).toBe('Z');
    expect(colToLetter(27)).toBe('AA');
    expect(colToLetter(702)).toBe('ZZ');
    expect(colToLetter(703)).toBe('AAA');
  });

  it('converts letters to column numbers', () => {
    expect(letterToCol('A')).toBe(1);
    expect(letterToCol('Z')).toBe(26);
    expect(letterToCol('AA')).toBe(27);
    expect(letterToCol('ZZ')).toBe(702);
    expect(letterToCol('AAA')).toBe(703);
  });

  it('parses cell addresses', () => {
    expect(parseAddress('A1')).toEqual({ col: 1, row: 1 });
    expect(parseAddress('B10')).toEqual({ col: 2, row: 10 });
    expect(parseAddress('AA100')).toEqual({ col: 27, row: 100 });
  });

  it('formats cell addresses', () => {
    expect(formatAddress(1, 1)).toBe('A1');
    expect(formatAddress(2, 10)).toBe('B10');
    expect(formatAddress(27, 100)).toBe('AA100');
  });

  it('parses ranges', () => {
    const range = parseRange('A1:C3');
    expect(range.start).toEqual({ col: 1, row: 1 });
    expect(range.end).toEqual({ col: 3, row: 3 });
  });

  it('throws on invalid address', () => {
    expect(() => parseAddress('123')).toThrow('Invalid cell address');
    expect(() => parseAddress('')).toThrow('Invalid cell address');
  });
});

describe('Workbook', () => {
  it('creates and retrieves worksheets', () => {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('Sheet1');

    expect(sheet.name).toBe('Sheet1');
    expect(sheet.id).toBe(1);
    expect(wb.worksheets).toHaveLength(1);
    expect(wb.getWorksheet('Sheet1')).toBe(sheet);
    expect(wb.getWorksheet(1)).toBe(sheet);
  });

  it('removes worksheets', () => {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('Sheet1');
    wb.removeWorksheet(sheet.id);
    expect(wb.worksheets).toHaveLength(0);
    expect(wb.getWorksheet('Sheet1')).toBeUndefined();
  });

  it('iterates over sheets', () => {
    const wb = new Workbook();
    wb.addWorksheet('A');
    wb.addWorksheet('B');

    const names: string[] = [];
    wb.eachSheet((ws) => names.push(ws.name));
    expect(names).toEqual(['A', 'B']);
  });

  it('sets workbook properties', () => {
    const wb = new Workbook();
    wb.creator = 'Test';
    wb.lastModifiedBy = 'Tester';
    expect(wb.creator).toBe('Test');
    expect(wb.lastModifiedBy).toBe('Tester');
  });
});

describe('Worksheet', () => {
  it('adds rows from array', () => {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('Data');

    sheet.addRow(['Alice', 30, true]);
    sheet.addRow(['Bob', 25, false]);

    expect(sheet.rowCount).toBe(2);
    expect(sheet.getCell('A1').value).toBe('Alice');
    expect(sheet.getCell('B1').value).toBe(30);
    expect(sheet.getCell('C2').value).toBe(false);
  });

  it('adds rows from objects using column keys', () => {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('Data');
    sheet.columns = [
      { header: 'Name', key: 'name', width: 20 },
      { header: 'Age', key: 'age', width: 10 },
    ];

    sheet.addRow({ name: 'Alice', age: 30 });

    // Row 1 has headers, row 2 has data
    expect(sheet.getCell('A1').value).toBe('Name');
    expect(sheet.getCell('A2').value).toBe('Alice');
    expect(sheet.getCell('B2').value).toBe(30);
  });

  it('adds multiple rows at once', () => {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('Data');

    sheet.addRows([
      ['a', 'b'],
      ['c', 'd'],
    ]);

    expect(sheet.rowCount).toBe(2);
    expect(sheet.getCell('A2').value).toBe('c');
  });

  it('gets rows and cells', () => {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('Data');

    const row = sheet.getRow(5);
    expect(row.number).toBe(5);

    const cell = row.getCell(3);
    expect(cell.col).toBe(3);
    expect(cell.row).toBe(5);
    expect(cell.address).toBe('C5');
  });

  it('iterates over rows', () => {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('Data');
    sheet.addRow(['a']);
    sheet.addRow(['b']);

    const rows: number[] = [];
    sheet.eachRow((row, num) => rows.push(num));
    expect(rows).toEqual([1, 2]);
  });

  it('counts rows and columns', () => {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('Data');
    sheet.addRow(['a', 'b', 'c']);
    sheet.addRow(['d']);

    expect(sheet.rowCount).toBe(2);
    expect(sheet.columnCount).toBe(3);
    expect(sheet.actualRowCount).toBe(2);
  });
});

describe('Cell', () => {
  it('detects value types', () => {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('Types');

    sheet.getCell('A1').value = null;
    expect(sheet.getCell('A1').type).toBe(ValueType.Null);

    sheet.getCell('A2').value = 'hello';
    expect(sheet.getCell('A2').type).toBe(ValueType.String);

    sheet.getCell('A3').value = 42;
    expect(sheet.getCell('A3').type).toBe(ValueType.Number);

    sheet.getCell('A4').value = true;
    expect(sheet.getCell('A4').type).toBe(ValueType.Boolean);

    sheet.getCell('A5').value = new Date();
    expect(sheet.getCell('A5').type).toBe(ValueType.Date);

    sheet.getCell('A6').value = { formula: 'SUM(A1:A5)' };
    expect(sheet.getCell('A6').type).toBe(ValueType.Formula);

    sheet.getCell('A7').value = { richText: [{ text: 'bold' }] };
    expect(sheet.getCell('A7').type).toBe(ValueType.RichText);

    sheet.getCell('A8').value = { text: 'click', hyperlink: 'http://example.com' };
    expect(sheet.getCell('A8').type).toBe(ValueType.Hyperlink);

    sheet.getCell('A9').value = { error: '#N/A' };
    expect(sheet.getCell('A9').type).toBe(ValueType.Error);
  });

  it('supports cell styling', () => {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('Styles');
    const cell = sheet.getCell('A1');

    cell.style = {
      font: { name: 'Arial', bold: true, size: 12 },
      alignment: { horizontal: 'center' },
    };

    expect(cell.style.font?.bold).toBe(true);
    expect(cell.style.alignment?.horizontal).toBe('center');
  });
});

describe('Merge cells', () => {
  it('merges and unmerges cells', () => {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('Merge');

    sheet.getCell('A1').value = 'merged';
    sheet.mergeCells('A1:B2');

    expect(sheet.getCell('A1').isMerged).toBe(false); // master is not "merged"
    expect(sheet.getCell('B1').isMerged).toBe(true);
    expect(sheet.getCell('B1').master).toBe(sheet.getCell('A1'));
    expect(sheet.mergedCells).toEqual(['A1:B2']);

    sheet.unMergeCells('A1:B2');
    expect(sheet.getCell('B1').isMerged).toBe(false);
    expect(sheet.mergedCells).toEqual([]);
  });
});

describe('Row', () => {
  it('reads values array (1-based)', () => {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('Data');
    const row = sheet.addRow(['a', 'b', 'c']);

    const vals = row.values;
    // Index 0 is undefined, 1='a', 2='b', 3='c'
    expect(vals[1]).toBe('a');
    expect(vals[2]).toBe('b');
    expect(vals[3]).toBe('c');
  });

  it('iterates cells', () => {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('Data');
    sheet.addRow(['x', 'y']);

    const cells: string[] = [];
    sheet.getRow(1).eachCell((cell) => {
      cells.push(cell.address);
    });
    expect(cells).toEqual(['A1', 'B1']);
  });

  it('sets row properties', () => {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('Data');
    const row = sheet.getRow(1);
    row.height = 20;
    row.hidden = true;
    expect(row.height).toBe(20);
    expect(row.hidden).toBe(true);
  });
});
