import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/model/workbook.js';
import { writeXlsx } from '../src/xlsx-writer.js';
import { readXlsx } from '../src/xlsx-reader.js';

/** Helper: write then read back a workbook */
function roundTrip(wb: Workbook): Workbook {
  const buf = writeXlsx(wb);
  return readXlsx(buf);
}

describe('XLSX Reader', () => {
  describe('basic round-trip', () => {
    it('reads back sheet names', () => {
      const wb = new Workbook();
      wb.addWorksheet('Alpha');
      wb.addWorksheet('Beta');
      const loaded = roundTrip(wb);
      expect(loaded.worksheets).toHaveLength(2);
      expect(loaded.worksheets[0].name).toBe('Alpha');
      expect(loaded.worksheets[1].name).toBe('Beta');
    });

    it('reads back string cell values', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      sheet.getCell('A1').value = 'hello';
      sheet.getCell('B1').value = 'world';
      const loaded = roundTrip(wb);
      const s = loaded.getWorksheet('S1')!;
      expect(s.getCell('A1').value).toBe('hello');
      expect(s.getCell('B1').value).toBe('world');
    });

    it('reads back number cell values', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      sheet.getCell('A1').value = 42;
      sheet.getCell('A2').value = 3.14159;
      const loaded = roundTrip(wb);
      const s = loaded.getWorksheet('S1')!;
      expect(s.getCell('A1').value).toBe(42);
      expect(s.getCell('A2').value).toBeCloseTo(3.14159);
    });

    it('reads back boolean cell values', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      sheet.getCell('A1').value = true;
      sheet.getCell('A2').value = false;
      const loaded = roundTrip(wb);
      const s = loaded.getWorksheet('S1')!;
      expect(s.getCell('A1').value).toBe(true);
      expect(s.getCell('A2').value).toBe(false);
    });

    it('reads back date cell values', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      const date = new Date(2024, 0, 15); // Jan 15, 2024
      sheet.getCell('A1').value = date;
      const loaded = roundTrip(wb);
      const s = loaded.getWorksheet('S1')!;
      const readDate = s.getCell('A1').value as Date;
      expect(readDate).toBeInstanceOf(Date);
      expect(readDate.getFullYear()).toBe(2024);
      expect(readDate.getMonth()).toBe(0);
      expect(readDate.getDate()).toBe(15);
    });

    it('reads back formula cells', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      sheet.getCell('A1').value = 10;
      sheet.getCell('A2').value = 20;
      sheet.getCell('A3').value = { formula: 'SUM(A1:A2)', result: 30 };
      const loaded = roundTrip(wb);
      const s = loaded.getWorksheet('S1')!;
      const fv = s.getCell('A3').value as { formula: string; result?: number };
      expect(fv.formula).toBe('SUM(A1:A2)');
      expect(fv.result).toBe(30);
    });

    it('reads back error cells', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      sheet.getCell('A1').value = { error: '#DIV/0!' as const };
      const loaded = roundTrip(wb);
      const s = loaded.getWorksheet('S1')!;
      const ev = s.getCell('A1').value as { error: string };
      expect(ev.error).toBe('#DIV/0!');
    });

    it('reads back rich text cells', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      sheet.getCell('A1').value = {
        richText: [
          { text: 'bold', font: { bold: true } },
          { text: ' normal' },
        ],
      };
      const loaded = roundTrip(wb);
      const s = loaded.getWorksheet('S1')!;
      const rv = s.getCell('A1').value as { richText: { text: string; font?: { bold?: boolean } }[] };
      expect(rv.richText).toHaveLength(2);
      expect(rv.richText[0].text).toBe('bold');
      expect(rv.richText[0].font?.bold).toBe(true);
      expect(rv.richText[1].text).toBe(' normal');
    });
  });

  describe('styles round-trip', () => {
    it('reads back font style', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      const cell = sheet.getCell('A1');
      cell.value = 'styled';
      cell.style = { font: { bold: true, name: 'Arial', size: 14 } };
      const loaded = roundTrip(wb);
      const s = loaded.getWorksheet('S1')!;
      const style = s.getCell('A1').style;
      expect(style.font?.bold).toBe(true);
      expect(style.font?.name).toBe('Arial');
      expect(style.font?.size).toBe(14);
    });

    it('reads back fill style', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      const cell = sheet.getCell('A1');
      cell.value = 'filled';
      cell.style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } } };
      const loaded = roundTrip(wb);
      const s = loaded.getWorksheet('S1')!;
      const fill = s.getCell('A1').style.fill;
      expect(fill).toBeDefined();
      expect((fill as any)?.pattern).toBe('solid');
      expect((fill as any)?.fgColor?.argb).toBe('FFFF0000');
    });

    it('reads back border style', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      const cell = sheet.getCell('A1');
      cell.value = 'bordered';
      cell.style = { border: { top: { style: 'thin', color: { argb: 'FF000000' } } } };
      const loaded = roundTrip(wb);
      const s = loaded.getWorksheet('S1')!;
      const border = s.getCell('A1').style.border;
      expect(border?.top?.style).toBe('thin');
      expect(border?.top?.color?.argb).toBe('FF000000');
    });

    it('reads back alignment', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      const cell = sheet.getCell('A1');
      cell.value = 'centered';
      cell.style = { alignment: { horizontal: 'center', wrapText: true } };
      const loaded = roundTrip(wb);
      const s = loaded.getWorksheet('S1')!;
      const aln = s.getCell('A1').style.alignment;
      expect(aln?.horizontal).toBe('center');
      expect(aln?.wrapText).toBe(true);
    });

    it('reads back number format', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      const cell = sheet.getCell('A1');
      cell.value = 1234.56;
      cell.style = { numFmt: '#,##0.00' };
      const loaded = roundTrip(wb);
      const s = loaded.getWorksheet('S1')!;
      expect(s.getCell('A1').style.numFmt).toBe('#,##0.00');
    });
  });

  describe('row properties round-trip', () => {
    it('reads back custom row height', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      const row = sheet.getRow(1);
      row.getCell(1).value = 'tall';
      row.height = 30;
      const loaded = roundTrip(wb);
      const s = loaded.getWorksheet('S1')!;
      expect(s.getRow(1).height).toBe(30);
    });

    it('reads back hidden rows', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      const row = sheet.getRow(1);
      row.getCell(1).value = 'hidden';
      row.hidden = true;
      const loaded = roundTrip(wb);
      const s = loaded.getWorksheet('S1')!;
      expect(s.getRow(1).hidden).toBe(true);
    });
  });

  describe('merged cells round-trip', () => {
    it('reads back merged cell ranges', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      sheet.getCell('A1').value = 'merged';
      sheet.mergeCells('A1:B2');
      const loaded = roundTrip(wb);
      const s = loaded.getWorksheet('S1')!;
      expect(s.mergedCells).toContain('A1:B2');
    });
  });

  describe('multiple sheets', () => {
    it('round-trips workbook with multiple sheets and different data', () => {
      const wb = new Workbook();
      const s1 = wb.addWorksheet('Numbers');
      s1.getCell('A1').value = 1;
      s1.getCell('A2').value = 2;
      const s2 = wb.addWorksheet('Strings');
      s2.getCell('A1').value = 'hello';
      s2.getCell('A2').value = 'world';

      const loaded = roundTrip(wb);
      expect(loaded.worksheets).toHaveLength(2);
      expect(loaded.getWorksheet('Numbers')!.getCell('A1').value).toBe(1);
      expect(loaded.getWorksheet('Strings')!.getCell('A1').value).toBe('hello');
    });
  });

  describe('Workbook API integration', () => {
    it('writeBuffer + readXlsx round-trips correctly', async () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('API');
      sheet.getCell('A1').value = 'test';
      sheet.getCell('B1').value = 42;

      const buf = await wb.xlsx.writeBuffer();
      const loaded = readXlsx(buf);
      expect(loaded.getWorksheet('API')!.getCell('A1').value).toBe('test');
      expect(loaded.getWorksheet('API')!.getCell('B1').value).toBe(42);
    });
  });
});
