import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/model/workbook.js';
import { writeXlsx } from '../src/xlsx-writer.js';
import { ZipReader } from '../src/zip/reader.js';
import { parseSax } from '../src/xml/parser.js';
import type { XmlAttribute } from '../src/xml/parser.js';

/** Helper: extract a file from the xlsx buffer */
function extractFile(buffer: Buffer, filename: string): string {
  const reader = new ZipReader(buffer);
  return reader.extract(filename).toString('utf8');
}

/** Helper: collect open tags from XML */
function collectTags(xml: string): { name: string; attrs: Record<string, string> }[] {
  const tags: { name: string; attrs: Record<string, string> }[] = [];
  parseSax(xml, {
    onOpenTag(name: string, attrs: XmlAttribute[]) {
      const a: Record<string, string> = {};
      for (const attr of attrs) a[attr.name] = attr.value;
      tags.push({ name, attrs: a });
    },
  });
  return tags;
}

describe('XLSX Writer', () => {
  describe('ZIP structure', () => {
    it('generates a valid ZIP with required OOXML parts', () => {
      const wb = new Workbook();
      wb.addWorksheet('Sheet1');
      const buf = writeXlsx(wb);

      const reader = new ZipReader(buf);
      const filenames = reader.getEntries().map((e) => e.filename);

      expect(filenames).toContain('[Content_Types].xml');
      expect(filenames).toContain('_rels/.rels');
      expect(filenames).toContain('xl/_rels/workbook.xml.rels');
      expect(filenames).toContain('xl/workbook.xml');
      expect(filenames).toContain('xl/styles.xml');
      expect(filenames).toContain('xl/worksheets/sheet1.xml');
    });

    it('includes sharedStrings.xml when strings are present', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('Sheet1');
      sheet.getCell('A1').value = 'hello';
      const buf = writeXlsx(wb);

      const reader = new ZipReader(buf);
      const filenames = reader.getEntries().map((e) => e.filename);
      expect(filenames).toContain('xl/sharedStrings.xml');
    });

    it('omits sharedStrings.xml when no strings', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('Sheet1');
      sheet.getCell('A1').value = 42;
      const buf = writeXlsx(wb);

      const reader = new ZipReader(buf);
      const filenames = reader.getEntries().map((e) => e.filename);
      expect(filenames).not.toContain('xl/sharedStrings.xml');
    });
  });

  describe('Content Types', () => {
    it('declares correct content types', () => {
      const wb = new Workbook();
      wb.addWorksheet('S1');
      wb.addWorksheet('S2');
      const buf = writeXlsx(wb);
      const xml = extractFile(buf, '[Content_Types].xml');

      expect(xml).toContain('spreadsheetml.sheet.main');
      expect(xml).toContain('spreadsheetml.styles');
      expect(xml).toContain('/xl/worksheets/sheet1.xml');
      expect(xml).toContain('/xl/worksheets/sheet2.xml');
    });
  });

  describe('Workbook XML', () => {
    it('lists sheets with correct names and rIds', () => {
      const wb = new Workbook();
      wb.addWorksheet('Data');
      wb.addWorksheet('Summary');
      const buf = writeXlsx(wb);
      const xml = extractFile(buf, 'xl/workbook.xml');

      const tags = collectTags(xml).filter((t) => t.name === 'sheet');
      expect(tags).toHaveLength(2);
      expect(tags[0].attrs.name).toBe('Data');
      expect(tags[0].attrs['r:id']).toBe('rId1');
      expect(tags[1].attrs.name).toBe('Summary');
      expect(tags[1].attrs['r:id']).toBe('rId2');
    });
  });

  describe('Cell values', () => {
    it('writes string cells as shared string references', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      sheet.getCell('A1').value = 'hello';
      sheet.getCell('A2').value = 'world';
      const buf = writeXlsx(wb);

      const sheetXml = extractFile(buf, 'xl/worksheets/sheet1.xml');
      const tags = collectTags(sheetXml);
      const cells = tags.filter((t) => t.name === 'c');
      expect(cells[0].attrs.t).toBe('s');
      expect(cells[1].attrs.t).toBe('s');

      // Verify shared strings
      const sstXml = extractFile(buf, 'xl/sharedStrings.xml');
      expect(sstXml).toContain('hello');
      expect(sstXml).toContain('world');
    });

    it('writes number cells without type attribute', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      sheet.getCell('A1').value = 42;
      sheet.getCell('A2').value = 3.14;
      const buf = writeXlsx(wb);

      const xml = extractFile(buf, 'xl/worksheets/sheet1.xml');
      const tags = collectTags(xml);
      const cells = tags.filter((t) => t.name === 'c');
      expect(cells[0].attrs.t).toBeUndefined();
      expect(cells[1].attrs.t).toBeUndefined();
      // Check values
      expect(xml).toContain('>42<');
      expect(xml).toContain('>3.14<');
    });

    it('writes boolean cells with t="b"', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      sheet.getCell('A1').value = true;
      sheet.getCell('A2').value = false;
      const buf = writeXlsx(wb);

      const xml = extractFile(buf, 'xl/worksheets/sheet1.xml');
      const tags = collectTags(xml);
      const cells = tags.filter((t) => t.name === 'c');
      expect(cells[0].attrs.t).toBe('b');
      expect(cells[1].attrs.t).toBe('b');
    });

    it('writes date cells as serial numbers with date format', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      sheet.getCell('A1').value = new Date(2024, 0, 1); // Jan 1, 2024
      const buf = writeXlsx(wb);

      const xml = extractFile(buf, 'xl/worksheets/sheet1.xml');
      const tags = collectTags(xml);
      const cell = tags.find((t) => t.name === 'c');
      // Should have a style (for date numFmt)
      expect(cell?.attrs.s).toBeDefined();
    });

    it('writes formula cells with <f> element', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      sheet.getCell('A1').value = { formula: 'SUM(B1:B10)', result: 100 };
      const buf = writeXlsx(wb);

      const xml = extractFile(buf, 'xl/worksheets/sheet1.xml');
      expect(xml).toContain('<f>SUM(B1:B10)</f>');
      expect(xml).toContain('>100<');
    });

    it('writes error cells with t="e"', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      sheet.getCell('A1').value = { error: '#DIV/0!' as const };
      const buf = writeXlsx(wb);

      const xml = extractFile(buf, 'xl/worksheets/sheet1.xml');
      const tags = collectTags(xml);
      const cell = tags.find((t) => t.name === 'c');
      expect(cell?.attrs.t).toBe('e');
    });

    it('writes rich text cells via shared strings', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      sheet.getCell('A1').value = {
        richText: [
          { text: 'bold', font: { bold: true } },
          { text: ' text' },
        ],
      };
      const buf = writeXlsx(wb);

      const sstXml = extractFile(buf, 'xl/sharedStrings.xml');
      expect(sstXml).toContain('<r>');
      expect(sstXml).toContain('<b/>');
    });

    it('writes hyperlink cells as shared strings', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      sheet.getCell('A1').value = { text: 'Click me', hyperlink: 'https://example.com' };
      const buf = writeXlsx(wb);

      const sstXml = extractFile(buf, 'xl/sharedStrings.xml');
      expect(sstXml).toContain('Click me');
    });
  });

  describe('Styles', () => {
    it('writes styles.xml with default styles', () => {
      const wb = new Workbook();
      wb.addWorksheet('S1');
      const buf = writeXlsx(wb);

      const xml = extractFile(buf, 'xl/styles.xml');
      expect(xml).toContain('<fonts');
      expect(xml).toContain('<fills');
      expect(xml).toContain('<borders');
      expect(xml).toContain('<cellXfs');
    });

    it('applies font style to cell', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      const cell = sheet.getCell('A1');
      cell.value = 42;
      cell.style = { font: { bold: true, name: 'Arial', size: 14 } };
      const buf = writeXlsx(wb);

      const stylesXml = extractFile(buf, 'xl/styles.xml');
      expect(stylesXml).toContain('<b/>');
      expect(stylesXml).toContain('val="Arial"');

      const sheetXml = extractFile(buf, 'xl/worksheets/sheet1.xml');
      const tags = collectTags(sheetXml);
      const cell2 = tags.find((t) => t.name === 'c');
      expect(cell2?.attrs.s).toBeDefined();
      expect(Number(cell2?.attrs.s)).toBeGreaterThan(0);
    });
  });

  describe('Merged cells', () => {
    it('writes mergeCells element', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      sheet.getCell('A1').value = 'merged';
      sheet.mergeCells('A1:B2');
      const buf = writeXlsx(wb);

      const xml = extractFile(buf, 'xl/worksheets/sheet1.xml');
      expect(xml).toContain('<mergeCells');
      expect(xml).toContain('ref="A1:B2"');
    });
  });

  describe('Row properties', () => {
    it('writes custom row height', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      const row = sheet.getRow(1);
      row.getCell(1).value = 'tall';
      row.height = 30;
      const buf = writeXlsx(wb);

      const xml = extractFile(buf, 'xl/worksheets/sheet1.xml');
      expect(xml).toContain('ht="30"');
      expect(xml).toContain('customHeight="1"');
    });

    it('writes hidden rows', () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S1');
      const row = sheet.getRow(1);
      row.getCell(1).value = 'hidden';
      row.hidden = true;
      const buf = writeXlsx(wb);

      const xml = extractFile(buf, 'xl/worksheets/sheet1.xml');
      expect(xml).toContain('hidden="1"');
    });
  });

  describe('Multiple sheets', () => {
    it('generates separate worksheet files for each sheet', () => {
      const wb = new Workbook();
      const s1 = wb.addWorksheet('First');
      const s2 = wb.addWorksheet('Second');
      s1.getCell('A1').value = 'sheet1';
      s2.getCell('A1').value = 'sheet2';
      const buf = writeXlsx(wb);

      const reader = new ZipReader(buf);
      const filenames = reader.getEntries().map((e) => e.filename);
      expect(filenames).toContain('xl/worksheets/sheet1.xml');
      expect(filenames).toContain('xl/worksheets/sheet2.xml');

      const xml1 = extractFile(buf, 'xl/worksheets/sheet1.xml');
      const xml2 = extractFile(buf, 'xl/worksheets/sheet2.xml');
      // Both should be valid XML
      expect(xml1).toContain('<?xml');
      expect(xml2).toContain('<?xml');
    });
  });

  describe('Workbook API integration', () => {
    it('writeBuffer returns a valid xlsx buffer', async () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('Test');
      sheet.getCell('A1').value = 'API test';
      const buf = await wb.xlsx.writeBuffer();

      expect(buf).toBeInstanceOf(Buffer);
      const reader = new ZipReader(buf);
      expect(reader.getEntries().length).toBeGreaterThan(0);
    });
  });
});
