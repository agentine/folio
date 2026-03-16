import { describe, it, expect } from 'vitest';
import { SharedStrings } from '../src/shared-strings.js';

describe('SharedStrings', () => {
  describe('add and lookup', () => {
    it('adds plain strings and returns indices', () => {
      const sst = new SharedStrings();
      expect(sst.add('hello')).toBe(0);
      expect(sst.add('world')).toBe(1);
      expect(sst.getText(0)).toBe('hello');
      expect(sst.getText(1)).toBe('world');
    });

    it('deduplicates identical strings', () => {
      const sst = new SharedStrings();
      const id1 = sst.add('hello');
      const id2 = sst.add('hello');
      expect(id1).toBe(id2);
      expect(sst.uniqueCount).toBe(1);
      expect(sst.count).toBe(2); // two references
    });

    it('handles empty string', () => {
      const sst = new SharedStrings();
      const id = sst.add('');
      expect(id).toBe(0);
      expect(sst.getText(0)).toBe('');
    });

    it('returns empty string for invalid index', () => {
      const sst = new SharedStrings();
      expect(sst.getText(99)).toBe('');
    });
  });

  describe('rich text', () => {
    it('adds rich text runs', () => {
      const sst = new SharedStrings();
      const runs = [
        { text: 'bold', font: { bold: true } },
        { text: ' normal' },
      ];
      const id = sst.addRichText(runs);
      const item = sst.get(id);
      expect(item?.text).toBe('bold normal');
      expect(item?.richText).toEqual(runs);
    });

    it('deduplicates identical rich text', () => {
      const sst = new SharedStrings();
      const runs = [{ text: 'hello', font: { bold: true } }];
      const id1 = sst.addRichText(runs);
      const id2 = sst.addRichText([{ text: 'hello', font: { bold: true } }]);
      expect(id1).toBe(id2);
      expect(sst.uniqueCount).toBe(1);
      expect(sst.count).toBe(2);
    });

    it('treats plain and rich text as different', () => {
      const sst = new SharedStrings();
      const plainId = sst.add('hello');
      const richId = sst.addRichText([{ text: 'hello' }]);
      expect(plainId).not.toBe(richId);
      expect(sst.uniqueCount).toBe(2);
    });
  });

  describe('count tracking', () => {
    it('tracks total references and unique count', () => {
      const sst = new SharedStrings();
      sst.add('a');
      sst.add('b');
      sst.add('a');
      sst.add('c');
      sst.add('b');
      expect(sst.uniqueCount).toBe(3);
      expect(sst.count).toBe(5);
    });
  });

  describe('XML round-trip', () => {
    it('round-trips plain strings through XML', () => {
      const sst = new SharedStrings();
      sst.add('hello');
      sst.add('world');
      sst.add('hello'); // dup
      const xml = sst.toXml();

      expect(xml).toContain('<?xml');
      expect(xml).toContain('<sst');
      expect(xml).toContain('count="3"');
      expect(xml).toContain('uniqueCount="2"');

      const parsed = SharedStrings.fromXml(xml);
      expect(parsed.uniqueCount).toBe(2);
      expect(parsed.count).toBe(3);
      expect(parsed.getText(0)).toBe('hello');
      expect(parsed.getText(1)).toBe('world');
    });

    it('round-trips rich text through XML', () => {
      const sst = new SharedStrings();
      sst.addRichText([
        { text: 'bold', font: { bold: true, name: 'Arial', size: 12 } },
        { text: ' and italic', font: { italic: true } },
      ]);
      const xml = sst.toXml();

      expect(xml).toContain('<r>');
      expect(xml).toContain('<rPr>');
      expect(xml).toContain('<b/>');

      const parsed = SharedStrings.fromXml(xml);
      expect(parsed.uniqueCount).toBe(1);
      const item = parsed.get(0);
      expect(item?.text).toBe('bold and italic');
      expect(item?.richText).toHaveLength(2);
      expect(item?.richText?.[0].font?.bold).toBe(true);
      expect(item?.richText?.[0].font?.name).toBe('Arial');
      expect(item?.richText?.[1].font?.italic).toBe(true);
    });

    it('preserves whitespace in strings', () => {
      const sst = new SharedStrings();
      sst.add('  leading');
      sst.add('trailing  ');
      const xml = sst.toXml();
      expect(xml).toContain('xml:space="preserve"');

      const parsed = SharedStrings.fromXml(xml);
      expect(parsed.getText(0)).toBe('  leading');
      expect(parsed.getText(1)).toBe('trailing  ');
    });

    it('handles empty shared strings table', () => {
      const sst = new SharedStrings();
      const xml = sst.toXml();
      const parsed = SharedStrings.fromXml(xml);
      expect(parsed.uniqueCount).toBe(0);
      expect(parsed.count).toBe(0);
    });

    it('handles special XML characters in strings', () => {
      const sst = new SharedStrings();
      sst.add('A & B < C > D "E" \'F\'');
      const xml = sst.toXml();

      const parsed = SharedStrings.fromXml(xml);
      expect(parsed.getText(0)).toBe('A & B < C > D "E" \'F\'');
    });

    it('round-trips rich text with full font properties', () => {
      const sst = new SharedStrings();
      sst.addRichText([{
        text: 'styled',
        font: {
          name: 'Times New Roman',
          size: 14,
          bold: true,
          italic: true,
          strike: true,
          underline: 'double',
          vertAlign: 'superscript',
          color: { argb: 'FFFF0000' },
          family: 1,
        },
      }]);
      const xml = sst.toXml();
      const parsed = SharedStrings.fromXml(xml);
      const font = parsed.get(0)?.richText?.[0].font;
      expect(font?.name).toBe('Times New Roman');
      expect(font?.size).toBe(14);
      expect(font?.bold).toBe(true);
      expect(font?.italic).toBe(true);
      expect(font?.strike).toBe(true);
      expect(font?.underline).toBe('double');
      expect(font?.vertAlign).toBe('superscript');
      expect(font?.color?.argb).toBe('FFFF0000');
      expect(font?.family).toBe(1);
    });
  });
});
