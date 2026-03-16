import { describe, it, expect } from 'vitest';
import { StyleRegistry, isDateNumFmt, BUILTIN_NUMFMTS } from '../src/style.js';
import type { Font, Fill, Border, Style } from '../src/model/types.js';

describe('StyleRegistry', () => {
  describe('constructor defaults', () => {
    it('initializes with default font, fills, border, and xf', () => {
      const reg = new StyleRegistry();
      // Default font: Calibri 11pt
      expect(reg.fonts).toHaveLength(1);
      expect(reg.getFont(0)).toEqual({ name: 'Calibri', size: 11 });
      // Two required fills: none and gray125
      expect(reg.fills).toHaveLength(2);
      expect((reg.getFill(0) as any).pattern).toBe('none');
      expect((reg.getFill(1) as any).pattern).toBe('gray125');
      // Default border
      expect(reg.borders).toHaveLength(1);
      // Default xf
      expect(reg.xfs).toHaveLength(1);
      expect(reg.getXf(0)).toEqual({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 });
    });
  });

  describe('font registration', () => {
    it('registers and deduplicates fonts', () => {
      const reg = new StyleRegistry();
      const font: Font = { name: 'Arial', size: 12, bold: true };
      const id1 = reg.registerFont(font);
      const id2 = reg.registerFont({ name: 'Arial', size: 12, bold: true });
      expect(id1).toBe(id2);
      expect(reg.fonts).toHaveLength(2); // default + 1
    });

    it('assigns different ids to different fonts', () => {
      const reg = new StyleRegistry();
      const id1 = reg.registerFont({ name: 'Arial', size: 12 });
      const id2 = reg.registerFont({ name: 'Times New Roman', size: 14 });
      expect(id1).not.toBe(id2);
      expect(reg.fonts).toHaveLength(3); // default + 2
    });

    it('stores all font properties', () => {
      const reg = new StyleRegistry();
      const font: Font = {
        name: 'Arial',
        size: 12,
        bold: true,
        italic: true,
        underline: 'double',
        strike: true,
        color: { argb: 'FFFF0000' },
        family: 2,
        vertAlign: 'superscript',
      };
      const id = reg.registerFont(font);
      const stored = reg.getFont(id);
      expect(stored).toEqual(font);
    });
  });

  describe('fill registration', () => {
    it('registers and deduplicates pattern fills', () => {
      const reg = new StyleRegistry();
      const fill: Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
      const id1 = reg.registerFill(fill);
      const id2 = reg.registerFill({ type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } });
      expect(id1).toBe(id2);
      expect(id1).toBeGreaterThanOrEqual(2); // after default fills
    });

    it('registers gradient fills', () => {
      const reg = new StyleRegistry();
      const fill: Fill = {
        type: 'gradient',
        gradient: 'angle',
        degree: 90,
        stops: [
          { position: 0, color: { argb: 'FFFF0000' } },
          { position: 1, color: { argb: 'FF0000FF' } },
        ],
      };
      const id = reg.registerFill(fill);
      expect(id).toBeGreaterThanOrEqual(2);
    });

    it('differentiates fills by color', () => {
      const reg = new StyleRegistry();
      const id1 = reg.registerFill({ type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } });
      const id2 = reg.registerFill({ type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00FF00' } });
      expect(id1).not.toBe(id2);
    });
  });

  describe('border registration', () => {
    it('registers and deduplicates borders', () => {
      const reg = new StyleRegistry();
      const border: Border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
      };
      const id1 = reg.registerBorder(border);
      const id2 = reg.registerBorder({
        top: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
      });
      expect(id1).toBe(id2);
    });

    it('handles diagonal borders', () => {
      const reg = new StyleRegistry();
      const id = reg.registerBorder({
        diagonal: { style: 'thin', color: { argb: 'FF000000' }, up: true, down: false },
      });
      expect(id).toBeGreaterThan(0);
      expect(reg.getBorder(id)?.diagonal?.up).toBe(true);
    });
  });

  describe('numFmt registration', () => {
    it('returns built-in numFmt IDs for known formats', () => {
      const reg = new StyleRegistry();
      expect(reg.registerNumFmt('General')).toBe(0);
      expect(reg.registerNumFmt('0.00')).toBe(2);
      expect(reg.registerNumFmt('#,##0')).toBe(3);
      expect(reg.registerNumFmt('@')).toBe(49);
    });

    it('assigns custom IDs starting at 164', () => {
      const reg = new StyleRegistry();
      const id = reg.registerNumFmt('#,##0.000');
      expect(id).toBe(164);
    });

    it('deduplicates custom formats', () => {
      const reg = new StyleRegistry();
      const id1 = reg.registerNumFmt('$#,##0.00');
      const id2 = reg.registerNumFmt('$#,##0.00');
      expect(id1).toBe(id2);
    });

    it('looks up format codes by ID', () => {
      const reg = new StyleRegistry();
      expect(reg.getNumFmtCode(0)).toBe('General');
      expect(reg.getNumFmtCode(14)).toBe('mm-dd-yy');
      const id = reg.registerNumFmt('yyyy-mm-dd');
      expect(reg.getNumFmtCode(id)).toBe('yyyy-mm-dd');
    });

    it('tracks custom numFmts separately', () => {
      const reg = new StyleRegistry();
      reg.registerNumFmt('$#,##0.00');
      reg.registerNumFmt('0.0%');
      expect(reg.customNumFmts.size).toBe(2);
    });
  });

  describe('style registration (cellXf)', () => {
    it('registers a composite style and deduplicates', () => {
      const reg = new StyleRegistry();
      const style: Style = {
        font: { name: 'Arial', size: 12, bold: true },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } },
        border: { top: { style: 'thin' } },
        alignment: { horizontal: 'center', wrapText: true },
        numFmt: '#,##0.00',
      };
      const id1 = reg.registerStyle(style);
      const id2 = reg.registerStyle({
        font: { name: 'Arial', size: 12, bold: true },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } },
        border: { top: { style: 'thin' } },
        alignment: { horizontal: 'center', wrapText: true },
        numFmt: '#,##0.00',
      });
      expect(id1).toBe(id2);
      expect(id1).toBeGreaterThan(0);
    });

    it('handles style with protection', () => {
      const reg = new StyleRegistry();
      const id = reg.registerStyle({
        protection: { locked: true, hidden: false },
      });
      const xf = reg.getXf(id);
      expect(xf?.protection).toEqual({ locked: true, hidden: false });
    });

    it('registers empty style as default xf', () => {
      const reg = new StyleRegistry();
      const id = reg.registerStyle({});
      expect(id).toBe(0); // matches default xf
    });
  });

  describe('resolveStyle', () => {
    it('round-trips a full style', () => {
      const reg = new StyleRegistry();
      const style: Style = {
        font: { name: 'Verdana', size: 10, italic: true },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00FF00' } },
        border: {
          top: { style: 'medium', color: { argb: 'FFFF0000' } },
          bottom: { style: 'thin' },
        },
        alignment: { vertical: 'middle', shrinkToFit: true },
        numFmt: '0.00%',
        protection: { locked: true },
      };
      const id = reg.registerStyle(style);
      const resolved = reg.resolveStyle(id);
      expect(resolved.font).toEqual(style.font);
      expect(resolved.fill).toEqual(style.fill);
      expect(resolved.border).toEqual(style.border);
      expect(resolved.alignment).toEqual(style.alignment);
      expect(resolved.numFmt).toBe('0.00%');
      expect(resolved.protection).toEqual(style.protection);
    });

    it('returns empty style for default xf', () => {
      const reg = new StyleRegistry();
      const resolved = reg.resolveStyle(0);
      // Default xf uses default font/fill/border/numfmt — resolveStyle skips defaults
      expect(resolved.font).toBeUndefined();
      expect(resolved.fill).toBeUndefined();
      expect(resolved.border).toBeUndefined();
      expect(resolved.numFmt).toBeUndefined();
    });

    it('returns empty style for invalid xf id', () => {
      const reg = new StyleRegistry();
      expect(reg.resolveStyle(999)).toEqual({});
    });
  });
});

describe('isDateNumFmt', () => {
  it('recognizes built-in date format IDs', () => {
    expect(isDateNumFmt(14)).toBe(true); // mm-dd-yy
    expect(isDateNumFmt(22)).toBe(true); // m/d/yy h:mm
    expect(isDateNumFmt(45)).toBe(true); // mm:ss
  });

  it('rejects non-date built-in IDs', () => {
    expect(isDateNumFmt(0)).toBe(false); // General
    expect(isDateNumFmt(1)).toBe(false); // 0
    expect(isDateNumFmt(4)).toBe(false); // #,##0.00
  });

  it('detects custom date format strings', () => {
    expect(isDateNumFmt(164, 'yyyy-mm-dd')).toBe(true);
    expect(isDateNumFmt(164, 'dd/mm/yyyy hh:mm:ss')).toBe(true);
    expect(isDateNumFmt(164, 'mmm-yy')).toBe(true);
  });

  it('rejects non-date custom format strings', () => {
    expect(isDateNumFmt(164, '#,##0.00')).toBe(false);
    expect(isDateNumFmt(164, '0%')).toBe(false);
    expect(isDateNumFmt(164, '@')).toBe(false);
  });
});

describe('BUILTIN_NUMFMTS', () => {
  it('contains standard formats', () => {
    expect(BUILTIN_NUMFMTS[0]).toBe('General');
    expect(BUILTIN_NUMFMTS[1]).toBe('0');
    expect(BUILTIN_NUMFMTS[14]).toBe('mm-dd-yy');
    expect(BUILTIN_NUMFMTS[49]).toBe('@');
  });
});
