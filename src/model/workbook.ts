import { Worksheet } from './worksheet.js';
import type { WorksheetOptions } from './worksheet.js';

export class Workbook {
  creator = '';
  lastModifiedBy = '';
  created: Date = new Date();
  modified: Date = new Date();

  private _worksheets: Map<number, Worksheet> = new Map();
  private _nextSheetId = 1;

  get worksheets(): Worksheet[] {
    return [...this._worksheets.values()];
  }

  addWorksheet(name: string, options?: WorksheetOptions): Worksheet {
    const id = this._nextSheetId++;
    const sheet = new Worksheet(id, name, options);
    this._worksheets.set(id, sheet);
    return sheet;
  }

  removeWorksheet(id: number): void {
    this._worksheets.delete(id);
  }

  getWorksheet(nameOrId: string | number): Worksheet | undefined {
    if (typeof nameOrId === 'number') {
      return this._worksheets.get(nameOrId);
    }
    for (const sheet of this._worksheets.values()) {
      if (sheet.name === nameOrId) return sheet;
    }
    return undefined;
  }

  eachSheet(callback: (worksheet: Worksheet, id: number) => void): void {
    for (const [id, sheet] of this._worksheets) {
      callback(sheet, id);
    }
  }

  // Stubs for Phase 2 I/O
  get xlsx(): { readFile: (filename: string) => Promise<void>; writeFile: (filename: string) => Promise<void>; writeBuffer: () => Promise<Buffer> } {
    return {
      readFile: async (_filename: string) => {
        throw new Error('XLSX reading not yet implemented');
      },
      writeFile: async (_filename: string) => {
        throw new Error('XLSX writing not yet implemented');
      },
      writeBuffer: async () => {
        throw new Error('XLSX writing not yet implemented');
      },
    };
  }

  get csv(): { readFile: (filename: string, options?: unknown) => Promise<void>; writeFile: (filename: string, options?: unknown) => Promise<void> } {
    return {
      readFile: async (_filename: string, _options?: unknown) => {
        throw new Error('CSV reading not yet implemented');
      },
      writeFile: async (_filename: string, _options?: unknown) => {
        throw new Error('CSV writing not yet implemented');
      },
    };
  }
}
