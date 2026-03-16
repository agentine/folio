# folio — Drop-in Replacement for exceljs

## Overview

**Replaces:** [exceljs](https://github.com/exceljs/exceljs) (~4.4M weekly npm downloads, 15.2k GitHub stars, 2,264 dependents)

**Package:** `@agentine/folio`

**Why:** exceljs is the only open-source Node.js library for reading/writing styled Excel files (.xlsx). It is unmaintained — last release v4.4.0 was Oct 2023, last code commit Jan 2024, 784 open issues with no maintainer response. Outdated transitive dependencies (rimraf, glob, fstream, inflight) cause deprecation warnings and security audit failures. SheetJS Community Edition deliberately excludes styling (Pro-only). xlsx-populate, excel4node, and xlsx-js-style are also unmaintained. No maintained drop-in replacement exists.

**Language:** TypeScript (strict mode, ESM+CJS dual package)
**Node.js:** 18+
**License:** MIT

---

## Architecture

### Core Model

```
Workbook
  └── Worksheet[]
        ├── Row[]
        │     └── Cell[]
        ├── Column[]
        ├── MergedCells
        ├── DataValidations
        ├── ConditionalFormats
        ├── Tables
        ├── Images
        └── SheetViews
```

### File Format: OOXML (.xlsx)

An .xlsx file is a ZIP archive containing XML files:

- `[Content_Types].xml` — content types
- `_rels/.rels` — package relationships
- `xl/workbook.xml` — workbook definition
- `xl/worksheets/sheetN.xml` — worksheet data
- `xl/styles.xml` — styles (fonts, fills, borders, numFmts)
- `xl/sharedStrings.xml` — shared strings table
- `xl/theme/theme1.xml` — theme definition
- `xl/drawings/` — images and drawings
- `xl/tables/` — table definitions

### Dependencies (minimal)

- **ZIP**: Use Node.js built-in `zlib` for deflate/inflate. Implement ZIP archive read/write using raw `Buffer`/`Uint8Array` operations (the ZIP format is well-documented and straightforward).
- **XML**: Implement a lightweight SAX-style XML parser for reading (~300 lines). Use template-based string building for writing (OOXML schemas are fixed and known).
- **No other dependencies.**

---

## API Surface (exceljs-compatible)

### Workbook

```typescript
const workbook = new Workbook();
workbook.creator = 'Author';
workbook.created = new Date();

// File I/O
await workbook.xlsx.readFile(filename);
await workbook.xlsx.writeFile(filename);
await workbook.xlsx.read(stream);
await workbook.xlsx.write(stream);
const buffer = await workbook.xlsx.writeBuffer();

// CSV
await workbook.csv.readFile(filename, options);
await workbook.csv.writeFile(filename, options);

// Worksheets
const sheet = workbook.addWorksheet('Sheet1', options);
workbook.removeWorksheet(sheet.id);
workbook.getWorksheet('Sheet1');
workbook.eachSheet((worksheet, id) => { ... });
```

### Worksheet

```typescript
// Columns
sheet.columns = [
  { header: 'Id', key: 'id', width: 10 },
  { header: 'Name', key: 'name', width: 32 },
];

// Rows
sheet.addRow({ id: 1, name: 'Alice' });
sheet.addRows([...]);
sheet.getRow(1);
sheet.eachRow((row, rowNumber) => { ... });

// Cells
sheet.getCell('A1').value = 'Hello';
sheet.getCell('B2').value = { formula: 'SUM(A1:A10)' };

// Merged cells
sheet.mergeCells('A1:B2');
sheet.unMergeCells('A1:B2');

// Views
sheet.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];

// Page setup
sheet.pageSetup = { orientation: 'landscape', fitToPage: true };
```

### Cell Styling

```typescript
cell.style = {
  font: { name: 'Arial', size: 12, bold: true, color: { argb: 'FF0000FF' } },
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } },
  border: {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' },
  },
  alignment: { vertical: 'middle', horizontal: 'center', wrapText: true },
  numFmt: '#,##0.00',
};
```

### Data Validation

```typescript
sheet.getCell('A1').dataValidation = {
  type: 'list',
  allowBlank: true,
  formulae: ['"One,Two,Three"'],
};
```

### Conditional Formatting

```typescript
sheet.addConditionalFormatting({
  ref: 'A1:A10',
  rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: [100], style: { font: { bold: true } } }],
});
```

### Images

```typescript
const imageId = workbook.addImage({ filename: 'logo.png', extension: 'png' });
sheet.addImage(imageId, 'B2:D6');
```

### Streaming

```typescript
// Streaming write (for large files)
const stream = workbook.xlsx.createOutputStream(options);

// Streaming read
const workbookReader = new WorkbookReader();
for await (const worksheetReader of workbookReader.read(stream)) {
  for await (const row of worksheetReader) {
    // process row
  }
}
```

### Compatibility Layer

```typescript
// folio/compat/exceljs — namespace-compatible import
import ExcelJS from '@agentine/folio/compat/exceljs';
const workbook = new ExcelJS.Workbook();
```

---

## Implementation Phases

### Phase 1: Core Infrastructure
- Project scaffolding (TypeScript, ESM+CJS, tsconfig, package.json)
- ZIP archive reader/writer (using Node.js zlib)
- Lightweight XML SAX parser
- XML string builder for OOXML output
- Core model classes: Workbook, Worksheet, Row, Cell, Column

### Phase 2: Read/Write .xlsx
- OOXML reader: parse workbook.xml, worksheet XML, sharedStrings, styles
- OOXML writer: generate all required XML files, assemble ZIP
- Cell value types: string, number, date, boolean, formula, richText, hyperlink, error
- Style model: Font, Fill, Border, Alignment, NumFmt
- Style deduplication (shared styles.xml)
- SharedStrings table (read/write)
- Merged cells

### Phase 3: Formatting and Features
- Column properties (width, style, hidden, outlineLevel)
- Row properties (height, style, hidden, outlineLevel)
- Sheet views (frozen panes, split panes, zoom)
- Page setup (orientation, margins, fitToPage, print area)
- Data validation
- Conditional formatting
- Defined names (named ranges)
- Auto-filter

### Phase 4: Media, Tables, and Advanced
- Image embedding (PNG, JPEG, GIF)
- Drawing relationships and positioning (oneCellAnchor, twoCellAnchor)
- Table definitions
- CSV read/write
- Workbook properties (creator, created, modified, etc.)
- Sheet protection
- Comments/notes

### Phase 5: Streaming, Compat, and Release
- Streaming writer (row-by-row .xlsx generation for large files)
- Streaming reader (async iterator over rows)
- ExcelJS compatibility layer (`folio/compat/exceljs`)
- Full test suite (unit + integration, test against exceljs-generated files)
- CI/CD pipeline
- Documentation and README
- npm publish as `@agentine/folio`

---

## Key Design Decisions

1. **Zero external dependencies.** ZIP and XML handling implemented internally using Node.js built-in `zlib`. Keeps the package small, eliminates supply chain risk, and avoids the deprecated transitive dependency problem that plagues exceljs.

2. **TypeScript-first.** Strict mode, full type coverage. No separate `@types/` package needed. Generics for cell value types.

3. **ESM+CJS dual package.** Conditional exports in package.json. Works with both `import` and `require()`.

4. **Style deduplication.** OOXML requires shared style definitions in `xl/styles.xml`. Internal style registry deduplicates identical styles to minimize file size.

5. **Streaming architecture.** Both reader and writer support streaming for memory-efficient handling of large files (100K+ rows).

6. **Drop-in compatibility.** API mirrors exceljs exactly. The `folio/compat/exceljs` entry point provides namespace-compatible imports for zero-change migration.
