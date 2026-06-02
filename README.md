# @royaltics/xlsx-styles

JS Excel file builder and reader with zero dependencies.

Read more in [xlsx-styles](https://github.com/royaltics-solutions/xlsx-styles.git).

> [!NOTE]
> This project is a fork of the original [xldx](https://github.com/yowainwright/xldx.git) project.


## Installation

```bash
pnpm add @royaltics/xlsx-styles
```

## Features

- **Zero dependencies** — no Java, no .NET, no native modules
- **XLSX reader** — read existing `.xlsx` files back into structured data or JSON
- **Pattern-based styling and theming** — zebra stripes, diff-based colors, custom pattern functions
- **Multi-sheet support** — creates, reads, and modifies multiple worksheets
- **Cell styling** — fonts, fills, borders, alignment, number formats
- **Auto column width** — smart width calculation via `max`, `avg`, or `median`
- **Pre-rows and cell merging** — prepend custom rows, merge cells across ranges
- **Freeze panes and grid line control**
- **Column-oriented data** — pass data as record arrays or column objects
- **Full type safety** — TypeScript types for every API surface
- **Works in browsers and node, bun, or deno** — platform-specific entry points
- **< 17KB minified** — tiny footprint
- **Round-trip fidelity** — write → read → write preserves shared strings, booleans, numbers, and dates

## Quick Start

```typescript
import { Xldx } from '@royaltics/xlsx-styles';

// Sample data
const data = [
  { id: 1, name: 'John', age: 30, status: 'active' },
  { id: 2, name: 'Jane', age: 25, status: 'inactive' },
  { id: 3, name: 'Bob', age: 35, status: 'active' }
];

// Create Excel builder
const xlsx = new Xldx(data);

// Create a sheet with columns
xlsx.createSheet(
  { name: 'Users' },
  { key: 'id', header: 'ID', width: 10 },
  { key: 'name', header: 'Name', width: 20 },
  { key: 'age', header: 'Age', width: 10 },
  { key: 'status', header: 'Status', width: 15 }
);

// node, bun, or deno - write to file
await xlsx.write('users.xlsx');

// Browser - trigger download
await xlsx.download('users.xlsx');

// Get raw data
const buffer = await xlsx.toBuffer(); // Buffer
const uint8Array = await xlsx.toUint8Array(); // Uint8Array
const blob = await xlsx.toBlob(); // Browser Blob
```

## Reading XLSX Files

### Read from file (node)

```typescript
import { readFile } from '@royaltics/xlsx-styles/server';

const result = await readFile('users.xlsx');
console.log(result.sheets[0].name);     // "Users"
console.log(result.sheets[0].data);     // raw row arrays
console.log(result.sheets[0].json);     // JSON objects
```

### Read from buffer/bytes

```typescript
import { Xldx } from '@royaltics/xlsx-styles';

const bytes = await fs.readFile('users.xlsx');
const result = await Xldx.read(bytes);

result.sheets.forEach(sheet => {
  console.log(sheet.name, sheet.json);
});
```

### Read options

```typescript
const result = await Xldx.read(bytes, {
  header: ['name', 'age', 'status'],   // custom column keys
  header: 0,                           // use row index 0 as headers
  blankrows: true,                     // skip completely empty rows
  defval: 'N/A',                       // default for empty cells
  range: 100,                          // read only first 100 rows
  range: { s: { r: 0, c: 0 }, e: { r: 50, c: 10 } }, // bounding box
  sheetRows: 50                        // max rows per sheet
});
```

### Low-level reader

```typescript
import { XlsxReader } from '@royaltics/xlsx-styles';

const reader = new XlsxReader(uint8Array);
const result = await reader.read();
```

## Advanced Usage

### Themes

```typescript
import { Xldx, themes } from '@royaltics/xlsx-styles';

const xlsx = new Xldx(data);
xlsx.setTheme(themes.dark);
```

### Pattern-Based Styling

```typescript
xlsx.createSheet(
  { name: 'StyledSheet' },
  {
    key: 'status',
    header: 'Status',
    patterns: {
      bgColorPattern: (context) => {
        if (context.value === 'active') {
          return { fill: { type: 'pattern', pattern: 'solid', fgColor: '90EE90FF' } };
        }
        return null;
      }
    }
  }
);
```

### Zebra Striping

```typescript
xlsx.createSheet(
  { name: 'ZebraSheet' },
  {
    key: 'data',
    patterns: {
      bgColorPattern: 'zebra' // Built-in zebra pattern
    }
  }
);
```

### Pre-rows and Merged Headers

```typescript
xlsx.createSheet(
  {
    name: 'Report',
    preRows: [['Monthly Report', '', '', '']],
    merges: ['A1:F1']  // merge first row across columns
  },
  { key: 'month', header: 'Month' },
  { key: 'amount', header: 'Amount' }
);
```

### Column-Oriented Data

```typescript
const xlsx = new Xldx({
  sheets: [
    {
      name: 'Sales',
      data: {
        product: ['Widget', 'Gadget', 'Doohickey'],
        revenue: [100, 200, 150]
      }
    }
  ]
});
```

## API

### Constructor

```typescript
new Xldx(data: DataRow[] | SheetsData, options?: XldxOptions)
```

### Methods

#### Writing / Export

- `setTheme(theme: ColorTheme): this` — Set the color theme
- `createSheet(options: SheetOptions, ...columns: ColumnDefinition[]): this` — Create a new worksheet
- `createSheets(sheets: Array<{options: SheetOptions; columns: ColumnDefinition[]}>): this` — Create multiple worksheets
- `toUint8Array(): Promise<Uint8Array>` — Generate Excel file as Uint8Array
- `toBuffer(): Promise<Buffer>` — Generate as Buffer (node/bun/deno; import from `xldx/server`)
- `toBlob(): Promise<Blob>` — Generate as Blob (browser; import from `xldx/browser`)
- `download(filename?: string): Promise<void>` — Trigger download (browser) or write (node)
- `write(filePath: string): Promise<void>` — Write to disk (node/bun/deno; import from `xldx/server`)
- `toJSON(): any` — Export workbook structure as JSON

#### Reading / Import

- `static read(data: Uint8Array | Buffer, options?: ReadOptions): Promise<{ sheets: SheetResult[] }>` — Read XLSX file with configurable parsing
- `static fromJSON(json: any): Xldx` — Create an Xldx instance from a JSON structure

#### Runtime Data Access

- `getSheetData(sheet: string | number): SheetDataAPI` — Access a sheet at runtime to read/update rows, columns, and styles

### `read()` return shape

```typescript
{
  sheets: [{
    name: string;          // worksheet name
    data: any[][];         // raw cell values (row-major)
    json: Record<string, any>[];  // converted to objects (if headers resolved)
  }]
}
```

### ReadOptions

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `header` | `string[] \| number` | — | Column headers: pass a string array, or row index to use as headers (0-based) |
| `blankrows` | `boolean` | `false` | Skip entirely empty rows in JSON output |
| `range` | `number \| { s: { r, c }, e: { r, c } }` | — | Limit rows: pass a number (max rows) or a bounding-box object |
| `defval` | `any` | `null` | Default value for empty/undefined cells |
| `raw` | `boolean` | `true` | Return raw values; `false` attempts formatted strings |
| `sheetRows` | `number` | — | Maximum rows to parse per sheet |

### Platform entry points

```typescript
import { Xldx }            from '@royaltics/xlsx-styles';       // auto-detects
import { Xldx }            from '@royaltics/xlsx-styles/browser'; // browser build
import { Xldx, readFile }  from '@royaltics/xlsx-styles/server';  // node build
```

### Export-only utilities

```typescript
import { XlsxWriter, XlsxReader } from '@royaltics/xlsx-styles';

const writer = new XlsxWriter();
writer.addWorksheet('Sheet1', [['A', 'B'], [1, 2]]);
const xlsx = writer.generate(); // Uint8Array

const reader = new XlsxReader(xlsx);
const result = await reader.read();
```

## Development

```bash
# Install dependencies
pnpm install

# Build
pnpm build

# Lint
pnpm lint

# Type check
pnpm typecheck
```

## License

MIT