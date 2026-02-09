# Excel Viewer Component - Implementation Guide

A reusable, generic Angular 14 component for displaying Excel files with support for images and dropdown validations.

## Table of Contents

- [Quick Start](#quick-start)
- [Installation](#installation)
- [Usage](#usage)
- [Configuration](#configuration)
- [Events](#events)
- [Public API](#public-api)
- [Architecture](#architecture)
- [Limitations](#limitations)

---

## Quick Start

```html
<app-excel-viewer
  [url]="'https://example.com/file.xlsx'"
  [config]="{ enableImages: true, enableDataValidation: true }"
  (loaded)="onLoaded($event)"
  (error)="onError($event)"
></app-excel-viewer>
```

---

## Installation

### 1. Install Dependencies

```bash
npm install @univerjs/presets @univerjs/preset-sheets-core @univerjs/preset-sheets-drawing @univerjs/preset-sheets-data-validation @zwight/luckyexcel exceljs
```

### 2. Update `angular.json`

Add the required CSS files to your styles array:

```json
{
  "styles": [
    "node_modules/@univerjs/preset-sheets-core/lib/index.css",
    "node_modules/@univerjs/preset-sheets-drawing/lib/index.css",
    "node_modules/@univerjs/preset-sheets-data-validation/lib/index.css",
    "src/styles.css"
  ],
  "allowedCommonJsDependencies": [
    "react-dom/client",
    "@univerjs/engine-render",
    "@zwight/exceljs",
    "dayjs",
    "papaparse",
    "exceljs"
  ]
}
```

### 3. Import the Module

```typescript
import { ExcelViewerModule } from './components/excel-viewer';

@NgModule({
  imports: [
    ExcelViewerModule,
    // ...
  ],
})
export class AppModule {}
```

---

## Usage

### Load from URL

```html
<app-excel-viewer [url]="excelUrl"></app-excel-viewer>
```

### Load from File (e.g., file input)

```html
<input type="file" (change)="onFileSelected($event)" accept=".xlsx,.xls" />
<app-excel-viewer [file]="selectedFile"></app-excel-viewer>
```

```typescript
selectedFile: File | null = null;

onFileSelected(event: Event): void {
  const input = event.target as HTMLInputElement;
  if (input.files?.length) {
    this.selectedFile = input.files[0];
  }
}
```

### Load from ArrayBuffer

```typescript
<app-excel-viewer [data]="excelArrayBuffer"></app-excel-viewer>
```

### Multiple Viewers on Same Page

Use unique `containerId` for each instance:

```html
<app-excel-viewer [url]="url1" containerId="viewer-1"></app-excel-viewer>
<app-excel-viewer [url]="url2" containerId="viewer-2"></app-excel-viewer>
```

---

## Configuration

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `enableImages` | `boolean` | `true` | Extract and display embedded images |
| `enableDataValidation` | `boolean` | `true` | Extract and apply dropdown lists |
| `showToolbar` | `boolean` | `true` | Show/hide the toolbar |
| `showFormulaBar` | `boolean` | `true` | Show/hide the formula bar |
| `showSheetTabs` | `boolean` | `true` | Show/hide sheet tabs |
| `editable` | `boolean` | `true` | Enable/disable editing |
| `locale` | `string` | `'en-US'` | UI language (`'en-US'`, `'zh-CN'`) |
| `zoom` | `number` | `100` | Initial zoom level (percentage) |
| `insertDelay` | `number` | `500` | Delay (ms) before inserting images/validations |

### Example

```typescript
config: ExcelViewerConfig = {
  enableImages: true,
  enableDataValidation: true,
  editable: false,  // Read-only mode
  locale: 'en-US',
  insertDelay: 300,
};
```

---

## Events

### `loaded`

Emitted when the Excel file is successfully loaded.

```typescript
interface ExcelLoadedEvent {
  sheetCount: number;      // Number of sheets
  sheetNames: string[];    // Names of all sheets
  imageCount: number;      // Number of images extracted
  validationCount: number; // Number of validations extracted
}
```

### `cellSelected`

Emitted when a cell is selected.

```typescript
interface CellSelectionEvent {
  sheetName: string;  // Current sheet name
  row: number;        // Row index (0-based)
  col: number;        // Column index (0-based)
  address: string;    // Cell address (e.g., "A1")
  value: any;         // Cell value
}
```

### `cellChanged`

Emitted when a cell value changes.

```typescript
interface CellChangeEvent {
  sheetName: string;
  row: number;
  col: number;
  address: string;
  oldValue: any;
  newValue: any;
}
```

### `error`

Emitted on any error.

```typescript
interface ExcelErrorEvent {
  type: 'load' | 'parse' | 'image' | 'validation' | 'unknown';
  message: string;
  error?: any;
}
```

### `loadingChange`

Emitted when loading state changes.

```typescript
(loadingChange)="onLoadingChange($event)"  // boolean
```

---

## Public API

Access the component instance via `@ViewChild`:

```typescript
@ViewChild(ExcelViewerComponent) excelViewer!: ExcelViewerComponent;
```

### Methods

| Method | Parameters | Returns | Description |
|--------|------------|---------|-------------|
| `getCellValue` | `row, col, sheetIndex?` | `any` | Get cell value |
| `setCellValue` | `row, col, value, sheetIndex?` | `void` | Set cell value |
| `getSelectedRange` | - | `any` | Get currently selected range |
| `getSheetNames` | - | `string[]` | Get all sheet names |
| `setActiveSheet` | `index` | `void` | Switch to sheet by index |
| `exportAsJson` | - | `any` | Export workbook as JSON |
| `reload` | - | `void` | Reload the Excel file |
| `dispose` | - | `void` | Clean up resources |

### Example

```typescript
// Read a cell value
const value = this.excelViewer.getCellValue(0, 0); // A1

// Update a cell
this.excelViewer.setCellValue(0, 0, 'New Value');

// Get all sheet names
const sheets = this.excelViewer.getSheetNames();

// Switch sheets
this.excelViewer.setActiveSheet(1); // Second sheet

// Export
const json = this.excelViewer.exportAsJson();
```

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                        Excel File (.xlsx)                           │
└─────────────────────────────────────────────────────────────────────┘
                                 │
                    ┌────────────┴────────────┐
                    ▼                         ▼
          ┌─────────────────┐       ┌─────────────────┐
          │   LuckyExcel    │       │    ExcelJS      │
          │                 │       │                 │
          │ • Cell data     │       │ • Images        │
          │ • Formatting    │       │ • Dropdowns     │
          │ • Formulas      │       │                 │
          │ • Merged cells  │       │                 │
          └────────┬────────┘       └────────┬────────┘
                   │                         │
                   └────────────┬────────────┘
                                ▼
                    ┌─────────────────────┐
                    │      Univerjs       │
                    │  (Spreadsheet UI)   │
                    └─────────────────────┘
```

### Why Two Libraries?

| Feature | LuckyExcel | ExcelJS |
|---------|------------|---------|
| Cell data & formatting | ✅ | ✅ |
| Formulas | ✅ | ✅ |
| Merged cells | ✅ | ✅ |
| **Images** | ❌ | ✅ |
| **Dropdowns** | ❌ | ✅ |
| Univer format output | ✅ | ❌ |

**LuckyExcel** converts Excel to Univer's format but doesn't handle images/dropdowns.
**ExcelJS** extracts images and validations but doesn't output Univer format.
**Together**, they provide complete Excel rendering.

### Data Flow

1. **Fetch** - Download Excel file as ArrayBuffer
2. **Extract** - ExcelJS extracts images and data validations
3. **Convert** - LuckyExcel converts to Univer format
4. **Render** - Univer displays the spreadsheet
5. **Enhance** - Images and dropdowns are added via Univer's Facade API

---

## Limitations

### Images
- Image positioning is approximate (based on cell coordinates)
- Very large images may affect performance

### Dropdowns
- Cell reference-based validations (e.g., `=$A$1:$A$10`) are not supported
- Only inline value lists work (e.g., `"Option1,Option2,Option3"`)

### General
- Charts are not supported
- Pivot tables are not supported
- Macros/VBA are not supported
- Password-protected files are not supported

---

## File Structure

```
src/app/components/excel-viewer/
├── index.ts                      # Public API exports
├── excel-viewer.module.ts        # Angular module
├── excel-viewer.component.ts     # Main component
├── excel-viewer.component.html   # Template
├── excel-viewer.component.css    # Styles
└── excel-viewer.types.ts         # TypeScript interfaces
```

---

## Troubleshooting

### CORS Issues

If loading from a different domain, ensure CORS is enabled on the server, or use a proxy:

**proxy.conf.json:**
```json
{
  "/api/excel": {
    "target": "https://your-excel-server.com",
    "secure": true,
    "changeOrigin": true,
    "pathRewrite": { "^/api/excel": "" }
  }
}
```

### Build Errors

If you see TypeScript errors about missing types, the component uses `// @ts-ignore` for some Univer imports. This is expected.

### Performance

For large Excel files (>5MB), consider:
- Setting `enableImages: false` if images aren't needed
- Using a loading indicator
- Implementing lazy loading

---

## License

MIT
