# XLSX POC - Development Timeline

## Phase 1: Initial Setup — 9 Feb 2026
- **Angular app scaffolding** with Excel viewer component
- **Univer integration** — core spreadsheet engine with preset-based configuration
- **LuckyExcel** for `.xlsx` → Univer format conversion
- **ExcelJS** for data validation extraction
- **Dev server proxy** for loading Excel files from API

---

## Phase 2: Core Excel Viewer Component — 9 Feb 2026
- Configurable inputs: `url`, `file`, `data` (ArrayBuffer), `config`, `containerId`
- Event emitters: `loaded`, `cellSelected`, `cellChanged`, `error`, `loadingChange`
- Public API methods: `getCellValue`, `setCellValue`, `getSelectedRange`, `getSheetNames`, `setActiveSheet`, `highlightRange`, `exportAsJson`, `reload`, `dispose`
- Workbook data sanitization (fixes negative column widths / row heights from LuckyExcel)
- Multi-sheet support with sheet tab navigation

---

## Phase 3: Read-Only Mode — 10 Feb 2026
- **Edit guard** — command interceptor blocking 79+ Univer mutation commands (cell content, clipboard, row/col ops, styling, sheet ops, merge, auto-fill, data validation)
- **Context menu hiding** — suppresses edit-related right-click menu items (cut, paste, insert, delete, clear, etc.)
- **Runtime toggle** — editable state can be switched at runtime via config change
- Default mode is read-only (`editable: false`)

---

## Phase 4: Image Support — 11 Feb 2026
- **LuckyExcel handles images natively** — images render with correct dimensions via Univer drawing preset
- **Removed duplicate ExcelJS image extraction** — was causing same image to appear twice with different sizes
- **Image context menu disabled in read-only mode** — hides edit/delete/crop/reset popup via CSS (`[data-u-comp="rect-popup"]`)
- **Image move/resize blocked in read-only mode**:
  - Drawing commands blocked (`delete-drawing`, `remove-sheet-image`, `set-sheet-image`, `move-drawing`, `set-drawing-arrange`, etc.)
  - Canvas transformer disabled (`attachTo` overridden to no-op) — prevents drag handles from appearing

---

## Phase 5: Data Validation Support — 11 Feb 2026
- Extracts list-type data validations from Excel via ExcelJS
- Applies dropdown validations to Univer cells via facade API
- Handles inline comma-separated values and skips cell-reference formulas
- Applied after configurable delay (`insertDelay`) to avoid race conditions with workbook creation

---

## Phase 6: Toolbar & Formula Bar Control — 11 Feb 2026
- `showToolbar` config option — hides/shows the header bar and toolbar ribbon
- `showFormulaBar` config option — hides/shows the formula bar
- Wired to Univer's `header`, `toolbar`, and `formulaBar` preset config flags
- Works independently of read-only mode

---

## Phase 7: Code Refactoring (npm-library-ready) — 11 Feb 2026
Restructured from a single 895-line component into a clean modular architecture:

```
excel-viewer/
├── index.ts                          # Public barrel exports
├── excel-viewer.module.ts            # Module + service providers
├── excel-viewer.component.ts         # Slim orchestrator (~500 lines)
├── excel-viewer.component.html
├── excel-viewer.component.css
├── excel-viewer.types.ts             # Config, event interfaces, defaults
├── constants/
│   ├── blocked-commands.ts           # 79+ blocked Univer command IDs
│   └── readonly-menu-overrides.ts    # Context menu hiding config
├── utils/
│   ├── cell-address.ts               # columnToLetter, letterToColumn, parseAddress
│   └── workbook-sanitizer.ts         # sanitizeWorkbookData
└── services/
    ├── data-validation.service.ts    # extract() + apply() for dropdowns
    └── edit-guard.service.ts         # apply(), remove(), disableDrawingInteraction()
```

- **Constants** extracted — blocked commands and menu overrides as standalone files
- **Utils** extracted — pure functions (cell address parsing, workbook sanitization) are framework-agnostic and independently testable
- **Services** extracted — `DataValidationService` and `EditGuardService` as Angular injectables
- **Dead test code removed** — `startHighlightTest`, `highlightTestInterval`
- Barrel exports expose public API; internals (constants, services) kept private

---

## Phase 8: Selection Confirm Popup (POC) — 18 Feb 2026
- `confirmSelection` config option — when `true`, shows a Select/Cancel popup on cell selection instead of emitting immediately
- Popup positioned near the user's click via mouse position tracking
- **Select** button emits `cellSelected` with the pending payload
- **Cancel** button (or backdrop click) dismisses without emitting
- When `false` (default), `cellSelected` emits immediately as before

---

## Key Dependencies
| Package | Version | Purpose |
|---|---|---|
| `@univerjs/presets` | 0.15.4 | Spreadsheet engine |
| `@univerjs/preset-sheets-core` | — | Core sheet preset |
| `@univerjs/preset-sheets-drawing` | — | Image/drawing support |
| `@univerjs/preset-sheets-data-validation` | — | Data validation support |
| `@zwight/luckyexcel` | 1.1.6 | Excel → Univer conversion |
| `exceljs` | 4.4.0 | Excel parsing (data validations) |
| `@angular/core` | 14.2.0 | Framework |

---

## Configuration Reference
| Option | Type | Default | Description |
|---|---|---|---|
| `enableImages` | boolean | `true` | Load drawing preset for image display |
| `enableDataValidation` | boolean | `true` | Extract and apply dropdown validations |
| `showToolbar` | boolean | `true` | Show/hide toolbar and header |
| `showFormulaBar` | boolean | `true` | Show/hide formula bar |
| `showSheetTabs` | boolean | `true` | Show/hide sheet tabs |
| `editable` | boolean | `false` | Enable/disable editing |
| `locale` | string | `en-US` | UI locale |
| `zoom` | number | `100` | Initial zoom % |
| `insertDelay` | number | `500` | Delay (ms) before applying validations |
| `confirmSelection` | boolean | `false` | Show Select/Cancel popup on cell click |
