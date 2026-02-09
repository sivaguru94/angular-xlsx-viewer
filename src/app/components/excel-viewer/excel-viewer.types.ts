/**
 * Configuration options for the Excel Viewer component
 */
export interface ExcelViewerConfig {
  /** Enable/disable image extraction and display (default: true) */
  enableImages?: boolean;

  /** Enable/disable dropdown/data validation support (default: true) */
  enableDataValidation?: boolean;

  /** Enable/disable toolbar (default: true) */
  showToolbar?: boolean;

  /** Enable/disable formula bar (default: true) */
  showFormulaBar?: boolean;

  /** Enable/disable sheet tabs (default: true) */
  showSheetTabs?: boolean;

  /** Enable/disable editing (default: true) */
  editable?: boolean;

  /** Locale for the spreadsheet (default: 'en-US') */
  locale?: 'en-US' | 'zh-CN';

  /** Initial zoom level in percentage (default: 100) */
  zoom?: number;

  /** Delay in ms before inserting images/validations (default: 500) */
  insertDelay?: number;
}

/**
 * Default configuration values
 */
export const DEFAULT_CONFIG: ExcelViewerConfig = {
  enableImages: true,
  enableDataValidation: true,
  showToolbar: true,
  showFormulaBar: true,
  showSheetTabs: true,
  editable: true,
  locale: 'en-US',
  zoom: 100,
  insertDelay: 500,
};

/**
 * Event emitted when the Excel file is loaded
 */
export interface ExcelLoadedEvent {
  /** Number of sheets in the workbook */
  sheetCount: number;

  /** Names of all sheets */
  sheetNames: string[];

  /** Number of images extracted */
  imageCount: number;

  /** Number of data validations extracted */
  validationCount: number;
}

/**
 * Event emitted when a cell is selected
 */
export interface CellSelectionEvent {
  /** Sheet name */
  sheetName: string;

  /** Row index (0-based) */
  row: number;

  /** Column index (0-based) */
  col: number;

  /** Cell address (e.g., "A1") */
  address: string;

  /** Cell value */
  value: any;
}

/**
 * Event emitted when a cell value changes
 */
export interface CellChangeEvent {
  /** Sheet name */
  sheetName: string;

  /** Row index (0-based) */
  row: number;

  /** Column index (0-based) */
  col: number;

  /** Cell address (e.g., "A1") */
  address: string;

  /** Previous value */
  oldValue: any;

  /** New value */
  newValue: any;
}

/**
 * Event emitted on errors
 */
export interface ExcelErrorEvent {
  /** Error type */
  type: 'load' | 'parse' | 'image' | 'validation' | 'unknown';

  /** Error message */
  message: string;

  /** Original error object */
  error?: any;
}

/**
 * Internal interface for extracted images
 */
export interface ExtractedImage {
  buffer: Buffer | ArrayBuffer;
  extension: string;
  sheetIndex: number;
  sheetName: string;
  col: number;
  row: number;
  colOffset: number;
  rowOffset: number;
  width: number;
  height: number;
}

/**
 * Internal interface for extracted data validations
 */
export interface ExtractedDataValidation {
  sheetIndex: number;
  sheetName: string;
  address: string;
  type: string;
  allowBlank: boolean;
  formulae: string[];
  showDropDown?: boolean;
  errorTitle?: string;
  error?: string;
  promptTitle?: string;
  prompt?: string;
}
