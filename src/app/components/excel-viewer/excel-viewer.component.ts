import {
  Component,
  Input,
  Output,
  EventEmitter,
  OnInit,
  OnDestroy,
  OnChanges,
  SimpleChanges,
  ElementRef,
  ViewChild,
} from '@angular/core';
import { createUniver, LocaleType, merge } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
// @ts-ignore - no types available
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing';
// @ts-ignore - no types available
import UniverPresetSheetsDrawingEnUS from '@univerjs/preset-sheets-drawing/locales/en-US';
// @ts-ignore - no types available
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation';
// @ts-ignore - no types available
import UniverPresetSheetsDataValidationEnUS from '@univerjs/preset-sheets-data-validation/locales/en-US';
import '@univerjs/sheets-drawing-ui/facade';
import '@univerjs/sheets-data-validation/facade';
import LuckyExcel from '@zwight/luckyexcel';

import {
  ExcelViewerConfig,
  DEFAULT_CONFIG,
  ExcelLoadedEvent,
  CellSelectionEvent,
  CellChangeEvent,
  ExcelErrorEvent,
  ExtractedDataValidation,
} from './excel-viewer.types';
import { READONLY_MENU_OVERRIDES } from './constants';
import { columnToLetter, parseAddress, sanitizeWorkbookData } from './utils';
import { DataValidationService } from './services/data-validation.service';
import { EditGuardService } from './services/edit-guard.service';

@Component({
  selector: 'app-excel-viewer',
  templateUrl: './excel-viewer.component.html',
  styleUrls: ['./excel-viewer.component.css'],
})
export class ExcelViewerComponent implements OnInit, OnDestroy, OnChanges {
  /**
   * URL to load the Excel file from.
   * Mutually exclusive with `file` input.
   */
  @Input() url: string | null = null;

  /**
   * File object to load.
   * Mutually exclusive with `url` input.
   */
  @Input() file: File | Blob | null = null;

  /**
   * ArrayBuffer of Excel data.
   * Mutually exclusive with `url` and `file` inputs.
   */
  @Input() data: ArrayBuffer | null = null;

  /**
   * Configuration options for the viewer.
   */
  @Input() config: ExcelViewerConfig = {};

  /**
   * Unique container ID (auto-generated if not provided).
   * Useful when multiple viewers are on the same page.
   */
  @Input() containerId: string = '';

  /**
   * Emitted when Excel file is successfully loaded.
   */
  @Output() loaded = new EventEmitter<ExcelLoadedEvent>();

  /**
   * Emitted when a cell is selected.
   */
  @Output() cellSelected = new EventEmitter<CellSelectionEvent>();

  /**
   * Emitted when a cell value changes.
   */
  @Output() cellChanged = new EventEmitter<CellChangeEvent>();

  /**
   * Emitted on any error.
   */
  @Output() error = new EventEmitter<ExcelErrorEvent>();

  /**
   * Emitted when loading state changes.
   */
  @Output() loadingChange = new EventEmitter<boolean>();

  @ViewChild('univerContainer', { static: true }) containerRef!: ElementRef;

  // Internal state
  loading = false;
  errorMessage: string | null = null;
  isReadOnly = false;

  private univerAPI: any;
  private univer: any;
  private mergedConfig: ExcelViewerConfig = DEFAULT_CONFIG;
  private instanceId: string;

  constructor(
    private dataValidationService: DataValidationService,
    private editGuardService: EditGuardService,
  ) {
    this.instanceId = `univer-${Math.random().toString(36).substr(2, 9)}`;
  }

  ngOnInit(): void {
    this.mergedConfig = { ...DEFAULT_CONFIG, ...this.config };
    this.isReadOnly = !this.mergedConfig.editable;
    if (!this.containerId) {
      this.containerId = this.instanceId;
    }
    this.initUniver();
    this.loadData();
  }

  ngOnChanges(changes: SimpleChanges): void {
    // Reload if source changes
    if (
      (changes['url'] || changes['file'] || changes['data']) &&
      !changes['url']?.firstChange &&
      !changes['file']?.firstChange &&
      !changes['data']?.firstChange
    ) {
      this.loadData();
    }

    // Update config
    if (changes['config'] && !changes['config'].firstChange) {
      const prevEditable = this.mergedConfig.editable;
      this.mergedConfig = { ...DEFAULT_CONFIG, ...this.config };
      this.isReadOnly = !this.mergedConfig.editable;

      // Toggle read-only at runtime if editable changed
      if (prevEditable !== this.mergedConfig.editable) {
        if (this.mergedConfig.editable) {
          this.editGuardService.remove();
        } else {
          this.editGuardService.apply(this.univerAPI);
          this.editGuardService.disableDrawingInteraction(this.univerAPI, this.univer);
        }
      }
    }
  }

  ngOnDestroy(): void {
    this.dispose();
  }

  /**
   * Disposes the Univer instance and cleans up resources.
   */
  dispose(): void {
    this.editGuardService.remove();
    if (this.univer) {
      this.univer.dispose();
      this.univer = null;
      this.univerAPI = null;
    }
  }

  /**
   * Reloads the Excel file from the current source.
   */
  reload(): void {
    this.loadData();
  }

  /**
   * Gets the value of a cell.
   */
  getCellValue(row: number, col: number, sheetIndex?: number): any {
    const workbook = this.univerAPI?.getActiveWorkbook?.();
    if (!workbook) return null;

    const sheet =
      sheetIndex !== undefined
        ? workbook.getSheets()[sheetIndex]
        : workbook.getActiveSheet();

    if (!sheet) return null;

    const range = sheet.getRange(row, col);
    return range?.getValue?.();
  }

  /**
   * Sets the value of a cell. No-op if editable is false.
   */
  setCellValue(row: number, col: number, value: any, sheetIndex?: number): void {
    if (!this.mergedConfig.editable) return;

    const workbook = this.univerAPI?.getActiveWorkbook?.();
    if (!workbook) return;

    const sheet =
      sheetIndex !== undefined
        ? workbook.getSheets()[sheetIndex]
        : workbook.getActiveSheet();

    if (!sheet) return;

    const range = sheet.getRange(row, col);
    range?.setValue?.(value);
  }

  /**
   * Gets the currently selected range.
   */
  getSelectedRange(): any {
    const workbook = this.univerAPI?.getActiveWorkbook?.();
    if (!workbook) return null;

    const sheet = workbook.getActiveSheet();
    return sheet?.getSelection?.()?.getActiveRange?.();
  }

  /**
   * Gets all sheet names.
   */
  getSheetNames(): string[] {
    const workbook = this.univerAPI?.getActiveWorkbook?.();
    if (!workbook) return [];

    return workbook.getSheets().map((s: any) => s.getSheetName());
  }

  /**
   * Switches to a specific sheet by index.
   */
  setActiveSheet(index: number): void {
    const workbook = this.univerAPI?.getActiveWorkbook?.();
    if (!workbook) return;

    const sheets = workbook.getSheets();
    if (sheets[index]) {
      sheets[index].activate();
    }
  }

  /**
   * Selects a range of cells (blue selection rectangle, same as user click/drag).
   * @param startCell Start cell address (e.g., "A1") or { row, col } (0-based)
   * @param endCell   End cell address (e.g., "C5") or { row, col } (0-based)
   * @param sheetIndex Sheet index (0-based). Uses active sheet if omitted.
   */
  highlightRange(
    startCell: string | { row: number; col: number },
    endCell: string | { row: number; col: number },
    sheetIndex?: number
  ): void {
    const workbook = this.univerAPI?.getActiveWorkbook?.();
    if (!workbook) return;

    const sheet =
      sheetIndex !== undefined
        ? workbook.getSheets()[sheetIndex]
        : workbook.getActiveSheet();
    if (!sheet) return;

    // Activate the target sheet if it's not already active
    if (sheetIndex !== undefined) {
      sheet.activate();
    }

    const start = typeof startCell === 'string' ? parseAddress(startCell) : startCell;
    const end = typeof endCell === 'string' ? parseAddress(endCell) : endCell;

    const numRows = end.row - start.row + 1;
    const numCols = end.col - start.col + 1;

    const range = sheet.getRange(start.row, start.col, numRows, numCols);
    range?.activate?.();
  }

  /**
   * Exports the current workbook as JSON (Univer format).
   */
  exportAsJson(): any {
    const workbook = this.univerAPI?.getActiveWorkbook?.();
    return workbook?.save?.();
  }

  // ─────────────────────────────────────────────────────────────────
  // Private Methods
  // ─────────────────────────────────────────────────────────────────

  private initUniver(): void {
    const corePresetConfig: any = {
      container: this.containerRef.nativeElement,
    };

    if (!this.mergedConfig.editable) {
      corePresetConfig.menu = READONLY_MENU_OVERRIDES;
    }

    const presets: any[] = [
      UniverSheetsCorePreset(corePresetConfig),
    ];

    if (this.mergedConfig.enableImages) {
      presets.push(UniverSheetsDrawingPreset());
    }

    if (this.mergedConfig.enableDataValidation) {
      presets.push(UniverSheetsDataValidationPreset());
    }

    const localeMap: Record<string, string> = {
      'en-US': LocaleType.EN_US,
      'zh-CN': LocaleType.ZH_CN,
    };

    const { univer, univerAPI } = createUniver({
      locale: (localeMap[this.mergedConfig.locale || 'en-US'] || LocaleType.EN_US) as any,
      locales: {
        [LocaleType.EN_US]: merge(
          {},
          UniverPresetSheetsCoreEnUS,
          UniverPresetSheetsDrawingEnUS,
          UniverPresetSheetsDataValidationEnUS
        ),
      },
      presets,
    });

    this.univer = univer;
    this.univerAPI = univerAPI;
  }

  private async loadData(): Promise<void> {
    if (!this.url && !this.file && !this.data) {
      return;
    }

    this.setLoading(true);
    this.errorMessage = null;

    try {
      let arrayBuffer: ArrayBuffer;

      if (this.data) {
        arrayBuffer = this.data;
      } else if (this.file) {
        arrayBuffer = await this.file.arrayBuffer();
      } else if (this.url) {
        const response = await fetch(this.url);
        if (!response.ok) {
          throw new Error(`Failed to fetch: ${response.statusText}`);
        }
        const blob = await response.blob();
        arrayBuffer = await blob.arrayBuffer();
      } else {
        throw new Error('No data source provided');
      }

      await this.parseAndLoadExcel(arrayBuffer);
    } catch (err: any) {
      this.handleError('load', err.message || 'Failed to load Excel file', err);
    } finally {
      this.setLoading(false);
    }
  }

  private async parseAndLoadExcel(arrayBuffer: ArrayBuffer): Promise<void> {
    let dataValidations: ExtractedDataValidation[] = [];

    if (this.mergedConfig.enableDataValidation) {
      try {
        dataValidations = await this.dataValidationService.extract(arrayBuffer);
      } catch (err) {
        this.handleError('parse', 'Error extracting data validations with ExcelJS', err);
      }
    }

    // Convert to File for LuckyExcel
    const blob = new Blob([arrayBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const file = new File([blob], 'workbook.xlsx', { type: blob.type });

    // Transform and load
    return new Promise((resolve, reject) => {
      LuckyExcel.transformExcelToUniver(
        file,
        async (univerData: any) => {
          if (!univerData) {
            reject(new Error('Failed to parse Excel data'));
            return;
          }

          sanitizeWorkbookData(univerData);
          this.univerAPI.createWorkbook(univerData);

          const workbook = this.univerAPI.getActiveWorkbook();
          const sheets = workbook?.getSheets() || [];
          const sheetNames = sheets.map((s: any) => s.getSheetName());

          this.setupEventListeners();

          // Apply validations and read-only guard after delay
          const delay = this.mergedConfig.insertDelay || 500;
          setTimeout(() => {
            if (dataValidations.length > 0) {
              this.dataValidationService
                .apply(this.univerAPI, dataValidations)
                .catch((err) =>
                  this.handleError('validation', 'Failed to apply data validations', err)
                );
            }

            if (!this.mergedConfig.editable) {
              this.editGuardService.apply(this.univerAPI);
              this.editGuardService.disableDrawingInteraction(this.univerAPI, this.univer);
            }
          }, delay);

          this.loaded.emit({
            sheetCount: sheets.length,
            sheetNames,
            validationCount: dataValidations.length,
          });

          resolve();
        },
        (err: any) => {
          this.handleError('parse', 'Failed to parse Excel file', err);
          reject(err);
        }
      );
    });
  }

  private setupEventListeners(): void {
    if (!this.univerAPI) return;

    const workbook = this.univerAPI.getActiveWorkbook();
    if (!workbook) return;

    workbook.onSelectionChange((selections: any[]) => {
      if (!selections || selections.length === 0) return;

      const sheet = workbook.getActiveSheet();
      if (!sheet) return;

      const range = selections[selections.length - 1];

      const startRow = range.startRow;
      const startCol = range.startColumn;
      const endRow = range.endRow ?? startRow;
      const endCol = range.endColumn ?? startCol;

      const startAddr = columnToLetter(startCol) + (startRow + 1);
      const endAddr = columnToLetter(endCol) + (endRow + 1);
      const address = startAddr === endAddr ? startAddr : `${startAddr}:${endAddr}`;

      const values: string[] = [];
      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          const val = sheet.getRange(r, c)?.getValue?.();
          if (val !== null && val !== undefined && val !== '') {
            values.push(String(val));
          }
        }
      }

      this.cellSelected.emit({
        sheetName: sheet.getSheetName(),
        startRow,
        startCol,
        endRow,
        endCol,
        address,
        value: values.join(' '),
      });
    });
  }

  private setLoading(value: boolean): void {
    this.loading = value;
    this.loadingChange.emit(value);
  }

  private handleError(type: ExcelErrorEvent['type'], message: string, err?: any): void {
    this.errorMessage = message;
    this.error.emit({ type, message, error: err });
    console.error(`[ExcelViewer] ${type}: ${message}`, err);
  }
}
