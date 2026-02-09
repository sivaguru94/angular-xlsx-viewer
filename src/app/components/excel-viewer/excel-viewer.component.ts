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
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation';
import '@univerjs/sheets-drawing-ui/facade';
import '@univerjs/sheets-data-validation/facade';
import LuckyExcel from '@zwight/luckyexcel';
import * as ExcelJS from 'exceljs';

import {
  ExcelViewerConfig,
  DEFAULT_CONFIG,
  ExcelLoadedEvent,
  CellSelectionEvent,
  CellChangeEvent,
  ExcelErrorEvent,
  ExtractedImage,
  ExtractedDataValidation,
} from './excel-viewer.types';

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

  private univerAPI: any;
  private univer: any;
  private mergedConfig: ExcelViewerConfig = DEFAULT_CONFIG;
  private instanceId: string;

  constructor() {
    this.instanceId = `univer-${Math.random().toString(36).substr(2, 9)}`;
  }

  ngOnInit(): void {
    this.mergedConfig = { ...DEFAULT_CONFIG, ...this.config };
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
      this.mergedConfig = { ...DEFAULT_CONFIG, ...this.config };
    }
  }

  ngOnDestroy(): void {
    this.dispose();
  }

  /**
   * Disposes the Univer instance and cleans up resources.
   */
  dispose(): void {
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
   * Sets the value of a cell.
   */
  setCellValue(row: number, col: number, value: any, sheetIndex?: number): void {
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
    const presets: any[] = [
      UniverSheetsCorePreset({
        container: this.containerRef.nativeElement,
      }),
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
        [LocaleType.EN_US]: merge({}, UniverPresetSheetsCoreEnUS),
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
    // Extract images and validations using ExcelJS
    let images: ExtractedImage[] = [];
    let dataValidations: ExtractedDataValidation[] = [];

    if (this.mergedConfig.enableImages || this.mergedConfig.enableDataValidation) {
      const extracted = await this.extractFromExcel(arrayBuffer);
      images = this.mergedConfig.enableImages ? extracted.images : [];
      dataValidations = this.mergedConfig.enableDataValidation
        ? extracted.dataValidations
        : [];
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

          // Sanitize column widths to prevent "column width is less than 0" warnings
          this.sanitizeWorkbookData(univerData);

          // Create workbook
          this.univerAPI.createWorkbook(univerData);

          // Get sheet info for event
          const workbook = this.univerAPI.getActiveWorkbook();
          const sheets = workbook?.getSheets() || [];
          const sheetNames = sheets.map((s: any) => s.getSheetName());

          // Setup event listeners
          this.setupEventListeners();

          // Insert images and validations after delay
          const delay = this.mergedConfig.insertDelay || 500;
          setTimeout(() => {
            if (images.length > 0) {
              this.insertImages(images);
            }
            if (dataValidations.length > 0) {
              this.applyDataValidations(dataValidations);
            }
          }, delay);

          // Emit loaded event
          this.loaded.emit({
            sheetCount: sheets.length,
            sheetNames,
            imageCount: images.length,
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

    // Selection change listener
    const hooks = this.univerAPI.getSheetHooks?.();
    if (hooks?.onSelectionChange) {
      hooks.onSelectionChange((selection: any) => {
        if (!selection) return;

        const workbook = this.univerAPI.getActiveWorkbook();
        const sheet = workbook?.getActiveSheet();
        if (!sheet) return;

        const range = selection.range;
        if (!range) return;

        const row = range.startRow;
        const col = range.startColumn;
        const address = this.columnToLetter(col) + (row + 1);
        const value = this.getCellValue(row, col);

        this.cellSelected.emit({
          sheetName: sheet.getSheetName(),
          row,
          col,
          address,
          value,
        });
      });
    }
  }

  private async extractFromExcel(arrayBuffer: ArrayBuffer): Promise<{
    images: ExtractedImage[];
    dataValidations: ExtractedDataValidation[];
  }> {
    const images: ExtractedImage[] = [];
    const dataValidations: ExtractedDataValidation[] = [];

    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(arrayBuffer);

      workbook.eachSheet((worksheet, sheetId) => {
        // Extract images
        if (this.mergedConfig.enableImages) {
          const sheetImages = worksheet.getImages();
          sheetImages.forEach((image) => {
            const img = workbook.getImage(Number(image.imageId));
            if (img?.buffer) {
              const range = image.range;
              const col = range.tl?.col ?? 0;
              const row = range.tl?.row ?? 0;

              let width = 200;
              let height = 150;
              if (range.br) {
                width = Math.max(
                  50,
                  ((range.br.col ?? col) - col) * 64 +
                    (range.br.nativeColOff ?? 0) / 9525
                );
                height = Math.max(
                  50,
                  ((range.br.row ?? row) - row) * 20 +
                    (range.br.nativeRowOff ?? 0) / 9525
                );
              }

              images.push({
                buffer: img.buffer as Buffer,
                extension: img.extension || 'png',
                sheetIndex: sheetId - 1,
                sheetName: worksheet.name,
                col: Math.floor(col),
                row: Math.floor(row),
                colOffset: (range.tl?.nativeColOff ?? 0) / 9525,
                rowOffset: (range.tl?.nativeRowOff ?? 0) / 9525,
                width,
                height,
              });
            }
          });
        }

        // Extract data validations
        if (this.mergedConfig.enableDataValidation) {
          const validationsMap = (worksheet as any).dataValidations?.model;
          if (validationsMap) {
            Object.entries(validationsMap).forEach(
              ([address, validation]: [string, any]) => {
                if (validation?.type === 'list') {
                  const formulae = validation.formulae || [];
                  let values: string[] = [];

                  if (formulae.length > 0) {
                    const formula = formulae[0];
                    if (typeof formula === 'string') {
                      if (formula.startsWith('"') || !formula.includes('$')) {
                        values = formula
                          .replace(/^"|"$/g, '')
                          .split(',')
                          .map((v: string) => v.trim());
                      } else {
                        values = [formula];
                      }
                    }
                  }

                  dataValidations.push({
                    sheetIndex: sheetId - 1,
                    sheetName: worksheet.name,
                    address,
                    type: validation.type,
                    allowBlank: validation.allowBlank !== false,
                    formulae: values,
                    showDropDown: validation.showDropDown !== false,
                    errorTitle: validation.errorTitle,
                    error: validation.error,
                    promptTitle: validation.promptTitle,
                    prompt: validation.prompt,
                  });
                }
              }
            );
          }
        }
      });
    } catch (err) {
      this.handleError('parse', 'Error extracting data with ExcelJS', err);
    }

    return { images, dataValidations };
  }

  private async insertImages(images: ExtractedImage[]): Promise<void> {
    const workbook = this.univerAPI?.getActiveWorkbook?.();
    if (!workbook) return;

    for (const img of images) {
      try {
        const sheets = workbook.getSheets();
        const targetSheet = sheets[img.sheetIndex] || workbook.getActiveSheet();
        if (!targetSheet) continue;

        const base64 = this.bufferToBase64(img.buffer);
        const mimeType = this.getMimeType(img.extension);
        const dataUrl = `data:${mimeType};base64,${base64}`;

        const imageBuilder = targetSheet.newOverGridImage?.();
        if (imageBuilder) {
          const sheetImage = await imageBuilder
            .setSource(dataUrl, 'base64')
            .setColumn(img.col)
            .setRow(img.row)
            .setWidth(img.width)
            .setHeight(img.height)
            .buildAsync();

          targetSheet.insertImages([sheetImage]);
        } else {
          await targetSheet.insertImage?.(dataUrl, img.col, img.row);
        }
      } catch (err) {
        this.handleError('image', `Failed to insert image at ${img.sheetName}`, err);
      }
    }
  }

  private async applyDataValidations(
    validations: ExtractedDataValidation[]
  ): Promise<void> {
    const workbook = this.univerAPI?.getActiveWorkbook?.();
    if (!workbook) return;

    for (const validation of validations) {
      try {
        const sheets = workbook.getSheets();
        const targetSheet =
          sheets[validation.sheetIndex] || workbook.getActiveSheet();
        if (!targetSheet) continue;

        const range = targetSheet.getRange(validation.address);
        if (!range) continue;

        const values = validation.formulae;
        if (values.length === 0) continue;

        // Skip cell reference based validations
        if (values.length === 1 && values[0].includes('$')) continue;

        const rule = this.univerAPI
          .newDataValidation()
          .requireValueInList(values)
          .setOptions({
            allowBlank: validation.allowBlank,
            showErrorMessage: !!validation.error,
            error: validation.error || 'Please select a value from the list',
            errorTitle: validation.errorTitle || 'Invalid Input',
          })
          .build();

        range.setDataValidation(rule);
      } catch (err) {
        this.handleError(
          'validation',
          `Failed to apply validation at ${validation.address}`,
          err
        );
      }
    }
  }

  private sanitizeWorkbookData(data: any): void {
    if (!data?.sheets) return;

    const sheets = data.sheets;
    const sheetEntries = typeof sheets === 'object' ? Object.values(sheets) : sheets;

    for (const sheet of sheetEntries) {
      if (!sheet) continue;

      // Fix negative or zero column widths
      if (sheet.columnData) {
        for (const key of Object.keys(sheet.columnData)) {
          const col = sheet.columnData[key];
          if (col && typeof col.w === 'number' && col.w <= 0) {
            col.w = 72; // Default column width in pixels
          }
        }
      }

      // Fix negative or zero default column width
      if (sheet.defaultColumnWidth !== undefined && sheet.defaultColumnWidth <= 0) {
        sheet.defaultColumnWidth = 72;
      }

      // Fix negative or zero row heights
      if (sheet.rowData) {
        for (const key of Object.keys(sheet.rowData)) {
          const row = sheet.rowData[key];
          if (row && typeof row.h === 'number' && row.h <= 0) {
            row.h = 20; // Default row height in pixels
          }
        }
      }
    }
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

  private bufferToBase64(buffer: Buffer | ArrayBuffer): string {
    const bytes = new Uint8Array(buffer);
    let binary = '';
    for (let i = 0; i < bytes.byteLength; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
  }

  private getMimeType(extension: string): string {
    const mimeTypes: Record<string, string> = {
      png: 'image/png',
      jpg: 'image/jpeg',
      jpeg: 'image/jpeg',
      gif: 'image/gif',
      bmp: 'image/bmp',
      webp: 'image/webp',
      svg: 'image/svg+xml',
    };
    return mimeTypes[extension.toLowerCase()] || 'image/png';
  }

  private columnToLetter(col: number): string {
    let letter = '';
    let temp = col;
    while (temp >= 0) {
      letter = String.fromCharCode((temp % 26) + 65) + letter;
      temp = Math.floor(temp / 26) - 1;
    }
    return letter;
  }
}
