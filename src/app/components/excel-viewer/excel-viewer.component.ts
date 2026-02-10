import {
  Component,
  Input,
  Output,
  EventEmitter,
  OnInit,
  OnDestroy,
  OnChanges,
  AfterViewInit,
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
export class ExcelViewerComponent implements OnInit, OnDestroy, OnChanges, AfterViewInit {
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
   * Emitted on any error.cellSelected
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
  private highlightTestInterval: any = null;
  private editGuardDisposable: any = null;

  constructor() {
    this.instanceId = `univer-${Math.random().toString(36).substr(2, 9)}`;
  }

  ngOnInit(): void {
    this.mergedConfig = { ...DEFAULT_CONFIG, ...this.config };
    console.log('mergedConfig', this.mergedConfig);
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

      // Toggle read-only at runtime if editable changed
      if (prevEditable !== this.mergedConfig.editable) {
        if (this.mergedConfig.editable) {
          this.removeEditGuard();
        } else {
          this.applyEditGuard();
        }
      }
    }
  }

  ngAfterViewInit(): void {
    // [TEST] Start random highlight cycle after a delay to let the sheet load
    // setTimeout(() => this.startHighlightTest(), 3000);
  }

  ngOnDestroy(): void {
    if (this.highlightTestInterval) {
      // clearInterval(this.highlightTestInterval);
    }
    this.dispose();
  }

  /**
   * Disposes the Univer instance and cleans up resources.
   */
  dispose(): void {
    this.removeEditGuard();
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

    const start = typeof startCell === 'string' ? this.parseAddress(startCell) : startCell;
    const end = typeof endCell === 'string' ? this.parseAddress(endCell) : endCell;

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

  /**
   * [TEST] Randomly selects a cell range every 10 seconds.
   * Remove this method and the ngAfterViewInit call when done testing.
   */
  private startHighlightTest(): void {
    const randomHighlight = () => {
      const workbook = this.univerAPI?.getActiveWorkbook?.();
      if (!workbook) return;

      const sheets = workbook.getSheets();
      if (!sheets || sheets.length === 0) return;

      const sheetIndex = Math.floor(Math.random() * sheets.length);

      // Random start cell within rows 0-14, cols 0-7
      const startRow = Math.floor(Math.random() * 15);
      const startCol = Math.floor(Math.random() * 8);
      // Random range size: 1-5 rows, 1-4 cols
      const endRow = startRow + Math.floor(Math.random() * 5);
      const endCol = startCol + Math.floor(Math.random() * 4);

      const startAddr = this.columnToLetter(startCol) + (startRow + 1);
      const endAddr = this.columnToLetter(endCol) + (endRow + 1);

      console.log(`[TEST] Selecting ${startAddr}:${endAddr} on sheet ${sheetIndex}`);

      this.highlightRange({ row: startRow, col: startCol }, { row: endRow, col: endCol }, sheetIndex);
    };

    // First selection immediately
    randomHighlight();
    // Then every 10 seconds
    this.highlightTestInterval = setInterval(randomHighlight, 10000);
  }

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

          // Insert images and validations after delay, then apply edit guard
          const delay = this.mergedConfig.insertDelay || 500;
          setTimeout(() => {
            if (images.length > 0) {
              this.insertImages(images);
            }
            if (dataValidations.length > 0) {
              this.applyDataValidations(dataValidations);
            }

            // Apply read-only AFTER insertions are done
            if (!this.mergedConfig.editable) {
              this.applyEditGuard();
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

    const workbook = this.univerAPI.getActiveWorkbook();
    if (!workbook) return;

    // Selection change listener — callback receives IRange[]
    workbook.onSelectionChange((selections: any[]) => {
      if (!selections || selections.length === 0) return;

      const sheet = workbook.getActiveSheet();
      if (!sheet) return;

      // Use the last (most recent) selection range
      const range = selections[selections.length - 1];

      const startRow = range.startRow;
      const startCol = range.startColumn;
      const endRow = range.endRow ?? startRow;
      const endCol = range.endColumn ?? startCol;

      const startAddr = this.columnToLetter(startCol) + (startRow + 1);
      const endAddr = this.columnToLetter(endCol) + (endRow + 1);
      const address = startAddr === endAddr ? startAddr : `${startAddr}:${endAddr}`;

      // Concatenate all cell values: left-to-right, top-to-bottom
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

  // Commands that mutate workbook content, structure, or formatting
  private static readonly BLOCKED_COMMANDS: Set<string> = new Set([
    // Cell content
    'sheet.command.set-range-values',
    'sheet.command.clear-selection-content',
    'sheet.command.clear-selection-format',
    'sheet.command.clear-selection-all',
    // Clipboard
    'sheet.command.cut',
    'sheet.command.paste',
    'sheet.command.paste-value',
    'sheet.command.paste-format',
    'sheet.command.paste-col-width',
    'sheet.command.paste-besides-border',
    'sheet.command.optional-paste',
    // Row operations
    'sheet.command.insert-row',
    'sheet.command.insert-row-before',
    'sheet.command.insert-row-after',
    'sheet.command.insert-row-by-range',
    'sheet.command.insert-multi-rows-above',
    'sheet.command.insert-multi-rows-after',
    'sheet.command.remove-row',
    'sheet.command.remove-row-by-range',
    'sheet.command.append-row',
    'sheet.command.move-rows',
    'sheet.command.set-row-height',
    'sheet.command.set-row-data',
    'sheet.command.delta-row-height',
    // Column operations
    'sheet.command.insert-col',
    'sheet.command.insert-col-before',
    'sheet.command.insert-col-after',
    'sheet.command.insert-col-by-range',
    'sheet.command.insert-multi-cols-before',
    'sheet.command.insert-multi-cols-right',
    'sheet.command.remove-col',
    'sheet.command.remove-col-by-range',
    'sheet.command.move-cols',
    'sheet.command.set-worksheet-col-width',
    'sheet.command.delta-column-width',
    // Range operations
    'sheet.command.delete-range-move-left',
    'sheet.command.delete-range-move-up',
    'sheet.command.delete-range-move-left-confirm',
    'sheet.command.delete-range-move-up-confirm',
    'sheet.command.insert-range-move-down',
    'sheet.command.insert-range-move-right',
    'sheet.command.insert-range-move-down-confirm',
    'sheet.command.insert-range-move-right-confirm',
    'sheet.command.move-range',
    'sheet.command.reorder-range',
    // Styling
    'sheet.command.set-style',
    'sheet.command.set-bold',
    'sheet.command.set-italic',
    'sheet.command.set-underline',
    'sheet.command.set-stroke',
    'sheet.command.set-font-family',
    'sheet.command.set-font-size',
    'sheet.command.set-text-color',
    'sheet.command.set-background-color',
    'sheet.command.set-vertical-text-align',
    'sheet.command.set-horizontal-text-align',
    'sheet.command.set-text-wrap',
    'sheet.command.set-text-rotation',
    'sheet.command.set-border',
    'sheet.command.set-border-position',
    'sheet.command.set-border-style',
    'sheet.command.set-border-color',
    'sheet.command.set-border-basic',
    // Sheet operations
    'sheet.command.insert-sheet',
    'sheet.command.remove-sheet',
    'sheet.command.remove-sheet-confirm',
    'sheet.command.set-worksheet-name',
    'sheet.command.set-worksheet-order',
    'sheet.command.set-worksheet-hidden',
    'sheet.command.copy-sheet',
    'sheet.command.set-tab-color',
    'sheet.command.set-workbook-name',
    // Merge
    'sheet.command.add-worksheet-merge',
    'sheet.command.add-worksheet-merge-all',
    'sheet.command.add-worksheet-merge-horizontal',
    'sheet.command.add-worksheet-merge-vertical',
    'sheet.command.remove-worksheet-merge',
    // Auto fill / format painter
    'sheet.command.auto-fill',
    'sheet.command.auto-clear-content',
    'sheet.command.refill',
    'sheet.command.apply-format-painter',
    // Data validation (dropdowns, checkboxes, rules)
    'sheet.command.addDataValidation',
    'sheet.command.remove-data-validation-rule',
    'sheet.command.remove-all-data-validation',
    'sheet.command.updateDataValidationRuleRange',
    'sheets.command.update-data-validation-setting',
    'sheets.command.update-data-validation-options',
    'sheets.command.clear-range-data-validation',
    'data-validation.command.addRuleAndOpen',
    'sheet.operation.show-data-validation-dropdown',
  ]);

  private applyEditGuard(): void {
    if (this.editGuardDisposable || !this.univerAPI) return;

    this.editGuardDisposable = this.univerAPI.onBeforeCommandExecute((command: any) => {
      if (ExcelViewerComponent.BLOCKED_COMMANDS.has(command.id)) {
        throw new Error('Read-only mode: editing is disabled');
      }
    });
  }

  private removeEditGuard(): void {
    if (this.editGuardDisposable) {
      this.editGuardDisposable.dispose();
      this.editGuardDisposable = null;
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
                buffer: img.buffer as unknown as Buffer,
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

  private letterToColumn(letters: string): number {
    let col = 0;
    for (let i = 0; i < letters.length; i++) {
      col = col * 26 + (letters.charCodeAt(i) - 64);
    }
    return col - 1; // 0-based
  }

  private parseAddress(address: string): { row: number; col: number } {
    const match = address.toUpperCase().match(/^([A-Z]+)(\d+)$/);
    if (!match) {
      throw new Error(`Invalid cell address: ${address}`);
    }
    return {
      col: this.letterToColumn(match[1]),
      row: parseInt(match[2], 10) - 1, // 0-based
    };
  }
}
