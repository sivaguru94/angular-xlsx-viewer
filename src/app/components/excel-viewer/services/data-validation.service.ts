import { Injectable } from '@angular/core';
import * as ExcelJS from 'exceljs';
import { ExtractedDataValidation } from '../excel-viewer.types';

@Injectable()
export class DataValidationService {
  /**
   * Extracts list-type data validations from an Excel file using ExcelJS.
   */
  async extract(arrayBuffer: ArrayBuffer): Promise<ExtractedDataValidation[]> {
    const dataValidations: ExtractedDataValidation[] = [];

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);

    workbook.eachSheet((worksheet, sheetId) => {
      const validationsMap = (worksheet as any).dataValidations?.model;
      if (!validationsMap) return;

      Object.entries(validationsMap).forEach(
        ([address, validation]: [string, any]) => {
          if (validation?.type !== 'list') return;

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
      );
    });

    return dataValidations;
  }

  /**
   * Applies extracted data validations to a Univer workbook via the facade API.
   */
  async apply(univerAPI: any, validations: ExtractedDataValidation[]): Promise<void> {
    const workbook = univerAPI?.getActiveWorkbook?.();
    if (!workbook) return;

    for (const validation of validations) {
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

      const rule = univerAPI
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
    }
  }
}
