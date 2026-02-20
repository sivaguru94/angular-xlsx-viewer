import { Component } from '@angular/core';
import {
  ExcelViewerConfig,
  ExcelLoadedEvent,
  CellSelectionEvent,
  ExcelErrorEvent,
} from './components/excel-viewer';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  // Excel file URL (proxied via Angular dev server)
  // PaymentAdvice.xlsx
  excelUrl = '/api/excel/Titan+Engineering+Planning+-+Q4+2025.xlsx';

  // Configuration for the Excel viewer
  viewerConfig: ExcelViewerConfig = {
    enableImages: true,
    enableDataValidation: true,
    showToolbar: false,
    showFormulaBar: false,
    showSheetTabs: true,
    editable: false,
    locale: 'en-US',
  };

  // Cells to highlight — driven by the parent, passed to the viewer via input
  highlightedCells: Array<{ row: number; col: number }> = [];

  // Event handlers
  onExcelLoaded(event: ExcelLoadedEvent): void {
    console.log('Excel loaded:', event);
    console.log(`Loaded ${event.sheetCount} sheets: ${event.sheetNames.join(', ')}`);
    console.log(`Found ${event.validationCount} validations`);

    // Verify: highlight (0,0) and (1,1), then clear after 3 s
    this.highlightedCells = [{ row: 0, col: 0 }, { row: 1, col: 1 }];
    setTimeout(() => (this.highlightedCells = []), 3000);
  }

  onCellSelected(event: CellSelectionEvent): void {
    console.log('Cell selected:', event);
  }

  onError(event: ExcelErrorEvent): void {
    console.error('Excel error:', event);
  }

  onLoadingChange(loading: boolean): void {
    console.log('Loading:', loading);
  }
}
