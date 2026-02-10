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
  excelUrl = '/api/excel/Titan+Engineering+Planning+-+Q4+2025.xlsx';

  // Configuration for the Excel viewer
  viewerConfig: ExcelViewerConfig = {
    enableImages: true,
    enableDataValidation: true,
    showToolbar: true,
    showFormulaBar: true,
    showSheetTabs: true,
    editable: false,
    locale: 'en-US',
  };

  // Event handlers
  onExcelLoaded(event: ExcelLoadedEvent): void {
    console.log('Excel loaded:', event);
    console.log(`Loaded ${event.sheetCount} sheets: ${event.sheetNames.join(', ')}`);
    console.log(`Found ${event.imageCount} images and ${event.validationCount} validations`);
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
