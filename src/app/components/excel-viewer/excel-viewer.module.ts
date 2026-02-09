import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ExcelViewerComponent } from './excel-viewer.component';

@NgModule({
  declarations: [ExcelViewerComponent],
  imports: [CommonModule],
  exports: [ExcelViewerComponent],
})
export class ExcelViewerModule {}
