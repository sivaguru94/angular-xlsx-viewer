import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ExcelViewerComponent } from './excel-viewer.component';
import { DataValidationService } from './services/data-validation.service';
import { EditGuardService } from './services/edit-guard.service';

@NgModule({
  declarations: [ExcelViewerComponent],
  imports: [CommonModule],
  exports: [ExcelViewerComponent],
  providers: [DataValidationService, EditGuardService],
})
export class ExcelViewerModule {}
