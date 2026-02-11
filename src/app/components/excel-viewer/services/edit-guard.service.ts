import { Injectable } from '@angular/core';
import { IRenderManagerService } from '@univerjs/engine-render';
import { BLOCKED_COMMANDS } from '../constants';

@Injectable()
export class EditGuardService {
  private editGuardDisposable: any = null;

  /**
   * Registers a command interceptor that blocks all mutating commands.
   */
  apply(univerAPI: any): void {
    if (this.editGuardDisposable || !univerAPI) return;

    this.editGuardDisposable = univerAPI.onBeforeCommandExecute((command: any) => {
      if (BLOCKED_COMMANDS.has(command.id)) {
        throw new Error('Read-only mode: editing is disabled');
      }
    });
  }

  /**
   * Removes the command interceptor, re-enabling editing.
   */
  remove(): void {
    if (this.editGuardDisposable) {
      this.editGuardDisposable.dispose();
      this.editGuardDisposable = null;
    }
  }

  /**
   * Disables the drawing transformer so images cannot be dragged or resized.
   * Non-critical — if it fails, command blocking still prevents persistence.
   */
  disableDrawingInteraction(univerAPI: any, univer: any): void {
    try {
      const workbook = univerAPI?.getActiveWorkbook?.();
      if (!workbook) return;

      const unitId = workbook.getId();
      const injector = univer?.__getInjector?.();
      if (!injector) return;

      const renderManagerService = injector.get(IRenderManagerService);
      if (!renderManagerService) return;

      const render = renderManagerService.getRenderById(unitId);
      const transformer = render?.scene?.getTransformer?.();
      if (transformer) {
        transformer.attachTo = () => transformer;
      }
    } catch {
      // Non-critical: images will still be non-editable via command blocking
    }
  }
}
