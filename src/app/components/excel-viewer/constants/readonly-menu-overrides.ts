/**
 * Univer context menu items to hide in read-only mode.
 * Passed to the UniverSheetsCorePreset `menu` config.
 */
export const READONLY_MENU_OVERRIDES: Record<string, { hidden: boolean }> = {
  // Clipboard
  'sheet.command.cut': { hidden: true },
  'sheet.command.paste': { hidden: true },
  'sheet.menu.paste-special': { hidden: true },
  // Insert
  'sheet.menu.cell-insert': { hidden: true },
  'sheet.menu.row-insert': { hidden: true },
  'sheet.menu.col-insert': { hidden: true },
  // Delete
  'sheet.menu.delete': { hidden: true },
  // Clear
  'sheet.menu.clear-selection': { hidden: true },
  'sheet.command.clear-selection-content': { hidden: true },
  'sheet.command.clear-selection-format': { hidden: true },
  'sheet.command.clear-selection-all': { hidden: true },
  // Other edit actions
  'sheet.command.set-range-values': { hidden: true },
  'sheet.menu.data-validation': { hidden: true },
  'sheet.contextMenu.permission': { hidden: true },
};
