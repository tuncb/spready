export const APP_MENU_ACTIONS = {
  cut: "edit:cut",
  cutValues: "edit:cut-values",
  copy: "edit:copy",
  copyValues: "edit:copy-values",
  addColumn: "sheet:add-column",
  addRow: "sheet:add-row",
  deleteSelection: "edit:delete-selection",
  deleteSheet: "sheet:delete-sheet",
  exportCsv: "csv:export",
  importCsv: "csv:import",
  newSheet: "sheet:new-sheet",
  openWorkbook: "workbook:open",
  paste: "edit:paste",
  pasteValues: "edit:paste-values",
  saveWorkbook: "workbook:save",
  saveWorkbookAs: "workbook:save-as",
} as const;

export type AppMenuAction =
  (typeof APP_MENU_ACTIONS)[keyof typeof APP_MENU_ACTIONS];
