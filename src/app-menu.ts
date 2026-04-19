export const APP_MENU_ACTIONS = {
  addColumn: "sheet:add-column",
  addRow: "sheet:add-row",
  deleteSheet: "sheet:delete-sheet",
  exportCsv: "csv:export",
  importCsv: "csv:import",
  newSheet: "sheet:new-sheet",
  openWorkbook: "workbook:open",
  saveWorkbook: "workbook:save",
  saveWorkbookAs: "workbook:save-as",
} as const;

export type AppMenuAction =
  (typeof APP_MENU_ACTIONS)[keyof typeof APP_MENU_ACTIONS];
