export const APP_MENU_ACTIONS = {
  copy: "edit:copy",
  copyValues: "edit:copy-values",
  addColumn: "sheet:add-column",
  addRow: "sheet:add-row",
  deleteSelection: "edit:delete-selection",
  deleteSheet: "sheet:delete-sheet",
  export: "export",
  import: "import",
  newSheet: "sheet:new-sheet",
  paste: "edit:paste",
  pasteValues: "edit:paste-values",
} as const;

export type AppMenuAction =
  (typeof APP_MENU_ACTIONS)[keyof typeof APP_MENU_ACTIONS];
