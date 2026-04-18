export const APP_MENU_ACTIONS = {
  addColumn: "sheet:add-column",
  addRow: "sheet:add-row",
  deleteSheet: "sheet:delete-sheet",
  export: "export",
  import: "import",
  newSheet: "sheet:new-sheet",
} as const;

export type AppMenuAction =
  (typeof APP_MENU_ACTIONS)[keyof typeof APP_MENU_ACTIONS];
